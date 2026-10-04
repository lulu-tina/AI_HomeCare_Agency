"""長照居家照顧派單系統 - 核心運算引擎

將原始三階段流程（適配度評分 -> OR-Tools 最佳化 -> DiD 效益回溯）拆成可
重複呼叫的函式，並把所有具派單政策意義的係數集中到 PipelineConfig，供
CLI (ai_caregiver_pipeline.py) 與網頁儀錶板 (app.py) 共用同一份邏輯。

另包含 BA 服務代碼併報法規防呆檢核、長照申報點數與居服員拆帳薪資試算、
星期／可服務時段／請假防呆的硬性條件，以及依日期逐日批次派單的輔助函式。
新增的欄位需求皆以 `.get()`／`pd.notna()` 方式取值，僅在資料表實際具備對應
欄位時才生效，因此仍可原封不動套用在既有的單日排班資料表格式上。
"""

import os
import re
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from contextvars import ContextVar
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Optional

import numpy as np
import pandas as pd
import requests
from care_policy import difficulty_reference, certificate_names, certificate_bonus, DEFAULT_CODE_CERTS, REGISTRATION_FIELDS, dementia_status, special_qualification_error
import streamlit as st
from ortools.linear_solver import pywraplp

DEFAULT_EXCEL_PATH = "./00_DB/AI_Caregiver_Allocation_Ultimate_Database.xlsx"
  
# OSRM 僅提供本版自架機車路網；公車／捷運使用其他交通來源。
OSRM_TIMEOUT_SECONDS = 3.0
OSRM_TABLE_TIMEOUT_SECONDS = 10.0
OSRM_TABLE_MAX_COORDS = 90  # 單次 /table 批次查詢座標數上限，避免超出公用伺服器限制

# Google Routes API
GOOGLE_ROUTES_URL = (
    "https://routes.googleapis.com/distanceMatrix/v2:computeRouteMatrix"
)


GOOGLE_ROUTES_TIMEOUT_SECONDS = 15.0

# 個資保護：
# - 排班核心只使用內部 ID，不以姓名作為 key。
# - 姓名只允許存在於「案家姓名對照／居服員姓名對照」顯示層。
# - 外部 Google Routes payload 採白名單驗證，只允許座標、時間、交通模式等必要欄位。
DIRECT_IDENTIFIER_COLUMNS = {
    "姓名", "個案姓名", "員工姓名", "居服員姓名",
    "身分證", "身分證字號", "身分證號", "手機", "手機號碼",
    "電話", "聯絡電話", "地址", "住址", "Email", "email",
    "電子郵件", "生日", "出生日期",
}

IDENTITY_SHEET_ALIASES = {
    "client": ["案家姓名對照", "個案姓名對照", "Client_Mapping", "Client_Name_Mapping"],
    "caregiver": ["居服員姓名對照", "員工姓名對照", "Caregiver_Mapping", "Caregiver_Name_Mapping"],
}

def _strip_direct_identifiers(frame, keep=()):
    """從排班運算資料移除直接識別欄位；保留內部 ID 與運算必要欄位。"""
    if not isinstance(frame, pd.DataFrame):
        return frame, []
    keep = set(keep)
    drop = [c for c in frame.columns if str(c).strip() in DIRECT_IDENTIFIER_COLUMNS and c not in keep]
    return frame.drop(columns=drop, errors="ignore"), [str(c) for c in drop]

def _validate_google_routes_payload(body):
    """外部 Google Routes 僅允許最小必要欄位，避免姓名／電話／地址等欄位誤送。"""
    allowed_top = {
        "origins", "destinations", "travelMode", "languageCode",
        "regionCode", "departureTime", "arrivalTime",
    }
    unexpected = set(body) - allowed_top
    if unexpected:
        raise ValueError("Google Routes 請求含未允許欄位：" + "、".join(sorted(unexpected)))

    forbidden_tokens = ("name", "姓名", "phone", "電話", "address", "地址", "email", "身分證")
    def walk(value):
        if isinstance(value, dict):
            for key, item in value.items():
                key_text = str(key).lower()
                if any(token.lower() in key_text for token in forbidden_tokens):
                    raise ValueError("Google Routes 請求偵測到直接識別欄位，已阻擋送出")
                walk(item)
        elif isinstance(value, (list, tuple)):
            for item in value:
                walk(item)
    walk(body)

# Google Routes 僅提供公車／捷運；機車固定走自架 OSRM。
GOOGLE_TRAVEL_MODE_MAP = {
    "大眾運輸": "TRANSIT",
}
# Google Transit 批次矩陣：TRANSIT 每個 request 最多 100 elements。
GOOGLE_TRANSIT_MAX_ELEMENTS_PER_REQUEST = 100
GOOGLE_TRANSIT_MAX_WORKERS = 6
PHASE2_CANDIDATE_CAP_PER_TASK = 5
# 每次排班的 Google elements 安全上限；達上限明確停止，不替換成直線估算。
_GOOGLE_QUERY_STATE = ContextVar("careflow_google_queries", default=None)
_GOOGLE_ALLOWED = ContextVar("careflow_google_allowed", default=True)
_GOOGLE_COLLECT = ContextVar("careflow_google_collect", default=None)

def set_google_queries_enabled(enabled):
    """網頁預設關閉；只有明確核對按鈕暫時允許付費查詢。"""
    _GOOGLE_ALLOWED.set(bool(enabled))


# 服務項目強度加權係數：依體力耗費強度分級，用於疲勞度模型。
SERVICE_INTENSITY_WEIGHT = {
    "重度移位": 1.5,  # 重度移位／肢體關節活動
    "餐食管灌": 1.2,  # 餐食照顧／管灌洗頭
    "一般照護": 0.8,  # 一般家務／陪伴看護
}

# 環境排斥條件對照表 (居服員排斥 -> 案家環境)，兩側皆為 Excel 原始字串，需完全相符
EXCLUSION_MAP = {
    "拒爬高樓(無電梯)": "傳統公寓無電梯",
    "拒寵物環境": "有養寵物",
    "拒菸害環境": "有抽菸",
}

# 法定 20 小時特照培訓：案家需求類別 -> 合格居服員「核心專長證照」集合。
# 居服員若不具備對應證照，該配對於 Phase 1 直接硬性剔除，不得派單。
SPECIAL_CERT_REQUIREMENTS = {
    "失智引導與精神陪伴": {"失智症照顧專長"},
}

# 星期中文名稱 -> ISO 星期數字（1=一 ... 7=日），供「可排班星期」防呆使用。
WEEKDAY_NAME_TO_NUM = {
    "星期一": 1, "星期二": 2, "星期三": 3, "星期四": 4,
    "星期五": 5, "星期六": 6, "星期日": 7,
}
WEEKDAY_NUM_TO_NAME = {v: k for k, v in WEEKDAY_NAME_TO_NUM.items()}

def get_weekday_name(date_value) -> str:
    """將日期值轉換為中文星期名稱（"星期一"...），供派單結果表格顯示用。

    無法解析（空值、格式不符）時回傳空字串，而非拋出例外，避免單一列的日期
    格式問題導致整張結果表格無法呈現。
    """
    if pd.isna(date_value) or str(date_value).strip() == "":
        return ""
    try:
        parsed = pd.to_datetime(date_value)
    except (ValueError, TypeError):
        return ""
    return WEEKDAY_NUM_TO_NAME.get(parsed.isoweekday(), "")

# 臺灣長照 2.0 常用 BA 服務代碼點數（點值，機構申報營收以此試算；1 點通常對應約 1 元）。
BA_UNIT_POINTS = {
    "BA01": 175,
    "BA02": 210,
    "BA03": 150,
    "BA04": 180,
    "BA05": 200,
    "BA07": 250,
    "BA08": 300,
    "BA09": 160,
    "BA10": 190,
}

# 無法由 Service_Code_1/2 判斷點數時（例如資料表未升級至含 BA 碼申報欄位），
# 退回以服務歷時概算點數：每 30 分鐘 = 150 點。
_FALLBACK_POINTS_PER_30MIN = 150.0

@dataclass
class PipelineConfig:
    """所有可調整的派單政策參數，預設值與原始腳本一致。"""

    # 轉場緩衝與交通時間模型
    buffer_mins: float = 15.0
    travel_min_per_km: float = 3.0

    # 勞動合規：累計連續工作達 continuous_work_limit_mins 分鐘，強制要求下一段任務
    # 與前段之間至少間隔 mandatory_break_mins 分鐘（見 Phase 2 的 _build_break_constraints）。
    continuous_work_limit_mins: float = 240.0
    mandatory_break_mins: float = 30.0

    # Phase 1：軟性適配度評分
    base_score: float = 60.0
    cert_bonus_dementia: float = 15.0
    cert_bonus_other: float = 10.0
    preferred_caregiver_bonus: float = 60.0
    previous_week_caregiver_bonus: float = 80.0
    continuity_performance_bonus: float = 15.0
    continuity_satisfaction_threshold: float = 4.3
    satisfaction_baseline: float = 4.0
    satisfaction_weight: float = 10.0
    travel_penalty_weight: float = 2.0
    travel_penalty_cap: float = 25.0
    # 工作負荷平衡：若資料未提供「當月可排總工時」，預設以 160 小時作為月可用工時基準
    fatigue_reference_hours: float = 160.0
    fatigue_weight: float = 10.0
    plan_load_mode: str = "hours"
    solver_time_limit_ms: int = 10000
    solver_relative_gap: float = 0.01
    service_unit_prices: dict = field(default_factory=dict)

    # Phase 1：車程上限硬性條件（None = 停用，不做硬性剔除）。
    # 照護連續性優先於車程限制：一般候選人受 max_travel_minutes 限制，
    # 但歷史首選居服員改用較寬鬆的 preferred_caregiver_max_travel_minutes
    # （或直接不受限，若該欄位亦為 None），避免熟悉的居服員被車程微幅超標而剔除。
    max_travel_minutes: Optional[float] = None
    preferred_caregiver_max_travel_minutes: Optional[float] = None

    # Phase 2：OR-Tools 目標函數權重
    urgent_priority_bonus: float = 50.0
    normal_priority_bonus: float = 20.0

    # 財務試算：長照申報點數換算居服員薪資的拆帳比例；
    # 僅供結果顯示與財務 KPI 試算，不參與派單決策。
    caregiver_salary_rate_per_point: float = 0.65

# ==========================================
# 資料載入
# ==========================================
def load_data(excel_path: str = DEFAULT_EXCEL_PATH, tasks_sheet_name: str = "Today_Pending_Tasks"):
    """讀取派單資料庫四張工作表。

    tasks_sheet_name 預設為現行單日排班格式的「Today_Pending_Tasks」；若改用含
    星期／日期欄位的月批次排班資料表，呼叫端可傳入 "Monthly_Pending_Tasks"。
    """
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"找不到檔案 {excel_path}，請確認檔案與腳本在同一目錄下。")

    df_cg = pd.read_excel(excel_path, sheet_name="Caregiver_Profiles")
    df_cl = pd.read_excel(excel_path, sheet_name="Client_Profiles")
    df_tasks = pd.read_excel(excel_path, sheet_name=tasks_sheet_name)
    xl = pd.ExcelFile(excel_path)
    df_hist = pd.read_excel(xl,sheet_name='Historical_Service_Logs') if 'Historical_Service_Logs' in xl.sheet_names else pd.DataFrame()

    tasks = merge_task_client_tables(df_tasks, df_cl)
    return df_cg, df_cl, df_tasks, df_hist, tasks

# ==========================================
# 機構三檔匯入：保留來源值，缺少資訊不偽造
# ==========================================
def _text(value):
    return '' if pd.isna(value) else str(value).replace('_x000D_', '').strip()


def _number(value, default=np.nan):
    try:
        number = float(value)
        return number if np.isfinite(number) else default
    except (TypeError, ValueError):
        return default


def _date(value):
    if not _text(value):
        return None
    parsed = pd.to_datetime(value, errors='coerce')
    return None if pd.isna(parsed) else parsed.date()


def _hhmm(value):
    import re
    match = re.fullmatch(r'(\d{1,2}):(\d{2})(?::\d{2})?', _text(value))
    if not match or int(match[1]) > 23 or int(match[2]) > 59:
        return None
    return f'{int(match[1]):02d}:{int(match[2]):02d}'


def parse_service_items(value):
    """保留任意數量服務碼與次數（含 BA05-1/BA15-1 等子碼）。"""
    import re
    text = _text(value).upper().replace('（', '(').replace('）', ')')
    pattern = re.compile(r'((?:BA|GA|SC|SA|AA)\d{2}(?:-\d+|[A-Z]\d*)?)\s*\(\s*(\d+(?:\.\d+)?)\s*\)')
    matches = list(pattern.finditer(text))
    remainder = pattern.sub('', text).strip(' ,，;；\n\t')
    if not matches or remainder:
        raise ValueError('服務項目格式無法完整解析，請使用 BA13(2) BA07(1) 格式')
    return [(m[1], float(m[2])) for m in matches]


def iter_service_codes(row):
    import re
    indices = sorted({int(m[1]) for key in row.keys()
                      if (m := re.match(r'^Service_Code_(\d+)(?:_[xy])?$', str(key)))})
    for index in indices:
        code = _get_task_field(row, f'Service_Code_{index}')
        if code is not None:
            yield str(code).strip().upper(), _number(_get_task_field(row, f'Units_{index}'))


def caregiver_has_cert(cg, cert):
    return cert in certificate_names(cg) or cert in _text(cg.get('核心專長證照')).split('、')


def merge_task_client_tables(tasks, clients):
    """多對一合併；任務座標優先，保留來源；重複鍵直接報錯。"""
    if clients['案家ID'].isna().any() or clients['案家ID'].duplicated().any():
        raise ValueError('案家ID 不可空白或重複；請先修正案家資料')
    clients = clients.copy()
    if any(k in clients for k in ['核定等級','失能程度','障礙程度','身障手冊','身障證明']):
        for field in difficulty_reference({}):
            clients[field] = clients.apply(lambda row: difficulty_reference(row)[field],axis=1)
    merged = tasks.merge(clients, on='案家ID', how='left', suffixes=('', '_案家'), validate='many_to_one', indicator=True)
    if merged['_merge'].eq('left_only').any():
        raise ValueError('任務有案家ID無法對應個案表，請確認三檔使用同一套代碼')
    for col in clients.columns:
        alternate = col + '_案家'
        if alternate in merged:
            merged[col] = merged[col].combine_first(merged[alternate])
    return merged.drop(columns='_merge')


def _read_agency_table(file_or_frame):
    if isinstance(file_or_frame, pd.DataFrame):
        frame = file_or_frame.copy()
    else:
        frame = pd.read_excel(file_or_frame, sheet_name=0, dtype=object)
    frame.columns = [str(c).strip() for c in frame.columns]
    return frame


def load_agency_data(employee_file, client_file, weekly_file, start_date, end_date,
                     daily_start='08:00', daily_end='21:00', daily_cap=8.0,
                     availability_confirmed=False, biweekly_anchor_confirmed=False):
    """將三個機構檔轉為既有引擎資料表；只產生排班草案，回傳待補資料清單。"""
    g, c, w = map(_read_agency_table, (employee_file, client_file, weekly_file))

    # 個資保護：姓名／電話／地址等直接識別欄位不進入排班核心。
    # 內部編號、案號、座標與排班必要欄位保留；姓名由獨立對照 sheet 只用於顯示。
    privacy_removed = {}
    g, privacy_removed["員工資料"] = _strip_direct_identifiers(g)
    c, privacy_removed["個案資料"] = _strip_direct_identifiers(c)
    w, privacy_removed["週服務計畫"] = _strip_direct_identifiers(w)

    for title, frame, required in [
        ('員工', g, ['編號', '性別', '休息日', '取得資格', '交通工具']),
        ('個案', c, ['內部案號']),
        ('週服務', w, ['案號', '星期', '開始時間', '結束時間', '服務分鐘', '頻率', '生效日', '服務項目'])]:
        missing = [col for col in required if col not in frame]
        if missing:
            raise ValueError(title + '檔缺少欄位：' + '、'.join(missing))
    start, end = _date(start_date), _date(end_date)
    if start is None or end is None or end < start or (end-start).days > 30:
        raise ValueError('請選擇有效日期範圍（最多31天）')
    if not _hhmm(daily_start) or not _hhmm(daily_end) or daily_end <= daily_start or daily_cap <= 0:
        raise ValueError('每日可服務時段及工時上限設定無效')
    issues = []
    def issue(source, rownum, reason):
        issues.append({'資料表': source, '來源列': rownum, '問題': reason})

    for source, removed_cols in privacy_removed.items():
        if removed_cols:
            issue(source, None, "個資保護：排班運算已移除直接識別欄位：" + "、".join(removed_cols))
    g['居服員ID'] = g['編號'].map(_text)
    if g['居服員ID'].eq('').any() or g['居服員ID'].duplicated().any():
        raise ValueError('員工編號空白或重複，不能安全串檔')
    c['案家ID'] = c['內部案號'].map(_text)
    for idx in c.index[c['案家ID'].eq('')]:
        issue('個案', int(idx)+2, '缺內部案號，暫不匯入；請補原案號後重新匯入')
    c = c.loc[c['案家ID'].ne('')].copy()
    if c['案家ID'].duplicated().any():
        raise ValueError('內部案號重複，不能安全串檔')
    g['常用交通工具'] = g['交通工具'].map(lambda v: '大眾運輸' if any(x in _text(v) for x in ['公車','捷運','大眾']) else _text(v))
    for frame, lon, lat in [(g,'服務起點_經度(家)','服務起點_緯度(家)'), (c,'服務地點_經度','服務地點_緯度')]:
        frame[lon] = pd.to_numeric(frame.get('經度 (WGS84_Lon)', pd.Series(np.nan,index=frame.index)), errors='coerce')
        frame[lat] = pd.to_numeric(frame.get('緯度 (WGS84_Lat)', pd.Series(np.nan,index=frame.index)), errors='coerce')
    def certs(v):
        text = _text(v); result=[]
        for needle, name in [('失智症照顧服務20小時','失智症照顧專長'),('精神疾病','精神疾病照顧專長'),
                             ('照顧服務員單一級','單一級照服證照'),('身心障礙支持','身心障礙支持服務核心課程'),('BA08','BA08足部照護')]:
            if needle in text: result.append(name)
        return '、'.join(result)
    g['核心專長證照'] = g['取得資格'].map(certs)
    def available_days(v):
        text=_text(v)
        days={WEEKDAY_NAME_TO_NUM.get(text)}
        if text.startswith('週') and text[-1:] in '一二三四五六日天':
            days={7 if text[-1:] in '日天' else '一二三四五六'.index(text[-1:])+1}
        if not text or None in days:
            return ''
        return ','.join(str(n) for n in range(1,8) if n not in days)
    g['可排班星期'] = g['休息日'].map(available_days)
    g['每日可服務時段_起'],g['每日可服務時段_迄'] = daily_start,daily_end
    g['每日工時上限(小時)'] = float(daily_cap)
    g['可服務資料已確認'] = bool(availability_confirmed)
    g['可服務設定來源'] = '操作者採用的排班草案假設；非機構提供的可服務時段'
    for field in ['當月累計服務時數(疲勞度)','歷史滿意度均值','具備重度移位體力(0/1)']:
        if field not in g: g[field] = np.nan
    if '失智症個案(0/1)' not in c: c['失智症個案(0/1)'] = np.nan
    g['本次計畫已排時數'] = 0.0
    g['本次計畫加權負荷時數'] = 0.0
    g['特殊排斥條件'] = g.get('特殊排斥條件', '')
    for field in ['指定居服員性別','特殊照護需求','案家環境特徵','歷史首選居服員ID']:
        if field not in c: c[field] = ''
    if '需重度移位協助(0/1)' not in c: c['需重度移位協助(0/1)'] = np.nan
    for field in ['照護難度參考係數','照護難度判斷依據','照護提示','照護難度規則版本']:
        c[field] = c.apply(lambda row: difficulty_reference(row)[field],axis=1)
    c['照護條件待確認'] = True
    g['請假紀錄'] = pd.Series([None]*len(g),index=g.index,dtype=object)
    for idx,row in g.iterrows():
        if _text(row.get('類別')) != '請假': continue
        begin,finish = _date(row.get('開始')),_date(row.get('結束'))
        parts = _text(row.get('時間')).split('-')
        t1,t2 = (_hhmm(p.strip()) for p in parts) if len(parts)==2 else (None,None)
        if begin is None or finish is None or finish < begin or t1 is None or t2 is None or (begin==finish and t2<=t1):
            issue('員工', int(idx)+2, '請假區間無效或零時數；此員工暫停派單，請修正來源請假欄位後重新匯入')
            g.at[idx,'可服務資料已確認'] = False
            continue
        g.at[idx,'請假紀錄'] = (begin.isoformat()+' '+t1,finish.isoformat()+' '+t2)
    known=set(c['案家ID']); expanded=[]
    for idx,row in w.iterrows():
        rownum=int(idx)+2; client=_text(row.get('案號'))
        if not client or client not in known:
            issue('週服務',rownum,'缺案號或無法對應個案表，暫不產生任務'); continue
        weekday=WEEKDAY_NAME_TO_NUM.get(_text(row.get('星期')))
        begin,finish = _date(row.get('生效日')),_date(row.get('到期日'))
        t1,t2=_hhmm(row.get('開始時間')),_hhmm(row.get('結束時間'))
        minutes=_number(row.get('服務分鐘'))
        if weekday is None or begin is None or t1 is None or t2 is None or t2<=t1 or not np.isfinite(minutes) or minutes<=0:
            issue('週服務',rownum,'星期、生效日、時間或服務分鐘無效，暫不產生任務');continue
        occupied=(datetime.strptime(t2,'%H:%M')-datetime.strptime(t1,'%H:%M')).total_seconds()/60
        if minutes!=occupied:
            issue('週服務',rownum,'服務分鐘與起迄時間不符；暫不產生任務以免低估工時');continue
        if _text(row.get('到期日')) and finish is None:
            issue('週服務',rownum,'到期日無法解析，暫不產生任務');continue
        frequency=_text(row.get('頻率')).replace('隔周', '隔週')
        if frequency not in ('每週','隔週'):
            issue('週服務',rownum,'頻率無法辨識，暫不產生任務');continue
        try: codes=parse_service_items(row.get('服務項目'))
        except ValueError as exc:
            issue('週服務',rownum,str(exc));continue
        if not codes or any(units<=0 for _,units in codes):
            issue('週服務',rownum,'服務次數必須大於零');continue
        for day in pd.date_range(start,end):
            day=day.date()
            if day.isoweekday()!=weekday or day<begin or (finish and day>finish): continue
            if frequency=='隔週' and ((day-timedelta(days=day.weekday()))-(begin-timedelta(days=begin.weekday()))).days//7%2: continue
            record=row.to_dict()
            record.update({'任務ID':f'W{rownum:04d}_{day:%Y%m%d}','案家ID':client,'日期':day.isoformat(),
                '星期':WEEKDAY_NUM_TO_NAME[weekday],'時間窗_開始':t1,'時間窗_結束':t2,
                '服務歷時(分鐘)':minutes,'服務歷時來源':'機構服務分鐘／起迄時間相符',
                '任務優先級':'一般','是否為週期性任務':True,'來源列':rownum,'資料來源':'機構三檔',
                '任務待確認事項':'照護需求、性別指定、環境限制及首選居服員未提供；座標精度未確認'})
            record['服務地點_經度']=_number(row.get('經度 (WGS84_Lon)'))
            record['服務地點_緯度']=_number(row.get('緯度 (WGS84_Lat)'))
            for i,(code,units) in enumerate(codes,1):
                record[f'Service_Code_{i}'],record[f'Units_{i}']=code,units
            expanded.append(record)
    task_columns=['任務ID','案家ID','日期','星期','時間窗_開始','時間窗_結束','服務歷時(分鐘)',
                  '任務優先級','是否為週期性任務','Service_Code_1','Units_1','資料來源']
    tasks=pd.DataFrame(expanded) if expanded else pd.DataFrame(columns=task_columns)
    if not tasks.empty:
        # 同案同日重疊可能是重複計畫或多項併單，保留待確認並阻擋，不自行合併或刪除。
        tasks['任務資料異常']=''
        for (_,day), group in tasks.groupby(['案家ID','日期']):
            ordered=group.sort_values('時間窗_開始')
            pairs=list(ordered.iterrows())
            for i,(i1,r1) in enumerate(pairs):
                for i2,r2 in pairs[i+1:]:
                    if r2['時間窗_開始'] < r1['時間窗_結束']:
                        for key in (i1,i2):tasks.at[key,'任務資料異常']='同案同日服務時段重疊，需人工確認'
        overlap_rows=tasks.loc[tasks['任務資料異常'].ne(''),'來源列'].unique()
        for rownum in overlap_rows:issue('週服務',int(rownum),'同案同日服務重疊；對應任務暫不派單')
    all_codes=sorted({code for _,row in tasks.iterrows() for code,_ in iter_service_codes(row)})
    master=pd.DataFrame({'系統代碼':all_codes,'CareFlow排班分鐘(暫定)':np.nan,'是否納入CareFlow':'待確認','單價':np.nan})
    master['必要證照'] = master['系統代碼'].map(DEFAULT_CODE_CERTS).fillna('')
    history=pd.DataFrame(columns=['居服員ID','案家ID','歷史媒合機制(Treatment)','不滿意導致提早結案(0/1)','案家滿意度(1-5)'])
    return {'Caregiver_Profiles':g,'Client_Profiles':c,'Tasks':tasks,'Historical_Service_Logs':history,
            'Service_Code':master,'Import_Issues':pd.DataFrame(issues,columns=['資料表','來源列','問題'])}


def load_agency_workbook(workbook_file, start_date, end_date, **kwargs):
    """單一 Excel：4 張運算工作表 + 2 張姓名對照；姓名對照不進入排班核心。"""
    names = {
        '員工資料': ['員工資料', '員工資料_去識別化', 'Caregiver_Profiles'],
        '個案資料': ['個案資料', '個案資料_去識別化', 'Client_Profiles'],
        '週服務計畫': ['週服務計畫', '周服務計畫', 'Weekly_Service_Plan'],
    }
    with pd.ExcelFile(workbook_file) as xl:
        selected = {}
        for role, candidates in names.items():
            matches = [name for name in candidates if name in xl.sheet_names]
            if len(matches) != 1:
                raise ValueError(f'請保留一張「{role}」工作表；目前找不到或有多張同類工作表')
            selected[role] = pd.read_excel(xl, sheet_name=matches[0], dtype=object)
        master_names = [n for n in xl.sheet_names if n.lower() == 'service_code']
        if len(master_names)>1: raise ValueError('Service_code 工作表重複，請只保留一張')
        master = pd.read_excel(xl, sheet_name=master_names[0], dtype=object) if master_names else None
        if master is not None:
            master = master.loc[:,~master.columns.astype(str).str.startswith('Unnamed:')].copy()
            master = master.dropna(subset=['系統代碼'])
            master['系統代碼']=master['系統代碼'].astype(str).str.strip().str.upper()
            if '必要證照' not in master: master['必要證照']=master['系統代碼'].map(DEFAULT_CODE_CERTS).fillna('')
    result = load_agency_data(selected['員工資料'], selected['個案資料'], selected['週服務計畫'],
                            start_date, end_date, **kwargs)
    if master is not None:
        result['Service_Code'] = master
        known=set(master['系統代碼'])
        used={code for _,row in result['Tasks'].iterrows() for code,units in iter_service_codes(row) if units>0}
        missing=sorted(used-known)
        if missing:
            notes=pd.DataFrame([{'資料表':'Service_code','來源列':None,'問題':'主檔缺少 '+code+'；主檔模式需補分鐘，不猜測或以其他碼替代'} for code in missing])
            result['Import_Issues']=pd.concat([result['Import_Issues'],notes],ignore_index=True)
    return result


def validate_dispatch_inputs(tasks, caregivers):
    for title,frame,key in [('任務',tasks,'任務ID'),('員工',caregivers,'居服員ID')]:
        if key not in frame or frame[key].isna().any() or frame[key].astype(str).str.strip().eq('').any() or frame[key].duplicated().any():
            raise ValueError(title+' ID 空白或重複，請修正後排班')
    for _,task in tasks.iterrows():
        begin,end=_hhmm(task.get('時間窗_開始')),_hhmm(task.get('時間窗_結束'))
        if begin is None or end is None or end<=begin:
            raise ValueError('任務起迄時間無效，請使用 HH:MM 且結束晚於開始')
        if not np.isfinite(_number(task.get('服務歷時(分鐘)'))) or _number(task.get('服務歷時(分鐘)'))<=0:
            raise ValueError('任務服務分鐘必須為有效正數')
        if task.get('資料來源')=='機構三檔':
            span=(datetime.strptime(end,'%H:%M')-datetime.strptime(begin,'%H:%M')).total_seconds()/60
            if span != _number(task.get('服務歷時(分鐘)')):
                raise ValueError('服務分鐘與起迄時間不符，請同步修正後排班')



def _get_task_field(row, base_col_name: str):
    """依序嘗試 base_col_name、base_col_name_x、base_col_name_y，取得第一個有值的欄位。

    新版月批次任務資料表的 Service_Code_1/2、Units_1/2 欄位，若在 Client_Profiles
    亦重複定義同名欄位，經 `tasks_df.merge(df_cl, on="案家ID")` 合併後，pandas 會
    自動將同名欄位改為 base_col_name_x（左表／任務本身）、base_col_name_y（右表／
    案家）。本函式確保無論合併後欄位是否被加上後綴，皆優先取任務本身、其次取
    案家層級的對應值，而不會因欄位改名誤判為「缺少代碼」。row 可為 Series 或 dict。
    """
    for col in (base_col_name, f"{base_col_name}_x", f"{base_col_name}_y"):
        value = row.get(col)
        if pd.notna(value) and str(value).strip():
            return value
    return None

def apply_service_duration(
    tasks_df: pd.DataFrame,
    service_code_df: pd.DataFrame,
) -> pd.DataFrame:
    """依 Service_Code Master 的分鐘數，重算每筆任務服務歷時。"""

    result = tasks_df.copy()
    if '服務歷時來源' in result:
        supplied = result['服務歷時來源'].fillna('').astype(str).str.startswith(('機構', '人工確認'))
        if '使用Service_code重算' in result:
            supplied &= ~result['使用Service_code重算'].fillna(False).astype(bool)
        if supplied.any():
            if supplied.all():
                minutes = pd.to_numeric(result['服務歷時(分鐘)'], errors='coerce')
                if minutes.isna().any() or minutes.le(0).any():
                    raise ValueError('機構或人工確認服務分鐘必須為正數')
                result['服務歷時計算明細'] = result['服務歷時來源']
                return result
            other = apply_service_duration(result.loc[~supplied].drop(columns='服務歷時來源'), service_code_df)
            own = apply_service_duration(result.loc[supplied], service_code_df)
            return pd.concat([own, other]).sort_index()

    required_columns = [
        "系統代碼",
        "CareFlow排班分鐘(暫定)",
        "是否納入CareFlow",
    ]
    missing_columns = [
        column
        for column in required_columns
        if column not in service_code_df.columns
    ]

    if missing_columns:
        raise ValueError(
            "Service_Code 缺少必要欄位："
            + "、".join(missing_columns)
        )

    master = service_code_df.copy()
    master = master.dropna(subset=["系統代碼"])
    master["系統代碼"] = (
        master["系統代碼"].astype(str).str.strip().str.upper()
    )

    duplicated_codes = master[
        master["系統代碼"].duplicated(keep=False)
    ]["系統代碼"].unique()

    if len(duplicated_codes) > 0:
        raise ValueError(
            "Service_Code Master 有重複代碼："
            + "、".join(duplicated_codes)
        )

    master = master.set_index("系統代碼")

    calculated_minutes = []
    required_certificates = []
    calculation_details = []

    for _, row in result.iterrows():
        task_id = row.get("任務ID", "未知任務")
        total_minutes = 0.0
        details = []
        certificates = set()

        for code_raw, units_raw in iter_service_codes(row):

            if pd.isna(code_raw) or str(code_raw).strip() == "":
                continue

            code = str(code_raw).strip().upper()

            if pd.isna(units_raw):
                raise ValueError(
                    f"任務 {task_id} 的 {code} 未填寫 Units"
                )

            try:
                units = float(units_raw)
            except (TypeError, ValueError):
                raise ValueError(
                    f"任務 {task_id} 的 {code} Units 不是有效數字"
                )

            if units < 0:
                raise ValueError(
                    f"任務 {task_id} 的 {code} Units 不可小於 0"
                )

            if units == 0:
                continue

            if code not in master.index:
                raise ValueError(
                    f"任務 {task_id} 的服務碼 {code} "
                    "不存在於 Service_Code Master"
                )

            master_row = master.loc[code]
            configured = _text(master_row.get('必要證照'))
            if configured:
                certificates.update(x.strip() for x in configured.split('、') if x.strip())
            elif code in DEFAULT_CODE_CERTS:
                certificates.add(DEFAULT_CODE_CERTS[code])
            status = str(
                master_row["是否納入CareFlow"]
            ).strip()

            # AA07～AA11是附加碼，不另增加服務分鐘
            if code.startswith("AA"):
                details.append(f"{code}×{units:g}=0")
                continue

            if status in {"否", "否/另模組"}:
                raise ValueError(
                    f"任務 {task_id} 使用目前不支援排班的服務碼 {code}"
                )

            minutes = master_row["CareFlow排班分鐘(暫定)"]

            if pd.isna(minutes):
                raise ValueError(
                    f"Service_Code Master 的 {code} "
                    "尚未設定CareFlow排班分鐘"
                )

            if not np.isfinite(float(minutes)) or float(minutes) <= 0:
                raise ValueError(f'{code} 的排班分鐘必須是有效正數')
            subtotal = float(minutes) * units
            total_minutes += subtotal
            details.append(
                f"{code}×{units:g}={subtotal:g}分鐘"
            )

        if total_minutes <= 0:
            raise ValueError(
                f"任務 {task_id} 無法計算出有效服務歷時"
            )

        required_certificates.append("、".join(sorted(certificates)))
        calculated_minutes.append(total_minutes)
        calculation_details.append(" + ".join(details))

    # 保留Excel原有值，方便比較
    if "服務歷時(分鐘)" in result.columns:
        result["原服務歷時(分鐘)"] = result["服務歷時(分鐘)"]

    result["服務歷時(分鐘)"] = calculated_minutes
    result["服務歷時計算明細"] = calculation_details
    result['服務碼必要證照'] = required_certificates
    result['服務歷時來源'] = 'Service_code最新分鐘×Units'
    if '使用Service_code重算' in result:
        for idx,row in result.iterrows():
            if not bool(row.get('使用Service_code重算')): continue
            begin=pd.Timestamp('2000-01-01 '+str(row['時間窗_開始']))
            finish=begin+pd.Timedelta(minutes=float(row['服務歷時(分鐘)']))
            if finish.date()!=begin.date(): raise ValueError('服務時間跨午夜，請拆分任務')
            result.at[idx,'原計畫結束時間']=row.get('原計畫結束時間',row['時間窗_結束'])
            result.at[idx,'時間窗_結束']=finish.strftime('%H:%M')
    return result
# ==========================================
# 申報法規防呆：BA 服務代碼併報合規檢核
# ==========================================
def check_ba_code_compatibility(task_row) -> Optional[str]:
    """檢查單一任務的 BA 服務代碼併報合規性（依長照給付支付基準併報規則）。

    task_row 可為 DataFrame 的一列（Series）或 dict，僅需能以 `.get()` 取出
    Service_Code_1 / Service_Code_2（或合併後帶 _x/_y 後綴的同名欄位，見
    `_get_task_field`）。回傳 None 代表未發現已知違規；否則回傳違規說明文字。
    本函式僅檢核並回報，不會自動剔除或修改任務——是否略過警告並放行進入排程，
    由呼叫端（例如居督於 app.py 的人工複核介面）決定。
    """
    codes = {code for code, _ in iter_service_codes(task_row)}

    if "BA01" in codes and ("BA07" in codes or "BA23" in codes):
        return "既有規則提示（需人工複核）：BA01（基本身體清潔）不可與 BA07/BA23（沐浴/洗頭）同時段併申報"

    if "BA02" in codes and len(codes) > 1:
        other_codes = codes - {"BA02"}
        if not other_codes.issubset({"BA22"}):
            return f"既有規則提示（需人工複核）：BA02（基本日常照顧）除 BA22 外不得與其他項目 ({other_codes}) 評定併用"

    if "BA01" in codes and "BA24" in codes:
        return "注意：BA01 與 BA24 同時段申報可能涉及排泄項目重複，請確認計畫書核定內容"

    return None

def validate_ba_codes(tasks_df: pd.DataFrame) -> pd.DataFrame:
    """對整批任務套用 check_ba_code_compatibility，回傳新增檢核欄位的副本。

    新增「BA代碼檢核異常」（違規說明文字或 None）與「含違規代碼」（布林值）兩欄，
    供派單前的法規防呆健檢（例如上傳資料後的即時檢核儀錶板）使用；不修改傳入的
    DataFrame，亦不會排除任何任務列。
    """
    result = tasks_df.copy()
    messages = [check_ba_code_compatibility(row) for _, row in result.iterrows()]
    result["BA代碼檢核異常"] = messages
    result["含違規代碼"] = [m is not None for m in messages]
    return result

# ==========================================
# 財務試算：長照申報點數（營收）與居服員拆帳薪資
# ==========================================
def service_price_map(master: pd.DataFrame) -> dict:
    """只讀機構主檔單價；未知、負值、非數字或重複碼不猜測。"""
    if not {'系統代碼', '目前給付價格(元)'}.issubset(master.columns):
        return {}
    frame = master[['系統代碼', '目前給付價格(元)']].copy()
    frame['系統代碼'] = frame['系統代碼'].astype(str).str.strip().str.upper()
    duplicate = set(frame.loc[frame['系統代碼'].duplicated(False), '系統代碼'])
    prices = {}
    for code, raw in frame.itertuples(index=False, name=None):
        price = _number(str(raw).replace(',', '').strip())
        if code and code not in duplicate and np.isfinite(price) and price >= 0:
            prices[code] = price
    return prices


def calculate_task_revenue_and_salary(tasks_df: pd.DataFrame, config: "PipelineConfig") -> pd.DataFrame:
    """單價×次數加總，再乘拆帳比例。機構資料僅採最新 Service_code 單價。"""
    result = tasks_df.copy()
    revenues, statuses = [], []
    for _, row in result.iterrows():
        prices = config.service_unit_prices
        if row.get('資料來源') != '機構三檔' and not prices:
            prices = BA_UNIT_POINTS  # 保留原模型相容性
        items = list(iter_service_codes(row))
        missing = sorted({code for code, units in items if code not in prices})
        invalid = any(not np.isfinite(units) or units < 0 for _, units in items)
        if not items or missing or invalid:
            revenues.append(np.nan)
            statuses.append('缺少單價：' + '、'.join(missing) if missing else '服務碼或次數無效')
        else:
            revenues.append(round(sum(prices[code] * units for code, units in items), 2))
            statuses.append('Service_code單價×次數試算')
    result['預估長照申報點數(營收)'] = revenues
    result['預估居服員拆帳薪資'] = (result['預估長照申報點數(營收)'] * config.caregiver_salary_rate_per_point).round(2)
    result['財務估算狀態'] = statuses
    return result

# ==========================================
# 共用工具
# ==========================================
def calc_distance_km(lat1, lon1, lat2, lon2):
    """以臺灣緯度估算直線距離 km（1度緯度約111km，1度經度約101km）。"""
    dlat = (lat1 - lat2) * 111.0
    dlon = (lon1 - lon2) * 101.0
    return np.sqrt(dlat**2 + dlon**2)

# 字典快取已查詢過的經緯度對 -> 路網時間(分鐘)，Phase 1 / Phase 2 共用同一份快取，
# 避免同一對座標於不同階段重複發送 API 請求。
_OSRM_TRAVEL_TIME_CACHE: dict = {}

def _round_coord(v: float) -> float:
    return round(float(v), 6)

def _osrm_fallback_minutes(lat1, lon1, lat2, lon2, travel_min_per_km, reason) -> float:
    """降級為 Haversine/歐式距離估算，並印出 Warning Log。"""
    print(
        f"[OSRM Warning] 路網查詢失敗 ({lat1},{lon1}) -> ({lat2},{lon2})，"
        f"降級為 Haversine/歐式距離估算: {reason}"
    )
    return calc_distance_km(lat1, lon1, lat2, lon2) * travel_min_per_km

def _osrm_endpoint(transport_mode='機車'):
    def setting(key, default):
        try:
            value = st.session_state.get(key)
        except Exception:
            value = None
        return str(value or os.getenv(key) or default).rstrip('/')
    if transport_mode == '機車':
        return setting('CAREFLOW_OSRM_MOTORCYCLE_HOST', 'http://127.0.0.1:5001'), 'motorcycle'
    raise ValueError('OSRM 機車路網僅支援機車；公車／捷運請使用 Google Routes。本版不提供其他交通方式。')


def _osrm_cache_key(lat1, lon1, lat2, lon2, mode):
    host, profile = _osrm_endpoint(mode)
    return (host, profile, mode, _round_coord(lat1), _round_coord(lon1), _round_coord(lat2), _round_coord(lon2))


def get_osrm_travel_time(lat1, lon1, lat2, lon2, travel_min_per_km=3.0, transport_mode='機車'):
    """查詢對應交通方式的路網；失敗時明確停止，不拿汽車或直線冒充機車。"""
    host, profile = _osrm_endpoint(transport_mode)
    if not all(np.isfinite(_number(v)) for v in (lat1,lon1,lat2,lon2)):
        raise ValueError('缺有效座標，無法估算交通')
    if lat1 == lat2 and lon1 == lon2:
        return 0.0
    key = _osrm_cache_key(lat1, lon1, lat2, lon2, transport_mode)
    if key in _OSRM_TRAVEL_TIME_CACHE:
        return _OSRM_TRAVEL_TIME_CACHE[key]
    try:
        resp = requests.get(f'{host}/route/v1/{profile}/{lon1},{lat1};{lon2},{lat2}', params={'overview':'false'}, timeout=OSRM_TIMEOUT_SECONDS)
        resp.raise_for_status()
        data = resp.json()
        if data.get('code') != 'Ok' or not data.get('routes'):
            raise ValueError('無可用路線：' + str(data.get('code')))
        minutes = float(data['routes'][0]['duration']) / 60.0
        if not np.isfinite(minutes) or minutes < 0:
            raise ValueError('回傳無效時間')
    except Exception as exc:
        raise RuntimeError(f'OSRM {transport_mode}路網查詢失敗。機車需先依 OSRM_機車啟動說明 建置服務；未使用汽車或直線替代。原因：{exc}') from exc
    _OSRM_TRAVEL_TIME_CACHE[key] = minutes
    return minutes

def _get_google_maps_api_key():
    """沿用環境變數或 Streamlit secrets，金鑰不寫入日誌、結果或壓縮包。"""
    key = os.getenv('GOOGLE_MAPS_API_KEY')
    if not key:
        try:
            key = st.secrets.get('GOOGLE_MAPS_API_KEY')
        except Exception:
            key = None
    return str(key).strip() if key else None


def google_key_is_configured():
    return bool(_get_google_maps_api_key())


def begin_travel_query_run(max_queries=12000):
    # count 以 route-matrix element 計，不是 HTTP request 數。
    # 僅當次計算共用，不跨使用者／排班批次共用。
    _GOOGLE_QUERY_STATE.set({'count': 0, 'limit': int(max_queries), 'memo': {}})


def _google_query_state():
    state = _GOOGLE_QUERY_STATE.get()
    if state is None:
        begin_travel_query_run()
        state = _GOOGLE_QUERY_STATE.get()
    return state


def _bucket_transit_datetime(value, bucket_minutes=15):
    """Phase 1 用 15 分鐘時槽合併 Google TRANSIT 查詢；最終已派路段仍用精確時間驗證。"""
    if value is None:
        return None
    ts = pd.Timestamp(value)
    if pd.isna(ts):
        return None
    if ts.tzinfo is None:
        ts = ts.tz_localize("Asia/Taipei")
    minute = (ts.minute // bucket_minutes) * bucket_minutes
    return ts.replace(minute=minute, second=0, microsecond=0)


def _google_timestamp(value):
    if value is None:
        return None
    ts = pd.Timestamp(value)
    if pd.isna(ts):
        return None
    if ts.tzinfo is None:
        ts = ts.tz_localize('Asia/Taipei')
    return ts.isoformat()


def _task_clock(task, clock):
    day = task.get('日期') if hasattr(task, 'get') else task
    if day is None or pd.isna(day) or str(day) in ('', '單日'):
        day = pd.Timestamp.now(tz='Asia/Taipei').date()
    return pd.Timestamp(str(pd.Timestamp(day).date()) + ' ' + pd.Timestamp(clock).strftime('%H:%M:%S'))


def _task_arrival(task):
    return _task_clock(task, pd.Timestamp('2000-01-01 ' + str(task['時間窗_開始'])))


def _travel_source(mode):
    provider = _get_travel_provider()
    if not _GOOGLE_ALLOWED.get() and (provider == 'google' or (provider == 'hybrid' and mode == '大眾運輸')):
        return 'Google Maps（公車／捷運）'
    if provider == 'hybrid':
        return 'OSRM 機車路網' if mode == '機車' else 'Google Maps（公車／捷運）'
    return {'osrm':'OSRM 機車路網', 'google':'Google Maps', 'local':'本機距離概算'}[provider]


def _get_travel_provider():
    """
    交通時間來源：
    1. 優先讀頁面交通選擇
    2. 其次讀環境變數或 Streamlit secrets
    3. 引擎未設定時使用本機概算，頁面預設混合模式
    """
    try:
        selected = st.session_state.get('careflow_travel_provider')
    except Exception:
        selected = None
    provider = selected or os.getenv("CAREFLOW_TRAVEL_PROVIDER")

    if not provider:
        try:
            provider = st.secrets.get(
                "CAREFLOW_TRAVEL_PROVIDER",
                "local",
            )
        except Exception:
            provider = "local"

    provider = str(provider).strip().lower()

    if provider not in ("google", "osrm", "hybrid", "local"):
        provider = "local"

    return provider

def _use_google_routes():
    # Google 僅處理大眾運輸；機車固定 OSRM，因此始終保留 OSRM 批次暖身。
    return False

def _parse_google_duration(duration_str):
    """
    Google duration 格式例如 '723s'、'723.5s'
    → 回傳分鐘。
    """
    if not duration_str:
        return None

    seconds = float(str(duration_str).rstrip("s"))
    return seconds / 60.0


def _google_transit_key(lat1, lon1, lat2, lon2, departure=None, arrival=None):
    return (
        _round_coord(lat1), _round_coord(lon1),
        _round_coord(lat2), _round_coord(lon2),
        "TRANSIT", departure, arrival,
    )


def _google_waypoint(lat, lon):
    return {
        'waypoint': {
            'location': {
                'latLng': {
                    'latitude': float(lat),
                    'longitude': float(lon),
                }
            }
        }
    }


def _fetch_google_transit_matrix_raw(origins, destinations, departure=None, arrival=None):
    """單一 Google ComputeRouteMatrix request；不直接修改 query state。"""
    api_key = _get_google_maps_api_key()
    if not api_key:
        raise RuntimeError(
            '公車／捷運需 Google Routes：請在 .streamlit/secrets.toml 設定 '
            'GOOGLE_MAPS_API_KEY，並啟用 Routes API 與帳務。'
        )

    elements = len(origins) * len(destinations)
    if elements <= 0 or elements > GOOGLE_TRANSIT_MAX_ELEMENTS_PER_REQUEST:
        raise ValueError(
            f'Google TRANSIT matrix elements 必須介於 1～'
            f'{GOOGLE_TRANSIT_MAX_ELEMENTS_PER_REQUEST}，目前為 {elements}'
        )

    body = {
        'origins': [_google_waypoint(*p) for p in origins],
        'destinations': [_google_waypoint(*p) for p in destinations],
        'travelMode': 'TRANSIT',
        'languageCode': 'zh-TW',
        'regionCode': 'TW',
    }
    if departure:
        body['departureTime'] = departure
    if arrival:
        body['arrivalTime'] = arrival

    _validate_google_routes_payload(body)
    headers = {
        'Content-Type': 'application/json',
        'X-Goog-Api-Key': api_key,
        # 只拿排班真正需要的欄位，降低回傳量與 latency。
        'X-Goog-FieldMask': 'originIndex,destinationIndex,status,condition,duration',
    }

    response = requests.post(
        GOOGLE_ROUTES_URL,
        headers=headers,
        json=body,
        timeout=GOOGLE_ROUTES_TIMEOUT_SECONDS,
    )
    if not response.ok:
        raise RuntimeError(
            f'Google Routes HTTP {response.status_code}；'
            '請檢查金鑰、Routes API、帳務與配額'
        )

    data = response.json()
    if not isinstance(data, list):
        raise RuntimeError('Google Routes matrix 回傳格式異常')

    result = {}
    for element in data:
        oi = element.get('originIndex')
        di = element.get('destinationIndex')
        if oi is None or di is None:
            continue

        key = _google_transit_key(
            *origins[int(oi)],
            *destinations[int(di)],
            departure=departure,
            arrival=arrival,
        )
        code = element.get('status', {}).get('code', 0)
        condition = element.get('condition')
        minutes = _parse_google_duration(element.get('duration'))

        if (
            code
            or condition != 'ROUTE_EXISTS'
            or minutes is None
            or not np.isfinite(minutes)
            or minutes < 0
        ):
            result[key] = float('inf')
        else:
            result[key] = float(minutes)

    # Google 可能不回傳不存在的 element；補成 inf，避免後面又逐筆重查。
    for origin in origins:
        for destination in destinations:
            key = _google_transit_key(
                *origin, *destination,
                departure=departure, arrival=arrival,
            )
            result.setdefault(key, float('inf'))

    return result


def _build_google_matrix_jobs(required_pairs):
    """把稀疏 pair 依相同 departure/arrival 分組，再貪婪打包成 <=100 elements 的矩陣。"""
    grouped = {}
    for lat1, lon1, lat2, lon2, departure, arrival in required_pairs:
        dep = _google_timestamp(departure) if departure is not None else None
        arr = _google_timestamp(arrival) if arrival is not None else None
        if dep and arr:
            raise ValueError('交通查詢只能指定出發或抵達時間其中一個')
        if not dep and not arr:
            dep = _google_timestamp(pd.Timestamp.now(tz='Asia/Taipei').floor('min'))
        group_key = (dep, arr)
        o = (_round_coord(lat1), _round_coord(lon1))
        d = (_round_coord(lat2), _round_coord(lon2))
        grouped.setdefault(group_key, {}).setdefault(d, set()).add(o)

    jobs = []
    for (dep, arr), by_dest in grouped.items():
        block_origins = set()
        block_dests = []

        def flush():
            nonlocal block_origins, block_dests
            if block_origins and block_dests:
                jobs.append((
                    sorted(block_origins),
                    list(block_dests),
                    dep,
                    arr,
                ))
            block_origins = set()
            block_dests = []

        for dest, origins in sorted(by_dest.items()):
            candidate_origins = block_origins | origins
            candidate_dest_count = len(block_dests) + 1
            candidate_elements = len(candidate_origins) * candidate_dest_count

            if block_dests and candidate_elements > GOOGLE_TRANSIT_MAX_ELEMENTS_PER_REQUEST:
                flush()
                candidate_origins = set(origins)

            # 單一 destination 的 origins 理論上最多就是居服員數；仍保守切塊。
            if len(candidate_origins) > GOOGLE_TRANSIT_MAX_ELEMENTS_PER_REQUEST:
                origin_list = sorted(candidate_origins)
                for start in range(0, len(origin_list), GOOGLE_TRANSIT_MAX_ELEMENTS_PER_REQUEST):
                    jobs.append((
                        origin_list[start:start + GOOGLE_TRANSIT_MAX_ELEMENTS_PER_REQUEST],
                        [dest],
                        dep,
                        arr,
                    ))
                continue

            block_origins = candidate_origins
            block_dests.append(dest)

        flush()

    return jobs


def prefetch_google_transit_pairs(required_pairs):
    """批次＋平行預抓 Google TRANSIT，結果寫入本次排班 memo。"""
    if not required_pairs:
        return

    state = _google_query_state()
    missing = []
    for lat1, lon1, lat2, lon2, departure, arrival in required_pairs:
        dep = _google_timestamp(departure) if departure is not None else None
        arr = _google_timestamp(arrival) if arrival is not None else None
        if not dep and not arr:
            dep = _google_timestamp(pd.Timestamp.now(tz='Asia/Taipei').floor('min'))
        key = _google_transit_key(lat1, lon1, lat2, lon2, dep, arr)
        if key not in state['memo']:
            missing.append((lat1, lon1, lat2, lon2, departure, arrival))

    if not missing:
        return

    jobs = _build_google_matrix_jobs(missing)
    total_elements = sum(len(origins) * len(destinations) for origins, destinations, _, _ in jobs)
    if state['count'] + total_elements > state['limit']:
        raise RuntimeError(
            f"本次排班預計需要新增 {total_elements} 個 Google Transit elements，"
            f"累計會超過安全上限 {state['limit']}。"
            "請縮小排班範圍或在交通進階設定提高安全上限。"
        )

    # 先保留額度，避免平行執行時競態。
    state['count'] += total_elements

    workers = min(GOOGLE_TRANSIT_MAX_WORKERS, max(1, len(jobs)))
    if workers == 1:
        for origins, destinations, dep, arr in jobs:
            state['memo'].update(
                _fetch_google_transit_matrix_raw(origins, destinations, dep, arr)
            )
        return

    errors = []
    with ThreadPoolExecutor(max_workers=workers) as pool:
        future_map = {
            pool.submit(
                _fetch_google_transit_matrix_raw,
                origins,
                destinations,
                dep,
                arr,
            ): (origins, destinations, dep, arr)
            for origins, destinations, dep, arr in jobs
        }
        for future in as_completed(future_map):
            try:
                state['memo'].update(future.result())
            except Exception as exc:
                errors.append(str(exc))

    if errors:
        raise RuntimeError('Google Transit 批次查詢失敗：' + errors[0])


def get_google_travel_time(lat1, lon1, lat2, lon2, transport_mode,
                           travel_min_per_km=3.0, departure_time=None, arrival_time=None):
    """機車固定 OSRM；大眾運輸固定 Google TRANSIT。"""
    if transport_mode not in ('機車', '大眾運輸'):
        raise ValueError('交通工具僅支援機車或公車／捷運，請修正員工資料。')
    if not all(np.isfinite(_number(v)) for v in (lat1, lon1, lat2, lon2)):
        raise ValueError('缺有效座標，無法估算交通')
    if lat1 == lat2 and lon1 == lon2:
        return 0.0

    if transport_mode == '機車':
        return get_osrm_travel_time(
            lat1, lon1, lat2, lon2,
            travel_min_per_km,
            transport_mode='機車',
        )

    if not _GOOGLE_ALLOWED.get():
        raise RuntimeError('大眾運輸固定使用 Google Routes TRANSIT，目前 Google 查詢未啟用。')

    departure = _google_timestamp(departure_time) if departure_time is not None else None
    arrival = _google_timestamp(arrival_time) if arrival_time is not None else None
    if departure and arrival:
        raise ValueError('交通查詢只能指定出發或抵達時間其中一個')
    if not departure and not arrival:
        departure = _google_timestamp(pd.Timestamp.now(tz='Asia/Taipei').floor('min'))

    key = _google_transit_key(
        lat1, lon1, lat2, lon2,
        departure=departure,
        arrival=arrival,
    )
    state = _google_query_state()
    if key not in state['memo']:
        # 少數未被 prefetch 命中的動態改派/插單，仍能單筆查詢。
        prefetch_google_transit_pairs([
            (lat1, lon1, lat2, lon2, departure_time, arrival_time)
        ])

    return state['memo'].get(key, float('inf'))


def _fetch_osrm_table_chunk(origins, destinations, travel_min_per_km, transport_mode='機車'):
    host, profile = _osrm_endpoint(transport_mode)
    coords = origins + destinations
    coord_str = ';'.join(f'{lon},{lat}' for lat,lon in coords)
    try:
        resp = requests.get(f'{host}/table/v1/{profile}/{coord_str}',
            params={'annotations':'duration','sources':';'.join(str(i) for i in range(len(origins))),
                    'destinations':';'.join(str(len(origins)+i) for i in range(len(destinations)))}, timeout=OSRM_TABLE_TIMEOUT_SECONDS)
        resp.raise_for_status(); data = resp.json()
        if data.get('code') != 'Ok':
            raise ValueError('批次查詢失敗：'+str(data.get('code')))
        durations = data['durations']
    except Exception as exc:
        raise RuntimeError(f'OSRM {transport_mode}路網尚未啟動或查詢失敗；請依 OSRM_機車啟動說明 建置服務。未改用汽車或直線。原因：{exc}') from exc
    for i,(lat1,lon1) in enumerate(origins):
        for j,(lat2,lon2) in enumerate(destinations):
            duration = durations[i][j]
            if duration is not None and np.isfinite(float(duration)) and float(duration) >= 0:
                _OSRM_TRAVEL_TIME_CACHE[_osrm_cache_key(lat1,lon1,lat2,lon2,transport_mode)] = float(duration)/60


def prefetch_osrm_travel_times(origins, destinations, travel_min_per_km=3.0, transport_mode='機車'):
    if _get_travel_provider() not in ('osrm','hybrid') or transport_mode != '機車':
        return
    origins = sorted({(_round_coord(lat),_round_coord(lon)) for lat,lon in origins if np.isfinite(_number(lat)) and np.isfinite(_number(lon))})
    destinations = sorted({(_round_coord(lat),_round_coord(lon)) for lat,lon in destinations if np.isfinite(_number(lat)) and np.isfinite(_number(lon))})
    # 兩邊都切塊，避免大量起點已超出 table 座標上限。
    chunk_size = OSRM_TABLE_MAX_COORDS // 2
    for i in range(0,len(origins),chunk_size):
        for j in range(0,len(destinations),chunk_size):
            left,right = origins[i:i+chunk_size],destinations[j:j+chunk_size]
            if any(_osrm_cache_key(*o,*d,transport_mode) not in _OSRM_TRAVEL_TIME_CACHE for o in left for d in right):
                _fetch_osrm_table_chunk(left,right,travel_min_per_km,transport_mode)

def calc_travel_minutes(
    lat1,
    lon1,
    lat2,
    lon2,
    config: "PipelineConfig",
    transport_mode="機車",
    departure_time=None,
    arrival_time=None,
) -> float:
    """
    兩點間轉場時間（分鐘，不含轉場緩衝）。

    依居服員「常用交通工具」分流：
    機車 -> 自架 OSRM motorcycle
    大眾運輸 -> Google Routes TRANSIT
    """

    return get_google_travel_time(
        lat1,
        lon1,
        lat2,
        lon2,
        transport_mode,
        config.travel_min_per_km,
        departure_time=departure_time, arrival_time=arrival_time,
    )

def get_service_intensity_weight(task) -> float:
    """依服務項目之體力耗費強度，回傳該任務對應的疲勞度加權係數。"""
    if task.get("需重度移位協助(0/1)") == 1:
        return SERVICE_INTENSITY_WEIGHT["重度移位"]
    if task.get("特殊照護需求") in ("餐食照顧/管灌", "管路安全與特殊日常照護"):
        return SERVICE_INTENSITY_WEIGHT["餐食管灌"]
    if pd.isna(task.get('需重度移位協助(0/1)')) and not _text(task.get('特殊照護需求')):
        return 1.0  # 中性計畫負荷權重，不代表已確認照護強度
    return SERVICE_INTENSITY_WEIGHT["一般照護"]

def parse_time(time_str):
    if pd.isna(time_str) or time_str == "無既定行程":
        return None, None
    parts = str(time_str).split("-")
    return (
        datetime.strptime(parts[0].strip(), "%H:%M"),
        datetime.strptime(parts[1].strip(), "%H:%M"),
    )

def _check_hard_constraints(
    task,
    cg,
    config: "PipelineConfig",
    travel_time_min: Optional[float] = None,
    is_preferred_caregiver: bool = False,
):
    """檢查單一 (任務, 居服員) 配對的硬性條件。全數通過回傳 None，否則回傳未通過原因。

    抽成獨立函式供 Phase 1 過濾與「原首選替換原因」診斷共用同一套判斷邏輯。

    星期對應／時間窗涵蓋／請假排除三項檢查，僅在 task／cg 實際具備對應欄位
    （可排班星期、每日可服務時段_起/迄、請假或不排班日期、星期、日期）時才生效
    ——以 `.get()` 取值，取不到（欄位不存在或為空）即直接跳過該檢查，因此舊版
    僅有「今日既定行程」欄位、不含這些欄位的資料錶行為完全不受影響。
    車程上限檢查同理，僅在呼叫端提供 travel_time_min 且 config.max_travel_minutes
    已設定（非 None）時才生效，預設為停用。
    """
    # 以下訊息刻意採「無主語」寫法（不寫「居服員」/「原居服員」），因為本函式同時
    # 供 Phase 1 一般候選人過濾（訊息僅供 reason_counts 內部除錯彙總）與
    # `_diagnose_caregiver_change` 診斷「原首選居服員」共用；後者會在回傳訊息前
    # 明確加上 `原首選居服員[ID]` 主語，若訊息本身也帶主語詞，會讓居督誤以為
    # 訊息在描述「獲派居服員」而非「原首選居服員」，見任務一問題分析。

    # Excel 的「指定居服員性別」是：男女不拘，但原本辨識：限男性限女性，所以修改如下
    if _text(task.get('任務資料異常')):
        return _text(task.get('任務資料異常'))
    if '可服務資料已確認' in cg and not _is_periodic_task(cg.get('可服務資料已確認')):
        return '可服務時段假設尚未採用或請假資料待修正'
    if not all(np.isfinite(_number(v)) for v in [task.get('服務地點_緯度'),task.get('服務地點_經度'),
                   cg.get('服務起點_緯度(家)'),cg.get('服務起點_經度(家)')]):
        return '個案或員工缺有效座標'
    for latitude,longitude in [(task.get('服務地點_緯度'),task.get('服務地點_經度')),
                               (cg.get('服務起點_緯度(家)'),cg.get('服務起點_經度(家)'))]:
        if not (20 <= _number(latitude) <= 27.5 and 118 <= _number(longitude) <= 123.5):
            return '座標不在預期臺灣範圍，請確認資料'
    if not np.isfinite(_number(cg.get('每日工時上限(小時)'))) or _number(cg.get('每日工時上限(小時)'))<=0:
        return '每日工時上限缺漏或無效'
    day = _date(task.get('日期'))
    expiry = _date(cg.get('長照服務人員到期日'))
    if day and expiry and day > expiry:
        return '長照服務人員證明已到期，需更新資格資料'
    leave = cg.get('請假紀錄')
    if day and isinstance(leave,(tuple,list)) and len(leave)==2:
        task_begin = pd.Timestamp(f"{day.isoformat()} {task['時間窗_開始']}")
        task_end = pd.Timestamp(f"{day.isoformat()} {task['時間窗_結束']}")
        if task_begin < pd.Timestamp(leave[1]) and task_end > pd.Timestamp(leave[0]):
            return '任務時段與請假區間重疊'
    needed = set(x for x in _text(task.get('服務碼必要證照')).split('、') if x)
    for code,units in iter_service_codes(task):
        if units>0 and code.upper() in DEFAULT_CODE_CERTS: needed.add(DEFAULT_CODE_CERTS[code.upper()])
    qualification_error = special_qualification_error(task,cg,needed)
    if qualification_error:return qualification_error
    req_gender = str(task.get("指定居服員性別", "")).strip()
    if req_gender in ("女", "限女性") and cg["性別"] != "女":
        return "案家指定女性居服員，性別不符"
    if req_gender in ("男", "限男性") and cg["性別"] != "男":
        return "案家指定男性居服員，性別不符"

    if task.get("需重度移位協助(0/1)") == 1 and cg.get("具備重度移位體力(0/1)") != 1:
        return "案家需重度移位協助，不具備相關體力條件"

    cg_excl = cg.get("特殊排斥條件", "")
    if cg_excl in EXCLUSION_MAP and task.get("案家環境特徵", "") == EXCLUSION_MAP[cg_excl]:
        return f"排斥「{EXCLUSION_MAP[cg_excl]}」環境條件"

    # 星期對應檢查 (Day-of-Week Matching)：僅當「星期」／「可排班星期」欄位確實存在於
    # 資料表中時才生效——用 `in task.index`／`in cg.index` 判斷「欄位是否存在」，而非
    # 用 `pd.notna()` 判斷「儲存格是否為空」。這兩者過去被混為一談：欄位整體不存在
    # （舊版資料表，理應跳過此檢查，維持向下相容）與欄位存在但「此列」資料缺漏或格式
    # 無法解析（新版資料表的資料品質問題）都會落入同一個 if 分支被直接跳過，導致
    # 可排班星期留空、或星期名稱格式不符的居服員被silently當作「全天候可排班」而通過
    # 硬性限制——這正是「居服員當日不在可排班星期名單內，系統卻仍將其派單」的根因。
    # 欄位存在時，資料缺漏或無法解析一律保守判定為不通過（fail-closed）並印出警告，
    # 而非靜默放行（fail-open）。
    has_weekday_cols = "星期" in task.index and "可排班星期" in cg.index
    if has_weekday_cols:
        task_weekday = task.get("星期")
        allowed_days_str = cg.get("可排班星期")
        cg_id_for_log = cg.get("居服員ID", "?")
        task_id_for_log = task.get("任務ID", "?")
        if pd.isna(task_weekday) or pd.isna(allowed_days_str) or str(allowed_days_str).strip() == "":
            print(
                f"[Hard Constraint Warning] 任務 {task_id_for_log} 或居服員 {cg_id_for_log} "
                f"的「星期」／「可排班星期」欄位資料缺漏，保守判定當日不可派單。"
            )
            return "星期或可排班星期資料缺漏，保守判定當日不可派單"

        task_wd_num = WEEKDAY_NAME_TO_NUM.get(str(task_weekday).strip())
        if task_wd_num is None:
            print(
                f"[Hard Constraint Warning] 任務 {task_id_for_log} 的「星期」欄位值"
                f"「{task_weekday}」無法辨識，保守判定當日不可派單。"
            )
            return "任務星期欄位格式無法辨識，保守判定當日不可派單"

        allowed_days = [int(d.strip()) for d in str(allowed_days_str).split(",") if d.strip().isdigit()]
        if not allowed_days:
            print(
                f"[Hard Constraint Warning] 居服員 {cg_id_for_log} 的「可排班星期」欄位值"
                f"「{allowed_days_str}」無法解析出任何星期，保守判定當日不可派單。"
            )
            return "可排班星期欄位格式無法解析，保守判定當日不可派單"

        if task_wd_num not in allowed_days:
            return "當日不排班（不在可排班星期名單）"

    # 時間窗涵蓋檢查 (Time Window Overlap)：僅當居服員有每日可服務時段欄位時生效
    cg_start_str = cg.get("每日可服務時段_起")
    cg_end_str = cg.get("每日可服務時段_迄")
    if pd.notna(cg_start_str) and pd.notna(cg_end_str):
        t_start = datetime.strptime(str(task["時間窗_開始"]).strip(), "%H:%M")
        t_end = datetime.strptime(str(task["時間窗_結束"]).strip(), "%H:%M")
        c_start = datetime.strptime(str(cg_start_str).strip(), "%H:%M")
        c_end = datetime.strptime(str(cg_end_str).strip(), "%H:%M")
        if not (t_start >= c_start and t_end <= c_end):
            return "任務時段超出每日可服務時段範圍"

    # 請假或不排班日期排除 (Leave Exclusion)：僅當任務有「日期」、居服員有請假日期欄位時生效
    leave_dates_str = cg.get("請假或不排班日期")
    task_date = task.get("日期")
    if pd.notna(leave_dates_str) and pd.notna(task_date) and str(leave_dates_str).strip() not in ("", "無"):
        leave_list = [d.strip() for d in str(leave_dates_str).split(",") if d.strip()]
        task_date_str = str(task_date).split(" ")[0]
        if task_date_str in leave_list:
            return "當日已有請假或不排班記錄"

    task_duration_hrs = task["服務歷時(分鐘)"] / 60.0
    if cg.get("今日已佔用工時(小時)", 0.0) + task_duration_hrs > cg["每日工時上限(小時)"]:
        return "今日工時已達每日上限"

    # 車程上限，惟「照護連續性」優先於「車程限制」：一般候選人受 max_travel_minutes
    # 限制（None = 不設限），但歷史首選居服員改採獨立的
    # preferred_caregiver_max_travel_minutes 上限——若該欄位亦為 None，代表歷史首選
    # 居服員完全不受車程上限約束，不會退回套用一般候選人的 max_travel_minutes，
    # 確保熟悉度高的居服員不因車程稍長就被硬性剔除。
    if travel_time_min is not None:
        cap = config.preferred_caregiver_max_travel_minutes if is_preferred_caregiver else config.max_travel_minutes
        if cap is not None and travel_time_min > cap:
            return f"預估車程({travel_time_min:.0f}分)超過上限({cap:.0f}分鐘)"

    required_certs = SPECIAL_CERT_REQUIREMENTS.get(task.get("特殊照護需求"))
    if required_certs and not any(caregiver_has_cert(cg,c) for c in required_certs):
        return "缺乏該需求類別之法定專長認證"

    return None

# 大眾運輸 Google 查詢優化：
# 先用「硬限制 + 直線距離」做便宜的候選縮小，再只對前 N 名呼叫 Google TRANSIT。
# 注意：直線距離只用於 shortlist，不作為真正交通時間。
TRANSIT_GOOGLE_CANDIDATE_TOP_K = 3

# ==========================================
# Phase 1: 適配度過濾與評分機制
# ==========================================
def run_phase1_matching(
    tasks: pd.DataFrame,
    df_cg: pd.DataFrame,
    config: PipelineConfig,
    previous_week_map: dict | None = None,
) -> pd.DataFrame:
    validate_dispatch_inputs(tasks, df_cg)
    previous_week_map = previous_week_map or {}
    match_results = []
    diagnostics = []

    # 進入逐筆比對迴圈前，先以單次 OSRM /table 批次查詢暖身快取
    # （居服員住家 x 任務地點），取代 N x M 次個別 HTTP 請求。
    # 只有未啟用 Google Routes 時，才預抓 OSRM 矩陣
    motorcycle_workers = df_cg[df_cg['常用交通工具'].astype(str).str.strip().eq('機車')]
    if not motorcycle_workers.empty:
        prefetch_osrm_travel_times(
            motorcycle_workers[['服務起點_緯度(家)','服務起點_經度(家)']].itertuples(index=False,name=None),
            tasks[['服務地點_緯度','服務地點_經度']].itertuples(index=False,name=None),
            config.travel_min_per_km, transport_mode='機車')

    caregiver_rows = list(df_cg.iterrows())
    transit_shortlists = {}
    transit_pairs = []
    for _, pre_task in tasks.iterrows():
        pre_tid = pre_task['任務ID']
        pre_client_lat = pre_task['服務地點_緯度']
        pre_client_lon = pre_task['服務地點_經度']
        pre_pref_cg = _text(pre_task.get('歷史首選居服員ID'))
        candidates = []
        for _, pre_cg in caregiver_rows:
            if str(pre_cg.get('常用交通工具', '機車')).strip() != '大眾運輸':
                continue
            pre_cg_id = pre_cg['居服員ID']
            pre_is_preferred = bool(pre_pref_cg) and pre_cg_id == pre_pref_cg
            if _check_hard_constraints(pre_task, pre_cg, config, None, pre_is_preferred):
                continue
            dist = calc_distance_km(
                pre_cg['服務起點_緯度(家)'], pre_cg['服務起點_經度(家)'],
                pre_client_lat, pre_client_lon,
            )
            if np.isfinite(dist):
                candidates.append((float(dist), pre_cg_id, pre_cg))
        candidates.sort(key=lambda x: x[0])
        chosen = candidates[:TRANSIT_GOOGLE_CANDIDATE_TOP_K]
        transit_shortlists[pre_tid] = {cg_id for _, cg_id, _ in chosen}
        arrival = _bucket_transit_datetime(_task_arrival(pre_task))
        for _, _, pre_cg in chosen:
            transit_pairs.append((
                pre_cg['服務起點_緯度(家)'], pre_cg['服務起點_經度(家)'],
                pre_client_lat, pre_client_lon,
                None, arrival,
            ))

    prefetch_google_transit_pairs(transit_pairs)

    for _, task in tasks.iterrows():
        t_id = task["任務ID"]
        c_id = task["案家ID"]
        client_lat = task["服務地點_緯度"]
        client_lon = task["服務地點_經度"]
        pref_cg = _text(task.get("歷史首選居服員ID"))
        previous_week_cg = _text(previous_week_map.get(c_id))
        req_type = _text(task.get("特殊照護需求"))

        matched_count_for_task = 0
        reason_counts: dict = {}

        transit_shortlist_ids = transit_shortlists.get(t_id, set())

        for _, cg in caregiver_rows:
            cg_id = cg["居服員ID"]
            is_preferred = bool(pref_cg) and cg_id == pref_cg
            is_previous_week = bool(previous_week_cg) and str(cg_id) == str(previous_week_cg)

            transport_mode = str(
                cg.get("常用交通工具", "機車")
            ).strip()

            if transport_mode not in ("機車", "大眾運輸"):
                raise ValueError("交通工具僅支援機車或公車／捷運，請修正員工資料。")

            early_reason = _check_hard_constraints(task, cg, config, None, is_preferred)
            if early_reason:
                reason_counts[early_reason] = reason_counts.get(early_reason, 0) + 1
                continue

            if transport_mode == "大眾運輸" and cg_id not in transit_shortlist_ids:
                reason_counts["大眾運輸候選距離預篩未進前3"] = (
                    reason_counts.get("大眾運輸候選距離預篩未進前3", 0) + 1
                )
                continue

            travel_time_min = calc_travel_minutes(
                cg["服務起點_緯度(家)"],
                cg["服務起點_經度(家)"],
                client_lat,
                client_lon,
                config,
                transport_mode=transport_mode,
                arrival_time=(
                    _bucket_transit_datetime(_task_arrival(task))
                    if transport_mode == "大眾運輸"
                    else _task_arrival(task)
                ),
            )

            if transport_mode == '大眾運輸' and not np.isfinite(travel_time_min):
                reason_counts['Google Transit 無可用路線／班次'] = reason_counts.get('Google Transit 無可用路線／班次', 0) + 1
                continue

            # --- Hard Constraints (硬性過濾，不合格者直接剔除) ---
            err_msg = _check_hard_constraints(task, cg, config, travel_time_min, is_preferred)
            if err_msg is not None:
                reason_counts[err_msg] = reason_counts.get(err_msg, 0) + 1
                continue

            matched_count_for_task += 1
            cert = cg["核心專長證照"]

            # --- Soft Match Scoring ---
            # 0. 基礎分
            base_score = config.base_score

            # 1. 專長匹配
            skill_bonus = 0.0

            if req_type == "失智引導與精神陪伴" and caregiver_has_cert(cg,"失智症照顧專長"):
                skill_bonus = config.cert_bonus_dementia

            elif (
                req_type in ["管路安全與特殊日常照護", "餐食照顧/管灌"]
                and caregiver_has_cert(cg, "單一級照服證照")
            ):
                skill_bonus = config.cert_bonus_other

            skill_bonus += certificate_bonus(task,cg)

            # 2. 照護連續性
            continuity_bonus = 0.0
            continuity_performance_bonus = 0.0
            previous_week_bonus = 0.0

            if is_previous_week:
                previous_week_bonus = config.previous_week_caregiver_bonus

            if is_preferred:
                continuity_bonus = config.preferred_caregiver_bonus

                if (
                    cg["歷史滿意度均值"]
                    >= config.continuity_satisfaction_threshold
                ):
                    continuity_performance_bonus = (
                        config.continuity_performance_bonus
                    )

            # 3. 歷史服務品質
            satisfaction = _number(cg.get('歷史滿意度均值'))
            satisfaction_adjustment = ((satisfaction - config.satisfaction_baseline) * config.satisfaction_weight
                                       if np.isfinite(satisfaction) else 0.0)

            # 4. 交通成本
            travel_penalty = min(
                travel_time_min * config.travel_penalty_weight,
                config.travel_penalty_cap
            )

            # 5. 疲勞 / 工作負荷
            intensity_weight = get_service_intensity_weight(task)
            actual_hours = _number(cg.get('當月累計服務時數(疲勞度)'))
            actual_penalty = (actual_hours * intensity_weight / config.fatigue_reference_hours * config.fatigue_weight
                              if np.isfinite(actual_hours) else 0.0)
            planned_hours = _number(cg.get('本次計畫已排時數'),0.0)
            weighted_hours = _number(cg.get('本次計畫加權負荷時數'),planned_hours)
            plan_load = weighted_hours if config.plan_load_mode=='weighted' else planned_hours
            plan_penalty = plan_load / config.fatigue_reference_hours * config.fatigue_weight
            fatigue_penalty = actual_penalty + plan_penalty

            # ==========================================
            # 最終適配度分數
            # ==========================================
            score = (
                base_score
                + skill_bonus
                + continuity_bonus
                + previous_week_bonus
                + continuity_performance_bonus
                + satisfaction_adjustment
                - travel_penalty
                - fatigue_penalty
            )

            # ==========================================
            # 保存結果
            # ==========================================
            match_results.append(
                {
                    "任務ID": t_id,
                    "案家ID": c_id,
                    "居服員ID": cg_id,

                    "適配度分數": round(max(score, 0), 2),

                    # Explainable AI：分數組成
                    "基礎分": round(base_score, 2),
                    "專長匹配加分": round(skill_bonus, 2),
                    "歷史首選加分": round(continuity_bonus, 2),
                    "上一週同案加分": round(previous_week_bonus, 2),
                    "連續性品質加分": round(
                        continuity_performance_bonus,
                        2
                    ),
                    "滿意度調整": round(
                        satisfaction_adjustment,
                        2
                    ),
                    "交通扣分": round(
                        travel_penalty,
                        2
                    ),
                    "工作負荷扣分": round(
                        fatigue_penalty,
                        2
                    ),

                    # Explainability 輔助資訊
                    "是否歷史首選": bool(is_preferred),
                    "是否上一週同案居服員": bool(is_previous_week),
                    "上一週居服員ID": previous_week_cg,

                    "當月累計服務時數": round(
                        float(
                            actual_hours
                        ),
                        1
                    ),

                    "疲勞資料狀態": '已提供累計工時' if np.isfinite(actual_hours) else '以本次已排工作量分配；未提供實際出勤工時',
                    "滿意度資料狀態": '已提供' if np.isfinite(satisfaction) else '未提供，不參與評分',
                    "本次計畫已排時數": round(planned_hours,2),
                    "本次計畫加權負荷時數": round(weighted_hours,2),
                    **{k:task.get(k) for k in ["照護難度參考係數","照護難度判斷依據","照護提示","服務碼必要證照"]},
                    "本次計畫負荷扣分": round(plan_penalty,2),
                    "實際工時負荷扣分": round(actual_penalty,2) if np.isfinite(actual_hours) else np.nan,
                    "資料來源": task.get('資料來源','原格式'),
                    "待確認事項": task.get('任務待確認事項',''),
                    "交通估算來源": _travel_source(transport_mode),
                    "失智症資料狀態": "已確認失智" if dementia_status(task) is True else "已確認非失智" if dementia_status(task) is False else "失智狀態未提供，需機構確認",
                    "特殊資格檢查": "依員工資格文字符合本次服務要求",
                    "服務強度係數": round(
                        float(intensity_weight),
                        2
                    ),

                    # 原本欄位
                    "預估交通時間(分)": round(
                        travel_time_min,
                        1
                    ),

                    "任務開始時間":
                        task["時間窗_開始"],

                    "任務結束時間":
                        task["時間窗_結束"],

                    "優先級":
                        task["任務優先級"],

                    "地點緯度":
                        client_lat,

                    "地點經度":
                        client_lon,

                    "具備失智症20小時認證(0/1)":
                        int(
                            caregiver_has_cert(cg, "失智症照顧專長")
                        ),

                    "具備精神疾病20小時認證(0/1)":
                        int(
                            caregiver_has_cert(cg, "精神疾病照顧專長")
                        ),

                    "常用交通工具": transport_mode,
                }
            )
        diagnostics.append({'任務ID':t_id,'候選數':matched_count_for_task,'剔除原因':str(reason_counts)})
    frame = pd.DataFrame(match_results)
    frame.attrs['matching_diagnostics'] = diagnostics
    return frame

def _diagnose_caregiver_change(
    task_row,
    pref_cg_id,
    assigned_cg_id,
    df_cg: pd.DataFrame,
    df_matches: pd.DataFrame,
    df_valid: pd.DataFrame,
    other_assigned_task_ids,
    task_times: dict,
    config: "PipelineConfig",
) -> str:
    """回傳「原首選替換原因」文字，說明歷史首選居服員為何未獲派本次任務。"""
    if pd.isna(pref_cg_id) or str(pref_cg_id).strip() == "":
        return ""
    if pref_cg_id == assigned_cg_id:
        return ""

    t_id = task_row["任務ID"]
    subject = f"原首選居服員[{pref_cg_id}]"

    pref_in_matches = (
        (df_matches["任務ID"] == t_id)
        & (df_matches["居服員ID"] == pref_cg_id)
    ).any()

    if not pref_in_matches:
        cg_rows = df_cg[df_cg["居服員ID"] == pref_cg_id]
        if cg_rows.empty:
            return f"{subject}資料異動，系統查無此居服員"
        cg_row = cg_rows.iloc[0]

        transport_mode = str(cg_row.get("常用交通工具", "機車")).strip()
        if transport_mode not in ("機車", "大眾運輸"):

            raise ValueError("交通工具僅支援機車或公車／捷運，請修正員工資料。")

        travel_time_min = calc_travel_minutes(
            cg_row["服務起點_緯度(家)"],
            cg_row["服務起點_經度(家)"],
            task_row["服務地點_緯度"],
            task_row["服務地點_經度"],
            config,
            transport_mode=transport_mode,
        )
        detail = (
            _check_hard_constraints(
                task_row,
                cg_row,
                config,
                travel_time_min,
                is_preferred_caregiver=True,
            )
            or "不符合硬性派單條件"
        )
        return f"{subject}{detail}"

    pref_in_valid = (
        (df_valid["任務ID"] == t_id)
        & (df_valid["居服員ID"] == pref_cg_id)
    ).any()
    if not pref_in_valid:
        return f"{subject}今日既定行程與本任務時段衝突"

    cg_rows = df_cg[df_cg["居服員ID"] == pref_cg_id]
    transport_mode = "機車"
    if not cg_rows.empty:
        transport_mode = str(cg_rows.iloc[0].get("常用交通工具", "機車")).strip()
    if transport_mode not in ("機車", "大眾運輸"):

        raise ValueError("交通工具僅支援機車或公車／捷運，請修正員工資料。")

    t_start, t_end, t_lat, t_lon = task_times[t_id]
    for other_t_id in other_assigned_task_ids:
        o_start, o_end, o_lat, o_lon = task_times[other_t_id]
        travel_mins = (
            calc_travel_minutes(
                t_lat,
                t_lon,
                o_lat,
                o_lon,
                config,
                transport_mode=transport_mode,
            )
            + config.buffer_mins
        )
        if not (
            t_end + timedelta(minutes=travel_mins) <= o_start
            or o_end + timedelta(minutes=travel_mins) <= t_start
        ):
            return f"{subject}該時段已媒合其他案家任務"

    return f"{subject}雖符合派單資格，惟系統整體最佳化後綜合適配分數較低，已改派其他居服員"

def _evaluate_reassignment(
    task_id,
    new_cg_id,
    df_tasks: pd.DataFrame,
    df_result_effective: pd.DataFrame,
    df_cg: pd.DataFrame,
    config: "PipelineConfig",
    task_locations: Optional[dict] = None,
    date_column: str = "日期",
    df_cl: Optional[pd.DataFrame] = None,
) -> dict:
    """評估把 task_id 改派給 new_cg_id 是否可行，回傳結構化結果供兩種呼叫端共用：

    - `check_reassignment_conflict`：改派當下的最終權威判定（僅需 available/detail）。
    - `rank_candidates_by_availability`：改派下拉選單的候選人清單排序與標籤
      （需要衝突任務的時段細節，才能顯示如「14:00-15:00 服務中」的具體標籤）。

    單一事實來源：兩種呼叫端看到的「是否可派」結論恆一致，不會有選單顯示可派、
    實際確認卻被拒絕的落差。

    回傳 dict 固定包含 "cg_id"、"available"、"reason"
    （None｜"task_not_found"｜"caregiver_not_found"｜"bad_time_format"｜
    "hard_constraint"｜"time_conflict"｜"hour_cap"）、"detail"（完整說明文字）；
    reason 為 "time_conflict" 時另附 "conflict_task_id"／"conflict_start"／"conflict_end"。

    `df_cl`（Client_Profiles）為選填：提供時會額外以 Phase 1 同一套
    `_check_hard_constraints`（性別限定、重度移位體力、環境排斥、可排班星期、
    每日可服務時段、請假日期、專長認證）評估 new_cg_id 是否符合本任務的硬性
    資格條件，失敗回傳 reason="hard_constraint" 並附上具體原因（例如「當日已有
    請假或不排班記錄」「缺乏該需求類別之法定專長認證」）——這隻影響下拉選單的
    排序與標籤（見 rank_candidates_by_availability），刻意不接在
    check_reassignment_conflict 的權威判定路徑上：本系統沒有「強制派單權限」機制，
    居督仍可能因臨時狀況刻意指派不符合建議條件的居服員，故只標記不隱藏、不封鎖。
    不提供 df_cl（預設 None）時完全跳過此檢查，行為與加入前相同。
    """
    result = {
        "cg_id": new_cg_id,
        "available": False,
        "reason": None,
        "detail": "",
        "conflict_task_id": None,
        "conflict_start": None,
        "conflict_end": None,
    }

    task_rows = df_tasks[df_tasks["任務ID"] == task_id]
    if task_rows.empty:
        result.update(reason="task_not_found", detail="找不到該任務資料，無法檢查衝突")
        return result
    task_row = task_rows.iloc[0]

    cg_rows = df_cg[df_cg["居服員ID"].astype(str) == str(new_cg_id)]
    if cg_rows.empty:
        result.update(reason="caregiver_not_found", detail=f"找不到居服員 {new_cg_id} 的資料")
        return result
    cg_row = cg_rows.iloc[0]

    # 改派時依該居服員的常用交通工具計算交通時間
    transport_mode = str(cg_row.get("常用交通工具", "機車")).strip()
    if transport_mode not in ("機車", "大眾運輸"):

        raise ValueError("交通工具僅支援機車或公車／捷運，請修正員工資料。")

    task_locations = task_locations or {}

    merged_task = task_row.copy()
    if df_cl is not None and not df_cl.empty and '案家ID' in task_row:
        cl_rows=df_cl[df_cl['案家ID'].astype(str)==str(task_row['案家ID'])]
        if not cl_rows.empty:
            merged_task=pd.Series({**cl_rows.iloc[0].to_dict(),**task_row.to_dict()})
    loc=task_locations.get(task_id)
    if loc:
        if pd.isna(merged_task.get('服務地點_緯度')):merged_task['服務地點_緯度']=loc[0]
        if pd.isna(merged_task.get('服務地點_經度')):merged_task['服務地點_經度']=loc[1]
    hard_reason=_check_hard_constraints(merged_task,cg_row,config,None,
                         _text(merged_task.get('歷史首選居服員ID'))==str(new_cg_id))
    if hard_reason is not None:
        result.update(reason='hard_constraint',detail=hard_reason)
        return result

    t_start, t_end = parse_time(f"{task_row['時間窗_開始']}-{task_row['時間窗_結束']}")
    if t_start is None:
        result.update(reason="bad_time_format", detail="任務時間格式無法解析，無法檢查衝突")
        return result

    if date_column in df_tasks.columns and pd.notna(task_row.get(date_column)):
        task_date = task_row[date_column]
        same_day_ids = set(df_tasks.loc[df_tasks[date_column] == task_date, "任務ID"])
    else:
        same_day_ids = set(df_tasks["任務ID"])  # 單日排程：所有任務視為同一天

    other_assigned = (
        df_result_effective[
            (df_result_effective["派單居服員"].astype(str) == str(new_cg_id))
            & (df_result_effective["任務ID"].isin(same_day_ids))
            & (df_result_effective["任務ID"] != task_id)
        ]
        if not df_result_effective.empty
        else df_result_effective
    )

    total_minutes = float(task_row["服務歷時(分鐘)"])

    if other_assigned is not None:
        for other_t_id in other_assigned["任務ID"]:
            other_rows = df_tasks[df_tasks["任務ID"] == other_t_id]
            if other_rows.empty:
                continue
            other_row = other_rows.iloc[0]
            o_start, o_end = parse_time(f"{other_row['時間窗_開始']}-{other_row['時間窗_結束']}")
            if o_start is None:
                continue

            travel_mins = config.buffer_mins
            t_loc = task_locations.get(task_id)
            o_loc = task_locations.get(other_t_id)
            if t_loc and o_loc and all(pd.notna(v) for v in (*t_loc, *o_loc)):
                travel_mins = _interval_travel((t_start,t_end,*t_loc),(o_start,o_end,*o_loc),config,transport_mode,task_row)

            if not np.isfinite(travel_mins) or not (
                t_end + timedelta(minutes=travel_mins) <= o_start
                or o_end + timedelta(minutes=travel_mins) <= t_start
            ):
                result.update(
                    reason="time_conflict",
                    conflict_task_id=other_t_id,
                    conflict_start=str(other_row["時間窗_開始"]),
                    conflict_end=str(other_row["時間窗_結束"]),
                    detail=(
                        f"與居服員 {new_cg_id} 當日另一任務（{other_t_id}，"
                        f"{other_row['時間窗_開始']}-{other_row['時間窗_結束']}）時間衝突"
                        + (f"（含轉場緩衝約 {travel_mins:.0f} 分鐘）" if np.isfinite(travel_mins) else "（服務時段重疊）")
                    ),
                )
                return result
            total_minutes += float(other_row["服務歷時(分鐘)"])

    daily_cap = cg_row.get("每日工時上限(小時)")
    if pd.notna(daily_cap) and total_minutes / 60.0 > daily_cap:
        result.update(
            reason="hour_cap",
            detail=(
                f"居服員 {new_cg_id} 改派後當日總工時將達 {total_minutes / 60.0:.1f} 小時，"
                f"超過每日上限 {daily_cap:.1f} 小時"
            ),
        )
        return result

    result["available"] = True
    return result

def check_reassignment_conflict(
    task_id,
    new_cg_id,
    df_tasks: pd.DataFrame,
    df_result_effective: pd.DataFrame,
    df_cg: pd.DataFrame,
    config: "PipelineConfig",
    task_locations: Optional[dict] = None,
    date_column: str = "日期",
    df_cl: Optional[pd.DataFrame] = None,
) -> Optional[str]:
    """檢查「快速改派」把 task_id 轉給 new_cg_id 是否會造成時間衝突或工時超標。

    供月曆視角一鍵調班（calendar_view.py）使用：`df_result_effective` 須為已套用
    居督覆寫後的『目前生效』派單結果（見 apply_overrides_to_result），確保衝突
    檢查基準與畫面顯示一致，不會用「AI 原始建議」誤判已被覆寫過的任務。

    衝突判定與 Phase 2 限制條件 2（同一居服員新任務時間不得重疊，須預留
    config.buffer_mins 轉場緩衝）採同一公式，僅多檢查每日工時上限。
    `task_locations`（任務ID -> (緯度, 經度)，通常取自 df_matches）用於估算轉場
    車程；缺少座標時保守僅以 config.buffer_mins 判斷重疊，不會略過檢查。

    回傳 None 表示可安全改派；否則回傳供 UI 顯示的錯誤說明文字。
    """
    result = _evaluate_reassignment(
        task_id, new_cg_id, df_tasks, df_result_effective, df_cg, config, task_locations, date_column, df_cl
    )
    return None if result["available"] else result["detail"]

def rank_candidates_by_availability(
    task_id,
    candidate_cg_ids,
    df_tasks: pd.DataFrame,
    df_result_effective: pd.DataFrame,
    df_cg: pd.DataFrame,
    config: "PipelineConfig",
    task_locations: Optional[dict] = None,
    date_column: str = "日期",
    df_cl: Optional[pd.DataFrame] = None,
) -> list:
    """供月曆快速改派下拉選單使用：把候選居服員依「該時段是否有空檔」排序。

    對 candidate_cg_ids 逐一呼叫 `_evaluate_reassignment`（與確認改派時
    check_reassignment_conflict 走同一套判定，避免選單顯示可派、實際確認卻被
    拒絕的落差），回傳依 available 由高到低排序（同組內維持 candidate_cg_ids
    原始順序）的 dict list，每筆結構同 `_evaluate_reassignment` 的回傳值。

    candidate_cg_ids 預期為「機構內全體居服員」（呼叫端不應預先用 Phase 1
    的硬性條件篩過一輪才傳進來，否則不符合資格的人會直接從清單消失，而非
    保留＋標記）；提供 `df_cl` 時，本函式會連同 Phase 1 的硬性資格條件
    （性別限定、重度移位體力、環境排斥、可排班星期、可服務時段、請假日期、
    專長認證）一併標記為 reason="hard_constraint"，而非讓呼叫端事先濾掉。
    刻意不將此檢查接上 check_reassignment_conflict（該函式呼叫
    _evaluate_reassignment 時不傳 df_cl）：本系統無強制派單權限機制，居督仍可
    在清楚看到標記後選擇覆寫，故硬性資格條件在此僅供標籤顯示，不封鎖改派。
    """
    evaluated = [
        _evaluate_reassignment(
            task_id, cg_id, df_tasks, df_result_effective, df_cg, config,
            task_locations, date_column, df_cl,
        )
        for cg_id in candidate_cg_ids
    ]
    return sorted(evaluated, key=lambda r: not r["available"])

def _build_cg_busy_blocks(df_cg: pd.DataFrame) -> dict:
    """建立居服員今日既定行程時間阻擋塊 (Time Blocks)，回傳 cg_id -> [(start, end, lat, lon), ...]。

    「今日既定行程」欄位為新版月批次排班資料表所無（改以可排班星期/請假日期取代），
    以 `.get()` 取值使兩種資料表格式皆可安全運作：欄位不存在時回傳 None，
    parse_time 會將 None 視為「無既定行程」而正確跳過。
    """
    cg_busy = {}
    for _, cg in df_cg.iterrows():
        cg_id = cg["居服員ID"]
        busy_intervals = []
        for i in [1, 2]:
            t1, t2 = parse_time(cg.get(f"今日既定行程{i}_時段"))
            if t1:
                busy_intervals.append(
                    (
                        t1,
                        t2,
                        cg.get(f"今日既定行程{i}_地點緯度"),
                        cg.get(f"今日既定行程{i}_地點經度"),
                    )
                )
        cg_busy[cg_id] = busy_intervals
    return cg_busy

def _build_break_constraints(solver, X: dict, df_valid: pd.DataFrame, cg_busy: dict, task_times: dict, config: "PipelineConfig") -> None:
    """規則1（勞動合規）：居服員累計連續工作達 config.continuous_work_limit_mins 分鐘，
    強制要求下一段任務與前段之間至少間隔 config.mandatory_break_mins 分鐘。

    做法：對每位居服員，把「今日既定行程」（固定發生）與「通過衝突檢查的候選新任務」
    （指派變數）依開始時間排序，切出「潛在連續鏈」——鏈內相鄰兩區塊的固定時間差
    < mandatory_break_mins。對鏈中每一個起點，往後累加工作分鐘數直到達到門檻，
    即對緊接在後的區塊加入限制式，禁止「該起點到達門檻的前綴」與「緊接的下一個
    候選任務」同時獲派（前綴若全為既有既定行程，則後續候選任務直接被禁止指派）。

    此為保守近似：以「潛在鏈上的固定時間」計算，而非僅以「實際獲派子集合」重新
    計算，在極少數「跳過鏈中某段候選任務即可讓實際連續工時縮短」的邊界情況下可能
    偏嚴（阻擋一個其實合規的組合），但休息規則屬勞動合規要求，寧可偏保守也不可
    漏判真違規；居督仍可透過既有人工覆寫機制（save_override_log）調整結果。
    """
    for cg_id in df_valid["居服員ID"].unique():
        blocks = []
        for b_start, b_end, _b_lat, _b_lon in cg_busy.get(cg_id, []):
            blocks.append(
                {"start": b_start, "end": b_end, "duration_min": (b_end - b_start).total_seconds() / 60.0, "var": None}
            )
        for t_id in df_valid[df_valid["居服員ID"] == cg_id]["任務ID"].unique():
            t_start, t_end, _, _ = task_times[t_id]
            blocks.append(
                {
                    "start": t_start,
                    "end": t_end,
                    "duration_min": (t_end - t_start).total_seconds() / 60.0,
                    "var": X[(t_id, cg_id)],
                }
            )
        blocks.sort(key=lambda b: b["start"])

        # 依固定時間切出潛在連續鏈：鏈內相鄰區塊時間差 < mandatory_break_mins
        i = 0
        n = len(blocks)
        while i < n:
            j = i + 1
            while j < n:
                gap_mins = (blocks[j]["start"] - blocks[j - 1]["end"]).total_seconds() / 60.0
                if gap_mins >= config.mandatory_break_mins:
                    break
                j += 1
            chain = blocks[i:j]

            # 對鏈中每個起點 k，找出往後累加達門檻的最短前綴 [k..m]，並限制其後一個
            # 候選任務不得與該前綴同時獲派。
            for k in range(len(chain)):
                cum = 0.0
                m = None
                for idx in range(k, len(chain)):
                    cum += chain[idx]["duration_min"]
                    if cum >= config.continuous_work_limit_mins:
                        m = idx
                        break
                if m is None or m + 1 >= len(chain):
                    continue

                extra = chain[m + 1]
                if extra["var"] is None:
                    continue  # 既有既定行程本身即固定發生，無法以指派變數禁止

                prefix_vars = [b["var"] for b in chain[k : m + 1] if b["var"] is not None]
                if not prefix_vars:
                    solver.Add(extra["var"] == 0)
                else:
                    solver.Add(sum(prefix_vars) + extra["var"] <= len(prefix_vars))

            i = j

# ==========================================
# 動態插單 MVP：臨時新增案件直接插入既有班表
# ==========================================
def find_insertion_candidates(
    new_task,
    current_tasks: pd.DataFrame,
    current_result: pd.DataFrame,
    df_cg: pd.DataFrame,
    config: "PipelineConfig",
    top_n: int = 3,
    date_column: str = "日期",
) -> pd.DataFrame:
    """找出「不移動既有班表」即可直接插入臨時任務的候選居服員。

    第一版 MVP 採 minimal-change insertion：
    1. 先沿用 Phase 1 的硬性條件與適配度評分。
    2. 對每位候選人只檢查新任務前一案 / 下一案的時間與交通銜接。
    3. 既有班表完全不移動；若需要挪班或交換任務，標記為不可直接插入。
    4. 再檢查插單後每日總工時是否超過居服員上限。

    Parameters
    ----------
    new_task:
        dict / pd.Series，需包含一般任務欄位及案家欄位（例如服務地點經緯度、
        指定居服員性別、特殊照護需求、服務歷時等）。
    current_tasks:
        目前排班任務資料，建議已與 Client_Profiles 合併，至少包含任務ID、時間窗、
        服務歷時與服務地點座標。
    current_result:
        目前生效派單結果（任務ID、派單居服員）。可先套用居督覆寫後再傳入。
    """
    if isinstance(new_task, dict):
        task_series = pd.Series(new_task)
    else:
        task_series = new_task.copy()

    required = [
        "任務ID", "案家ID", "時間窗_開始", "時間窗_結束",
        "服務歷時(分鐘)", "服務地點_緯度", "服務地點_經度",
    ]
    missing = [c for c in required if c not in task_series.index or pd.isna(task_series.get(c))]
    if missing:
        raise ValueError("臨時任務缺少必要欄位：" + "、".join(missing))

    def _same_date(v1, v2):
        if pd.isna(v1) or pd.isna(v2):
            return True
        try:
            return pd.to_datetime(v1).date() == pd.to_datetime(v2).date()
        except Exception:
            return str(v1).split(" ")[0] == str(v2).split(" ")[0]

    def _parse_hhmm(value):
        return datetime.strptime(str(value).strip(), "%H:%M")

    new_start = _parse_hhmm(task_series["時間窗_開始"])
    new_end = _parse_hhmm(task_series["時間窗_結束"])
    if new_end <= new_start:
        raise ValueError("臨時任務結束時間必須晚於開始時間")

    # 動態插單的「工時」採實際排班佔用時間，較符合督導直覺：
    # 08:00-11:00 = 180 分鐘 = 3 小時。
    # Service Code Master 的「服務歷時(分鐘)」仍保留，供服務碼／申報邏輯使用，
    # 不再拿來顯示動態插單後的實際排班工時。
    new_schedule_minutes = (new_end - new_start).total_seconds() / 60.0

    # Phase 1 的每日工時硬限制也應使用實際排班佔用時間。
    one_task = pd.DataFrame([task_series.to_dict()])
    one_task.loc[:, "服務歷時(分鐘)"] = new_schedule_minutes
    phase1 = run_phase1_matching(one_task, df_cg, config)

    output_columns = [
        "排名", "居服員ID", "可直接插入", "適配度分數",
        "待接任務時段", "重疊任務與時段",
        "前一任務（結束）", "下一任務（開始）",
        "前段交通時間(分)", "後段交通時間(分)", "插單後排班工時(小時)",
        "擾動成本", "推薦原因", "不可插入原因",
    ]
    if phase1.empty:
        return pd.DataFrame(columns=output_columns)

    task_lookup = {}
    if current_tasks is not None and not current_tasks.empty:
        for _, r in current_tasks.iterrows():
            task_lookup[r["任務ID"]] = r

    result_rows = []
    for _, match in phase1.iterrows():
        cg_id = match["居服員ID"]
        cg_rows = df_cg[df_cg["居服員ID"] == cg_id]
        if cg_rows.empty:
            continue
        cg = cg_rows.iloc[0]

        transport_mode = str(cg.get("常用交通工具", "機車")).strip()
        if transport_mode not in ("機車", "大眾運輸"):

            raise ValueError("交通工具僅支援機車或公車／捷運，請修正員工資料。")

        assigned = []
        if current_result is not None and not current_result.empty:
            cg_result = current_result[current_result["派單居服員"] == cg_id]
            for _, assigned_row in cg_result.iterrows():
                tid = assigned_row["任務ID"]
                trow = task_lookup.get(tid)
                if trow is None:
                    continue
                if date_column in task_series.index and date_column in trow.index:
                    if not _same_date(task_series.get(date_column), trow.get(date_column)):
                        continue
                try:
                    start = _parse_hhmm(trow["時間窗_開始"])
                    end = _parse_hhmm(trow["時間窗_結束"])
                except Exception:
                    continue
                assigned.append(
                    {
                        "task_id": tid,
                        "start": start,
                        "end": end,
                        "lat": trow.get("服務地點_緯度"),
                        "lon": trow.get("服務地點_經度"),
                        # 既有任務的排班工時直接取開始～結束時間，
                        # 避免 Service Code 分鐘數與實際排班時段不同時造成誤解。
                        "duration": (end - start).total_seconds() / 60.0,
                    }
                )

        assigned.sort(key=lambda x: x["start"])

        overlap = [
            a for a in assigned
            if not (a["end"] <= new_start or new_end <= a["start"])
        ]

        prev_tasks = [a for a in assigned if a["end"] <= new_start]
        next_tasks = [a for a in assigned if a["start"] >= new_end]
        prev_task = max(prev_tasks, key=lambda x: x["end"]) if prev_tasks else None
        next_task = min(next_tasks, key=lambda x: x["start"]) if next_tasks else None

        incoming = None
        outgoing = None
        reason = ""

        overlap_details = []
        for conflict in overlap:
            overlap_start = max(new_start, conflict['start'])
            overlap_end = min(new_end, conflict['end'])
            overlap_mins = (overlap_end - overlap_start).total_seconds() / 60
            overlap_details.append(
                f"{conflict['task_id']}（{conflict['start']:%H:%M}–{conflict['end']:%H:%M}），"
                f"重疊 {overlap_start:%H:%M}–{overlap_end:%H:%M}，共 {overlap_mins:g} 分鐘"
            )
        if overlap:
            reason = f"待接任務 {new_start:%H:%M}–{new_end:%H:%M} 與既有任務重疊：" + '；'.join(overlap_details)
        else:
            # 前一案 -> 新案
            if prev_task is not None:
                if pd.notna(prev_task["lat"]) and pd.notna(prev_task["lon"]):
                    incoming = calc_travel_minutes(
                        prev_task["lat"], prev_task["lon"],
                        task_series["服務地點_緯度"], task_series["服務地點_經度"],
                        config, transport_mode=transport_mode, departure_time=_task_clock(task_series,prev_task['end'] + timedelta(minutes=config.buffer_mins)),
                    )
                    required_gap = incoming + config.buffer_mins
                else:
                    required_gap = config.buffer_mins
                actual_gap = (new_start - prev_task["end"]).total_seconds() / 60.0
                if actual_gap < required_gap:
                    reason = (
                        f"前一任務 {prev_task['task_id']} 於 {prev_task['end']:%H:%M} 結束，"
                        f"待接任務 {new_start:%H:%M} 開始；可用 {actual_gap:g} 分，"
                        f"至少需 {required_gap:.1f} 分（車程 {format(incoming, '.1f') if incoming is not None else '未取得'}＋緩衝 {config.buffer_mins:g}）；"
                        f"不足 {required_gap-actual_gap:.1f} 分鐘"
                    )

            # 新案 -> 下一案
            if not reason and next_task is not None:
                if pd.notna(next_task["lat"]) and pd.notna(next_task["lon"]):
                    outgoing = calc_travel_minutes(
                        task_series["服務地點_緯度"], task_series["服務地點_經度"],
                        next_task["lat"], next_task["lon"],
                        config, transport_mode=transport_mode, departure_time=_task_clock(task_series,new_end + timedelta(minutes=config.buffer_mins)),
                    )
                    required_gap = outgoing + config.buffer_mins
                else:
                    required_gap = config.buffer_mins
                actual_gap = (next_task["start"] - new_end).total_seconds() / 60.0
                if actual_gap < required_gap:
                    reason = (
                        f"待接任務 {new_end:%H:%M} 結束，下一任務 {next_task['task_id']} 於 {next_task['start']:%H:%M} 開始；"
                        f"可用 {actual_gap:g} 分，至少需 {required_gap:.1f} 分"
                        f"（車程 {format(outgoing, '.1f') if outgoing is not None else '未取得'}＋緩衝 {config.buffer_mins:g}）；不足 {required_gap-actual_gap:.1f} 分鐘"
                    )

        existing_minutes = sum(a["duration"] for a in assigned)
        used_hours = float(cg.get("今日已佔用工時(小時)", 0.0) or 0.0)
        after_hours = used_hours + (existing_minutes + new_schedule_minutes) / 60.0
        cap = cg.get("每日工時上限(小時)")
        if pd.notna(cap) and after_hours > float(cap):
            cap_reason = f"插單後當日總工時 {after_hours:.1f} 小時，超過上限 {float(cap):.1f} 小時"
            reason = reason + '；另有：' + cap_reason if reason else cap_reason

        available = reason == ""
        if available:
            parts = ["可直接插入，不需移動既有班表"]
            if prev_task is not None:
                parts.append(
                    f"前一案 {prev_task['task_id']} 於 {prev_task['end'].strftime('%H:%M')} 結束"
                )
            if next_task is not None:
                parts.append(
                    f"下一案 {next_task['task_id']} 於 {next_task['start'].strftime('%H:%M')} 開始"
                )
            if bool(match.get("是否歷史首選", False)):
                parts.append("為案家歷史首選")
            recommendation = "；".join(parts)
        else:
            recommendation = ""

        result_rows.append(
            {
                "居服員ID": cg_id,
                "可直接插入": available,
                "適配度分數": float(match.get("適配度分數", 0)),
                "待接任務時段": f"{new_start:%H:%M}–{new_end:%H:%M}",
                "重疊任務與時段": '；'.join(overlap_details) if overlap_details else '無時段重疊',
                "前一任務（結束）": (
                    f"{prev_task['task_id']}｜{prev_task['end'].strftime('%H:%M')}"
                    if prev_task else "無"
                ),
                "下一任務（開始）": (
                    f"{next_task['task_id']}｜{next_task['start'].strftime('%H:%M')}"
                    if next_task else "無"
                ),
                "前段交通時間(分)": round(float(incoming), 1) if incoming is not None else None,
                "後段交通時間(分)": round(float(outgoing), 1) if outgoing is not None else None,
                "插單後排班工時(小時)": round(after_hours, 2),
                "擾動成本": 0 if available else 100,
                "推薦原因": recommendation,
                "不可插入原因": reason,
            }
        )

    result_df = pd.DataFrame(result_rows)
    if result_df.empty:
        return pd.DataFrame(columns=output_columns)

    result_df = result_df.sort_values(
        by=["可直接插入", "擾動成本", "適配度分數"],
        ascending=[False, True, False],
        kind="stable",
    ).reset_index(drop=True)

    # 第一版主要顯示可直接插入的 Top N；若完全沒有可插入者，保留前 N 筆失敗原因供診斷。
    available_df = result_df[result_df["可直接插入"]].head(top_n).copy()
    if not available_df.empty:
        result_df = available_df
    else:
        result_df = result_df.head(top_n).copy()

    result_df.insert(0, "排名", range(1, len(result_df) + 1))
    return result_df[output_columns]

def calculate_schedule_travel(result_df, tasks, caregivers, config):
    """依每位員工、每日有效派單順序計算家→首案→後續案；不含返家。"""
    result = result_df.copy().reset_index(drop=True)
    if result.empty:
        return result
    task_map = tasks.set_index('任務ID').to_dict('index')
    caregiver_map = caregivers.set_index('居服員ID').to_dict('index')
    if '配對起點車程(分)' not in result:
        result['配對起點車程(分)'] = result.get('預估車程(分)', np.nan)
    result['預估車程(分)'] = np.nan
    result['交通起點'] = ''
    result['交通終點'] = ''
    result['交通路段'] = ''
    result['路段緩衝(分)'] = 0.0
    result['含緩衝交通時間(分)'] = np.nan
    result['交通計算狀態'] = ''
    result['交通日期'] = ''
    groups = {}
    for index, row in result.iterrows():
        task = task_map.get(row['任務ID'], {})
        day = _text(task.get('日期')) or _text(row.get('排班日期')) or '單日'
        start = pd.to_datetime(_text(task.get('時間窗_開始')), format='%H:%M', errors='coerce')
        groups.setdefault((row['派單居服員'], day), []).append((start, str(row['任務ID']), index))
    route_cache = {}
    for (cg_id, day), entries in groups.items():
        cg = caregiver_map.get(cg_id, {})
        mode = _text(cg.get('常用交通工具')) or '機車'
        previous = (_number(cg.get('服務起點_緯度(家)')), _number(cg.get('服務起點_經度(家)')))
        origin = '員工家／服務起點'
        previous_end = None
        entries.sort(key=lambda entry: (pd.isna(entry[0]), entry[0] if pd.notna(entry[0]) else pd.Timestamp.max, entry[1]))
        for order, (start, task_id, index) in enumerate(entries):
            task = task_map.get(result.at[index, '任務ID'], {})
            destination = (_number(task.get('服務地點_緯度')), _number(task.get('服務地點_經度')))
            label = f"{_text(task.get('案家ID'))}（{task_id}）"
            buffer = 0.0 if order == 0 else config.buffer_mins
            result.loc[index, ['交通起點','交通終點','交通路段','交通日期']] = [origin,label,origin+' → '+label,day]
            result.at[index,'路段緩衝(分)'] = buffer
            if pd.isna(start):
                result.at[index,'交通計算狀態'] = '缺服務開始時間，無法確認順序'
            elif order > 0 and mode == '大眾運輸' and (previous_end is None or pd.isna(previous_end)):
                result.at[index,'交通計算狀態'] = '缺前案結束時間，無法查詢公車／捷運班次'
            elif not all(np.isfinite(value) for value in (*previous, *destination)):
                result.at[index,'交通計算狀態'] = '缺起點或終點座標，無法計算'
            else:
                arrival = _task_clock(task,start) if order == 0 else None
                departure = _task_clock(task,previous_end + timedelta(minutes=buffer)) if previous_end is not None and pd.notna(previous_end) else None
                key = (*previous, *destination, mode, str(departure), str(arrival))
                if key not in route_cache:
                    route_cache[key] = calc_travel_minutes(*previous, *destination, config, transport_mode=mode, departure_time=departure, arrival_time=arrival)
                minutes = route_cache[key]
                result.at[index,'預估車程(分)'] = round(minutes, 2)
                result.at[index,'含緩衝交通時間(分)'] = round(minutes + buffer, 2)
                result.at[index,'交通計算狀態'] = '依當日服務順序估算'
                result.at[index,'交通估算來源'] = _travel_source(mode)
                result.at[index,'交通查詢時間'] = _google_timestamp(departure or arrival) or ''
                if order > 0 and previous_end is not None and pd.notna(previous_end):
                    result.at[index,'轉場時段檢查'] = ('轉場時間不足，需調班' if previous_end + timedelta(minutes=minutes + buffer) > start else '轉場時間足夠')
                else:
                    result.at[index,'轉場時段檢查'] = '首案，請確認出門時間'
            previous, origin = destination, label
            previous_end = pd.to_datetime(_text(task.get('時間窗_結束')),format='%H:%M',errors='coerce')
    return result


def verify_schedule_google_travel(result_df, tasks, caregivers, config, max_queries=100):
    """只核對已排路段；先計數，超限不發送任何 Google 請求。"""
    allowed = _GOOGLE_ALLOWED.set(True)
    queries = set()
    collect = _GOOGLE_COLLECT.set(queries)
    try:
        preview = calculate_schedule_travel(result_df, tasks, caregivers, config)
        if not preview.empty and preview['預估車程(分)'].isna().any():
            raise ValueError('部分路段缺座標或時間，請先補齊；尚未發送 Google 查詢。')
        if len(queries) > max_queries:
            raise ValueError(f'已排班需要 {len(queries)} 筆 Google 查詢，超過上限 {max_queries}；尚未發送任何 Google 查詢。')
        _GOOGLE_COLLECT.reset(collect)
        collect = None
        begin_travel_query_run(max_queries)
        return calculate_schedule_travel(result_df, tasks, caregivers, config)
    finally:
        if collect is not None:
            _GOOGLE_COLLECT.reset(collect)
        _GOOGLE_ALLOWED.reset(allowed)


def summarize_schedule_travel(result_df):
    """每日交通與緩衝分開加總，缺一段則完整總計保持未知。"""
    if result_df.empty:
        return pd.DataFrame()
    rows = []
    for (cg_id, day), group in result_df.groupby(['派單居服員','交通日期'], sort=True):
        missing = int(group['預估車程(分)'].isna().sum())
        rows.append({'居服員ID':cg_id,'日期':day,'路段數':len(group),
            '交通合計(分)': group['預估車程(分)'].sum() if not missing else np.nan,
            '轉場緩衝合計(分)':group['路段緩衝(分)'].sum(),
            '含緩衝合計(分)':group['含緩衝交通時間(分)'].sum() if not missing else np.nan,
            '缺資料路段數':missing})
    return pd.DataFrame(rows)


# ==========================================
# Phase 2: 時空路徑衝突過濾 + OR-Tools 多目標最佳化
# ==========================================
def _interval_travel(first, second, config, mode, task):
    a_start,a_end,a_lat,a_lon = first
    b_start,b_end,b_lat,b_lon = second
    if a_end <= b_start:
        origin,destination,depart = (a_lat,a_lon),(b_lat,b_lon),a_end
    elif b_end <= a_start:
        origin,destination,depart = (b_lat,b_lon),(a_lat,a_lon),b_end
    else:
        return float('inf')
    return calc_travel_minutes(*origin,*destination,config,transport_mode=mode,
                               departure_time=_task_clock(task,depart + timedelta(minutes=config.buffer_mins))) + config.buffer_mins


def run_phase2_optimization(
    df_matches: pd.DataFrame,
    tasks: pd.DataFrame,
    df_cg: pd.DataFrame,
    config: PipelineConfig,
    extra_busy_blocks: Optional[dict] = None,
):
    """執行 Phase 2 時空衝突過濾與 OR-Tools 最佳化派單。

    extra_busy_blocks（cg_id -> [(start, end, lat, lon), ...]，可選）用於「週期性任務
    優先」排班：由呼叫端（見 `_run_periodic_then_adhoc`）將前一輪已鎖定的週期性任務
    指派結果，轉為額外的忙碌時間區塊注入本輪臨時單次任務的衝突檢查，使臨時任務
    只能競爭週期性任務排定後剩餘的時段，不會與其重疊。
    """
    # 速度優化：Phase 1 已完成完整評分，Phase 2 每個任務只保留最高分候選，
    # 大幅縮小 pairwise 時空衝突與 OR-Tools 變數數量。
    if not df_matches.empty and '適配度分數' in df_matches.columns:
        df_matches = (
            df_matches.sort_values(['任務ID', '適配度分數'], ascending=[True, False])
            .groupby('任務ID', sort=False, as_index=False, group_keys=False)
            .head(PHASE2_CANDIDATE_CAP_PER_TASK)
            .reset_index(drop=True)
        )

    # 建立居服員今日既定行程時間阻擋塊 (Time Blocks)。
    transport_by_cg={r['居服員ID']:r.get('常用交通工具','機車') for _,r in df_cg.iterrows()}
    cg_busy = _build_cg_busy_blocks(df_cg)
    if extra_busy_blocks:
        for cg_id, blocks in extra_busy_blocks.items():
            cg_busy.setdefault(cg_id, []).extend(blocks)

    # 建立任務時間表
    task_times = {}
    task_rows = tasks.set_index('任務ID').to_dict('index')
    for _, task in tasks.iterrows():
        t_id = task["任務ID"]
        t1 = datetime.strptime(task["時間窗_開始"].strip(), "%H:%M")
        t2 = datetime.strptime(task["時間窗_結束"].strip(), "%H:%M")
        task_times[t_id] = (t1, t2, task["服務地點_緯度"], task["服務地點_經度"])

    # 進入衝突檢查迴圈前，先以 OSRM /table 批次查詢暖身快取（任務地點 x 既定行程地點、
    # 任務地點 x 任務地點），取代逐筆個別 HTTP 請求。
    task_coords = [(lat, lon) for _, _, lat, lon in task_times.values()]
    busy_coords = [
        (b_lat, b_lon) for intervals in cg_busy.values() for (_, _, b_lat, b_lon) in intervals
    ]
    # 機車：OSRM table 一次批次暖身。
    if '機車' in set(str(v).strip() for v in transport_by_cg.values()):
        if busy_coords:
            prefetch_osrm_travel_times(task_coords, busy_coords, config.travel_min_per_km, transport_mode='機車')
        prefetch_osrm_travel_times(task_coords, task_coords, config.travel_min_per_km, transport_mode='機車')

    # UltraFast：Phase 2 不預查所有大眾運輸候選的兩兩交通。
    # 先求解，再只對真正被選中的相鄰轉場做 Google TRANSIT 精確驗證。

    # 過濾掉與既定行程衝突（含轉場緩衝時間）的配對
    valid_rows = []
    for _, row in df_matches.iterrows():
        t_id = row["任務ID"]
        cg_id = row["居服員ID"]
        t_start, t_end, t_lat, t_lon = task_times[t_id]

        conflict = False
        mode = str(transport_by_cg.get(cg_id, "機車")).strip()
        for b_start, b_end, b_lat, b_lon in cg_busy[cg_id]:
            if mode == "大眾運輸":
                if not (t_end <= b_start or b_end <= t_start):
                    conflict = True
                    break
            else:
                travel_mins = _interval_travel(
                    task_times[t_id],
                    (b_start, b_end, b_lat, b_lon),
                    config,
                    mode,
                    task_rows[t_id],
                )
                if not np.isfinite(travel_mins) or not (
                    t_end + timedelta(minutes=travel_mins) <= b_start
                    or b_end + timedelta(minutes=travel_mins) <= t_start
                ):
                    conflict = True
                    break

        if not conflict:
            valid_rows.append(row.to_dict())

    df_valid = pd.DataFrame(valid_rows)

    result = {
        "df_valid": df_valid,
        'df_matches': df_matches,
        'matching_diagnostics':df_matches.attrs.get('matching_diagnostics',[]),
        "status": None,
        "df_result": pd.DataFrame(),
        "assigned_count": 0,
    }

    if df_valid.empty:
        return result

    # 建立 OR-Tools 混合整數規劃 (MIP) 求解器
    solver = pywraplp.Solver.CreateSolver("SCIP")
    if solver is None:
        raise RuntimeError('OR-Tools SCIP 求解器無法啟動')
    solver.SetTimeLimit(config.solver_time_limit_ms)
    solver.SetSolverSpecificParametersAsString(f"limits/gap = {config.solver_relative_gap}")
    tasks_by_cg = df_valid.groupby("居服員ID", sort=False)["任務ID"].agg(list).to_dict()
    caregivers_by_task = df_valid.groupby("任務ID", sort=False)["居服員ID"].agg(list).to_dict()
    match_lookup = {(r["任務ID"], r["居服員ID"]): r for r in df_valid.to_dict("records")}
    pair_travel_cache = {}

    X = {}
    for _, row in df_valid.iterrows():
        X[(row["任務ID"], row["居服員ID"])] = solver.IntVar(
            0, 1, f"x_{row['任務ID']}_{row['居服員ID']}"
        )

    # 限制條件 1：每個任務至多隻能派給一位居服員
    for t_id in df_valid["任務ID"].unique():
        solver.Add(
            sum(
                X[(t_id, cg_id)]
                for cg_id in caregivers_by_task[t_id]
            )
            <= 1
        )

    # 限制條件 2：同一居服員若被派兩個新任務，時間不能重疊且須預留轉場緩衝
    for cg_id in df_valid["居服員ID"].unique():
        cg_tasks = tasks_by_cg[cg_id]
        for i in range(len(cg_tasks)):
            for j in range(i + 1, len(cg_tasks)):
                t1_id, t2_id = cg_tasks[i], cg_tasks[j]
                t1_start, t1_end, t1_lat, t1_lon = task_times[t1_id]
                t2_start, t2_end, t2_lat, t2_lon = task_times[t2_id]

                mode = str(transport_by_cg.get(cg_id, "機車")).strip()
                if mode == "大眾運輸":
                    if not (t1_end <= t2_start or t2_end <= t1_start):
                        solver.Add(X[(t1_id, cg_id)] + X[(t2_id, cg_id)] <= 1)
                else:
                    cache_key = (t1_id, t2_id, mode)
                    if cache_key not in pair_travel_cache:
                        pair_travel_cache[cache_key] = _interval_travel(
                            task_times[t1_id], task_times[t2_id],
                            config, mode, task_rows[t1_id]
                        )
                    travel_mins = pair_travel_cache[cache_key]
                    if not np.isfinite(travel_mins) or not (
                        t1_end + timedelta(minutes=travel_mins) <= t2_start
                        or t2_end + timedelta(minutes=travel_mins) <= t1_start
                    ):
                        solver.Add(X[(t1_id, cg_id)] + X[(t2_id, cg_id)] <= 1)

    # 限制條件 2b：規則1（4小時/30分鐘休息）——累計連續工作達門檻須強制安插休息。
    _build_break_constraints(solver, X, df_valid, cg_busy, task_times, config)

    # 限制條件 3：居服員今日新派任務總歷時 + 既有已佔用工時，不得超過每日工時上限。
    # Phase 1 僅逐筆過濾單一任務是否超時，無法阻擋「多筆任務加總後超派」的組合，
    # 此為求解器層級的產能限制式，修復該缺口。
    task_duration_hrs = {
        row["任務ID"]: row["服務歷時(分鐘)"] / 60.0 for _, row in tasks.iterrows()
    }
    cg_capacity = {
        row["居服員ID"]: (row.get("今日已佔用工時(小時)", 0.0), row["每日工時上限(小時)"])
        for _, row in df_cg.iterrows()
    }
    for cg_id in df_valid["居服員ID"].unique():
        used_hours, cap_hours = cg_capacity.get(cg_id, (0.0, float("inf")))
        cg_task_ids = tasks_by_cg[cg_id]
        solver.Add(
            sum(X[(t_id, cg_id)] * task_duration_hrs[t_id] for t_id in cg_task_ids)
            + used_hours
            <= cap_hours
        )

    # 財務試算：每筆任務的長照申報點數（營收）與居服員拆帳薪資，
    # 僅供結果顯示與財務 KPI 試算，不參與派單目標函數。
    tasks_with_revenue = calculate_task_revenue_and_salary(tasks, config)
    task_revenue_map = {
        row["任務ID"]: (row["預估長照申報點數(營收)"], row["預估居服員拆帳薪資"])
        for _, row in tasks_with_revenue.iterrows()
    }

    # Phase 2 目標函數：
    # 在通過所有硬性條件與時空衝突檢查的候選方案中，
    # 最大化 Phase 1 適配度分數與任務優先權。
    # 車程已納入 Phase 1 適配度計算，不重複扣分；
    # 財務營收僅供結果顯示與 KPI 試算，不參與派單決策。
    #
    # 「照護連續性」優先於「車程限制」的原則亦須貫徹到此目標函數層級：若僅單純調高
    # 反而會讓車程較遠但為案家歷史首選的居服員，在整體最佳化階段被距離較近的陌生
    # 居服員取代——這違背了 Phase 1 刻意給予首選居服員高額連續性加分的用意（其車程
    # 成本已由 Phase 1 的 travel_penalty_cap 合理封頂）。因此歷史首選居服員的配對在
    objective = solver.Objective()
    for _, row in df_valid.iterrows():
        t_id = row["任務ID"]
        cg_id = row["居服員ID"]
        w_match = row["適配度分數"]
        priority_bonus = (
            config.urgent_priority_bonus
            if "高" in str(row["優先級"])
            else config.normal_priority_bonus
        )

        coeff = w_match  + priority_bonus 
        objective.SetCoefficient(X[(t_id, cg_id)], coeff)

    objective.SetMaximization()

    caregiver_home = {
        r["居服員ID"]: (
            _number(r.get("服務起點_緯度(家)")),
            _number(r.get("服務起點_經度(家)")),
        )
        for _, r in df_cg.iterrows()
    }

    def _selected_transit_cuts():
        selected_by_cg = {}
        for (tid, cid), var in X.items():
            if (
                var.solution_value() > 0.5
                and str(transport_by_cg.get(cid, "機車")).strip() == "大眾運輸"
            ):
                selected_by_cg.setdefault(cid, []).append(tid)

        required = []
        checks = []

        for cid, tids in selected_by_cg.items():
            tids = sorted(tids, key=lambda tid: task_times[tid][0])
            if not tids:
                continue

            first = tids[0]
            f_start, _, f_lat, f_lon = task_times[first]
            h_lat, h_lon = caregiver_home.get(cid, (np.nan, np.nan))
            if all(np.isfinite(_number(v)) for v in (h_lat, h_lon, f_lat, f_lon)):
                arrival = _task_clock(task_rows[first], f_start)
                required.append((h_lat, h_lon, f_lat, f_lon, None, arrival))
                checks.append(("first", first, cid, h_lat, h_lon, f_lat, f_lon, arrival))

            for left, right in zip(tids, tids[1:]):
                _, l_end, l_lat, l_lon = task_times[left]
                r_start, _, r_lat, r_lon = task_times[right]
                if l_end > r_start:
                    checks.append(("pair_overlap", left, right, cid))
                    continue
                depart = _task_clock(
                    task_rows[left],
                    l_end + timedelta(minutes=config.buffer_mins),
                )
                required.append((l_lat, l_lon, r_lat, r_lon, depart, None))
                checks.append(
                    ("pair", left, right, cid, l_end, r_start,
                     l_lat, l_lon, r_lat, r_lon, depart)
                )

            for tid in tids:
                t_start, t_end, t_lat, t_lon = task_times[tid]
                for b_start, b_end, b_lat, b_lon in cg_busy.get(cid, []):
                    if t_end <= b_start:
                        depart = _task_clock(
                            task_rows[tid],
                            t_end + timedelta(minutes=config.buffer_mins),
                        )
                        required.append((t_lat, t_lon, b_lat, b_lon, depart, None))
                        checks.append(
                            ("busy_after", tid, cid, t_end, b_start,
                             t_lat, t_lon, b_lat, b_lon, depart)
                        )
                    elif b_end <= t_start:
                        depart = _task_clock(
                            task_rows[tid],
                            b_end + timedelta(minutes=config.buffer_mins),
                        )
                        required.append((b_lat, b_lon, t_lat, t_lon, depart, None))
                        checks.append(
                            ("busy_before", tid, cid, b_end, t_start,
                             b_lat, b_lon, t_lat, t_lon, depart)
                        )
                    else:
                        checks.append(("busy_overlap", tid, cid))

        prefetch_google_transit_pairs(required)

        cuts = []
        for chk in checks:
            kind = chk[0]
            if kind == "pair_overlap":
                _, left, right, cid = chk
                cuts.append(("pair", left, right, cid))
                continue
            if kind == "busy_overlap":
                _, tid, cid = chk
                cuts.append(("single", tid, cid))
                continue
            if kind == "first":
                _, tid, cid, olat, olon, dlat, dlon, arrival = chk
                mins = calc_travel_minutes(
                    olat, olon, dlat, dlon, config,
                    transport_mode="大眾運輸",
                    arrival_time=arrival,
                )
                if not np.isfinite(mins):
                    cuts.append(("single", tid, cid))
                continue

            if kind == "pair":
                _, left, right, cid, source_end, target_start, olat, olon, dlat, dlon, depart = chk
                mins = calc_travel_minutes(
                    olat, olon, dlat, dlon, config,
                    transport_mode="大眾運輸",
                    departure_time=depart,
                )
                total = mins + config.buffer_mins
                if (
                    not np.isfinite(mins)
                    or source_end + timedelta(minutes=total) > target_start
                ):
                    cuts.append(("pair", left, right, cid))
                continue

            if kind in ("busy_after", "busy_before"):
                _, tid, cid, source_end, target_start, olat, olon, dlat, dlon, depart = chk
                mins = calc_travel_minutes(
                    olat, olon, dlat, dlon, config,
                    transport_mode="大眾運輸",
                    departure_time=depart,
                )
                total = mins + config.buffer_mins
                if (
                    not np.isfinite(mins)
                    or source_end + timedelta(minutes=total) > target_start
                ):
                    cuts.append(("single", tid, cid))
                continue

        return list(dict.fromkeys(cuts))

    status = solver.Solve()

    for _round in range(3):
        if status not in (pywraplp.Solver.OPTIMAL, pywraplp.Solver.FEASIBLE):
            break
        cuts = _selected_transit_cuts()
        if not cuts:
            break

        added = 0
        for cut in cuts:
            if cut[0] == "single":
                _, tid, cid = cut
                var = X.get((tid, cid))
                if var is not None:
                    solver.Add(var <= 0)
                    added += 1
            else:
                _, left, right, cid = cut
                v1, v2 = X.get((left, cid)), X.get((right, cid))
                if v1 is not None and v2 is not None:
                    solver.Add(v1 + v2 <= 1)
                    added += 1

        if added == 0:
            break
        status = solver.Solve()

    result["status"] = status

    if status in (pywraplp.Solver.OPTIMAL, pywraplp.Solver.FEASIBLE):
        assigned_count = 0
        results = []
        for (t_id, cg_id), var in X.items():
            if var.solution_value() > 0.5:
                assigned_count += 1
                row_data = match_lookup[(t_id, cg_id)]
                revenue, salary = task_revenue_map.get(t_id, (0.0, 0.0))
                results.append(
                    {
                        "任務ID": t_id,
                        "案家ID": row_data["案家ID"],
                        "派單居服員": cg_id,
                        "適配分數": row_data["適配度分數"],
                        "預估車程(分)": row_data["預估交通時間(分)"],
                        "服務時段": f"{row_data['任務開始時間']}-{row_data['任務結束時間']}",
                        "任務優先級": row_data["優先級"],
                        "地點緯度": row_data["地點緯度"],
                        "地點經度": row_data["地點經度"],
                        "預估長照申報點數(營收)": revenue,
                        "預估居服員拆帳薪資": salary,
                        "基礎分": row_data.get("基礎分", 0),

                        "專長匹配加分": row_data.get("專長匹配加分", 0),
                        "歷史首選加分": row_data.get("歷史首選加分", 0),
                        "上一週同案加分": row_data.get("上一週同案加分", 0),
                        "連續性品質加分": row_data.get("連續性品質加分", 0),
                        "滿意度調整": row_data.get("滿意度調整", 0),
                        "交通扣分": row_data.get("交通扣分", 0),
                        "工作負荷扣分": row_data.get("工作負荷扣分", 0),
                        "是否歷史首選": row_data.get("是否歷史首選", False),
                        "是否上一週同案居服員": row_data.get("是否上一週同案居服員", False),
                        "上一週居服員ID": row_data.get("上一週居服員ID", ""),
                        "當月累計服務時數": row_data.get("當月累計服務時數", 0),
                        "服務強度係數": row_data.get("服務強度係數", 1),
                        **{k:row_data.get(k) for k in ['疲勞資料狀態','滿意度資料狀態','本次計畫已排時數',
                           '失智症資料狀態','特殊資格檢查','本次計畫加權負荷時數','照護難度參考係數','照護難度判斷依據','照護提示','服務碼必要證照','本次計畫負荷扣分','實際工時負荷扣分','資料來源','待確認事項','交通估算來源']},
                    }
                )

        # 每位居服員本次新指派到的任務清單，供「原首選替換原因」判斷同時段衝突用
        assigned_task_ids_by_cg = {}
        for row in results:
            assigned_task_ids_by_cg.setdefault(row["派單居服員"], []).append(row["任務ID"])

        for row in results:
            t_id = row["任務ID"]
            task_row = tasks[tasks["任務ID"] == t_id].iloc[0]
            pref_cg_id = task_row["歷史首選居服員ID"]
            other_task_ids = [
                tid for tid in assigned_task_ids_by_cg.get(pref_cg_id, []) if tid != t_id
            ]
            row["原首選替換原因"] = _diagnose_caregiver_change(
                task_row,
                pref_cg_id,
                row["派單居服員"],
                df_cg,
                df_matches,
                df_valid,
                other_task_ids,
                task_times,
                config,
            )

        df_result = pd.DataFrame(results)
        if not df_result.empty:
            df_result = df_result.sort_values(by="任務ID")
        result["df_result"] = calculate_schedule_travel(df_result, tasks, df_cg, config)
        result["assigned_count"] = assigned_count

    return result

PERIODIC_TASK_COLUMN = "是否為週期性任務"

def _is_periodic_task(value) -> bool:
    """將「是否為週期性任務」欄位值正規化為布林值。

    Excel 布林儲存格經 pandas 讀入後，實際型別可能是原生 Python bool（勾選格式）、
    字串 "TRUE"/"FALSE"（文字格式儲存格）、或 1/0；本函式統一正規化為 bool，
    空白（NaN）保守視為非週期性（臨時單次）任務。
    """
    if pd.isna(value):
        return False
    if isinstance(value, str):
        return value.strip().upper() == "TRUE"
    return bool(value)

def _split_periodic_tasks(tasks: pd.DataFrame, periodic_column: str = PERIODIC_TASK_COLUMN):
    """依 periodic_column 將 tasks 拆分為 (週期性任務, 臨時單次任務)。

    欄位不存在時回傳 (None, None)，代表呼叫端應維持原本單一批次排班邏輯
    （向下相容不含此欄位的舊版資料表）。
    """
    if periodic_column not in tasks.columns:
        return None, None
    is_periodic = tasks[periodic_column].apply(_is_periodic_task)
    return tasks[is_periodic].copy(), tasks[~is_periodic].copy()

def _run_periodic_then_adhoc(
    periodic_tasks: pd.DataFrame,
    adhoc_tasks: pd.DataFrame,
    df_cg: pd.DataFrame,
    config: PipelineConfig,
    previous_week_map: dict | None = None,
) -> dict:
    """規則2：週期性排班優先於臨時單次排班。

    先只用週期性任務跑一輪 Phase 1 + Phase 2，鎖定基礎班表；再將該輪指派結果轉為
    額外忙碌時間區塊與已佔用工時，注入臨時單次任務的第二輪 Phase 1 + Phase 2，
    使臨時任務只能競爭週期性任務排定後「剩餘的居服員產能與時段」，不會與其重疊
    或反過來排擠週期性任務。
    """
    empty_result = {"df_valid": pd.DataFrame(), "status": None, "df_result": pd.DataFrame(), "assigned_count": 0,
                    'df_matches':pd.DataFrame(),'matching_diagnostics':[]}

    periodic_result = empty_result
    if not periodic_tasks.empty:
        periodic_matches = run_phase1_matching(periodic_tasks, df_cg, config, previous_week_map=previous_week_map)
        periodic_result = run_phase2_optimization(periodic_matches, periodic_tasks, df_cg, config)

    df_cg_adhoc = df_cg.copy()
    extra_busy_blocks: dict = {}
    df_periodic_result = periodic_result["df_result"]
    if not df_periodic_result.empty:
        if "今日已佔用工時(小時)" not in df_cg_adhoc.columns:
            df_cg_adhoc["今日已佔用工時(小時)"] = 0.0

        periodic_tasks_by_id = periodic_tasks.set_index("任務ID")
        hours_by_cg: dict = {}
        for _, row in df_periodic_result.iterrows():
            t_id = row["任務ID"]
            cg_id = row["派單居服員"]
            task_row = periodic_tasks_by_id.loc[t_id]
            t_start = datetime.strptime(str(task_row["時間窗_開始"]).strip(), "%H:%M")
            t_end = datetime.strptime(str(task_row["時間窗_結束"]).strip(), "%H:%M")
            extra_busy_blocks.setdefault(cg_id, []).append(
                (t_start, t_end, task_row["服務地點_緯度"], task_row["服務地點_經度"])
            )
            hours_by_cg[cg_id] = hours_by_cg.get(cg_id, 0.0) + task_row["服務歷時(分鐘)"] / 60.0

        for cg_id, hrs in hours_by_cg.items():
            mask = df_cg_adhoc["居服員ID"] == cg_id
            df_cg_adhoc.loc[mask, "今日已佔用工時(小時)"] = (
                df_cg_adhoc.loc[mask, "今日已佔用工時(小時)"].fillna(0.0) + hrs
            )

    adhoc_result = empty_result
    if not adhoc_tasks.empty:
        adhoc_matches = run_phase1_matching(adhoc_tasks, df_cg_adhoc, config, previous_week_map=previous_week_map)
        adhoc_result = run_phase2_optimization(
            adhoc_matches, adhoc_tasks, df_cg_adhoc, config, extra_busy_blocks=extra_busy_blocks
        )

    df_valid = pd.concat([periodic_result["df_valid"], adhoc_result["df_valid"]], ignore_index=True)
    df_result = pd.concat([df_periodic_result, adhoc_result["df_result"]], ignore_index=True)
    if not df_result.empty:
        df_result = calculate_schedule_travel(df_result, pd.concat([periodic_tasks, adhoc_tasks], ignore_index=True), df_cg, config)

    return {
        "df_valid": df_valid,
        "df_result": df_result,
        "status": adhoc_result["status"] if not adhoc_tasks.empty else periodic_result["status"],
        "assigned_count": periodic_result["assigned_count"] + adhoc_result["assigned_count"],
        'df_matches':pd.concat([periodic_result['df_matches'],adhoc_result['df_matches']],ignore_index=True),
        'matching_diagnostics':periodic_result['matching_diagnostics']+adhoc_result['matching_diagnostics'],
    }

def _client_primary_caregiver_map(result_df: pd.DataFrame) -> dict:
    if result_df is None or result_df.empty:
        return {}
    required = {"案家ID", "派單居服員"}
    if not required.issubset(result_df.columns):
        return {}
    tmp = result_df[["案家ID", "派單居服員"]].dropna().copy()
    tmp["案家ID"] = tmp["案家ID"].astype(str).str.strip()
    tmp["派單居服員"] = tmp["派單居服員"].astype(str).str.strip()
    tmp = tmp[(tmp["案家ID"] != "") & (tmp["派單居服員"] != "") & (tmp["派單居服員"] != "未指派")]
    if tmp.empty:
        return {}
    counts = (
        tmp.groupby(["案家ID", "派單居服員"], as_index=False)
        .size()
        .sort_values(["案家ID", "size", "派單居服員"], ascending=[True, False, True], kind="stable")
    )
    top = counts.drop_duplicates("案家ID", keep="first")
    return dict(zip(top["案家ID"], top["派單居服員"]))



def _task_recurrence_key(task_or_id) -> str:
    """同一週服務列跨週的穩定 key；例如 W0411_20261003 -> W0411。"""
    if isinstance(task_or_id, pd.Series):
        source_row = _text(task_or_id.get("來源列"))
        if source_row:
            return "SRC:" + source_row
        task_id = _text(task_or_id.get("任務ID"))
    elif isinstance(task_or_id, dict):
        source_row = _text(task_or_id.get("來源列"))
        if source_row:
            return "SRC:" + source_row
        task_id = _text(task_or_id.get("任務ID"))
    else:
        task_id = _text(task_or_id)
    return re.sub(r"_\d{8}$", "", task_id)


def _week_exact_assignment_map(result_df: pd.DataFrame) -> dict:
    """將一週最終班表轉成「週期任務 key -> 居服員ID」。"""
    if result_df is None or result_df.empty:
        return {}
    output = {}
    for _, row in result_df.iterrows():
        task_id = _text(row.get("任務ID"))
        cg_id = _text(row.get("派單居服員"))
        if not task_id or not cg_id or cg_id == "未指派":
            continue
        key = _task_recurrence_key(task_id)
        if key:
            output[key] = cg_id
    return output


def _build_busy_blocks_from_assignments(result_df: pd.DataFrame, tasks: pd.DataFrame) -> tuple[dict, dict]:
    """已鎖定班表 -> 忙碌時段 + 每位居服員工時。"""
    blocks, hours = {}, {}
    if result_df is None or result_df.empty:
        return blocks, hours
    task_lookup = tasks.set_index("任務ID").to_dict("index")
    for _, row in result_df.iterrows():
        task = task_lookup.get(row.get("任務ID"))
        if not task:
            continue
        cg_id = row.get("派單居服員")
        start = datetime.strptime(str(task["時間窗_開始"]).strip(), "%H:%M")
        end = datetime.strptime(str(task["時間窗_結束"]).strip(), "%H:%M")
        blocks.setdefault(cg_id, []).append(
            (start, end, task.get("服務地點_緯度"), task.get("服務地點_經度"))
        )
        hours[cg_id] = hours.get(cg_id, 0.0) + float(task.get("服務歷時(分鐘)", 0)) / 60.0
    return blocks, hours


def _carried_result_rows(carried_pairs: list[tuple[str, str]], day_tasks: pd.DataFrame,
                         config: PipelineConfig) -> pd.DataFrame:
    """把已驗證可沿用的任務製成標準派單結果列。"""
    if not carried_pairs:
        return pd.DataFrame()
    task_lookup = day_tasks.set_index("任務ID").to_dict("index")
    financial = calculate_task_revenue_and_salary(day_tasks, config)
    finance_lookup = financial.set_index("任務ID").to_dict("index")
    rows = []
    for task_id, cg_id in carried_pairs:
        task = task_lookup[task_id]
        fin = finance_lookup.get(task_id, {})
        rows.append({
            "任務ID": task_id,
            "案家ID": task.get("案家ID"),
            "派單居服員": cg_id,
            "適配分數": float(config.base_score + config.previous_week_caregiver_bonus),
            "預估車程(分)": np.nan,
            "服務時段": f"{task.get('時間窗_開始')}-{task.get('時間窗_結束')}",
            "任務優先級": task.get("任務優先級", "一般"),
            "地點緯度": task.get("服務地點_緯度"),
            "地點經度": task.get("服務地點_經度"),
            "預估長照申報點數(營收)": fin.get("預估長照申報點數(營收)", np.nan),
            "預估居服員拆帳薪資": fin.get("預估居服員拆帳薪資", np.nan),
            "基礎分": config.base_score,
            "專長匹配加分": 0.0,
            "歷史首選加分": 0.0,
            "上一週同案加分": config.previous_week_caregiver_bonus,
            "連續性品質加分": 0.0,
            "滿意度調整": 0.0,
            "交通扣分": 0.0,
            "工作負荷扣分": 0.0,
            "是否歷史首選": False,
            "是否上一週同案居服員": True,
            "上一週居服員ID": cg_id,
            "排班來源": "上一週班表沿用",
            "原首選替換原因": "上一週班表通過硬限制與時空檢查，直接沿用",
        })
    return pd.DataFrame(rows)


def _attempt_carry_forward(day_tasks: pd.DataFrame, periodic_tasks: pd.DataFrame,
                           df_cg: pd.DataFrame, config: PipelineConfig,
                           previous_exact_map: dict) -> tuple[pd.DataFrame, pd.DataFrame, list]:
    """先沿用上一週同一週期任務；不符合硬限制/衝突者回到待排池。"""
    if periodic_tasks is None or periodic_tasks.empty or not previous_exact_map:
        return pd.DataFrame(), periodic_tasks.copy() if periodic_tasks is not None else pd.DataFrame(), []

    task_locations = {
        row["任務ID"]: (row.get("服務地點_緯度"), row.get("服務地點_經度"))
        for _, row in day_tasks.iterrows()
    }
    locked_preview = pd.DataFrame(columns=["任務ID", "案家ID", "派單居服員"])
    carried_pairs, failed = [], []

    ordered = periodic_tasks.copy()
    ordered["_sort_start"] = pd.to_datetime(ordered["時間窗_開始"], format="%H:%M", errors="coerce")
    ordered = ordered.sort_values(["_sort_start", "任務ID"], kind="stable")

    for _, task in ordered.iterrows():
        key = _task_recurrence_key(task)
        cg_id = _text(previous_exact_map.get(key))
        if not cg_id:
            failed.append((task["任務ID"], "上一週無對應任務"))
            continue
        evaluation = _evaluate_reassignment(
            task["任務ID"], cg_id, day_tasks, locked_preview, df_cg, config,
            task_locations=task_locations, date_column="日期", df_cl=None,
        )
        if evaluation.get("available"):
            carried_pairs.append((task["任務ID"], cg_id))
            locked_preview = pd.concat([
                locked_preview,
                pd.DataFrame([{"任務ID": task["任務ID"], "案家ID": task.get("案家ID"), "派單居服員": cg_id}])
            ], ignore_index=True)
        else:
            failed.append((task["任務ID"], evaluation.get("detail") or evaluation.get("reason") or "不符合沿用條件"))

    carried_ids = {task_id for task_id, _ in carried_pairs}
    pending = periodic_tasks[~periodic_tasks["任務ID"].isin(carried_ids)].copy()
    return _carried_result_rows(carried_pairs, day_tasks, config), pending, failed


# ==========================================
# 全月/按日批次派單：對每個日期各自執行 Phase 1 + Phase 2
# ==========================================

def _fast_weekly_key(task_or_id) -> str:
    # 跨週固定服務 key 一律優先使用任務ID去掉日期尾碼。
    # W0411_20261101 -> W0411
    # W0411_20261108 -> W0411
    if isinstance(task_or_id, pd.Series):
        task_id = _text(task_or_id.get("任務ID"))
        src = _text(task_or_id.get("來源列"))
    elif isinstance(task_or_id, dict):
        task_id = _text(task_or_id.get("任務ID"))
        src = _text(task_or_id.get("來源列"))
    else:
        task_id = _text(task_or_id)
        src = ""

    if task_id:
        return re.sub(r"_\d{8}$", "", task_id)

    return ("SRC:" + src) if src else ""


def _fast_exact_map(result_df: pd.DataFrame, tasks: pd.DataFrame | None = None) -> dict:
    if result_df is None or result_df.empty:
        return {}
    task_lookup = {}
    if tasks is not None and not tasks.empty and "任務ID" in tasks.columns:
        task_lookup = tasks.set_index("任務ID").to_dict("index")
    out = {}
    for _, row in result_df.iterrows():
        task_id = _text(row.get("任務ID"))
        cg_id = _text(row.get("派單居服員"))
        if not task_id or not cg_id or cg_id == "未指派":
            continue
        task = task_lookup.get(task_id)
        key = _fast_weekly_key(task if task is not None else task_id)
        if key:
            out[key] = cg_id
    return out


def _fast_carried_rows(carried_pairs, day_tasks, config):
    if not carried_pairs:
        return pd.DataFrame()
    task_lookup = day_tasks.set_index("任務ID").to_dict("index")
    financial = calculate_task_revenue_and_salary(day_tasks, config)
    finance_lookup = financial.set_index("任務ID").to_dict("index")
    rows = []
    for task_id, cg_id in carried_pairs:
        t = task_lookup[task_id]
        fin = finance_lookup.get(task_id, {})
        rows.append({
            "任務ID": task_id,
            "案家ID": t.get("案家ID"),
            "派單居服員": cg_id,
            "適配分數": float(config.base_score + config.previous_week_caregiver_bonus),
            "預估車程(分)": np.nan,
            "服務時段": f"{t.get('時間窗_開始')}-{t.get('時間窗_結束')}",
            "任務優先級": t.get("任務優先級", "一般"),
            "地點緯度": t.get("服務地點_緯度"),
            "地點經度": t.get("服務地點_經度"),
            "預估長照申報點數(營收)": fin.get("預估長照申報點數(營收)", np.nan),
            "預估居服員拆帳薪資": fin.get("預估居服員拆帳薪資", np.nan),
            "基礎分": config.base_score,
            "專長匹配加分": 0.0,
            "歷史首選加分": 0.0,
            "上一週同案加分": config.previous_week_caregiver_bonus,
            "連續性品質加分": 0.0,
            "滿意度調整": 0.0,
            "交通扣分": 0.0,
            "工作負荷扣分": 0.0,
            "是否歷史首選": False,
            "是否上一週同案居服員": True,
            "上一週居服員ID": cg_id,
            "排班來源": "上一週班表直接沿用",
            "原首選替換原因": "上一週固定班延續",
        })
    return pd.DataFrame(rows)


def _fast_basic_carry_check(task, cg, config):
    return _check_hard_constraints(task, cg, config, None, True)


def _fast_carry_day(day_tasks, periodic_tasks, df_cg, config, previous_exact_map):
    if periodic_tasks is None or periodic_tasks.empty or not previous_exact_map:
        return pd.DataFrame(), periodic_tasks.copy() if periodic_tasks is not None else pd.DataFrame(), []

    cg_lookup = df_cg.set_index("居服員ID").to_dict("index")
    ordered = periodic_tasks.copy()
    ordered["_start_dt"] = pd.to_datetime(ordered["時間窗_開始"], format="%H:%M", errors="coerce")
    ordered = ordered.sort_values(["_start_dt", "任務ID"], kind="stable")

    carried, failures = [], []
    occupied, hours = {}, {}

    for _, task in ordered.iterrows():
        task_id = task["任務ID"]
        key = _fast_weekly_key(task)
        cg_id = _text(previous_exact_map.get(key))

        if not cg_id:
            failures.append((task_id, "上一週無對應任務"))
            continue

        cg = cg_lookup.get(cg_id)
        if cg is None:
            failures.append((task_id, "上一週居服員目前不存在於員工資料"))
            continue

        reason = _fast_basic_carry_check(task, pd.Series(cg), config)
        if reason:
            failures.append((task_id, reason))
            continue

        try:
            start = datetime.strptime(str(task["時間窗_開始"]).strip(), "%H:%M")
            end = datetime.strptime(str(task["時間窗_結束"]).strip(), "%H:%M")
        except Exception:
            failures.append((task_id, "服務時段無法解析"))
            continue

        if any(start < old_end and end > old_start for old_start, old_end in occupied.get(cg_id, [])):
            failures.append((task_id, "與本週已沿用任務時段重疊"))
            continue

        dur = float(task.get("服務歷時(分鐘)", 0) or 0) / 60.0
        after = hours.get(cg_id, 0.0) + dur
        cap = cg.get("每日工時上限(小時)")
        if pd.notna(cap) and after > float(cap):
            failures.append((task_id, f"沿用後每日工時超過上限 {float(cap):g} 小時"))
            continue

        carried.append((task_id, cg_id))
        occupied.setdefault(cg_id, []).append((start, end))
        hours[cg_id] = after

    carried_ids = {x[0] for x in carried}
    pending = periodic_tasks[~periodic_tasks["任務ID"].isin(carried_ids)].copy()
    return _fast_carried_rows(carried, day_tasks, config), pending, failures



def run_monthly_batch_dispatch(
    tasks: pd.DataFrame,
    df_cg: pd.DataFrame,
    config: PipelineConfig,
    date_column: str = "日期",
    progress_callback=None,
    previous_week_map: dict | None = None,
    previous_week_assignment_map: dict | None = None,
    carry_forward: bool = True,
) -> dict:
    if tasks.empty:
        return {
            "daily_results": {},
            "df_result_all": pd.DataFrame(),
            "total_assigned_count": 0,
            "total_task_count": 0,
            "df_matches_all": pd.DataFrame(),
            "matching_diagnostics": [],
            "carried_count": 0,
            "ai_replanned_count": 0,
            "carry_failures": pd.DataFrame(),
            "latest_week_assignment_map": {},
            "latest_week_client_map": {},
        }

    work = tasks.copy()
    work["_dispatch_date"] = pd.to_datetime(work[date_column], errors="coerce")
    valid_dates = sorted(work["_dispatch_date"].dropna().dt.date.unique())
    if not valid_dates:
        raise ValueError("任務資料沒有有效日期。")

    range_start = min(valid_dates)
    daily_results = {}
    result_frames, match_frames, diagnostics = [], [], []
    carry_failures = []
    total_carried = 0
    total_ai_replanned = 0

    previous_exact_seed = dict(previous_week_assignment_map or {})
    previous_client_seed = dict(previous_week_map or {})
    week_results = {}

    df_cg_working = df_cg.copy()
    if "本次計畫已排時數" not in df_cg_working.columns:
        df_cg_working["本次計畫已排時數"] = 0.0
    if "本次計畫加權負荷時數" not in df_cg_working.columns:
        df_cg_working["本次計畫加權負荷時數"] = 0.0

    for day_idx, target_date in enumerate(valid_dates):
        if progress_callback:
            progress_callback(day_idx, len(valid_dates), target_date)

        day_tasks = work[work["_dispatch_date"].dt.date == target_date].drop(
            columns=["_dispatch_date"], errors="ignore"
        ).copy()
        if day_tasks.empty:
            continue

        week_index = (target_date - range_start).days // 7

        if week_index == 0:
            current_exact = previous_exact_seed
            current_client = previous_client_seed
        else:
            prev_frames = week_results.get(week_index - 1, [])
            if prev_frames:
                prev_week_df = pd.concat(prev_frames, ignore_index=True)
                current_exact = _fast_exact_map(prev_week_df)
                current_client = _client_primary_caregiver_map(prev_week_df)
            else:
                current_exact = {}
                current_client = {}

        periodic_tasks, adhoc_tasks = _split_periodic_tasks(day_tasks)
        if periodic_tasks is None:
            periodic_tasks = day_tasks.copy()
            adhoc_tasks = day_tasks.iloc[0:0].copy()

        if carry_forward and current_exact:
            carried_df, periodic_pending, failures = _fast_carry_day(
                day_tasks, periodic_tasks, df_cg_working, config, current_exact
            )
        else:
            carried_df = pd.DataFrame()
            periodic_pending = periodic_tasks.copy()
            failures = []

        for task_id, reason in failures:
            if reason != "上一週無對應任務":
                carry_failures.append({"日期": str(target_date), "任務ID": task_id, "原因": reason})

        total_carried += len(carried_df)

        extra_busy = {}
        used_hours = {}
        if not carried_df.empty:
            task_lookup = day_tasks.set_index("任務ID").to_dict("index")
            for _, r in carried_df.iterrows():
                t = task_lookup.get(r["任務ID"])
                if not t:
                    continue
                cg_id = r["派單居服員"]
                s = datetime.strptime(str(t["時間窗_開始"]).strip(), "%H:%M")
                e = datetime.strptime(str(t["時間窗_結束"]).strip(), "%H:%M")
                extra_busy.setdefault(cg_id, []).append(
                    (s, e, t.get("服務地點_緯度"), t.get("服務地點_經度"))
                )
                used_hours[cg_id] = used_hours.get(cg_id, 0.0) + float(
                    t.get("服務歷時(分鐘)", 0) or 0
                ) / 60.0

        df_cg_ai = df_cg_working.copy()
        if "今日已佔用工時(小時)" not in df_cg_ai.columns:
            df_cg_ai["今日已佔用工時(小時)"] = 0.0

        for cg_id, hrs in used_hours.items():
            mask = df_cg_ai["居服員ID"].astype(str) == str(cg_id)
            base = pd.to_numeric(
                df_cg_ai.loc[mask, "今日已佔用工時(小時)"], errors="coerce"
            ).fillna(0.0)
            df_cg_ai.loc[mask, "今日已佔用工時(小時)"] = base + hrs

        empty = {
            "df_valid": pd.DataFrame(),
            "status": None,
            "df_result": pd.DataFrame(),
            "assigned_count": 0,
            "df_matches": pd.DataFrame(),
            "matching_diagnostics": [],
        }

        periodic_result = empty.copy()
        if not periodic_pending.empty:
            periodic_matches = run_phase1_matching(
                periodic_pending, df_cg_ai, config, previous_week_map=current_client
            )
            periodic_result = run_phase2_optimization(
                periodic_matches,
                periodic_pending,
                df_cg_ai,
                config,
                extra_busy_blocks=extra_busy,
            )

        locked_after_periodic = pd.concat(
            [carried_df, periodic_result["df_result"]],
            ignore_index=True,
            sort=False,
        )

        adhoc_result = empty.copy()
        if adhoc_tasks is not None and not adhoc_tasks.empty:
            task_lookup = day_tasks.set_index("任務ID").to_dict("index")
            busy2, used_hours2 = {}, {}

            for _, r in locked_after_periodic.iterrows():
                t = task_lookup.get(r["任務ID"])
                if not t:
                    continue
                cg_id = r["派單居服員"]
                s = datetime.strptime(str(t["時間窗_開始"]).strip(), "%H:%M")
                e = datetime.strptime(str(t["時間窗_結束"]).strip(), "%H:%M")
                busy2.setdefault(cg_id, []).append(
                    (s, e, t.get("服務地點_緯度"), t.get("服務地點_經度"))
                )
                used_hours2[cg_id] = used_hours2.get(cg_id, 0.0) + float(
                    t.get("服務歷時(分鐘)", 0) or 0
                ) / 60.0

            df_cg_adhoc = df_cg_working.copy()
            if "今日已佔用工時(小時)" not in df_cg_adhoc.columns:
                df_cg_adhoc["今日已佔用工時(小時)"] = 0.0

            for cg_id, hrs in used_hours2.items():
                mask = df_cg_adhoc["居服員ID"].astype(str) == str(cg_id)
                base = pd.to_numeric(
                    df_cg_adhoc.loc[mask, "今日已佔用工時(小時)"], errors="coerce"
                ).fillna(0.0)
                df_cg_adhoc.loc[mask, "今日已佔用工時(小時)"] = base + hrs

            adhoc_matches = run_phase1_matching(
                adhoc_tasks, df_cg_adhoc, config, previous_week_map=current_client
            )
            adhoc_result = run_phase2_optimization(
                adhoc_matches,
                adhoc_tasks,
                df_cg_adhoc,
                config,
                extra_busy_blocks=busy2,
            )

        total_ai_replanned += (
            periodic_result["assigned_count"] + adhoc_result["assigned_count"]
        )

        df_day = pd.concat(
            [carried_df, periodic_result["df_result"], adhoc_result["df_result"]],
            ignore_index=True,
            sort=False,
        )

        if not df_day.empty:
            df_day["排班日期"] = target_date
            result_frames.append(df_day)
            week_results.setdefault(week_index, []).append(df_day.copy())

        day_matches = pd.concat(
            [periodic_result["df_matches"], adhoc_result["df_matches"]],
            ignore_index=True,
            sort=False,
        )

        daily_results[target_date] = {
            "df_result": df_day,
            "assigned_count": len(df_day),
            "task_count": len(day_tasks),
            "carried_count": len(carried_df),
            "ai_replanned_count": (
                periodic_result["assigned_count"] + adhoc_result["assigned_count"]
            ),
            "pending_for_ai_count": int(len(periodic_pending) + (len(adhoc_tasks) if adhoc_tasks is not None else 0)),
            "df_matches": day_matches,
            "matching_diagnostics": (
                periodic_result["matching_diagnostics"]
                + adhoc_result["matching_diagnostics"]
            ),
        }

        if not day_matches.empty:
            match_frames.append(day_matches)
        diagnostics.extend(daily_results[target_date]["matching_diagnostics"])

    if progress_callback:
        progress_callback(len(valid_dates), len(valid_dates), None)

    df_result_all = (
        pd.concat(result_frames, ignore_index=True)
        if result_frames else pd.DataFrame()
    )

    if not df_result_all.empty:
        df_result_all = calculate_schedule_travel(
            df_result_all, tasks, df_cg, config
        )

    latest_week_assignment_map = {}
    latest_week_client_map = {}
    if week_results:
        latest_idx = max(week_results)
        latest_df = pd.concat(week_results[latest_idx], ignore_index=True)
        latest_week_assignment_map = _fast_exact_map(latest_df)
        latest_week_client_map = _client_primary_caregiver_map(latest_df)

    return {
        "daily_results": daily_results,
        "df_result_all": df_result_all,
        "total_assigned_count": len(df_result_all),
        "total_task_count": len(tasks),
        "df_matches_all": (
            pd.concat(match_frames, ignore_index=True)
            if match_frames else pd.DataFrame()
        ),
        "matching_diagnostics": diagnostics,
        "carried_count": total_carried,
        "ai_replanned_count": total_ai_replanned,
        "carry_rate": (
            round(total_carried / max(len(tasks), 1) * 100.0, 1)
            if len(tasks) else 0.0
        ),
        "carry_failures": pd.DataFrame(carry_failures),
        "latest_week_assignment_map": latest_week_assignment_map,
        "latest_week_client_map": latest_week_client_map,
    }

def run_phase3_did(df_hist: pd.DataFrame) -> dict:
    required=['歷史媒合機制(Treatment)','案家滿意度(1-5)','不滿意導致提早結案(0/1)']
    if df_hist.empty or any(c not in df_hist for c in required):
        return {'available':False,'reason':'未提供歷史服務紀錄，無法估計滿意度差異或DiD效益',
                'ai_group':pd.DataFrame(),'human_group':pd.DataFrame(),
                **{k:np.nan for k in ['ai_sat','human_sat','ai_dropout','human_dropout','uplift_sat','uplift_dropout']}}
    ai_group = df_hist[df_hist["歷史媒合機制(Treatment)"] == 1]
    human_group = df_hist[df_hist["歷史媒合機制(Treatment)"] == 0]

    ai_sat = ai_group["案家滿意度(1-5)"].mean()
    human_sat = human_group["案家滿意度(1-5)"].mean()
    ai_dropout = ai_group["不滿意導致提早結案(0/1)"].mean()
    human_dropout = human_group["不滿意導致提早結案(0/1)"].mean()

    return {
        'available': not ai_group.empty and not human_group.empty,
        'reason': '僅兩組平均差，缺前後期不能視為DiD因果效益',
        "ai_group": ai_group,
        "human_group": human_group,
        "ai_sat": ai_sat,
        "human_sat": human_sat,
        "ai_dropout": ai_dropout,
        "human_dropout": human_dropout,
        "uplift_sat": ai_sat - human_sat,
        "uplift_dropout": human_dropout - ai_dropout,
    }

# ==========================================
# 居督人工覆寫稽覈日誌 (Human-in-the-Loop override audit log)
# ==========================================
OVERRIDE_LOG_PATH = os.path.join("output_results", "supervisor_override_log.csv")

OVERRIDE_LOG_COLUMNS = [
    "時間戳記", "任務ID", "案家ID", "AI推薦居服員ID", "居督指定居服員ID", "變更原因",
]

def save_override_log(
    task_id,
    client_id,
    ai_cg_id,
    supervisor_cg_id,
    reason: str,
    log_path: str = OVERRIDE_LOG_PATH,
) -> None:
    """將居督一筆人工覆寫紀錄以附加 (append) 方式寫入 CSV 稽覈日誌。

    每次呼叫寫入一列；檔案不存在時先建立目錄與標題列。
    """
    os.makedirs(os.path.dirname(log_path), exist_ok=True)
    entry = pd.DataFrame(
        [{
            "時間戳記": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "任務ID": task_id,
            "案家ID": client_id,
            "AI推薦居服員ID": ai_cg_id,
            "居督指定居服員ID": supervisor_cg_id,
            "變更原因": reason,
        }],
        columns=OVERRIDE_LOG_COLUMNS,
    )
    write_header = not os.path.exists(log_path)
    entry.to_csv(log_path, mode="a", header=write_header, index=False, encoding="utf-8-sig")

def load_override_log(log_path: str = OVERRIDE_LOG_PATH) -> pd.DataFrame:
    """讀取現有稽覈日誌；檔案不存在時回傳空的標準欄位 DataFrame。"""
    if not os.path.exists(log_path):
        return pd.DataFrame(columns=OVERRIDE_LOG_COLUMNS)
    return pd.read_csv(log_path, encoding="utf-8-sig")

def _privacy_mask_name(name):
    """姓名顯示最小化：僅保留第一個字，其餘統一遮罩為 ○。
    例如：王O明 / 王○明 / 王小明 -> 王○○。
    """
    text = _text(name)
    if not text:
        return ""
    compact = "".join(text.split())
    if len(compact) <= 1:
        return "○"
    return compact[0] + "○" * (len(compact) - 1)


def load_identity_mapping(source, kind, sheet_name=None):
    """讀取姓名對照，只回傳 ID→遮罩姓名 dict；原姓名不併入排班 DataFrame、不送外部 API。"""
    if kind not in ("caregiver", "client"):
        raise ValueError("對照類型必須為 caregiver 或 client")
    id_col, name_col = ("編號", "姓名") if kind == "caregiver" else ("案號", "個案姓名")
    aliases = IDENTITY_SHEET_ALIASES[kind]
    mappings = {}

    with pd.ExcelFile(source) as excel:
        if sheet_name is not None:
            candidates = [sheet_name] if sheet_name in excel.sheet_names else []
        else:
            candidates = [s for s in aliases if s in excel.sheet_names]

        if len(candidates) == 0:
            raise ValueError("找不到姓名對照工作表；可接受：" + "、".join(aliases))
        if len(candidates) > 1:
            raise ValueError("姓名對照工作表重複：" + "、".join(candidates))

        sheet = candidates[0]
        frame = pd.read_excel(excel, sheet_name=sheet, dtype=str)
        if not {id_col, name_col}.issubset(frame.columns):
            raise ValueError(f"工作表「{sheet}」需包含「{id_col}」「{name_col}」")

        # 對照 sheet 僅允許 ID + 姓名，其他欄位不進入記憶體映射。
        frame = frame[[id_col, name_col]].copy()
        for _, row in frame.iterrows():
            identifier, name = _text(row[id_col]), _text(row[name_col])
            if not identifier and not name:
                continue
            if not identifier or not name:
                raise ValueError(f"對照表「{sheet}」有空白 {id_col} 或 {name_col}，請補齊")
            masked_name = _privacy_mask_name(name)
            if identifier in mappings and mappings[identifier] != masked_name:
                raise ValueError(f"同一 ID {identifier} 對應不同姓名，請確認後再載入")
            # 個資最小化：mapping 僅保留遮罩後姓名，不保留原始／部分揭露姓名。
            mappings[identifier] = masked_name
    return mappings


def load_identity_mappings_from_workbook(source):
    """從同一份 6-sheet Excel 讀兩張姓名對照；只供本次 session 顯示／匯出使用。"""
    return {
        "clients": load_identity_mapping(source, "client"),
        "caregivers": load_identity_mapping(source, "caregiver"),
    }

def identity_display(frame, caregivers=None, clients=None):
    """只產生顯示／匯出副本，不更改用於計算或人工改派的 ID。"""
    if not isinstance(frame,pd.DataFrame):return frame
    caregivers,clients = caregivers or {},clients or {}
    out=frame.copy()
    caregiver_cols={'居服員ID','編號','派單居服員','原派居服員','新派居服員','最終派單居服員','AI建議居服員','歷史首選居服員ID','請假居服員'}
    client_cols={'案家ID','案號','內部案號'}
    for col in out.columns:
        mapping = caregivers if col in caregiver_cols else (clients if col in client_cols else None)
        if mapping is not None:
            out[col]=out[col].map(lambda v:f'{_text(v)}｜{mapping[_text(v)]}' if _text(v) in mapping else v)
        elif col in {'交通起點','交通終點','交通路段'} and clients:
            pattern=r'(?<![A-Za-z0-9])('+'|'.join(re.escape(k) for k in sorted(clients,key=len,reverse=True))+r')(?![A-Za-z0-9])'
            out[col]=out[col].map(lambda v:re.sub(pattern,lambda m:m[1]+'｜'+clients[m[1]],v) if isinstance(v,str) else v)
    return out
