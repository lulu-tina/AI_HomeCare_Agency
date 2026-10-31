"""長照居家照顧 AI 派單系統 - 網頁工作台

執行方式：

    streamlit run app.py

功能模組（側邊欄切換；每頁只呈現該任務需要的內容）：

    01 資料匯入與排班設定：Excel 匯入、排班期間、交通方式、排班規則及限制條件、
       服務碼組合檢查與資料表編輯。
    02 智慧排班作業：AI 自動排班、日曆班表、人工調整、個案與居服員媒合與候選人推薦。
    03 衝突與異常處理：班次重疊、轉場不足、未成功派班、條件不符的分類與明細；
       未派案家整案處理（替代方案）；臨時異動（新增、請假、改期、延遲）。
    04 排班結果與覆核：排班完成率、已排與未排班次、警示案件、最終班表與匯出、
       人工覆寫紀錄。
    05 人力與效益分析（主管）：工時、負荷分布、轉場時間、人力配置、服務產出與營運效益。

排班結果、交通與薪資等計算在每次畫面重跑時共用（不論目前在哪個模組），只有畫面
呈現依模組切換；設定類元件的值保存在 session，換頁不會遺失或清掉已完成的班表。

核心運算邏輯（適配度評分、OR-Tools 最佳化、財務點數試算、BA 代碼防呆、DiD 效益
回溯）皆定義於 caregiver_engine.py 等模組，本檔僅負責資料編輯介面與結果呈現，不重複
實作演算法；衝突與異常的分類整理在 issue_center.py。
"""

# ---- 版本一致性檢查：只覆蓋部分檔案會讓畫面與排班邏輯不一致（例如畫面是新的、引擎是舊的）。----
import pathlib as _pathlib
import streamlit as _guard_st
def _check_install():
    folder = _pathlib.Path(__file__).resolve().parent
    try:
        from build_info import CAREFLOW_BUILD, REQUIRED_MODULES
    except ImportError:
        return None, ['build_info.py（缺少，請把整個資料夾的檔案一起覆蓋）']
    stale = []
    for name in REQUIRED_MODULES:
        path = folder / (name + '.py')
        text = path.read_text(encoding='utf-8', errors='ignore') if path.exists() else ''
        if "CAREFLOW_BUILD = '%s'" % CAREFLOW_BUILD not in text:
            stale.append(name + '.py' + ('' if path.exists() else '（找不到）'))
    return CAREFLOW_BUILD, stale
CAREFLOW_BUILD, _stale_files = _check_install()
if _stale_files:
    _guard_st.set_page_config(page_title="CareFlow｜需要更新檔案", page_icon="⚠️", layout="centered")
    _guard_st.error('偵測到有檔案還是舊版，為避免排班結果不一致，已暫停啟動。')
    _guard_st.markdown('請把下列檔案換成新版（最簡單：**整個資料夾的 .py 檔全部覆蓋**）：\n\n' + '\n'.join('- `%s`' % n for n in _stale_files))
    _guard_st.stop()

from dataclasses import replace
from monthly_scheduler import run_fixed_period_dispatch
from pending_clients import (pending_client_summary, client_patterns, preview_client_assignment, plan_team, edited_pattern_table,
                             preview_with_edits, split_points, default_split_minutes, unit_minutes_from_master, pattern_token)
from pending_advice import apply_suggestion, empty_edits, edits_summary, change_log_rows, diagnose, build_plans
from address_geocoder import get_geocoder, GeoResult
from dispatch_message import parse_dispatch_message, redact_message
from temp_case import slots_to_frame, build_temp_tasks, search_slots, single_caregiver_plan
from eligibility import explain_slot, screen_caregivers, client_profile, simulate_worker, month_overtime_summary, worker_status_table
from source_fields import source_primary_map, normalize_clients
from agency_rules import payroll_tables, primary_guard, leave_overlap, consecutive_workdays_error
from osrm_client import ensure_ready
from schedule_output import build_schedule_output, schedule_excel_bytes, MAIN_COLUMNS
from dispatch_workflow import WorkflowStore, attach_blocks, blocked_reason, render_leave
from substitute_rank import RULE_TEXT as SUBSTITUTE_RULE_TEXT
from issue_center import unassigned_issues, assigned_overlaps, transition_issues, issue_summary, GROUP_NAMES as ISSUE_GROUP_NAMES


from datetime import datetime
import hashlib
import html
import io
import csv as csv_module
import math
import os
import pathlib
import re
import statistics

import matplotlib.font_manager as fm

import matplotlib.pyplot as plt

import pandas as pd

import seaborn as sns

import streamlit as st

from ortools.linear_solver import pywraplp

try:
    from calendar_view import apply_overrides_to_result, render_calendar_overview, render_task_override_picker
except ModuleNotFoundError as exc:
    if exc.name != 'calendar_view':
        raise
    # 保留原模組時使用完整月曆；未提供時仍可匯入、排班與檢視草案。
    def apply_overrides_to_result(frame, overrides):
        result=frame.copy()
        for task_id, override in overrides.items():
            cg_id=override.get('cg_id')
            if not result.empty and result['任務ID'].eq(task_id).any():
                if cg_id is None: result=result.loc[result['任務ID'].ne(task_id)].copy()
                else: result.loc[result['任務ID'].eq(task_id),'派單居服員']=cg_id
            elif cg_id is not None:
                result=pd.concat([result,pd.DataFrame([{'任務ID':task_id,'派單居服員':cg_id}])],ignore_index=True)
        return result
    def render_calendar_overview(result, tasks, caregivers, **kwargs):
        st.info('未提供 calendar_view.py，顯示草案明細；保留原檔即可恢復月曆與改派介面。')
        show_table(apply_overrides_to_result(result,kwargs.get('overrides',{})),width='stretch',hide_index=True)
    def render_task_override_picker(*args, **kwargs):
        st.info('月曆人工覆寫介面需原本的 calendar_view.py；本次未修改該模組。')

from caregiver_engine import (

    _check_hard_constraints,

    DEFAULT_EXCEL_PATH,
    google_key_is_configured,
    begin_travel_query_run,
    set_google_queries_enabled,
    load_agency_workbook,
    load_identity_mapping,
    load_identity_mappings_from_workbook,
    identity_display,
    merge_task_client_tables,

    PipelineConfig,
    service_price_map,
    calculate_task_revenue_and_salary,
    calculate_schedule_travel,
    summarize_schedule_travel,

    apply_service_duration,

    calc_travel_minutes,
    check_reassignment_conflict,
    find_insertion_candidates,

    get_weekday_name,

    load_override_log,

    rank_candidates_by_availability,

    run_monthly_batch_dispatch,

    run_phase1_matching,

    run_phase2_optimization,

    run_phase3_did,

    save_override_log,

    validate_ba_codes,

)

_caregiver_names = {}
_client_names = {}
_show_names = False
_export_names = False

def show_table(data, *args, **kwargs):
    display = identity_display(data, _caregiver_names, _client_names) if _show_names else data
    return st.dataframe(display, *args, **kwargs)

def download_button_quiet(label, data, file_name=None, mime=None, key=None, **kwargs):
    """下載按鈕：點下去只下載、不重跑整頁（Streamlit 1.43 以上；舊版自動退回一般按鈕）。"""
    try:
        return st.download_button(label, data, file_name, mime, key=key, on_click='ignore', **kwargs)
    except (TypeError, ValueError):
        return st.download_button(label, data, file_name, mime, key=key, **kwargs)

def caregiver_label(value):
    identifier = str(value).strip()
    return identifier + '｜' + _caregiver_names[identifier] if _show_names and identifier in _caregiver_names else identifier

def client_label(value):
    identifier = str(value).strip()
    return identifier + '｜' + _client_names[identifier] if _show_names and identifier in _client_names else identifier

def export_frame(data):
    return identity_display(data, _caregiver_names, _client_names) if _export_names else data

def _label_ids_plain(text):
    """把文字裡的居服員編號換成顯示用名稱（遵守遮罩設定）。"""
    out = str(text)
    if _show_names:
        for _cid in sorted(_caregiver_names, key=len, reverse=True):
            out = re.sub(r'(?<![A-Za-z0-9])' + re.escape(str(_cid)) + r'(?![A-Za-z0-9])', caregiver_label(_cid), out)
    return out

# ==========================================

# 跨平臺中文字型處理（避免 st.pyplot 圖表出現亂碼）

# ==========================================

# 直接隨 repo 附帶一份開源中文字型（Noto Sans TC, SIL Open Font License），

# 因為系統套件（packages.txt: fonts-noto-cjk）是否真的裝進 Streamlit Cloud 容器、

# 以及字型註冊名稱是否符合猜測，皆不受我們控制；自帶字型檔可確保任何部署環境

# 都一定找得到，不必依賴系統/雲端環境是否裝好中文字型。

_BUNDLED_FONT_PATH = pathlib.Path(__file__).parent / "assets" / "fonts" / "NotoSansTC-Regular.ttf"

def _frame_digest(frame):
    """Fast content fingerprint of a DataFrame (vectorised); used to skip unchanged recomputation."""
    if frame is None or len(frame) == 0:
        return '0'
    return hashlib.sha256(pd.util.hash_pandas_object(frame.astype(str), index=False).values.tobytes()
                          + str(list(frame.columns)).encode()).hexdigest()


def setup_chinese_font():

    # 依平臺猜測系統字型名稱不可靠：猜錯字型 matplotlib 不會報錯，只會靜默退回

    # DejaVu Sans（無中文字形，顯示為方框）。優先使用自帶字型，系統字型僅作備援。

    bundled_name = None

    if _BUNDLED_FONT_PATH.exists():

        fm.fontManager.addfont(str(_BUNDLED_FONT_PATH))

        bundled_name = fm.FontProperties(fname=str(_BUNDLED_FONT_PATH)).get_name()

    candidates = [

        "Microsoft JhengHei", "Microsoft YaHei", "SimHei",  # Windows

        "PingFang TC", "PingFang SC", "Heiti TC", "Arial Unicode MS",  # macOS

        "Noto Sans CJK TC", "Noto Sans CJK SC", "Noto Sans TC",  # Linux / Streamlit Cloud

        "WenQuanYi Micro Hei", "WenQuanYi Zen Hei",

    ]

    available = {f.name for f in fm.fontManager.ttflist}

    found = [name for name in candidates if name in available]

    ordered = ([bundled_name] if bundled_name else []) + found + ["DejaVu Sans"]

    seen = set()

    plt.rcParams["font.sans-serif"] = [n for n in ordered if not (n in seen or seen.add(n))]

    plt.rcParams["axes.unicode_minus"] = False

# ==========================================

# 色彩定義（來源：dataviz 色票，固定順序、非隨機挑色）

# ==========================================

BLUE = "#2a78d6"

ORANGE = "#eb6834"

MUTED = "#898781"

GOOD = "#0ca30c"

FIXED_REQUIRED_SHEETS = ["Caregiver_Profiles", "Client_Profiles", "Historical_Service_Logs","Service_Code"]

# 任務工作表名稱彈性相容：優先偵測月批次格式，其次退回現行單日格式。

TASKS_SHEET_CANDIDATES = ["Monthly_Pending_Tasks", "Today_Pending_Tasks"]

REQUIRED_COLUMNS = {

    "Caregiver_Profiles": [

        "居服員ID", "性別", "服務起點_經度(家)", "服務起點_緯度(家)", "核心專長證照",

        "每日工時上限(小時)", "當月累計服務時數(疲勞度)", "歷史滿意度均值",

        "具備重度移位體力(0/1)", "特殊排斥條件",

    ],

    "Client_Profiles": [

        "案家ID", "指定居服員性別", "服務地點_經度", "服務地點_緯度", "特殊照護需求",

        "案家環境特徵", "需重度移位協助(0/1)", "歷史首選居服員ID",

    ],

    "Historical_Service_Logs": [

        "居服員ID", "案家ID", "歷史媒合機制(Treatment)", "不滿意導致提早結案(0/1)", "案家滿意度(1-5)",

    ],

    "Service_Code": [

        "系統代碼",

        "CareFlow排班分鐘(暫定)",

        "是否納入CareFlow",

    ],

}

# 任務工作表的必要欄位為兩種格式共通的欄位；「今日既定行程」「可排班星期」等

# 各版本專屬欄位皆為選填——caregiver_engine 已能在欄位缺席時安全跳過對應檢查。

REQUIRED_TASKS_COLUMNS = ["任務ID", "案家ID", "時間窗_開始", "時間窗_結束", "任務優先級", "Service_Code_1", "Units_1",]

#移除："服務歷時(分鐘)"因為它之後由程式計算，不再是工程師必須提供的原始欄位。

#Service_Code_2、Units_2不用列為必填，因為有些任務只有一個服務碼。

st.set_page_config(page_title="CareFlow｜智慧居服排班", page_icon="🌿", layout="wide")
try:
    # 「載入示範案例」與換頁保存設定會先寫入 session 再畫元件；關掉 Streamlit 的重複設定提醒，避免畫面多出黃色警告。
    from streamlit import config as _st_config
    _st_config.set_option('global.disableWidgetStateDuplicationWarning', True)
except Exception:
    pass
from careflow_ui import (render_brand, section_title, render_summary, chips, flow_cards, file_status,
                         MODULES, MODULE_KEYS, module_label, module_header, module_nav_buttons, subhead,
                         go_to_module, empty_module, big_tabs,
                         reason_picker, reason_category, compose_reason, REASON_HINT, CHANGE_REASONS)

# ==========================================

# 功能模組導覽：依排班作業流程分成五個頁面，每頁只呈現該任務需要的內容

# ==========================================

M_IMPORT, M_SCHEDULE, M_ISSUES, M_REVIEW, M_ANALYTICS = MODULE_KEYS

defaults = PipelineConfig()
# 固定交通來源：機車=OSRM；大眾運輸=Google Routes TRANSIT。
set_google_queries_enabled(True)

_LOCAL_OSRM = 'http://127.0.0.1:5001'
_CLOUD_OSRM = 'https://careflow-osrm-motorcycle.onrender.com'
from osrm_client import is_remote as _osrm_is_remote
_configured_osrm = os.getenv('CAREFLOW_OSRM_MOTORCYCLE_HOST')
if not _configured_osrm:
    try:
        _configured_osrm = st.secrets.get('CAREFLOW_OSRM_MOTORCYCLE_HOST')
    except Exception:
        _configured_osrm = None
_configured_remote = bool(_configured_osrm) and _osrm_is_remote(str(_configured_osrm))
_configured_local = bool(_configured_osrm) and not _configured_remote
if _configured_remote:
    _CLOUD_OSRM = str(_configured_osrm).rstrip('/')
_OSRM_CHOICES = {'本機 Docker': _LOCAL_OSRM, '雲端 Render': _CLOUD_OSRM}

# 設定類元件只畫在「01 資料匯入與排班設定」；切換到其他模組時元件不會出現，
# Streamlit 會清掉沒畫出的元件狀態。先把值轉存為一般 session 狀態（官方建議做法），
# 換頁不會讓排班期間、交通或規則設定回到預設值，也不會因此清掉已完成的班表。
_today_tw = pd.Timestamp.now(tz='Asia/Taipei').date()
_SETTING_DEFAULTS = {
    'cf_module': M_IMPORT,
    'agency_period': '一個月',
    'agency_range_start': _today_tw.replace(day=1),
    'cfg_buffer_mins': float(defaults.buffer_mins),
    'cfg_travel_penalty_weight': float(defaults.travel_penalty_weight),
    'google_query_limit': 12000,
    'transit_fallback_enabled': True,
    'osrm_source_choice': '本機 Docker' if _configured_local else '雲端 Render',
}
_PERSISTED_KEYS = list(_SETTING_DEFAULTS) + ['agency_range_end', '_period_signature', 'show_identity_names', 'export_identity_mode']
if st.session_state.get('_google_limit_default_version') != 4:
    st.session_state['google_query_limit'] = 12000
    st.session_state['_google_limit_default_version'] = 4
for _key in _PERSISTED_KEYS:
    if _key in st.session_state:
        st.session_state[_key] = st.session_state[_key]
    elif _key in _SETTING_DEFAULTS:
        st.session_state[_key] = _SETTING_DEFAULTS[_key]
if st.session_state.get('cf_module') not in MODULE_KEYS:
    st.session_state['cf_module'] = M_IMPORT
_nav_request = st.session_state.pop('_cf_nav_request', None)
if _nav_request in MODULE_KEYS:
    st.session_state['cf_module'] = _nav_request
if st.session_state.get('osrm_source_choice') not in _OSRM_CHOICES:
    st.session_state['osrm_source_choice'] = _SETTING_DEFAULTS['osrm_source_choice']


def reset_policy_defaults():
    for key in ['cfg_buffer_mins','cfg_travel_penalty_weight','cfg_urgent_priority_bonus','cfg_continuity_bonus','cfg_previous_week_bonus','cfg_skill_bonus','cfg_sa10_price','cfg_hourly_wage','cfg_transition_wage','cfg_daily_cap','cfg_monthly_ot','google_query_limit']:
        st.session_state.pop(key, None)
    for key in ['cfg_buffer_mins', 'cfg_travel_penalty_weight', 'google_query_limit']:
        st.session_state[key] = _SETTING_DEFAULTS[key]


def _module_badge(key):
    if key == M_ISSUES and st.session_state.get('last_result') is not None:
        pending = st.session_state.get('pending_client_summary')
        count = 0 if pending is None else len(pending)
        return f'{count} 案' if count else None
    return None


with st.sidebar:
    st.markdown("### 🌿 CareFlow")
    st.caption("智慧居服排班工作台")
    st.markdown('<div class="cf-group">排班作業 · 居督</div>', unsafe_allow_html=True)
    _module = st.radio('功能模組', MODULE_KEYS, key='cf_module', label_visibility='collapsed',
                       format_func=lambda k: module_label(k, _module_badge(k)))
    st.caption('01→04 依排班流程操作；05 為主管管理分析，不影響班表。')
    st.divider()
    _sidebar_wake = st.container()
    _sidebar_status = st.container()
    _sidebar_privacy = st.container()

render_brand(CAREFLOW_BUILD, compact=_module != M_IMPORT)

institution_key = '中化'  # 固定居服員與請假記錄依此名稱保存；不再需要在畫面上設定。
sa10_price = defaults.sa10_hourly_price
hourly_wage = defaults.hourly_wage
transition_wage = defaults.transition_wage
max_daily_hours = defaults.max_daily_hours
max_monthly_ot = defaults.max_monthly_overtime_hours
st.session_state['careflow_travel_provider'] = 'hybrid'
provider = 'hybrid'
google_query_limit = int(st.session_state['google_query_limit'])
st.session_state['CAREFLOW_OSRM_MOTORCYCLE_HOST'] = _OSRM_CHOICES[st.session_state['osrm_source_choice']]
# 交通主機喚醒：打開網頁時就在背景先叫醒 Render（每個連線一次），等上傳完資料通常已經醒了。
import traffic_wake
_osrm_host = st.session_state['CAREFLOW_OSRM_MOTORCYCLE_HOST']
# 開網頁時自動叫醒；之後只要有人在操作、距上次確認超過 10 分鐘（stale）就在背景再叫一次，
# 讓主機在使用期間不會睡著。叫不醒（failed）只在開網頁時自動重試，避免一直打。
_wake_state = traffic_wake.status(_osrm_host)['status']
if _wake_state in ('idle', 'stale') or (_wake_state == 'failed' and not st.session_state.get('_osrm_auto_wake')):
    traffic_wake.start(_osrm_host)
st.session_state['_osrm_auto_wake'] = True
with _sidebar_wake:
    traffic_wake.panel(_osrm_host, 'sidebar', compact=True)
import caregiver_engine as _engine_mod
# 只設定「這個連線」的開關，不改全站共用的預設值（見 caregiver_engine.set_transit_fallback_enabled）。
_engine_mod.set_transit_fallback_enabled(bool(st.session_state['transit_fallback_enabled']))

previous_week_bonus = 0.0
continuity_bonus = defaults.preferred_caregiver_bonus
# 同案居服員安排固定採成功率最高的設定：主責＋最多 2 位備援；時段不自動挪動（要改時間請改 Excel）。
enable_backup = True
max_backups = 2
enable_support = True
time_flex = 0
buffer_mins = float(st.session_state['cfg_buffer_mins'])
travel_penalty_weight = float(st.session_state['cfg_travel_penalty_weight'])
urgent_priority_bonus = defaults.urgent_priority_bonus
skill_bonus = defaults.cert_bonus_dementia


def render_dispatch_settings():
    """01 模組：交通方式、排班規則及限制條件（值已在頁首讀入，這裡只負責顯示與調整）。"""
    subhead('交通方式與排班規則')
    _transport_col, _rule_col = st.columns(2, gap='large')
    with _transport_col.container(border=True):
        st.markdown('**🛵 交通方式**')
        st.caption('機車走自架 OSRM 機車路網；公車／捷運走 Google Routes TRANSIT，依員工資料的交通工具自動分流。')
        st.radio('機車路網', list(_OSRM_CHOICES), horizontal=True, key='osrm_source_choice',
                 help='預設雲端 Render。機構內部正式使用時，可用 Docker 自架機車路網（見 osrm_motorcycle 資料夾），案家座標不會傳到外部。')
        if st.session_state['osrm_source_choice'] == '本機 Docker' and not _configured_local:
            st.warning('本機 Docker 需要在自己電腦自架機車路網（安裝 Docker 後執行 osrm_motorcycle 資料夾內的腳本）。'
                       '雲端版或尚未架設時請選「雲端 Render」，否則交通時間會算不出來。')
        _host = _OSRM_CHOICES[st.session_state['osrm_source_choice']]
        st.caption(f"機車路網連線：{_host}（{'外部主機' if _osrm_is_remote(_host) else '本機'}）")
        traffic_wake.panel(_host, 'settings')
        st.checkbox('Google 查無公車／捷運路線時，改用估算（標示為估算）', key='transit_fallback_enabled',
                    help='短程（約 1～3 公里）常查不到公車／捷運路線，大眾運輸居服員會因此被判「無法取得交通時間」。'
                         '開啟後：直線 2 公里內以步行估算；更遠以機車路網時間×1.5＋候車 10 分鐘估算，並在逐段交通的「交通估算來源」標明。')
        with st.expander('進階：Google 查詢上限與連線檢查', expanded=False):
            st.number_input('每次排班 Google Transit elements 安全上限', min_value=100, max_value=100000, step=500, key='google_query_limit',
                            help='Google ComputeRouteMatrix 依 origins × destinations 計算 elements。系統會批次查詢並快取；這是安全上限，不代表一定會用到這麼多。')
            st.caption('此按鈕只確認連線，不驗證交通準確度；檢查使用公開街道座標。OSRM／Google 會將座標送往指定路網服務；大眾運輸不再使用本機直線概算。')
            if google_key_is_configured():
                st.caption('Google 金鑰：已讀取設定（未驗證 API 權限）。')
            else:
                st.warning('未讀到 Google 金鑰：本機請檢查 .streamlit/secrets.toml；雲端請檢查 Manage app → Settings → Secrets 的 GOOGLE_MAPS_API_KEY。')
    with _rule_col.container(border=True):
        st.markdown('**📏 排班規則及限制條件**')
        st.slider("任務間轉場緩衝時間（分鐘）", 0.0, 60.0, step=1.0, key="cfg_buffer_mins",
                  help="兩個服務任務之間，除實際交通時間外額外預留的緩衝時間，用於停車、步行、上下樓、門禁、服務紀錄及臨時延誤。")
        st.slider("車程扣分權重（分／分鐘）", 0.0, 5.0, step=0.1, key="cfg_travel_penalty_weight",
                  help="每一分鐘預估車程對適配度分數的扣分幅度。（目前車程扣分最高為 25 分，避免距離因素過度凌駕照護連續性。）")
        st.button('🔄 還原預設值', key='reset_policy_defaults', width='stretch', on_click=reset_policy_defaults)
        st.caption('優先排固定的歷史首選；首選排不進的服務才由同案備援承接；主責＋備援都排不進的單次服務，再另選合格居服員臨時支援（不列入固定備援）。請假代班不改固定人選。')
        st.markdown(
            f"- 可服務時段 08:00～21:00；服務加轉場總工時每日上限 {max_daily_hours:g} 小時，超過 8 小時計加班\n"
            f"- 月加班上限 {max_monthly_ot:g} 小時；不得連續上班超過規定天數\n"
            "- 必要證照、性別指定、可服務星期、請假、交通與休息皆為硬性條件\n"
            "- 隔週服務以生效日所在週起算，每隔一週安排")
        with st.expander('排班規則與計算說明', expanded=False):
            st.caption('先排每週固定班表；隔週服務與計畫生效異動另建週型。後續沿用週型，日期例外、公車／捷運及月加班另行核對。')
            st.caption('服務時間採週服務計畫的原服務分鐘；同案接續服務不另加轉場緩衝。不同案家轉場的交通與工作緩衝納入工時。')
            st.info('資格直接讀取員工「取得資格」。個案「是否失智」填是／否；BA08須有足部照護訓練。歷史首選填員工編號，匯入後直接沿用；年資由到職日計算作為加分。')
            st.caption('難度規則為機構參考：核定等級與失能／障礙程度取最大係數；身障類別只提供提示。')
            st.caption('SA10全部自費，一單位以350元估價；其餘項目依Service_code單價。服務薪資自動依機構時薪、轉場工資與加班制度估算。')

config = PipelineConfig(

    buffer_mins=buffer_mins,

    travel_penalty_weight=travel_penalty_weight,

    urgent_priority_bonus=urgent_priority_bonus,

    preferred_caregiver_bonus=continuity_bonus,
    previous_week_caregiver_bonus=previous_week_bonus,
    enable_backup_caregiver=enable_backup,
    max_backup_caregivers=int(max_backups),
    enable_support_caregiver=enable_support,
    time_flex_minutes=float(time_flex),

    cert_bonus_dementia=skill_bonus,

    cert_bonus_other=skill_bonus,

    institution_key=institution_key.strip() or '中化',
    hourly_wage=hourly_wage,
    transition_wage=transition_wage,
    sa10_hourly_price=sa10_price,
    max_daily_hours=max_daily_hours,
    max_monthly_overtime_hours=max_monthly_ot,

)

# ==========================================

# 資料載入（先於區塊①②，供 Phase 0 健檢與編輯區共用）

# ==========================================

input_mode = '機構資料（單一Excel／三個工作表）'

@st.cache_data(show_spinner="讀取 Excel 資料中...")

def _read_raw_sheets(file_or_path):

    xl = pd.ExcelFile(file_or_path)

    missing_sheets = [s for s in ['Caregiver_Profiles','Client_Profiles'] if s not in xl.sheet_names]

    if missing_sheets:

        raise ValueError(f"檔案缺少必要工作表：{'、'.join(missing_sheets)}")

    tasks_sheet_name = next((s for s in TASKS_SHEET_CANDIDATES if s in xl.sheet_names), None)

    if tasks_sheet_name is None:

        raise ValueError(f"檔案缺少任務工作表，需具備「{'」或「'.join(TASKS_SHEET_CANDIDATES)}」其中之一。")

    sheets = {name: pd.read_excel(xl, sheet_name=name) if name in xl.sheet_names else pd.DataFrame(columns=REQUIRED_COLUMNS[name]) for name in FIXED_REQUIRED_SHEETS}

    sheets["Tasks"] = pd.read_excel(xl, sheet_name=tasks_sheet_name)

    column_checks = dict(REQUIRED_COLUMNS)

    column_checks["Tasks"] = REQUIRED_TASKS_COLUMNS

    for name, required_cols in column_checks.items():
        if name in ('Historical_Service_Logs','Service_Code') and sheets[name].empty:
            continue

        missing_cols = [c for c in required_cols if c not in sheets[name].columns]

        if missing_cols:

            label = tasks_sheet_name if name == "Tasks" else name

            raise ValueError(f"工作表「{label}」缺少必要欄位：{'、'.join(missing_cols)}")

    return sheets, tasks_sheet_name


@st.cache_data(show_spinner="整理排班資料中...", max_entries=3)
def _load_agency_cached(workbook_bytes, start_date, end_date, daily_start, daily_end, daily_cap, confirmed, anchor):
    """Excel 展開成任務很耗時；檔案與日期沒變就直接沿用，每次操作畫面都不必重做。"""
    import io as _io
    return load_agency_workbook(_io.BytesIO(workbook_bytes), start_date, end_date, daily_start=daily_start, daily_end=daily_end,
                                daily_cap=daily_cap, availability_confirmed=confirmed, biweekly_anchor_confirmed=anchor)


# ---------- 自動存檔（第一階段）：重新整理或隔天打開可以還原整份作業進度 ----------
import session_store


def _restore_snapshot(file_hash):
    """按鈕回呼（畫面重跑前執行）：還原上次的排班結果、人工修改、Excel 與設定。"""
    try:
        loaded = session_store.load(institution_key, file_hash)
    except Exception as exc:
        st.session_state['_restore_error'] = '還原失敗：' + str(exc)
        return
    if not loaded:
        st.session_state['_restore_error'] = '找不到這份存檔，可能已被較新的進度取代。'
        return
    state, file_name, file_bytes, source = loaded
    for key, value in state.items():
        st.session_state[key] = value
    if file_bytes:
        st.session_state['_agency_upload'] = {'name': file_name, 'bytes': file_bytes}
    st.session_state['_data_source_key'] = source        # 讓「換檔案就清空」的檢查認得這份資料
    st.session_state['_restore_offer_done'] = True
    st.session_state['_autosave_token'] = None
    st.session_state['_restore_notice'] = f'已還原 {file_name} 的作業進度。'
    if st.session_state.get('last_result') is not None and st.session_state.get('cf_module') == M_IMPORT:
        st.session_state['cf_module'] = M_SCHEDULE


def _skip_restore():
    st.session_state['_restore_offer_done'] = True


def _restore_offer(found, key):
    """顯示「還原上次進度」卡片。"""
    with st.container(border=True):
        st.markdown(f"**💾 有上次未完成的排班作業**　·　{html.escape(str(found.get('file_name') or 'Excel'))}　·　"
                    f"最後存檔 {found['saved_at']}")
        if found.get('summary'):
            st.caption(found['summary'])
        _r1, _r2, _ = st.columns([1.3, 1, 2])
        _r1.button('↩️ 還原上次進度', type='primary', key=f'restore_{key}', on_click=_restore_snapshot,
                   args=(found['file'],), width='stretch')
        _r2.button('不用，重新開始', key=f'skip_restore_{key}', on_click=_skip_restore, width='stretch')
        st.caption('還原內容：排班結果、人工改派與調整紀錄、臨時插單、整案調整、排班期間與設定。請假與固定人選本來就會保存。')


if st.session_state.get('_restore_error'):
    st.error(st.session_state.pop('_restore_error'))
if st.session_state.get('_restore_notice'):
    st.success('💾 ' + st.session_state.pop('_restore_notice'))


class _StoredWorkbook(io.BytesIO):
    """上傳的 Excel 轉存在 session：切換到其他模組時上傳元件不會畫出，檔案仍可繼續使用。"""
    def __init__(self, data, name):
        super().__init__(data)
        self.name = name


def _remove_stored_workbook():
    st.session_state.pop('_agency_upload', None)
    st.session_state['_agency_upload_nonce'] = st.session_state.get('_agency_upload_nonce', 0) + 1


def _stored_workbook():
    stored = st.session_state.get('_agency_upload')
    return _StoredWorkbook(stored['bytes'], stored['name']) if stored else None


try:
    if input_mode == '機構資料（單一Excel／三個工作表）':
        local_dir=pathlib.Path(os.getenv('CAREFLOW_AGENCY_DIR','./00_DB'))
        local_file=local_dir/'居服機構資料_三工作表.xlsx'
        _has_default = local_file.exists()
        if _module == M_IMPORT:
            module_header(M_IMPORT)
            subhead('匯入排班資料')
            _stored_now = st.session_state.get('_agency_upload')
            _upload_key = 'agency_workbook' + (f"_{st.session_state['_agency_upload_nonce']}" if st.session_state.get('_agency_upload_nonce') else '')
            if _stored_now is None:
                if st.session_state.get('last_result') is None:
                    flow_cards(1)
                st.markdown('<div class="cf-upload-card"><h3>📂 上傳排班用 Excel</h3><p>拖曳檔案到下方，或按「Browse files」選擇；檔案只在本次作業中使用。</p></div>', unsafe_allow_html=True)
                agency_file=st.file_uploader('選擇 Excel 檔（.xlsx）', type=['xlsx'], key=_upload_key, label_visibility='collapsed')
            else:
                _swap_col, _remove_col = st.columns([4, 1], vertical_alignment='center')
                with _swap_col.expander(f"📂 目前使用：{_stored_now['name']}（要換檔案請展開）", expanded=False):
                    agency_file=st.file_uploader('更換 Excel 檔（.xlsx）', type=['xlsx'], key=_upload_key, label_visibility='collapsed')
                _remove_col.button('移除檔案', key='remove_agency_workbook', on_click=_remove_stored_workbook, width='stretch')
            if agency_file is not None:
                _new_bytes = agency_file.getvalue()
                if _stored_now is None or _stored_now['bytes'] != _new_bytes or _stored_now['name'] != agency_file.name:
                    st.session_state['_agency_upload'] = {'name': agency_file.name, 'bytes': _new_bytes}
        agency_source=_stored_workbook() or (local_file if _has_default else None)
        if agency_source is None:
            try:
                _latest = None if st.session_state.get('_restore_offer_done') else session_store.info(institution_key)
            except Exception:
                _latest = None
            if _latest:
                _restore_offer(_latest, 'nofile')
            if _module == M_IMPORT:
                with st.expander('Excel 需要哪些工作表？'):
                    st.write('請包含：員工資料、個案資料、週服務計畫、Service_code、案家姓名對照、居服員姓名對照。')
            else:
                module_header(_module)
                empty_module(_module, '尚未匯入排班資料，請先到「01 資料匯入與排班設定」上傳機構 Excel。',
                             target=M_IMPORT, button='前往「01 資料匯入與排班設定」')
            with _sidebar_status:
                st.markdown('<div class="cf-status">尚未匯入資料</div>', unsafe_allow_html=True)
            st.stop()
        # 排班期間：值保存在 session（換頁不會遺失）；元件只在 01 模組顯示。
        period = st.session_state.get('agency_period') or '一個月'
        range_start = st.session_state.get('agency_range_start') or _SETTING_DEFAULTS['agency_range_start']
        if _module == M_IMPORT:
            _period_box = st.container(border=True)
            period = _period_box.selectbox('排班範圍', ['一個月', '一年', '自訂'], key='agency_period')
            day1,day2=_period_box.columns(2)
            range_start = day1.date_input('排班開始日', value=None, key='agency_range_start') or _SETTING_DEFAULTS['agency_range_start']
        automatic_end=(pd.Timestamp(range_start)+pd.DateOffset(years=1) if period=='一年' else pd.Timestamp(range_start)+pd.DateOffset(months=1))-pd.Timedelta(days=1)
        period_signature=(period,str(range_start))
        if st.session_state.get('_period_signature')!=period_signature:
            st.session_state['agency_range_end']=automatic_end.date()
            st.session_state['_period_signature']=period_signature
        if _module == M_IMPORT:
            range_end=day2.date_input('排班結束日',value=None,key='agency_range_end',disabled=period!='自訂')
        else:
            range_end=st.session_state.get('agency_range_end')
        if range_end is None or range_end < range_start or (range_end-range_start).days>365:
            if _module != M_IMPORT:
                module_header(_module)
            st.error('日期範圍需為1～366天，結束日不得早於開始日。' + ('' if _module == M_IMPORT else '請到「01 資料匯入與排班設定」修正排班期間。'))
            st.stop()
        available_start=pd.Timestamp('08:00').time()
        available_end=pd.Timestamp('21:00').time()
        cap=config.max_daily_hours
        confirmed=True
        anchor=True
        if _module == M_IMPORT:
            _period_box.caption(f'預設可服務時段08:00～21:00，服務加轉場總工時每日上限{cap:g}小時；超過8小時計加班。休息日與請假排除，可於員工表逐人調整。')
            _period_box.caption('隔週服務自動以生效日所在週起算，每隔一週安排；不早於生效日，並配合原服務星期與到期日。')
        _agency_bytes = agency_source.getvalue() if hasattr(agency_source,'getvalue') else pathlib.Path(agency_source).read_bytes()
        sheets=_load_agency_cached(_agency_bytes,range_start,range_end,available_start.strftime('%H:%M'),
            available_end.strftime('%H:%M'),cap,confirmed,anchor)
        # 機構資料固定採用的排班參數（原本寫在「排班規則與計算說明」展開區內，換頁時也必須套用）。
        config.plan_load_mode='hours'
        config.fatigue_reference_hours=40.0
        config.fatigue_weight=10.0
        timing_mode='機構原計畫分鐘'
        sheets['Tasks']['使用Service_code重算']=False
        sheets['Import_Issues']=sheets['Import_Issues'].loc[sheets['Import_Issues']['資料表'].ne('Service_code')].copy()
        tasks_sheet_name='機構週服務展開任務'
        digest=hashlib.sha256()
        digest.update(_agency_bytes)
        digest.update(str((range_start,range_end,available_start,available_end,cap,confirmed,anchor,timing_mode)).encode())
        digest.update(sheets['Service_Code'].to_json(force_ascii=False).encode())
        source_key='agency::'+digest.hexdigest()
        _source_name = getattr(agency_source,'name',None) or pathlib.Path(str(agency_source)).name
        if _module == M_IMPORT:
            file_status(_source_name, len(sheets['Caregiver_Profiles']), len(sheets['Client_Profiles']), len(sheets['Tasks']))
            st.button('✅ 資料讀取完成，前往「02 智慧排班作業」開始排班 →', key='import_go_schedule',
                      on_click=go_to_module, args=(M_SCHEDULE,), width='stretch')
            if not sheets['Import_Issues'].empty:
                st.warning('以下資料需補正；缺鍵、無效服務列不產生任務，重疊服務任務保留但暫不派單。')
                show_table(sheets['Import_Issues'],hide_index=True,width='stretch')
        with _sidebar_status:
            st.markdown(f'<div class="cf-status">📄 <b>{html.escape(str(_source_name))}</b><br>'
                        f'📅 {range_start:%Y/%m/%d} ～ {range_end:%Y/%m/%d}<br>'
                        f'👥 {len(sheets["Caregiver_Profiles"]):,} 位居服員 · {len(sheets["Client_Profiles"]):,} 位個案<br>'
                        f'🗓️ {len(sheets["Tasks"]):,} 筆服務</div>', unsafe_allow_html=True)
except Exception as exc:
    st.error('資料匯入失敗：'+str(exc))
    st.stop()

# 個資保護：姓名對照直接從同一份 6-sheet Excel 載入。
# 姓名只存在於本次 Streamlit session 的顯示 mapping，不併入排班資料表、不寫入 Google/OSRM request。
try:
    identity_maps = load_identity_mappings_from_workbook(agency_source)
    _client_names = identity_maps["clients"]
    _caregiver_names = identity_maps["caregivers"]
except Exception as exc:
    _client_names, _caregiver_names = {}, {}
    if _module == M_IMPORT:
        st.warning("姓名對照未載入，系統仍可只用代碼排班：" + str(exc))

with _sidebar_privacy.expander("🔒 個資顯示設定", expanded=False):
    have_names = bool(_client_names or _caregiver_names)
    st.caption(
        "姓名匯入後會立即最小化遮罩（例如王O明→王○○），只供本頁顯示／選擇性匯出；排班、OR-Tools、Google Routes、OSRM 與稽核 log "
        "均使用 ID。Google Routes 僅收到必要的座標、時間與交通模式。"
    )
    if have_names and st.session_state.get('_identity_display_version') != 2:
        st.session_state['show_identity_names'] = True
        st.session_state['_identity_display_version'] = 2
    _show_names = st.checkbox(
        "畫面顯示遮罩姓名＋ID",
        key="show_identity_names",
        disabled=not have_names,
        help="已載入姓名對照時預設開啟，方便辨識案家與居服員；可隨時關閉。",
    ) and have_names
    _export_names = (
        st.radio(
            "匯出內容",
            ["代碼版", "遮罩姓名版（ID＋王○○）"],
            index=0,
            horizontal=True,
            key="export_identity_mode",
            disabled=not have_names,
        ) == "遮罩姓名版（ID＋王○○）"
        and have_names
    )
    if have_names:
        st.caption(
            f"姓名對照已隔離並遮罩：{len(_client_names)} 個案、{len(_caregiver_names)} 位居服員。"
            "畫面可顯示遮罩姓名；匯出仍可選擇代碼版或遮罩姓名版。"
        )

# 僅在資料來源變更時（首次載入／換檔案／原檔被覆寫）重設編輯區，避免使用者的編輯內容被覆蓋

if st.session_state.get("_data_source_key") != source_key:

    st.session_state["_data_source_key"] = source_key
    for stale in ['last_result','overrides','dynamic_insertions']:
        st.session_state.pop(stale,None)
    for key in list(st.session_state):
        if key.startswith('editor_'):
            del st.session_state[key]

    st.session_state["edit_cg"] = sheets["Caregiver_Profiles"].copy()

    st.session_state["edit_cl"] = sheets["Client_Profiles"].copy()

    st.session_state["edit_tasks"] = sheets["Tasks"].copy()

    st.session_state["data_hist"] = sheets["Historical_Service_Logs"].copy()

    st.session_state["edit_service_code"] = (

    sheets["Service_Code"].copy()

    )

    st.session_state["tasks_sheet_name"] = tasks_sheet_name

_file_hash = session_store.file_key(_agency_bytes)
if st.session_state.get('last_result') is not None:
    st.session_state['_restore_offer_done'] = True
elif not st.session_state.get('_restore_offer_done'):
    try:
        _found = session_store.info(institution_key, _file_hash)
    except Exception:
        _found = None
    if _found:
        _restore_offer(_found, 'file')

# ==========================================

# 區塊①：Phase 0 申報法規防呆健檢

# ==========================================

if _module == M_IMPORT:
    render_dispatch_settings()
    subhead('確認排班資料')
    st.caption('需要調整時，展開下方資料表直接編輯；資料或規則修改後，需回到「02 智慧排班作業」重新排班。')
    with st.expander('服務碼組合檢查與提醒', expanded=False):

        st.caption(

            "依目前資料提示程式既有服務碼組合規則；規則版本未核定，需由機構複核。"

            "本健檢僅檢核與提示、不會排除任何任務，亦不影響「02 智慧排班作業」的派單運算。"

        )

        try:

            _ba_key = (_frame_digest(st.session_state["edit_tasks"]), _frame_digest(st.session_state["edit_cl"]))
            if st.session_state.get('_ba_preview_key') != _ba_key:
                st.session_state['_ba_preview'] = validate_ba_codes(merge_task_client_tables(st.session_state["edit_tasks"], st.session_state["edit_cl"]))
                st.session_state['_ba_preview_key'] = _ba_key
            _validated_preview = st.session_state['_ba_preview']

            _total_tasks_preview = len(_validated_preview)

            _violation_count = int(_validated_preview["含違規代碼"].sum())

            _qualified_count = _total_tasks_preview - _violation_count

            h1, h2, h3 = st.columns(3)

            h1.metric("總任務數", f"{_total_tasks_preview}")

            h2.metric("沒有組合提醒的任務數", f"{_qualified_count}")

            h3.metric("⚠️ 需要檢查服務碼組合的任務數", f"{_violation_count}")

            if _violation_count > 0:

                st.warning(

                    f"偵測到 {_violation_count} 筆任務的 BA 服務代碼併報疑似違反長照給付支付基準規則，"

                    "建議於下方「檢視／編輯」資料表修正服務代碼，或由居督複核後仍可略過警告繼續派單。"

                )

                show_table(

                    _validated_preview.loc[_validated_preview["含違規代碼"], ["任務ID", "案家ID", "BA代碼檢核異常"]],

                    width="stretch",

                    hide_index=True,

                )

            else:

                st.success("目前未命中程式既有服務碼提示；不代表已完成申報合規審核。")

        except Exception as e:

            st.info(f"暫無法執行 BA 代碼健檢（請確認任務與案家資料的「案家ID」欄位可正常對應）：{e}")

    # ==========================================

    # 區塊②：Excel 載入與現況檢視／動態編輯

    # ==========================================

    with st.expander("檢視／編輯居服員、案家與服務任務", expanded=False):

        def render_editable_sheet(session_key: str):

            with st.expander("➕ 新增欄位"):

                c1, c2 = st.columns([3, 1])

                new_col_name = c1.text_input("欄位名稱", key=f"newcol_input_{session_key}", label_visibility="collapsed",

                                              placeholder="輸入新欄位名稱")

                if c2.button("新增欄位", key=f"newcol_btn_{session_key}", width="stretch"):

                    df_current = st.session_state[session_key]

                    if not new_col_name:

                        st.warning("請先輸入欄位名稱。")

                    elif new_col_name in df_current.columns:

                        st.warning(f"欄位「{new_col_name}」已存在。")

                    else:

                        df_current[new_col_name] = None

                        st.session_state[session_key] = df_current

                        st.rerun()

            df = st.session_state[session_key]
            if session_key == 'edit_cg' and '取得資格' in df:
                from care_policy import certificate_names
                df=df.copy()
                df['核心專長證照']=df.apply(lambda row:'、'.join(sorted(certificate_names(row))),axis=1)

            transport_columns = {}
            if session_key == 'edit_cg':
                transport_columns['常用交通工具'] = st.column_config.SelectboxColumn('常用交通工具', options=['機車','大眾運輸'], required=True, help='大眾運輸代表公車／捷運。')
                st.caption('資格直接採「取得資格」欄位；移除失智訓練文字後，該員工將不能承接已標記失智的個案。衍生的核心專長欄位只供檢視。')
            if session_key == 'edit_cl':
                if '是否失智' in df:
                    transport_columns['是否失智']=st.column_config.SelectboxColumn('是否失智',options=['是','否'],required=True)
                    transport_columns['失智症個案(0/1)']=None
                if '歷史首選' in df:
                    transport_columns['歷史首選居服員ID']=None
            edited = st.data_editor(

                df,
                column_config=transport_columns,
                disabled=['核心專長證照'] if session_key == 'edit_cg' else [],

                num_rows="dynamic",

                width="stretch",

                key=f"editor_{session_key}_{len(df.columns)}",

            )

            st.session_state[session_key] = edited

        tab_cg, tab_cl, tab_tasks, tab_service_code = st.tabs([

            "居服員",

            "案家",

            "服務任務",

             "服務碼",

        ])

        with tab_cg:

            st.caption("可直接編輯儲存格、新增／刪除列（勾選列號後按 Delete），或透過「新增欄位」新增自訂欄位。")

            render_editable_sheet("edit_cg")

        with tab_cl:

            st.caption("可直接編輯儲存格、新增／刪除列，或透過「新增欄位」新增自訂欄位。")

            render_editable_sheet("edit_cl")

        with tab_tasks:

            st.caption("可直接編輯儲存格、新增／刪除列，或透過「新增欄位」新增自訂欄位。")

            render_editable_sheet("edit_tasks")

        with tab_service_code:

            st.caption('機構資料採已提供服務分鐘，不用未核定的Master覆蓋。此處保留服務碼清單；原版資料仍可依Master計算。')

            render_editable_sheet("edit_service_code")
    module_nav_buttons(M_IMPORT)


def build_previous_week_assignment_map_from_csv(uploaded_file):
    if uploaded_file is None:
        return {}, {}, {}
    df = pd.read_csv(uploaded_file)
    required = {"任務ID", "派單居服員"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError("上一週班表缺少必要欄位：" + "、".join(sorted(missing)))
    df = df.copy()
    df["任務ID"] = df["任務ID"].astype(str).str.strip()
    df["派單居服員"] = df["派單居服員"].astype(str).str.strip()
    df = df[
        (df["任務ID"] != "")
        & (df["派單居服員"] != "")
        & (df["派單居服員"] != "未指派")
    ].copy()
    if df.empty:
        raise ValueError("上一週班表沒有可用的已指派資料。")
    df["_key"] = df["任務ID"].str.replace(r"_\d{8}$", "", regex=True)
    df = df.drop_duplicates("_key", keep="last")
    exact_map = dict(zip(df["_key"], df["派單居服員"]))
    client_map = {}
    if "案家ID" in df.columns:
        tmp = df[["案家ID", "派單居服員"]].dropna().copy()
        tmp["案家ID"] = tmp["案家ID"].astype(str).str.strip()
        tmp = tmp[tmp["案家ID"] != ""]
        if not tmp.empty:
            counts = (
                tmp.groupby(["案家ID", "派單居服員"], as_index=False)
                .size()
                .sort_values(
                    ["案家ID", "size", "派單居服員"],
                    ascending=[True, False, True],
                    kind="stable",
                )
                .drop_duplicates("案家ID", keep="first")
            )
            client_map = dict(
                zip(counts["案家ID"].astype(str), counts["派單居服員"].astype(str))
            )
    return exact_map, client_map, {
        "rows": len(df),
        "caregivers": df["派單居服員"].nunique(),
    }


# 區塊③：一鍵執行與成果儀錶板

# ==========================================

# 區塊③（02 智慧排班作業）：一鍵執行；計算全模組共用，畫面只在 02 模組呈現
st.session_state['edit_cl'] = normalize_clients(st.session_state['edit_cl'], set(st.session_state['edit_cg']['居服員ID'].astype(str)))
def _current_input_signature():
    signature = hashlib.sha256()
    for _key in ['edit_cg','edit_cl','edit_tasks','edit_service_code']:
        signature.update(st.session_state[_key].to_json(date_format='iso',default_handler=str).encode())
    config.service_unit_prices = service_price_map(st.session_state['edit_service_code'])
    signature.update(repr(replace(config, hourly_wage=0.0, transition_wage=0.0, sa10_hourly_price=0.0,
        client_primary_map={}, primary_profiles={}, work_day_minutes={})).encode())
    signature.update(st.session_state.get('careflow_travel_provider','local').encode())
    signature.update(repr(sorted({**__import__('caregiver_engine').TRANSIT_FALLBACK,
                                  'enabled': __import__('caregiver_engine').transit_fallback_enabled()}.items())).encode())
    signature.update(st.session_state.get('CAREFLOW_OSRM_MOTORCYCLE_HOST',os.getenv('CAREFLOW_OSRM_MOTORCYCLE_HOST','https://careflow-osrm-motorcycle.onrender.com')).encode())
    return signature.hexdigest()

_live_signature = _current_input_signature()
if 'last_result' in st.session_state and st.session_state['last_result'].get('input_signature') != _live_signature:
    st.warning('資料或排班參數已修改，請重新執行；舊派單結果已清除。')
    st.session_state.pop('last_result',None)
    st.session_state.pop('overrides',None)


_workflow_store = WorkflowStore(config.institution_key)
config.client_primary_map = source_primary_map(st.session_state['edit_cl'],
    set(st.session_state['edit_cg']['居服員ID'].astype(str)), _workflow_store.primary_map())
config.primary_profiles = {str(r['居服員ID']):r.to_dict() for _,r in attach_blocks(st.session_state['edit_cg'],_workflow_store).iterrows()}
if _module == M_SCHEDULE:
    module_header(M_SCHEDULE)
    subhead('AI 自動排班')
    st.caption(f'已讀取 {len(config.client_primary_map)} 個案家的首選；Excel「歷史首選」優先，直接排班即可。')

_tasks_preview_df = st.session_state["edit_tasks"]

has_date_column = '日期' in _tasks_preview_df.columns
new_dispatch_completed = False

if _module == M_SCHEDULE:
    if not has_date_column:
        st.info("目前任務沒有日期欄位，使用一般 AI 派單。")

        if st.button("開始 AI 排班", type="primary", width="stretch"):
            begin_travel_query_run(google_query_limit)

            df_cg = attach_blocks(st.session_state["edit_cg"], _workflow_store)
            df_cl = st.session_state["edit_cl"].copy()
            df_hist = st.session_state["data_hist"].copy()
            df_tasks = st.session_state["edit_tasks"].copy()
            df_service_code = st.session_state["edit_service_code"].copy()

            try:
                df_tasks = apply_service_duration(df_tasks, df_service_code)
                tasks = merge_task_client_tables(df_tasks, df_cl)
                df_matches = run_phase1_matching(tasks, df_cg, config)
                phase2 = run_phase2_optimization(df_matches, tasks, df_cg, config)
                did = run_phase3_did(df_hist)
            except Exception as e:
                st.error(f"⚠️ 派單運算過程發生錯誤：{e}")
                st.stop()

            st.session_state["last_result"] = {
                "input_signature": _live_signature,
                "df_tasks": df_tasks,
                "df_tasks_all": df_tasks.copy(),
                "df_cg": df_cg,
                "df_cl": df_cl,
                "df_matches": df_matches,
                "phase2": phase2,
                "did": did,
                "schedule_failed_dates": [],
                "weekly_workflow": False,
            }
            st.session_state["overrides"] = {}
            new_dispatch_completed = True

    else:
        all_dates_preview = pd.to_datetime(_tasks_preview_df['日期'], errors='coerce').dropna()
        if all_dates_preview.empty:
            st.warning('目前沒有有效日期可排班。')
        else:
            st.write(f'本次期間：{range_start} ～ {range_end}，共 {len(_tasks_preview_df)} 筆服務。')
            st.caption('先依歷史首選排固定週型，再延展至期間結束。隔週與計畫異動分別建立週型；請假、資格日期、月加班及公車／捷運按實際日期核對。')
            if traffic_wake.is_ready(_osrm_host):
                st.session_state.pop('_schedule_osrm_error', None)     # 主機已醒，舊的失敗說明不再顯示
            traffic_wake.panel(_osrm_host, 'schedule')
            if st.session_state.get('_schedule_osrm_error'):
                st.error('⚠️ ' + st.session_state['_schedule_osrm_error'])
                with st.expander('技術細節（給工程師）'):
                    st.code(st.session_state.get('_schedule_osrm_detail', ''))
            if st.button('開始排班 · 全期間', type='primary', width='stretch', key='run_full_period'):
                st.session_state.pop('_schedule_osrm_error', None)
                st.session_state.pop('_schedule_osrm_detail', None)
                begin_travel_query_run(google_query_limit)
                df_cg = attach_blocks(st.session_state['edit_cg'], _workflow_store)
                df_cl = st.session_state['edit_cl'].copy()
                df_tasks_all = apply_service_duration(st.session_state['edit_tasks'].copy(), st.session_state['edit_service_code'].copy())
                progress = st.progress(0.0, text='建立案家固定人選…')
                try:
                    _routine_signature = hashlib.sha256((df_cg.to_json(date_format='iso',default_handler=str)+
                        repr(replace(config,client_primary_map={},primary_profiles={},work_day_minutes={},dispatch_cache={}))+
                        st.session_state.get('careflow_travel_provider','local')+
                        st.session_state.get('CAREFLOW_OSRM_MOTORCYCLE_HOST',os.getenv('CAREFLOW_OSRM_MOTORCYCLE_HOST','https://careflow-osrm-motorcycle.onrender.com'))).encode()).hexdigest()
                    _routine_cache = st.session_state.get('_routine_templates',{})
                    # 連續上班天數硬性規則（AI 絕對不可排到連續第7天）需要跨月回溯：讀本期間開始日
                    # 往前10天、本機構已確定的上班日當種子，傳入排班引擎逐日累計檢查。
                    _consecutive_seed = _workflow_store.recent_workdays(range_start, lookback=10)
                    period_result = run_fixed_period_dispatch(df_tasks_all, df_cg, df_cl, config,
                        primary_map=config.client_primary_map, progress=lambda value,text: progress.progress(value,text=text),
                        routine_templates=_routine_cache.get('templates') if _routine_cache.get('signature')==_routine_signature else None,
                        worker_workday_seed=_consecutive_seed)
                    st.session_state['_routine_templates']={'signature':_routine_signature,'templates':period_result['_day_templates']}
                    # 批次完成即寫回持久帳本，供下一次（含跨月）排班回溯「連續上班天數」。
                    if not period_result['df_result'].empty:
                        _workflow_store.record_workdays(zip(period_result['df_result']['派單居服員'], period_result['df_result']['排班日期']))
                    _moved = period_result.get('time_shifts')
                    if _moved is not None and not _moved.empty:
                        _new_times = _moved.set_index('任務ID')
                        _hit = df_tasks_all['任務ID'].isin(_new_times.index)
                        df_tasks_all.loc[_hit,'時間窗_開始'] = df_tasks_all.loc[_hit,'任務ID'].map(_new_times['新開始'])
                        df_tasks_all.loc[_hit,'時間窗_結束'] = df_tasks_all.loc[_hit,'任務ID'].map(_new_times['新結束'])
                    did = run_phase3_did(st.session_state['data_hist'].copy())
                except Exception as exc:
                    _friendly = traffic_wake.friendly_error(exc)
                    if _friendly:
                        # 交通主機沒醒：改用白話說明，並自動開始在背景喚醒，醒了再按一次即可。
                        traffic_wake.start(_osrm_host)
                        st.session_state['_schedule_osrm_error'] = _friendly
                        st.session_state['_schedule_osrm_detail'] = str(exc)
                        st.rerun()
                    st.error('排班失敗：'+str(exc))
                    st.stop()
                st.session_state['last_result'] = {
                    'input_signature': _live_signature, 'df_tasks': df_tasks_all.copy(),
                    'df_tasks_all': df_tasks_all.copy(), 'df_cg': df_cg, 'df_cl': df_cl,
                    'df_matches': period_result['df_matches'], 'phase2': period_result,
                    'did': did, 'schedule_failed_dates': [], 'weekly_workflow': False,
                    'agency_monthly_workflow': True, 'range_start': range_start, 'range_end': range_end,
                    'route_result': period_result['route_result'], 'client_primary_map': period_result['client_primary_map'],
                    'client_backup_map': dict(period_result.get('client_backup_map') or {}),
                    'task_support_map': dict(period_result.get('task_support_map') or {}),
                    '_seed_routes_available': True,
                    'extended_to_end': True,
                }
                st.session_state['overrides'] = {}
                st.session_state['dynamic_insertions'] = []
                new_dispatch_completed = True
                st.rerun()


if "last_result" in st.session_state:

    res = st.session_state["last_result"]
    res["df_cg"] = attach_blocks(res["df_cg"], _workflow_store)
    config.client_primary_map = dict(res.get('client_primary_map', _workflow_store.primary_map()))
    # 備援名單必須跟著班表還原：每次畫面重跑 config 都是新的，少了這行，備援接手的服務會被當成「非固定人選」退回，成功率就憑空掉。
    config.client_backup_map = dict(res.get('client_backup_map') or (res.get('phase2') or {}).get('client_backup_map') or {})
    # 臨時支援同理：只對該次服務有效，也要跟著班表還原。
    config.task_support_map = dict(res.get('task_support_map') or (res.get('phase2') or {}).get('task_support_map') or {})
    config.primary_profiles = {str(r['居服員ID']):r.to_dict() for _,r in res['df_cg'].iterrows()}
    config.work_day_minutes = dict(res.get('work_day_minutes', {}))
    # 人工插單／人工改派的「連續上班」提醒種子：本期間開始日往前10天本機構已確定的上班日。
    # 只在此處重算一次（本次畫面重跑共用），每個 config 實例都會被下方各入口讀取。
    _period_first_day = pd.to_datetime(res['df_tasks'].get('日期'), errors='coerce').min() if '日期' in res['df_tasks'].columns else None
    config.worker_workday_seed = _workflow_store.recent_workdays(_period_first_day, lookback=10) if pd.notna(_period_first_day) else {}

    df_tasks = res["df_tasks"]

    df_matches = res["df_matches"]

    phase2 = res["phase2"]

    did = res["did"]

    schedule_failed_dates = res.get("schedule_failed_dates") or []

    # 財務設定只改金額，保留已完成的排班結果。
    if not phase2['df_result'].empty:
        financials = calculate_task_revenue_and_salary(res['df_tasks'], config).set_index('任務ID')
        for column in ['預估服務營收', '預估居服員薪資', '財務估算狀態']:
            phase2['df_result'][column] = phase2['df_result']['任務ID'].map(financials[column])
    df_result = phase2["df_result"]

    assigned_count = phase2["assigned_count"]

    solver_ok = phase2["status"] in (pywraplp.Solver.OPTIMAL,pywraplp.Solver.FEASIBLE)

    # overrides／assigned_map 提前到此處初始化（原僅存在於④區塊），因為月曆視角

    # 「一鍵調班」與④區塊「居督人工覆寫」共用同一份 overrides 狀態與同一套

    # save_override_log 稽覈紀錄，月曆總覽（KPI 之後即會渲染）需要在此之前就能

    # 讀寫這份狀態，兩處纔不會各自維護一份互不同步的覆寫紀錄。

    st.session_state.setdefault("overrides", {})

    overrides = st.session_state["overrides"]

    # Persisted unavailability applies to every entrance, including an old baseline.
    _wf_people = {str(r['居服員ID']): r for _, r in res['df_cg'].iterrows()}
    _wf_effective = apply_overrides_to_result(df_result, overrides)
    _wf_assignments = dict(zip(_wf_effective.get('任務ID', []), _wf_effective.get('派單居服員', [])))
    for _, _wf_task in df_tasks.iterrows():
        _wf_tid = str(_wf_task['任務ID'])
        _wf_cg = str(_wf_assignments.get(_wf_task['任務ID'], ''))
        _wf_block = blocked_reason(_wf_task, _wf_people[_wf_cg]) if _wf_cg in _wf_people else None
        if _wf_cg in _wf_people and not _wf_block:
            _wf_block = primary_guard(_wf_task, _wf_people[_wf_cg], config) or (leave_overlap(_wf_task, _wf_people[_wf_cg]) and '已登錄請假，服務待安排')
        if _wf_block:
            overrides[_wf_tid] = {'cg_id': None, 'reason': _wf_block}
            _wf_record = _workflow_store.record(_wf_task, '')
            _workflow_store.save(_wf_tid, {**_wf_record, 'note': _wf_block})
    for _wf_item in st.session_state.get('dynamic_insertions', []):
        _wf_cg = str(_wf_item.get('cg_id') or '')
        if _wf_cg in _wf_people and (blocked_reason(_wf_item['task'], _wf_people[_wf_cg]) or primary_guard(_wf_item['task'], _wf_people[_wf_cg], config) or leave_overlap(_wf_item['task'], _wf_people[_wf_cg])):
            _wf_item['cg_id'] = None
    _wf_rejected = []
    _wf_saved_records = _workflow_store.all_records()
    for _, _wf_task in df_tasks.iterrows():
        _wf_saved = _wf_saved_records.get(str(_wf_task['任務ID']),{})
        _wf_cg = str(_wf_assignments.get(_wf_task['任務ID'], '') or '')
        # An explicitly rejected unchanged assignment must not return on reload.
        if _wf_saved.get('rejected_signature') == __import__('dispatch_workflow').task_signature(_wf_task, _wf_cg):
            overrides[str(_wf_task['任務ID'])] = {'cg_id':None, 'reason':_wf_saved.get('note','拒接／撤回')}

    assigned_map = dict(zip(df_result["任務ID"], df_result["派單居服員"])) if not df_result.empty else {}

    task_locations = (

        df_matches.drop_duplicates("任務ID").set_index("任務ID")[["地點緯度", "地點經度"]].apply(tuple, axis=1).to_dict()

        if not df_matches.empty

        else {}

    )

    def _workflow_frames():
        tasks = df_tasks.copy()
        effective = apply_overrides_to_result(df_result, overrides)
        for item in st.session_state.get('dynamic_insertions', []):
            tid = str(item['task']['任務ID'])
            if tid not in set(tasks['任務ID'].astype(str)):
                tasks = pd.concat([tasks, pd.DataFrame([item['task']])], ignore_index=True)
            effective = effective[effective['任務ID'].astype(str) != tid] if '任務ID' in effective else effective
            if item.get('cg_id'):
                effective = pd.concat([effective, pd.DataFrame([{'任務ID':tid,'派單居服員':item['cg_id']}])],ignore_index=True)
        for pending in _workflow_store.pending():
            if str(pending['任務ID']) not in set(tasks['任務ID'].astype(str)):
                tasks=pd.concat([tasks,pd.DataFrame([pending])],ignore_index=True)
        return tasks, effective

    def _consecutive_days_notice(cg_id, day, tasks_frame, result_frame):
        """人工改派／人工插單共用：只提醒不擋。合併跨月種子（config.worker_workday_seed）與
        目前有效班表，判斷指派 cg_id 到 day 是否會連續上班達上限+1天；與 AI 硬性規則
        （agency_rules.consecutive_workdays_error）同一顆函式、同一套門檻與文案。"""
        day = pd.to_datetime(day, errors='coerce')
        if pd.isna(day) or not cg_id:
            return None
        workdays = {str(k): set(v) for k, v in (getattr(config, 'worker_workday_seed', {}) or {}).items()}
        if tasks_frame is not None and not tasks_frame.empty and result_frame is not None and not result_frame.empty and '日期' in tasks_frame.columns:
            merged = result_frame.merge(tasks_frame[['任務ID', '日期']], on='任務ID', how='left')
            for _cg, _day in zip(merged['派單居服員'], merged['日期']):
                if pd.isna(_day):
                    continue
                workdays.setdefault(str(_cg), set()).add(pd.Timestamp(_day).strftime('%Y-%m-%d'))
        prior_seed = config.worker_workdays if hasattr(config, 'worker_workdays') else None
        config.worker_workdays = workdays
        try:
            return consecutive_workdays_error(cg_id, day.strftime('%Y-%m-%d'), config)
        finally:
            if prior_seed is None:
                del config.worker_workdays
            else:
                config.worker_workdays = prior_seed

    def _remember_owner(task, cg):
        client=str(task.get('案家ID',''))
        if cg and client and not config.client_primary_map.get(client):
            config.client_primary_map[client]=str(cg)
            res['client_primary_map']=dict(config.client_primary_map)

    def _remember_support(task, cg):
        """固定人選以外的人接下這一次服務時，記為「臨時支援（僅本次）」，重新整理後不會被退回。"""
        from caregiver_engine import _fixed_team
        if not cg:
            return
        task = task.to_dict() if hasattr(task, 'to_dict') else dict(task)
        team = _fixed_team(task, config)
        if team and str(cg) not in team:
            config.task_support_map[str(task.get('任務ID'))] = str(cg)
            res['task_support_map'] = dict(config.task_support_map)

    def _quick_reassign(task_id, new_cg_id, reason, apply_scope=None):

        """統一的居督覆寫寫入入口：即時衝突檢查通過後，寫入 overrides 狀態與稽覈

        日誌；回傳 None 表示成功，否則回傳供呼叫端 Modal 顯示的錯誤訊息。

        供月曆「居服員 x 日期」改派 Modal、月曆「未派單案件」處置，以及④區塊

        「居督人工覆寫」任務搜尋器共用（calendar_view._task_override_dialog／

        _quick_reassign_dialog）：三者差異只在 task_id 原本是否已有指派、觸發

        來源不同，apply_overrides_to_result 本就支援替未指派任務新增覆寫列，

        不需要另外實作一套指派邏輯；reason 由呼叫端的原因選擇器收集（必填）。

        """

        check_tasks, df_result_effective = _workflow_frames()

        conflict_msg = check_reassignment_conflict(

            task_id, new_cg_id, check_tasks, df_result_effective, res.get("df_cg", pd.DataFrame()),

            config, task_locations=task_locations, date_column="日期", df_cl=res.get('df_cl'),

        )

        if conflict_msg:

            return conflict_msg

        task_rows = check_tasks[check_tasks["任務ID"] == task_id]

        cl_id = task_rows.iloc[0]["案家ID"] if not task_rows.empty else ""

        ai_cg = assigned_map.get(task_id)

        ai_label = str(ai_cg) if ai_cg is not None else "未指派"

        save_override_log(task_id, cl_id, ai_label, new_cg_id if new_cg_id is not None else "未指派", reason)

        overrides[task_id] = {"cg_id": new_cg_id, "reason": reason}
        _remember_support(task_rows.iloc[0] if not task_rows.empty else {'任務ID': task_id, '案家ID': cl_id}, new_cg_id)

        if not config.client_primary_map.get(str(cl_id)):
            config.client_primary_map[str(cl_id)] = str(new_cg_id)
            res['client_primary_map'] = dict(config.client_primary_map)
        st.session_state.pop('last_override_propagation', None)
        if new_cg_id and not task_rows.empty and '日期' in task_rows.columns:
            _notice = _consecutive_days_notice(new_cg_id, task_rows.iloc[0]['日期'], check_tasks, df_result_effective)
            if _notice:
                st.session_state.setdefault('consecutive_days_notices', []).append(_notice)
        return None


    def _clear_override(task_id):

        """清除單一任務的居督覆寫，還原為 AI 建議；與原④區塊「清除覆寫」行為

        相同，不寫入稽覈日誌（稽覈日誌只記錄實際發生過的覆寫變更）。"""

        overrides.pop(task_id, None)

    def _list_candidates(task_id, candidate_cg_ids):

        """供月曆快速改派下拉選單使用：把候選居服員依「該時段是否有空檔」排序＋標籤。

        與 _quick_reassign 共用同一套 caregiver_engine.rank_candidates_by_availability／

        check_reassignment_conflict 判定邏輯（同源於 _evaluate_reassignment），確保

        選單顯示「可派單」的候選人在按下確認改派時不會被判定衝突而拒絕。

        另外傳入 df_cl，讓排序結果一併標記 Phase 1 硬性資格條件（性別、重度移位、

        環境排斥、可排班星期、可服務時段、請假、專長認證）不符者，而不是讓這些人

        因為候選名單只取全體居服員（見 calendar_view._quick_reassign_dialog）而

        「看似可派但其實從未檢查資格」——這裡刻意不影響 _quick_reassign 的

        check_reassignment_conflict（未傳 df_cl，維持原本只擋時間衝突／工時超額的

        權威判定範圍），因為本系統無強制派單權限機制，居督看到標記後仍可自行覆寫。

        """

        check_tasks, df_result_effective = _workflow_frames()

        return rank_candidates_by_availability(

            task_id, candidate_cg_ids, check_tasks, df_result_effective, res.get("df_cg", pd.DataFrame()),

            config, task_locations=task_locations, date_column="日期",

            df_cl=res.get("df_cl", pd.DataFrame()),

        )


    def _calendar_move(draft, reason, scope):
        """Preview confirmation: validate first, then change a single task atomically."""
        task_id = str(draft.get("task_id", ""))
        new_cg = str(draft.get("cg_id", ""))
        found = df_tasks["任務ID"].astype(str).eq(task_id)
        if found.sum() != 1:
            return "找不到唯一的服務任務。"
        old = df_tasks.loc[found].iloc[0]
        try:
            new_day = pd.Timestamp(draft["date"]).date()
            first = datetime.strptime(draft["start"], "%H:%M")
            last = datetime.strptime(draft["end"], "%H:%M")
        except (KeyError, TypeError, ValueError):
            return "日期或時段格式不正確。"
        old_start = pd.to_datetime(old.get("時間窗_開始"), errors="coerce")
        old_end = pd.to_datetime(old.get("時間窗_結束"), errors="coerce")
        if pd.isna(old_start) or pd.isna(old_end):
            return "原服務時段無法辨識，請改用人工改派。"
        if (last - first).total_seconds() != (old_end - old_start).total_seconds():
            return "拖曳須維持原服務分鐘，請重新操作。"
        if first.time() < pd.Timestamp("08:00").time() or last.time() > pd.Timestamp("21:00").time() or first >= last:
            return "服務須在 08:00–21:00 內，且結束時間晚於開始。"
        original_day = pd.to_datetime(old.get("日期"), errors="coerce")
        if pd.isna(original_day):
            return "原服務日期無法辨識。"
        allowed_dates = pd.to_datetime(df_tasks["日期"], errors="coerce").dropna().dt.date
        if allowed_dates.empty or not allowed_dates.min() <= new_day <= allowed_dates.max():
            return "新日期須在目前已產生的班表期間內。"
        time_changed = (new_day != original_day.date() or draft["start"] != str(old["時間窗_開始"])[:5])
        if time_changed and scope != "僅本次":
            return "調整日期或時段時，請選『僅本次』。"
        if not time_changed:
            st.session_state["override_apply_scope"] = scope
            return _quick_reassign(task_id, new_cg, reason)
        proposed = df_tasks.copy()
        proposed.loc[found, "日期"] = new_day.isoformat()
        proposed.loc[found, "時間窗_開始"] = draft["start"]
        proposed.loc[found, "時間窗_結束"] = draft["end"]
        proposed.loc[found, "星期"] = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"][new_day.weekday()]
        all_tasks, effective = _workflow_frames()
        extras=all_tasks.loc[~all_tasks['任務ID'].isin(proposed['任務ID'])]
        check_proposed=pd.concat([proposed,extras],ignore_index=True)
        try:
            conflict = check_reassignment_conflict(
                task_id, new_cg, check_proposed, effective, res.get("df_cg", pd.DataFrame()),
                config, task_locations=task_locations, date_column="日期", df_cl=res.get("df_cl"),
            )
        except Exception as exc:
            return "無法完成衝突檢查：" + str(exc)
        if conflict:
            return conflict
        old_cg = effective.loc[effective["任務ID"].astype(str).eq(task_id), "派單居服員"]
        old_label = str(old_cg.iloc[0]) if not old_cg.empty else "未指派"
        try:
            save_override_log(task_id, old.get("案家ID", ""), old_label, new_cg,
                              f"{reason}｜時段 {original_day.date()} {old['時間窗_開始']}-{old['時間窗_結束']} → {new_day} {draft['start']}-{draft['end']}")
        except Exception as exc:
            return "寫入改班紀錄失敗：" + str(exc)
        res["df_tasks"] = proposed
        if "df_tasks_all" in res:
            all_tasks = res["df_tasks_all"].copy()
            all_mask = all_tasks["任務ID"].astype(str).eq(task_id)
            for col in ("日期", "時間窗_開始", "時間窗_結束", "星期"):
                if col in all_tasks:
                    all_tasks.loc[all_mask, col] = proposed.loc[found, col].iloc[0]
            res["df_tasks_all"] = all_tasks
        overrides[task_id] = {"cg_id": new_cg, "reason": reason}
        _remember_support(proposed.loc[found].iloc[0], new_cg)
        result_mask = phase2["df_result"]["任務ID"].astype(str).eq(task_id)
        if result_mask.any():
            if "服務時段" in phase2["df_result"]:
                phase2["df_result"].loc[result_mask, "服務時段"] = draft["start"] + "-" + draft["end"]
            if "排班日期" in phase2["df_result"]:
                phase2["df_result"].loc[result_mask, "排班日期"] = new_day
        res.pop("route_signature", None)
        res.pop("google_verified_signature", None)
        return None

    # 排班過程的提醒先收集起來，分別在「02 智慧排班作業」與「04 排班結果與覆核」的警示區呈現。
    _schedule_notes = []
    if res.get('agency_monthly_workflow'):
        # 固定居服員／備援名單只有員編，不另外呈現；排班結果與班表直接顯示在下方。
        _shift_table = (res.get('phase2') or {}).get('time_shifts')
        if _shift_table is not None and not _shift_table.empty:
            _schedule_notes.append(('caption', f'有 {len(_shift_table)} 次服務的開始時間在設定的彈性範圍內挪動（服務分鐘不變），班表已採用新時間。'))

    if not solver_ok:
        if schedule_failed_dates:
            _schedule_notes.append(('warning', f"以下 {len(schedule_failed_dates)} 個日期無法求得最佳解（其餘日期仍正常派單）："
                                               f"{', '.join(str(d) for d in schedule_failed_dates)}"))
        else:
            _schedule_notes.append(('warning', "OR-Tools 無法在目前資料與參數下找到最佳解，請檢查資料是否有效或調整參數。"))

    _operational_tasks = df_tasks.copy()
    _operational_assignments = apply_overrides_to_result(df_result, overrides)
    for _insertion in st.session_state.get('dynamic_insertions', []):
        _inserted = _insertion['task']
        _operational_tasks = pd.concat([_operational_tasks, pd.DataFrame([_inserted])], ignore_index=True)
        _operational_assignments = pd.concat([_operational_assignments,
            pd.DataFrame([{'任務ID': _inserted['任務ID'], '派單居服員': _insertion.get('cg_id')}])], ignore_index=True)
    _operational_tasks = _operational_tasks.drop_duplicates('任務ID', keep='last')
    _operational_assignments = _operational_assignments.drop_duplicates('任務ID', keep='last')
    if '派單居服員' in _operational_assignments:
        _operational_assignments = _operational_assignments.loc[
            _operational_assignments['派單居服員'].notna() &
            ~_operational_assignments['派單居服員'].astype(str).isin(['', '未指派', 'None', 'nan'])].copy()
    # 臨時新增的案家不在案家表內，其任務在建立時已帶齊案家屬性與座標：只合併原有任務，再接上臨時任務。
    _inserted_ids = {str(i['task']['任務ID']) for i in st.session_state.get('dynamic_insertions', [])}
    _is_inserted = _operational_tasks['任務ID'].astype(str).isin(_inserted_ids)
    travel_tasks = merge_task_client_tables(df_tasks.loc[~df_tasks['任務ID'].astype(str).isin(_inserted_ids)], res['df_cl'])
    if _is_inserted.any():
        # 欄位維持與原任務表一致：多出欄位會讓每一列的簽章改變，造成整個班表交通重算。
        travel_tasks = pd.concat([travel_tasks, _operational_tasks.loc[_is_inserted].reindex(columns=travel_tasks.columns)], ignore_index=True)
    route_signature = hashlib.sha256(repr((
        _operational_assignments[['任務ID', '派單居服員']].sort_values('任務ID').to_dict('records'),
        travel_tasks.sort_values('任務ID').to_dict('records'), res['df_cg'].to_dict('records'),
        {k: v for k, v in vars(config).items() if k not in ('work_day_minutes', 'dispatch_cache')}
    )).encode()).hexdigest()

    from schedule_checks import incremental_travel
    try:
        # 指派、任務、居服員、設定都沒變（例如只是開關視窗、切換日期）就直接沿用上次的逐段交通，
        # 不再逐日重新雜湊與合併；有異動時只重算受影響的「居服員×日期」。
        if (res.get('route_signature') == route_signature and isinstance(res.get('route_result'), pd.DataFrame)
                and not res['route_result'].empty and '預估車程(分)' in res['route_result']):
            res.pop('_seed_routes_available', None)
        else:
            res['route_result']=incremental_travel(
                _operational_assignments,travel_tasks,res['df_cg'],config,
                res.setdefault('_daily_route_cache',{}),calculate_schedule_travel,
                initial_routes=res.get('route_result') if res.pop('_seed_routes_available',False) else None)
            if res.get('route_signature') != route_signature:
                res.pop('google_verified_signature',None)
            res['route_signature']=route_signature
    except Exception as exc:
        st.error('交通暫未計算，班表仍可檢視與下載；估計交通留空。原因：'+str(exc))
        res['route_result'] = pd.DataFrame()
    df_route_result = res.get('route_result', pd.DataFrame()).copy()
    _route_columns = ['任務ID', '交通日期', '派單居服員', '交通路段', '預估車程(分)',
                      '路段緩衝(分)', '含緩衝交通時間(分)', '交通計算狀態']
    df_route_result = df_route_result.reindex(columns=list(dict.fromkeys(list(df_route_result.columns) + _route_columns)))
    df_route_result['預估車程(分)'] = pd.to_numeric(df_route_result['預估車程(分)'], errors='coerce')
    daily_travel = summarize_schedule_travel(df_route_result)

    _effective_for_kpi = _operational_assignments.copy()
    _task_ids = set(df_tasks['任務ID'].dropna().astype(str))
    if {'任務ID', '派單居服員'}.issubset(_effective_for_kpi.columns):
        _valid_assignee = _effective_for_kpi['派單居服員'].notna() & ~_effective_for_kpi['派單居服員'].astype(str).str.strip().isin(['', '未指派', 'nan', 'None'])
        _effective_for_kpi = _effective_for_kpi.loc[_valid_assignee & _effective_for_kpi['任務ID'].astype(str).isin(_task_ids)].drop_duplicates('任務ID', keep='last').copy()
        _assigned_ids = set(_effective_for_kpi['任務ID'].astype(str))
    else:
        _assigned_ids = set()
    assigned_count = len(_assigned_ids)
    total_tasks = len(_task_ids)
    assign_rate = (assigned_count / total_tasks * 100) if total_tasks else 0.0

    travel_complete = not df_route_result.empty and df_route_result['預估車程(分)'].notna().all()
    avg_transition = df_route_result['預估車程(分)'].mean()

    avg_score = df_result["適配分數"].mean() if not df_result.empty else 0.0

    _route_digest = _frame_digest(df_route_result)
    _payroll_key = (route_signature, _route_digest)
    _payroll_cached = res.get('_payroll_cache')
    if _payroll_cached and _payroll_cached[0] == _payroll_key:
        task_payroll, daily_payroll, monthly_payroll, work_ledger = (
            _payroll_cached[1][0].copy(), _payroll_cached[1][1].copy(), _payroll_cached[1][2].copy(), dict(_payroll_cached[1][3]))
    else:
        task_payroll, daily_payroll, monthly_payroll, work_ledger = payroll_tables(
            travel_tasks, _operational_assignments, res['df_cg'], df_route_result, config)
        res['_payroll_cache'] = (_payroll_key, (task_payroll.copy(), daily_payroll.copy(), monthly_payroll.copy(), dict(work_ledger)))
    config.work_day_minutes = work_ledger
    res['work_day_minutes'] = work_ledger
    total_revenue = task_payroll['預估服務營收'].sum(min_count=1) if not task_payroll.empty else 0.0
    finance_partial = not task_payroll.empty and task_payroll['預估服務營收'].isna().any()
    salary_partial = not daily_payroll.empty and daily_payroll['預估居服員薪資'].isna().any()
    total_salary = (daily_payroll['預估居服員薪資'].sum() if not salary_partial else float('nan')) if not daily_payroll.empty else 0.0
    if input_mode=='機構資料（單一Excel／三個工作表）':
        _schedule_notes.append(('warning', '此結果為排班草案，需確認可服務時段、照護條件及座標精度後使用。'))
    if phase2['status'] == pywraplp.Solver.FEASIBLE:
        _schedule_notes.append(('info', '目前為固定人選的可行班表草案；尚未經居服員確認，也未證明全域最優。'))
    diagnostics=df_matches.attrs.get('matching_diagnostics',[])
    _scope_dates = pd.to_datetime(df_tasks.get('日期', pd.Series(dtype=object)), errors='coerce').dropna()
    _scope_label = f"{_scope_dates.min():%Y-%m-%d} ～ {_scope_dates.max():%Y-%m-%d}" if not _scope_dates.empty else '目前任務'
    _pending_report = df_tasks.loc[~df_tasks['任務ID'].astype(str).isin(_assigned_ids)].drop_duplicates('任務ID').copy()
    _pending_diagnostics = list(diagnostics)
    for _tid,_override in overrides.items():
        if not _override.get('cg_id'):
            _pending_diagnostics.append({'任務ID':str(_tid),'問題分類':'人工待處理','剔除原因':_override.get('reason','')})
    _pending_clients = pending_client_summary(df_tasks,_effective_for_kpi,_pending_diagnostics,config.client_primary_map)
    st.session_state['pending_client_summary'] = _pending_clients

    # 衝突與異常：未派服務分類、已派服務的班次重疊與轉場不足（只整理資料，不重新排班）。
    try:
        _issue_unassigned = unassigned_issues(_pending_report, _pending_diagnostics)
        _issue_overlaps = assigned_overlaps(_operational_tasks, _operational_assignments)
        _issue_transitions = transition_issues(df_route_result)
    except Exception as exc:
        _issue_unassigned, _issue_overlaps, _issue_transitions = pd.DataFrame(columns=['問題類型', '案家ID']), None, None
        _schedule_notes.append(('warning', '衝突與異常整理暫時無法完成，班表不受影響。原因：' + str(exc)))
    _issue_table = issue_summary(_issue_unassigned, _issue_overlaps, _issue_transitions)
    _issue_count = lambda name: int(_issue_table.loc[_issue_table['問題類型'].eq(name), ['已派需覆核', '未派服務']].sum(axis=1).sum())
    _review_needed = int(_issue_table['已派需覆核'].sum())

    with _sidebar_status:
        st.markdown(f'<div class="cf-status" style="margin-top:6px">✅ 排班完成率 <b>{assign_rate:.1f}%</b>（{assigned_count:,}／{total_tasks:,}）<br>'
                    f'⚠️ 未派 <b>{total_tasks - assigned_count:,}</b> 筆 · {len(_pending_clients)} 個案家'
                    + (f'<br>🔁 已派需覆核 <b>{_review_needed}</b> 筆' if _review_needed else '') + '</div>', unsafe_allow_html=True)

    def _render_kpis():
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("排班完成率", f"{assign_rate:.1f}%", help=f"{assigned_count} / {total_tasks} 筆任務成功指派（派單成功率）")
        k2.metric("已排／全部服務", f"{assigned_count:,} ／ {total_tasks:,}", help="已納入人工改派，包含整個選定期間。")
        k3.metric("未排服務", f"{total_tasks - assigned_count:,} 筆",
                  delta=f"{len(_pending_clients)} 個案家待處理" if len(_pending_clients) else "全部已派",
                  delta_color="off", help="每個仍有未派服務的案家只計一筆待辦；整案套用會處理同案全部星期與時段。")
        k4.metric("平均每段交通", f"{avg_transition:.1f} 分" if travel_complete else ("尚未計算" if df_route_result.empty else "路段資料待補齊"),
                  help="依每位員工每天服務順序：服務起點→首案，前案→下一案。不含轉場緩衝及末案返家。")
        st.caption(f"統計範圍：{_scope_label}　·　平均適配得分 {avg_score:.1f} 分")

    def _render_schedule_notes():
        for _level, _text in _schedule_notes:
            getattr(st, _level)(_text)

    def _issue_go_button(key):
        _open = len(_pending_clients) + _review_needed
        if _open:
            st.button(f"⚠️ {len(_pending_clients)} 個案家未排完、{_review_needed} 筆已派服務需覆核 → 前往「03 衝突與異常處理」",
                      key=key, on_click=go_to_module, args=(M_ISSUES,), width='stretch')
        else:
            st.success('沒有未派服務，也沒有需要覆核的衝突。')

    _issue_view = None
    if _module == M_SCHEDULE:
        subhead('排班成果')
        _render_kpis()
        _render_schedule_notes()
        _issue_go_button('schedule_go_issues')
    elif _module == M_REVIEW:
        module_header(M_REVIEW)
        subhead('排班完成率')
        _render_kpis()
    elif _module == M_ISSUES:
        module_header(M_ISSUES)
        _i1, _i2, _i3, _i4, _i5 = st.columns(5)
        _i1.metric('未成功派班', f"{total_tasks - assigned_count:,} 筆", delta=f"{len(_pending_clients)} 個案家", delta_color='off',
                   help='尚未排入班表的服務；下方「未派案家與替代方案」逐案處理。')
        _i2.metric('班次重疊', f"{_issue_count('班次重疊'):,} 筆", help='已派服務中同一人時段重疊，加上因時段重疊而排不進的服務。')
        _i3.metric('轉場不足', f"{_issue_count('轉場不足'):,} 筆", help='已派服務轉場時間不足或交通無法計算，加上因轉場空檔不足而排不進的服務。')
        _i4.metric('條件不符', f"{_issue_count('條件不符'):,} 筆", help='證照／訓練、資格期限、可服務星期不符而排不進的服務。')
        _i5.metric('工時與休息', f"{_issue_count('工時與休息'):,} 筆", help='每日工時、休息或連續上班天數限制而排不進的服務。')
        _issue_views = ['未派案家與替代方案', '衝突與條件明細', '臨時異動']
        _issue_view = big_tabs(_issue_views, 'cf_issue_view', icons=['🧩', '🔍', '⚡'])
        if _issue_view == '衝突與條件明細':
            st.caption('依問題類型整理；「已派需覆核」是已在班表上但需要居督確認的服務，「未派服務」是尚未排入的服務。')
            show_table(_issue_table, hide_index=True, width='stretch',
                       column_config={'說明': st.column_config.TextColumn(width='medium'),
                                      '建議處理': st.column_config.TextColumn(width='large')})
            if _issue_overlaps is not None and not _issue_overlaps.empty:
                subhead(f'班次重疊（已派）· {len(_issue_overlaps)} 筆')
                _ov = _issue_overlaps.copy()
                _ov['居服員'] = _ov['居服員'].map(caregiver_label)
                for _col in ('案家ID', '重疊案家ID'):
                    _ov[_col] = _ov[_col].map(client_label)
                show_table(_ov, hide_index=True, width='stretch')
                st.caption('處理方式：到「02 智慧排班作業」的人工調整把其中一筆改派給其他人。')
            if _issue_transitions is not None and not _issue_transitions.empty:
                subhead(f'轉場不足／交通待確認（已派）· {len(_issue_transitions)} 筆')
                _tv = _issue_transitions.copy()
                _tv['派單居服員'] = _tv['派單居服員'].map(lambda v: caregiver_label(v) if pd.notna(v) else v)
                show_table(_tv, hide_index=True, width='stretch', column_config={'交通路段': st.column_config.TextColumn(width='large')})
                st.caption('轉場時間不足：調整前後案時段或改派順路的居服員；交通待確認：核對案家座標與路網連線。')
            if not _issue_unassigned.empty:
                subhead(f'未成功派班 · {len(_issue_unassigned)} 筆')
                _group_pick = st.multiselect('問題類型', [g for g in ISSUE_GROUP_NAMES if g in set(_issue_unassigned['問題類型'])],
                                             key='issue_group_filter', placeholder='全部問題類型')
                _uv = _issue_unassigned if not _group_pick else _issue_unassigned.loc[_issue_unassigned['問題類型'].isin(_group_pick)]
                _uv = _uv.assign(案家ID=_uv['案家ID'].map(client_label), 原因=_uv['原因'].map(lambda t: _label_ids_plain(t)))
                show_table(_uv, hide_index=True, width='stretch',
                           column_config={'原因': st.column_config.TextColumn(width='large'),
                                          '建議處理': st.column_config.TextColumn(width='medium')})
                download_button_quiet('下載衝突與異常清單 CSV', export_frame(_issue_unassigned).to_csv(index=False).encode('utf-8-sig'),
                                      '衝突與異常清單.csv', 'text/csv', key='download_issue_list')
                st.caption('替代方案：切換到「未派案家與替代方案」，系統會逐案試算前後挪時間、換人、拆成兩段等方案。')
            if diagnostics:
                with st.expander('候選配對及未派單原因（排班引擎原始診斷）'):
                    show_table(pd.DataFrame(diagnostics),hide_index=True,width='stretch')
            if (_issue_overlaps is None or _issue_overlaps.empty) and (_issue_transitions is None or _issue_transitions.empty) and _issue_unassigned.empty:
                st.success('目前沒有班次重疊、轉場不足或未成功派班的服務。')

    if True:
        _unit_minutes = unit_minutes_from_master(st.session_state.get('edit_service_code'))
        _limit_mins = float(getattr(config, 'continuous_work_limit_mins', 240.0))

        def _label_ids(text):
            """把文字裡的居服員編號換成顯示用名稱（遵守遮罩設定）。"""
            out = str(text)
            if _show_names:
                for _cid in sorted(_caregiver_names, key=len, reverse=True):
                    out = re.sub(r'(?<![A-Za-z0-9])' + re.escape(str(_cid)) + r'(?![A-Za-z0-9])', caregiver_label(_cid), out)
            return out

        _sig_memo = []

        def _base_signature():
            if _sig_memo:
                return _sig_memo[0]
            _sig_memo.append(hashlib.sha256(repr((
                _frame_digest(_operational_assignments[['任務ID', '派單居服員']]),
                _frame_digest(df_tasks[[c for c in ['任務ID', '案家ID', '日期', '時間窗_開始', '時間窗_結束'] if c in df_tasks]]),
                sorted((str(k), str(v)) for k, v in config.client_primary_map.items()),
                sorted((str(k), tuple(map(str, v))) for k, v in (getattr(config, 'client_backup_map', None) or {}).items()),
                sorted((str(k), str(v)) for k, v in (getattr(config, 'task_support_map', None) or {}).items()),
                _frame_digest(res['df_cg']), config.max_monthly_overtime_hours)).encode()).hexdigest())
            return _sig_memo[0]

        import uuid as _uuid
        import dataclasses as _dc
        import analysis_bg
        _bg_session = res.setdefault('_bg_id', _uuid.uuid4().hex)
        _empty_key = repr(empty_edits())
        _bg_ctx = []

        def _bg_context():
            """背景分析用的資料快照（算的是「現在這一刻」的班表，之後班表變了就會被丟掉）。"""
            if not _bg_ctx:
                _cfg = _dc.replace(config, client_primary_map=dict(config.client_primary_map),
                                   client_backup_map=dict(getattr(config, 'client_backup_map', None) or {}),
                                   task_support_map=dict(getattr(config, 'task_support_map', None) or {}),
                                   work_day_minutes=dict(config.work_day_minutes), dispatch_cache={},
                                   primary_profiles=dict(getattr(config, 'primary_profiles', None) or {}))
                _bg_ctx.append({'tasks': df_tasks.copy(), 'assignments': _operational_assignments.copy(), 'caregivers': res['df_cg'].copy(),
                                'clients': res['df_cl'].copy(), 'config': _cfg, 'unit_minutes': _unit_minutes, 'limit': _limit_mins})
            return _bg_ctx[0]

        def _bg_key(case, worker, stage):
            return (_bg_session, _base_signature(), str(case), str(worker), stage)

        def _bg_submit(items):
            """把未派案家的分析排進背景：先全部算「最優先方案」（卡片用），再補齊其他方案。已排過的不重複。"""
            _sig = _base_signature()
            analysis_bg.forget_other_signatures(_bg_session, _sig)
            _items = [(str(c), str(w)) for c, w in items if str(c) and str(w) and str(w) != 'None']
            if len(_items) > 1:
                # 先把所有未派案家的交通一次備齊（打包成最少的路網請求），之後每個案家的分析就不必各自等路網。
                analysis_bg.submit((_bg_session, _sig, '__prefetch__', '', 'p'),
                                   lambda _cases=[c for c, _ in _items]: analysis_bg.prefetch_case(
                                       _bg_context()['tasks'], _bg_context()['caregivers'], _bg_context()['clients'],
                                       _bg_context()['config'], _cases))
            for _c, _w in _items:
                _memo = analysis_bg.memo_for(_bg_session, _sig, _c, _w)
                analysis_bg.submit(_bg_key(_c, _w, 'q'),
                                   lambda _c=_c, _w=_w, _memo=_memo: analysis_bg.compute_entry(_bg_context(), _c, _w, empty_edits(), True, _memo))
            for _c, _w in _items:
                _memo = analysis_bg.memo_for(_bg_session, _sig, _c, _w)

                def _full(_c=_c, _w=_w, _memo=_memo):
                    _q = analysis_bg.get(_bg_key(_c, _w, 'q'))
                    _seed = _q['entry'] if _q and _q['state'] == 'done' else None
                    return analysis_bg.compute_entry(_bg_context(), _c, _w, empty_edits(), False, _memo, existing=_seed)
                analysis_bg.submit(_bg_key(_c, _w, 'f'), _full)

        def _bg_entry(case, worker):
            _f, _q = analysis_bg.get(_bg_key(case, worker, 'f')), analysis_bg.get(_bg_key(case, worker, 'q'))
            for _slot in (_f, _q):
                if _slot and _slot['state'] == 'done' and _slot['entry'] is not None:
                    return _slot['entry']
            return None

        def _bg_state(case, worker):
            """'done'／'running'（排隊或計算中）／'error'／'none'"""
            _slots = [analysis_bg.get(_bg_key(case, worker, st_)) for st_ in ('q', 'f')]
            if any(x and x['state'] == 'error' for x in _slots) and _bg_entry(case, worker) is None:
                return 'error'
            if _bg_entry(case, worker) is not None:
                return 'done'
            return 'running' if any(x and x['state'] in ('queued', 'running') for x in _slots) else 'none'

        def _analysis(case, edits_obj, worker, quick=False):
            """整案預覽＋白話診斷＋已驗證的方案。
            沒有任何手動調整時由背景算好直接取用（還沒算好回傳 None）；有調整時當場算（只算最優先的下一步，約 2～3 秒）。"""
            _sig = _base_signature()
            case, worker = str(case), str(worker)
            if repr(edits_obj) == _empty_key:
                _bg_submit([(case, worker)])
                _hit = _bg_entry(case, worker)
                if _hit is not None or (_bg_state(case, worker) in ('running', 'none') and (case, worker) in _bg_pending):
                    return _hit
            store = res.setdefault('_plan_cache', {})
            if store.get('_sig') != _sig:
                store.clear()
                store['_sig'] = _sig
            key = (case, worker, repr(edits_obj))
            if key not in store:
                store[key] = analysis_bg.compute_entry(_bg_context(), case, worker, edits_obj, True,
                                                       analysis_bg.memo_for(_bg_session, _sig, case, worker))
            return store[key]

        def _cached_analysis(case, edits_obj, worker):
            if repr(edits_obj) == _empty_key:
                return _bg_entry(case, worker)
            store = res.get('_plan_cache') or {}
            return store.get((str(case), str(worker), repr(edits_obj))) if store.get('_sig') == _base_signature() else None

        _bg_pending = [(str(r['案家ID']), str(config.client_primary_map.get(str(r['案家ID'])) or ''))
                       for r in _pending_clients.to_dict('records')[:6]]
        _bg_pending = [(c, w) for c, w in _bg_pending if w]
        analysis_bg.forget_other_signatures(_bg_session, _base_signature())
        # 背景分析等整頁畫完才開始排（不跟畫面搶 CPU）；卡片先顯示「分析中」，之後每 3 秒自動刷新到算完為止。
        _bg_busy = any(_bg_state(_c, _w) != 'done' and _bg_state(_c, _w) != 'error' for _c, _w in _bg_pending)
        st.session_state['_pending_timer'] = _bg_busy
        st.session_state['_pending_tick'] = 0

        def _pending_panel():
            """整案處理：先看這個案家需要什麼，再看每位居服員符不符合，最後才選人預覽。只重跑這一區。"""
            st.session_state['_pending_tick'] = st.session_state.get('_pending_tick', 0) + 1
            if st.session_state.get('_pending_timer'):
                if st.session_state['_pending_tick'] >= 2:
                    _bg_submit(_bg_pending)   # 保險：萬一整頁沒畫到最後，定時刷新時也會開始分析
                if all(_bg_state(_c, _w) in ('done', 'error') for _c, _w in _bg_pending):
                    st.session_state['_pending_timer'] = False
                    st.rerun()
            _has_pending = not _pending_clients.empty
            _panel_title = (f'未派案家 · {len(_pending_clients)} 個案家待整案處理' if _has_pending
                            else '整案調整（進階）· 目前沒有未派案家，仍可重新安排任何案家')
            def _edit_apply_many(case, applies):
                """按鈕回呼：把一個方案（一到數個調整）記到這個案家的調整清單。"""
                store = st.session_state.setdefault('pending_edits', {})
                current = store.get(str(case)) or empty_edits()
                for item in applies:
                    current = apply_suggestion(current, item)
                store[str(case)] = current
                st.session_state.pop('pending_client_preview', None)

            def _edit_apply(case, apply):
                _edit_apply_many(case, [apply])

            def _edits_has_any(edits_obj):
                return bool(edits_obj.get('shifts') or edits_obj.get('splits') or edits_obj.get('moves') or edits_obj.get('workers') or edits_obj.get('pattern_workers'))

            def _edit_clear(case):
                st.session_state.setdefault('pending_edits', {}).pop(str(case), None)
                st.session_state.pop('pending_client_preview', None)

            def _choose_case(case):
                st.session_state['pending_client_choice'] = str(case)

            def _commit_case(_case, _preview, _reason, _edits_used, main=None):
                """把整案預覽的結果寫進班表（已派／固定人選／調整紀錄）。簡單流程與進階流程共用。"""
                _base_rows = client_patterns(df_tasks, _case)
                _accepted_workers = [str(w) for w in _preview['accepted']['派單居服員']]
                if not main:
                    _counts = pd.Series(_accepted_workers).value_counts()
                    _tied = sorted(_counts[_counts == _counts.max()].index) if len(_counts) else []
                    main = str(_preview['worker']) if str(_preview['worker']) in _tied else (_tied[0] if _tied else str(_preview['worker']))
                main = str(main)
                _workflow_store.set_reviewed_primary(_case, main, _reason)
                _team_others = tuple(sorted(({str(w) for w in _preview.get('team', ())} | set(_accepted_workers)) - {main, ''}))
                _chosen = dict(zip(_preview['accepted']['任務ID'], _preview['accepted']['派單居服員']))
                _failed = dict(zip(_preview['failed']['任務ID'], _preview['failed']['原因']))
                _case_rows = _preview['tasks'].loc[_preview['tasks']['案家ID'].astype(str).eq(str(_case))].to_dict('records')
                for _task in _case_rows:
                    _tid = str(_task['任務ID'])
                    overrides[_tid] = {'cg_id': _chosen.get(_tid), 'reason': _failed.get(_tid, _reason)}
                def _slim(_frame, _ref):
                    """預覽表已併入案家欄位（多出約 25 欄）；寫回時只留原本的任務欄位，
                    否則每一列的交通簽章都會改變，整份班表的路線被迫全部重算（存檔後白畫面很久的主因）。"""
                    _cl_cols = set(res['df_cl'].columns) if res.get('df_cl') is not None else set()
                    _cols = [c for c in _ref.columns if c in _frame.columns] + [c for c in _frame.columns if c not in _ref.columns and c not in _cl_cols]
                    return _frame[_cols].copy()
                res['df_tasks'] = _slim(_preview['tasks'], df_tasks)
                _weekly_changes = res.setdefault('weekly_plan_changes', {})
                for _task in _case_rows:
                    if '#' in str(_preview.get('pattern_of', {}).get(str(_task['任務ID']), '')):
                        continue  # 拆成兩段的時段不是「改週計畫時間」，Excel 要另外改成兩列
                    _source_row = pd.to_numeric(_task.get('來源列'), errors='coerce')
                    if pd.notna(_source_row):
                        _weekly_changes[int(_source_row)] = {'開始時間': _task['時間窗_開始'], '結束時間': _task['時間窗_結束']}
                res['df_tasks_all'] = _slim(_preview['tasks'], res.get('df_tasks_all') if res.get('df_tasks_all') is not None else df_tasks)
                st.session_state['edit_tasks'] = _slim(_preview['tasks'], st.session_state.get('edit_tasks', df_tasks))
                for _frame in [st.session_state['edit_cl'], res['df_cl']]:
                    _mask = _frame['案家ID'].astype(str).eq(str(_case))
                    for _field in ['歷史首選', '歷史首選居服員ID']:
                        _frame.loc[_mask, _field] = main
                    if '歷史備援' not in _frame:
                        _frame['歷史備援'] = ''
                    _frame['歷史備援'] = _frame['歷史備援'].astype(object)
                    _frame.loc[_mask, '歷史備援'] = '、'.join(_team_others)
                res.setdefault('whole_client_log', []).extend(change_log_rows(
                    _case, _edits_used, _base_rows, main, _team_others, _reason))
                st.session_state.setdefault('pending_edits', {}).pop(str(_case), None)
                for _stale in ('_advice_cache', '_plan_cache', '_screen_cache', '_trial_cache', '_team_cache', 'route_signature', 'google_verified_signature'):
                    res.pop(_stale, None)
                config.client_primary_map = {**dict(_preview['primary_map']), str(_case): main}
                config.client_backup_map = {**config.client_backup_map, str(_case): _team_others}
                res['client_primary_map'] = dict(config.client_primary_map)
                res['client_backup_map'] = dict(config.client_backup_map)
                res['phase2']['client_primary_map'] = dict(config.client_primary_map)
                res['phase2']['client_backup_map'] = dict(config.client_backup_map)
                res['input_signature'] = _current_input_signature()
                st.session_state.pop('pending_client_preview', None)
                st.session_state['whole_client_notice'] = (
                    f"✅ 已處理 {client_label(_case)}：{len(_chosen)} 筆服務由 {'、'.join(caregiver_label(str(w)) for w in sorted(set(_chosen.values())))} 承接"
                    + (f"；還有 {len(_failed)} 筆仍待處理。" if _failed else "；這個案家已全部排進，已從待處理清單移除。")
                    + " 固定人選已更新。調整只存在這次操作中，要長期保留請下載「整案調整紀錄」並同步修改 Excel。")
                st.rerun()

            with st.expander(_panel_title, expanded=_has_pending):
                _done_cases = sorted({str(r.get('案號')) for r in (res.get('whole_client_log') or [])})
                if _done_cases:
                    st.success('本次已處理完成的案家：' + '、'.join(client_label(c) for c in _done_cases) + '（已寫入班表，不會再出現在待處理清單）')
                if _has_pending:
                    st.markdown('##### ① 先看：每個未派案家為什麼排不進、怎麼處理')
                    st.caption('系統已替每個案家用同一套規則（資格、時段、交通、休息、連續工作、月工時）逐項試算過。'
                               '點「處理這一案」→ 選一個建議方案 → 確認寫入，就會從這份清單消失。')
                    _rows_pending = _pending_clients.to_dict('records')
                    if _bg_busy:
                        st.caption('⏳ 系統正在背景分析（一次一個案家）；算好的會自動換成原因與建議，這段時間你可以先看班表、做別的事。')
                    _chosen_now = str(st.session_state.get('pending_client_choice') or '')
                    for _r in _rows_pending:
                        _cid = str(_r['案家ID'])
                        _own = str(config.client_primary_map.get(_cid) or '')
                        _hit = _cached_analysis(_cid, empty_edits(), _own) if _own else None
                        with st.container(border=True):
                            _c1, _c2, _c3 = st.columns([2.2, 5.6, 1.6], vertical_alignment='center')
                            _c1.markdown(f"**{client_label(_cid)}**")
                            _c1.caption(f"未派 {int(_r['未派服務數'])}／{int(_r['全部服務數'])} 次" + (f"　固定居服員 {caregiver_label(_own)}" if _own else '　尚未指定固定居服員'))
                            if _hit:
                                for _g in _hit['diag'][:2]:
                                    _c2.markdown(f"❗ **{_g['時段']}**（{'、'.join(_g['日期'][:4])}{'…' if len(_g['日期']) > 4 else ''}）{_g['次數']} 次：{_label_ids(_g['原因'])}")
                                if _hit['plans']:
                                    _top = _hit['plans'][0]
                                    _c2.markdown(f"💡 **建議**：{_label_ids(_top['標題'])}　" + ('✅ 套用後全部排進' if _top['仍未派'] == 0 else f"⚠️ 套用後仍剩 {_top['仍未派']} 次"))
                                else:
                                    _c2.markdown('⚠️ 目前找不到單一方案（人力或規則限制），點進去看詳情與補救方向。')
                            else:
                                _stt = ('running' if _bg_state(_cid, _own) == 'none' else _bg_state(_cid, _own)) if _own else 'none'
                                _c2.caption(('⏳ 分析中…　' if _stt == 'running' else '') + str(_r.get('未派原因') or '待核對')
                                            + ('' if _stt == 'running' else '　（點「處理這一案」後會算出原因與建議）'))
                            _c3.button('處理中 ▸' if _chosen_now == _cid else '處理這一案 →', key=f'pending_card_{_cid}', on_click=_choose_case, args=(_cid,),
                                       disabled=_chosen_now == _cid, type='secondary' if _chosen_now == _cid else 'primary', width='stretch')
                    download_button_quiet('下載未派案家清單 CSV', export_frame(_pending_clients).to_csv(index=False).encode('utf-8-sig'),
                                          '未派案家清單.csv', 'text/csv', key='download_pending_clients')
                else:
                    st.success('🎉 目前沒有未派案家。若想重新安排某個案家（例如前後移時間、拆成兩段），勾選下面「也列出已排班的案家」。')

                st.markdown('##### ② 選擇要處理的案家（也可以直接按上面卡片的「處理這一案」）')
                _show_all = st.checkbox('也列出已排班的案家（重新安排整案、拆成兩段、前後移動時間）', value=not _has_pending,
                                        key='pending_show_all', disabled=not _has_pending)
                _case_options = _pending_clients['案家ID'].astype(str).tolist() if _has_pending else []
                if _show_all or not _has_pending:
                    _case_options += [c for c in sorted(res['df_cl']['案家ID'].astype(str)) if c not in set(_case_options)]
                if not _case_options:
                    st.info('目前沒有案家資料。')
                    return
                if st.session_state.get('pending_client_choice') not in _case_options:
                    st.session_state['pending_client_choice'] = _case_options[0]
                _case = st.selectbox('案家', _case_options, format_func=lambda v: client_label(v), key='pending_client_choice')
                _edits_all = st.session_state.setdefault('pending_edits', {})
                _edits = _edits_all.setdefault(str(_case), empty_edits())
                _unit_minutes = unit_minutes_from_master(st.session_state.get('edit_service_code'))
                _case_people = res['df_cg']['居服員ID'].astype(str).tolist()
                _case_owner = str(config.client_primary_map.get(str(_case)) or '')
                _case_backups = tuple(str(x) for x in (config.client_backup_map.get(str(_case)) or ()))
                _merged_all = merge_task_client_tables(df_tasks, res['df_cl'])
                _client_tasks = _merged_all.loc[_merged_all['案家ID'].astype(str).eq(str(_case))]
                _need = client_profile(_client_tasks)
                _need_chips = [f"{_need['visits']} 次服務・約 {_need['hours']:g} 小時", f"每週 {_need['weekday_text']}",
                               f"{_need['earliest']}～{_need['latest']}", f"{_need['distinct_slots']} 種固定時段"]
                _need_chips += [f"需 {cert}" for cert in _need['certs']]
                _gender_chip = {'女': '限女性居服員', '男': '限男性居服員', '限女性': '限女性居服員', '限男性': '限男性居服員'}.get(_need['gender'])
                if _gender_chip:
                    _need_chips.append(_gender_chip)
                if _need['heavy']:
                    _need_chips.append('需重度移位體力')
                if _need['dementia']:
                    _need_chips.append('失智個案')
                chips(_need_chips, 'mint')
                _why = {cert: '、'.join(reasons) for cert, reasons in _need['certs'].items()}
                st.caption(
                    f"目前固定居服員：{caregiver_label(_case_owner) if _case_owner else '尚未指定'}"
                    + (f"　·　備援：{'、'.join(caregiver_label(x) for x in _case_backups)}" if _case_backups else '')
                    + ('　·　證照要求來源：' + '；'.join(f'{k}（{v}）' for k, v in _why.items()) if _why else ''))

                _screen_cache = res.setdefault('_screen_cache', {})
                _screen_key = (str(_case), route_signature)
                if _screen_key not in _screen_cache:
                    _screen_cache.clear()
                    _screen_cache[_screen_key] = screen_caregivers(_client_tasks, res['df_cg'], config, _case_owner, _case_backups,
                                                                    _operational_assignments, _merged_all)
                _screen, _screen_detail = _screen_cache[_screen_key]

                # ───────── ③ 處理：白話診斷 → 已驗證的方案 → 確認寫入 ─────────
                _base_rows = client_patterns(df_tasks, _case)
                _plabel = {str(r['模式ID']): f"{r['星期'].replace('星期', '每週')} {r['開始時間']}～{r['結束時間']}"
                           for r in _base_rows.to_dict('records')}
                _worker0 = _case_owner or next((str(r['居服員ID']) for _, r in _screen.iterrows() if r['_ok'] > 0), '')
                st.markdown(f'##### ③ 處理：{client_label(_case)}')
                _state = None
                if not _worker0:
                    st.warning('這個案家還沒有固定居服員，也沒有資格符合的人；請展開最下方「進階」或補充人力。')
                else:
                    try:
                        with st.spinner('正在核對整案並試算可行方案…'):
                            _state = _analysis(_case, _edits, _worker0)
                    except Exception as exc:
                        st.error('分析失敗：' + str(exc)[:160])
                    if _state is None:
                        st.info('⏳ 這個案家正在背景分析原因與解法，算好會自動顯示（通常幾秒到十幾秒）；期間可以先看上面的班表。')
                if _state is not None:
                    _pv = _state['preview']
                    _m1, _m2, _m3 = st.columns(3)
                    _m1.metric('這案服務筆數', _pv['total'], help='拆成兩段後，每一段各算一筆。')
                    _m2.metric('目前可排進', len(_pv['accepted']))
                    _m3.metric('排不進', len(_pv['failed']))
                    _who_now = _pv['accepted']['派單居服員'].astype(str).value_counts()
                    if len(_who_now):
                        st.caption('目前各人承接：' + '；'.join(f'{caregiver_label(str(w))} {n} 筆' for w, n in _who_now.items()))
                    if _pv['failed'].empty:
                        st.success('✅ 整案全部排得進去。' + ('確認沒問題就按下方「確認寫入班表」。' if _edits_has_any(_edits) else '（這個案家不需要處理；若想調整時間或分段，用下方「我想自己調整」。）'))
                    else:
                        st.markdown('**❗ 哪裡排不進、為什麼**')
                        for _g in _state['diag']:
                            with st.container(border=True):
                                st.markdown(f"**{_g['時段']}**　{_g['次數']} 次排不進（{'、'.join(_g['日期'])}）")
                                st.write('原因：' + _label_ids(_g['原因']))
                                st.caption('可以怎麼做：' + _g['方向'])
                        st.markdown('**💡 建議方案**（每個方案都已用完整規則驗證過）')
                        if _state['plans']:
                            for _i, _p in enumerate(_state['plans']):
                                with st.container(border=True):
                                    st.markdown(f"**方案 {_i + 1}**　{'＋'.join(_p['種類'])}")
                                    st.write(_label_ids(_p['標題']))
                                    st.caption('✅ 套用後，這案全部排得進去' if _p['仍未派'] == 0 else
                                               f"⚠️ 套用後仍有 {_p['仍未派']} 筆排不進（比現在少 {len(_pv['failed']) - _p['仍未派']} 筆），套用後會再給下一步建議")
                                    if _p['需確認']:
                                        st.caption('📞 ' + '；'.join(_p['需確認']))
                                    st.button('套用此方案', key=f'pending_plan_{_case}_{_i}', on_click=_edit_apply_many, args=(_case, _p['applies']),
                                              type='primary' if _i == 0 else 'secondary')
                        else:
                            st.warning('在目前的人力與規則下，系統試過「前後挪時間、換天、換人、拆成兩段」都沒有辦法讓它排進去。'
                                       '可以：① 增加居服員（例如補機動人員）② 到「02 智慧排班作業」的人工調整，把衝突的既有班改給別人 ③ 與案家協調服務時間。')
                    _lines = edits_summary(_edits, _plabel)
                    if _lines:
                        st.markdown('**目前已選的調整（還沒寫入班表）**')
                        for _line in _lines:
                            st.caption('・' + _label_ids(_line))
                        st.button('↩️ 全部重來', key=f'pending_edit_clear_{_case}', on_click=_edit_clear, args=(_case,))
                    if len(_pv['accepted']):
                        st.markdown('**④ 確認寫入班表**')
                        _simple_reason = reason_picker(f'pending_simple_reason_{_case}', '修改原因（必選，會寫進調整紀錄）')
                        _kinds_used = bool(_edits.get('shifts') or _edits.get('moves') or _edits.get('splits'))
                        _people_changed = bool(_edits.get('pattern_workers') or _edits.get('workers') or _edits.get('splits'))
                        _notes = (['要先和案家確認新的服務時間／分段'] if _kinds_used else []) + (['要通知新負責的居服員'] if _people_changed else [])
                        if _notes:
                            st.caption('📞 ' + '；'.join(_notes))
                        _go = ('✅ 確認寫入班表' if _pv['failed'].empty else f"先寫入可排的 {len(_pv['accepted'])} 筆（其餘 {len(_pv['failed'])} 筆仍待處理）")
                        if st.button(_go, type='primary', key=f'pending_simple_commit_{_case}'):
                            if not _simple_reason:
                                st.warning(REASON_HINT)
                            else:
                                try:
                                    _commit_case(_case, _pv, _simple_reason, _edits)
                                except Exception as exc:
                                    st.error('寫入失敗：' + str(exc))
                        st.caption('寫入後：這案服務變成「已派」，上方待處理清單與未派數字立刻更新。調整只存在這次操作中（重新按「開始排班」會從 Excel 重算），'
                                   '所以請同步下載最下方「整案調整紀錄」，照它修改 Excel。')

                with st.expander('✋ 我想自己調整（時間前後移、拆成兩段）— 系統沒建議、或想用自己的方式', expanded=False):
                    _edit_pid = st.selectbox('要調整哪個每週固定時段', list(_plabel), format_func=lambda p: _plabel.get(p, p), key=f'pending_edit_pid_{_case}')
                    _sample = next((t for t in _client_tasks.to_dict('records') if pattern_token(t) == _edit_pid), None)
                    _col_shift, _col_split = st.columns(2)
                    with _col_shift:
                        st.markdown('**⏱️ 整個時段前後移（服務分鐘不變）**')
                        _step = int(st.number_input('移動幾分鐘', min_value=5, max_value=240, value=30, step=5, key=f'pending_edit_step_{_case}'))
                        _b1, _b2 = st.columns(2)
                        _b1.button(f'後移 {_step} 分', key=f'pending_edit_later_{_case}', on_click=_edit_apply,
                                   args=(_case, {'type': 'shift', 'pid': _edit_pid, 'minutes': _step}))
                        _b2.button(f'前移 {_step} 分', key=f'pending_edit_earlier_{_case}', on_click=_edit_apply,
                                   args=(_case, {'type': 'shift', 'pid': _edit_pid, 'minutes': -_step}))
                        st.caption('例如 15:30～17:30 後移 30 分 → 16:00～18:00，仍是 120 分鐘。每次按都會累加，按完上方會立刻重新試算。')
                    with _col_split:
                        st.markdown('**✂️ 每次都拆成兩段，各自指派**')
                        _points = []
                        if _sample is not None:
                            try:
                                _points = split_points(pd.Series(_sample), _unit_minutes)
                            except ValueError as exc:
                                st.caption(f'無法拆段：{exc}')
                        if _points and _edit_pid in (_edits.get('splits') or {}):
                            st.caption('這個時段已拆成兩段；若要重拆，請先按上方「全部重來」。')
                        elif _points:
                            _total = int(round(float(_sample['服務歷時(分鐘)'])))
                            _dflt = default_split_minutes(pd.Series(_sample), _unit_minutes)
                            _first = st.selectbox('第一段長度', _points, index=_points.index(_dflt),
                                                  format_func=lambda m: f'第一段 {m} 分｜第二段 {_total - m} 分', key=f'pending_edit_first_{_case}_{_edit_pid}')
                            _owner_now = (_edits.get('pattern_workers') or {}).get(_edit_pid) or _worker0
                            _first_pick = st.selectbox('第一段負責居服員', _case_people, index=_case_people.index(_owner_now) if _owner_now in _case_people else 0,
                                                       format_func=caregiver_label, key=f'pending_edit_firstw_{_case}_{_edit_pid}')
                            _second_default = next((x for x in _case_people if x != _first_pick), _first_pick)
                            _second = st.selectbox('第二段負責居服員', _case_people, index=_case_people.index(_second_default),
                                                   format_func=caregiver_label, key=f'pending_edit_second_{_case}_{_edit_pid}')
                            _gap = 30 if _second == _first_pick else 0
                            st.caption('兩段同一位居服員：中間自動休息 30 分（整體結束時間延後 30 分）。' if _gap else '兩位各做一段：不必休息，原時段完整保留。')
                            st.button('✂️ 拆成兩段', key=f'pending_edit_split_{_case}_{_edit_pid}', on_click=_edit_apply,
                                      args=(_case, {'type': 'split', 'pid': _edit_pid, 'first': int(_first), 'gap': _gap,
                                                    'first_worker': _first_pick, 'second_worker': _second}))

                with st.expander('🔧 進階：自己挑人、逐位檢查、分給兩位以上（一般不需要打開）', expanded=False):
                    st.markdown('##### 進階 A．誰符合？逐位檢查資格')
                    _ok_total = int((_screen['結論'] == '✅ 資格全符合').sum())
                    st.caption(f"{len(_screen)} 位居服員中，{_ok_total} 位「資格」全部符合（證照、可排班星期、可服務時段、性別、體力、請假）。"
                               "接著要看與其他案家的時間衝突、休息、每日工時與月加班，請按下方「完整試算」。")
                    _table_cols = ['居服員ID', '結論', '資格可排', '證照／訓練', '排班日', '可服務時段', '性別／體力／環境', '請假／到期',
                                   '本月已排(時)', '加上此案(時)', '住家距離(公里)', '與此案關係']
                    show_table(_screen[_table_cols].rename(columns={'居服員ID': '居服員'}), hide_index=True, width='stretch',
                               column_config={
                                   '證照／訓練': st.column_config.TextColumn(width='medium'),
                                   '排班日': st.column_config.TextColumn(width='medium'),
                                   '可服務時段': st.column_config.TextColumn(width='medium'),
                                   '本月已排(時)': st.column_config.NumberColumn(help='目前已排給這位居服員的服務時數（不含轉場）'),
                                   '加上此案(時)': st.column_config.NumberColumn(help='接手整案後的服務時數；是否超過加班上限，請看下方完整試算'),
                                   '住家距離(公里)': st.column_config.NumberColumn(help='居服員住家到案家的直線距離'),
                               })
                    _blocked = _screen[_screen['阻擋原因'].ne('—')]
                    if not _blocked.empty:
                        with st.expander(f'查看 {len(_blocked)} 位居服員「不符合」的具體原因與次數'):
                            show_table(_blocked[['居服員ID', '結論', '資格可排', '阻擋原因']].rename(columns={'居服員ID': '居服員'}),
                                       hide_index=True, width='stretch')

                    st.markdown('##### 進階 B．完整試算：加入現有班表後真的排得進嗎？')
                    _eligible = [str(r['居服員ID']) for _, r in _screen.iterrows() if r['_ok'] > 0]
                    _cap = max(1, len(_eligible))
                    if _cap > 1:
                        _limit = st.slider('試算人數（依資格與距離由近到遠；預設全部資格符合者，避免漏看）', 1, _cap, _cap, key='pending_trial_n')
                    else:
                        _limit = 1
                    _trial_cache = res.setdefault('_trial_cache', {})
                    _trial_key = (str(_case), route_signature, _limit)
                    if st.button('🧪 完整試算', key='pending_trial_run', disabled=not _eligible,
                                 help='逐位用整案預覽核對：同日其他案家、交通與緩衝、連續工作與休息、每日工時、月加班。'):
                        _rows = []
                        _months = sorted({str(d)[:7] for d in _client_tasks['日期']})
                        _default_patterns = client_patterns(df_tasks, _case)
                        _bar = st.progress(0.0, text='試算中…')
                        for _n, _worker in enumerate(_eligible[:_limit], start=1):
                            _bar.progress((_n - 1) / _limit, text=f'試算 {caregiver_label(_worker)}（{_n}/{_limit}）…')
                            try:
                                _pv = preview_client_assignment(df_tasks, _operational_assignments, res['df_cg'], res['df_cl'], config,
                                                                _case, _worker, _default_patterns)
                                _row = simulate_worker(_pv, dict(config.work_day_minutes), _worker, _months, config)
                                _row['_accepted'] = frozenset(str(t) for t in _pv['accepted']['任務ID'])
                            except Exception as exc:
                                _row = {'居服員ID': _worker, '試算結果': '無法試算', '可安排': 0, '仍未派': _need['visits'],
                                        '總次數': _need['visits'], '月加班(時)': '—', '主要原因': str(exc)[:80], '_accepted': frozenset()}
                            _rows.append(_row)
                        _bar.empty()
                        _trial_cache.clear()
                        _trial_cache[_trial_key] = pd.DataFrame(_rows)
                    _trial = _trial_cache.get(_trial_key)
                    _total_visits = int(_need['visits'])
                    _status = worker_status_table(_screen, _trial, _total_visits)
                    _state_of = dict(zip(_status['居服員ID'], _status['狀態']))
                    _team_plan = None
                    if _trial is not None and not _trial.empty:
                        _best = _trial.sort_values(['可安排'], ascending=False, kind='stable').iloc[0]
                        if int(_best['可安排']) == int(_best['總次數']):
                            _full = [str(x) for x in _trial.loc[_trial['可安排'] == _trial['總次數'], '居服員ID']]
                            st.success(f"{len(_full)} 位居服員可以一個人接下整案全部 {_total_visits} 次：{'、'.join(caregiver_label(x) for x in _full[:6])}"
                                       + ('…' if len(_full) > 6 else ''))
                        else:
                            _plan_key = ('team', str(_case), route_signature, _limit)
                            _plan_cache = res.setdefault('_team_cache', {})
                            if _plan_key not in _plan_cache:
                                with st.spinner('沒有人能一人接下全部，正在組合多位居服員…'):
                                    _plan_cache.clear()
                                    _plan_cache[_plan_key] = plan_team(
                                        df_tasks, _operational_assignments, res['df_cg'], res['df_cl'], config, _case,
                                        client_patterns(df_tasks, _case),
                                        {str(r['居服員ID']): set(r['_accepted']) for _, r in _trial.iterrows()}, preferred=str(_best['居服員ID']))
                            _team_plan = _plan_cache[_plan_key]
                            if int(_best['可安排']) > 0:
                                st.warning(f"試算的人選都無法一個人接下全部；單人最多是 {caregiver_label(str(_best['居服員ID']))}，可排 {int(_best['可安排'])}/{_total_visits} 次。"
                                           + (f" 改由多位分工，系統組合出 {_team_plan['accepted']}/{_team_plan['total']} 次（見下方④）。" if _team_plan['accepted'] > int(_best['可安排']) else ''))
                            else:
                                st.error('試算的人選都排不進；請看「主要原因」，或回到上方「③ 處理」用建議方案（前後挪時間、換人、拆成兩段）。')
                    st.caption('狀態說明：✅ 全部可排　⚠️ 只能排其中幾次　❌ 資格符合但時間排不進　🔘 資格符合、尚未試算　⛔ 資格不符（證照、性別、排班日等）。下方「進階 C」的選單用同一套狀態。')
                    show_table(_status[['居服員ID', '狀態', '資格', '月加班(時)', '主要原因']].rename(columns={'居服員ID': '居服員', '資格': '資格可排'}),
                               hide_index=True, width='stretch', column_config={'主要原因': st.column_config.TextColumn(width='large')})

                    st.markdown('##### 進階 C．自己選人（可分給兩位以上），預覽並套用')
                    _order = [str(x) for x in _status['居服員ID']] + [p for p in _case_people if p not in set(_status['居服員ID'])]

                    def _worker_label(v):
                        return f"{caregiver_label(v)}　{_state_of.get(v, '')}"
                    _have_full = _trial is not None and not _trial.empty and bool((_trial['可安排'] == _trial['總次數']).any())
                    _modes = ['一位居服員包下整案', '分給兩位以上（依下表指定）']
                    _mode_key = 'pending_mode_' + str(_case)
                    _has_split = bool(_edits.get('splits') or _edits.get('workers') or _edits.get('pattern_workers'))
                    _default_multi = (_has_split or (_team_plan is not None and _team_plan['accepted'] > 0 and not _have_full))
                    _mode = st.radio('安排方式', _modes, index=1 if _default_multi else 0, horizontal=True, key=_mode_key + ('_m' if _default_multi else '_s'))
                    _multi = _mode == _modes[1] or _has_split
                    _suggest = _team_plan if (_multi and _team_plan is not None) else None
                    _main_default = (_suggest['main'] if _suggest and _suggest['main'] else (_order[0] if _order else ''))
                    _main_options = [x for x in _order if x]
                    _case_worker = st.selectbox('主責居服員（固定人選）', _main_options, index=_main_options.index(_main_default) if _main_default in _main_options else 0,
                                                format_func=_worker_label,
                                                key='pending_client_worker_' + str(_case) + ('_t' if _trial is not None else '_n') + ('_m' if _multi else '_s') + str(_main_default))
                    _main_state = _state_of.get(_case_worker, '')
                    if _main_state.startswith('⛔'):
                        st.error(f"{caregiver_label(_case_worker)} 資格不符，排不進這個案家：{str(_status.loc[_status['居服員ID'] == _case_worker, '主要原因'].iloc[0])}"
                                 "　（仍可按預覽看逐筆原因，但不會被排入。）")
                    elif _main_state.startswith('❌'):
                        st.warning(f"{caregiver_label(_case_worker)} 在目前班表上排不進任何一次；請改選其他人，或調整下表時段。")
                    _patterns_base = edited_pattern_table(df_tasks, _case, _edits)
                    _suggest_map = (_suggest or {}).get('pattern_workers', {})

                    def _default_worker(pid):
                        base = pid.split('#')[0]
                        if '#' in pid:
                            split = _edits['splits'].get(base, {})
                            picked = split.get('second_worker') if pid.endswith('#2') else split.get('first_worker')
                            if picked:
                                return picked
                        return ((_edits.get('pattern_workers') or {}).get(pid)
                                or ((_suggest_map.get(pid) or _suggest_map.get(base)) if _multi else None) or _case_worker)
                    _patterns_base['負責居服員'] = [_default_worker(str(t)) for t in _patterns_base['模式ID']]
                    _editor_key = (f"pending_client_patterns_{_case}_{_mode_key}_{_case_worker}_{abs(hash(tuple(sorted(_suggest_map.items()))))}"
                                   f"_{abs(hash(repr(_edits)))}")
                    if _multi:
                        st.caption('每一列是一個每週固定時段；在「負責居服員」欄指定誰來做（可全部同一人，也可分給不同人）。可修改開始／結束時間（維持服務分鐘）。')
                    else:
                        st.caption('所有星期與時段固定同一人；可修改下表起訖時間（維持服務分鐘）。')
                    _case_patterns = st.data_editor(
                        _patterns_base, hide_index=True, width='stretch', disabled=['模式ID', '來源列', '星期', '服務分鐘', '段'] + ([] if _multi else ['負責居服員']),
                        column_config={'負責居服員': st.column_config.SelectboxColumn('負責居服員', options=_main_options, required=True,
                                                                                  help='資格不符或排不進的人不會被排入，預覽會列出原因。')},
                        key=_editor_key)
                    _pattern_workers = ({str(r['模式ID']): str(r['負責居服員']) for _, r in _case_patterns.iterrows()}
                                        if _multi else {})

                    _task_workers = {**(dict((_suggest or {}).get('task_workers', {})) if _multi else {}), **(_edits.get('workers') or {})}
                    if _task_workers:
                        with st.expander(f'另有 {len(_task_workers)} 個個別日期由系統建議改派他人（固定時段內少數日期衝突）'):
                            show_table(pd.DataFrame([{'任務ID': k, '改派給': v} for k, v in sorted(_task_workers.items())]).rename(columns={'改派給': '居服員'}),
                                       hide_index=True, width='stretch')
                    _team_set = sorted({_case_worker, *_pattern_workers.values(), *_task_workers.values()})
                    _bad = [w for w in _team_set if _state_of.get(w, '').startswith('⛔')]
                    if _bad:
                        st.error('以下人選資格不符，不會被排入：' + '、'.join(caregiver_label(w) for w in _bad))
                    _case_reason = reason_picker('pending_client_reason_' + str(_case), '整案安排原因（必選）')
                    _case_preview_clicked = st.button('預覽整案排班', type='primary', key='pending_preview_btn')
                    _case_signature = hashlib.sha256(repr((_case, _case_worker, _case_patterns.to_dict('records'), sorted(_task_workers.items()),
                        _operational_assignments[['任務ID', '派單居服員']].to_dict('records'),
                        df_tasks.to_dict('records'), res['df_cg'].to_dict('records'),
                        config.client_primary_map, config.max_monthly_overtime_hours, config.buffer_mins, repr(_edits))).encode()).hexdigest()
                    if _case_preview_clicked:
                        try:
                            with st.spinner('核對整案日期、時段、資格、交通與工時…'):
                                _preview = preview_client_assignment(df_tasks, _operational_assignments, res['df_cg'],
                                    res['df_cl'], config, _case, _case_worker, _case_patterns, _pattern_workers, _task_workers,
                                    splits=_edits.get('splits'), moves=_edits.get('moves'), unit_minutes=_unit_minutes)
                            st.session_state['pending_client_preview'] = {'signature': _case_signature, 'result': _preview, 'reason': _case_reason or ''}
                        except Exception as exc:
                            st.session_state.pop('pending_client_preview', None)
                            st.error('整案預覽失敗：' + str(exc))
                    _stored = st.session_state.get('pending_client_preview')
                    if _stored and _stored['signature'] == _case_signature:
                        _preview = _stored['result']
                        _p1, _p2, _p3 = st.columns(3)
                        _p1.metric('整案服務次數', _preview['total'])
                        _p2.metric('可安排', len(_preview['accepted']))
                        _p3.metric('仍未派', len(_preview['failed']))
                        _who = _preview['accepted']['派單居服員'].astype(str).value_counts()
                        if len(_who):
                            st.caption('各人承接：' + '；'.join(f'{caregiver_label(str(w))} {n} 次' for w, n in _who.items()))
                        _months_now = sorted({str(d)[:7] for d in _client_tasks['日期']})
                        for _member in sorted(set(_who.index)):
                            st.caption(f'{caregiver_label(str(_member))} 月加班（接手前→後／上限）：' + month_overtime_summary(
                                _preview.get('ledger', {}), dict(config.work_day_minutes), str(_member), _months_now, config))
                        if not _preview['failed'].empty:
                            _grouped = _preview['failed'].assign(原因分類=_preview['failed']['原因'].map(lambda t: str(t).split('；')[0][:60]))
                            st.error('以下服務仍排不進去（資格不符、時間衝突、休息或工時超標的人不會被硬排）：')
                            show_table(_grouped.groupby('原因分類').agg(次數=('任務ID', 'size'), 例如日期=('日期', 'first')).reset_index().sort_values('次數', ascending=False),
                                       hide_index=True, width='stretch')
                            with st.expander('逐筆明細'):
                                show_table(_preview['failed'], hide_index=True, width='stretch')
                            st.info('建議方案請看上方「③ 處理」區；這裡是進階手動模式：可改選其他主責、在上表改開始時間，或到「02 智慧排班作業」的人工調整，把衝突的既有班改給別人。')
                        _apply_label = ('套用整案排班' if _preview['failed'].empty else
                                        f"套用可排的 {len(_preview['accepted'])} 次（其餘 {len(_preview['failed'])} 次仍待處理）")
                        if st.button(_apply_label, type='primary', key='apply_pending_client',
                                     disabled=_preview['accepted'].empty):
                            if not _case_reason:
                                st.warning(REASON_HINT)
                            else:
                                try:
                                    _commit_case(_case, _preview, _case_reason, _edits, main=_case_worker)
                                except Exception as exc:
                                    st.error('整案套用失敗：' + str(exc))
                _log_rows = res.get('whole_client_log') or []
                if _log_rows:
                    with st.expander(f'本次整案調整紀錄 · {len(_log_rows)} 筆（可用來同步更新原始 Excel）'):
                        _log_frame = pd.DataFrame(_log_rows)
                        _log_frame.insert(0, '#', range(1, len(_log_frame) + 1))
                        show_table(_log_frame, hide_index=True, width='stretch')
                        download_button_quiet('下載整案調整紀錄 CSV', _log_frame.to_csv(index=False).encode('utf-8-sig'),
                                              '整案調整紀錄.csv', 'text/csv', key='download_whole_client_log')
                        st.caption('原始 Excel 不會被自動改寫；若這些調整要長期保留，請照這份紀錄修改 Excel 後重新匯入。')

        if _module == M_ISSUES and _issue_view == '未派案家與替代方案':
            if st.session_state.get('whole_client_notice'):
                st.success(st.session_state.pop('whole_client_notice'))
            st.fragment(run_every=3 if _bg_busy else None)(_pending_panel)()
            if not _pending_clients.empty:
                with st.expander(f'查看 {len(_pending_report)} 次未派服務明細',expanded=False):
                    _diag_map = {str(d.get('任務ID')):d for d in _pending_diagnostics}
                    _pending_report['未派原因與處理提示'] = [_diag_map.get(str(t),{}).get('剔除原因','待核對') for t in _pending_report['任務ID']]
                    _pending_columns = [c for c in ['日期','任務ID','案家ID','時間窗_開始','時間窗_結束','Service_Code_1','未派原因與處理提示'] if c in _pending_report]
                    show_table(_pending_report[_pending_columns],hide_index=True,width='stretch')
                    download_button_quiet('下載未派任務診斷 CSV',export_frame(_pending_report[_pending_columns]).to_csv(index=False).encode('utf-8-sig'),
                                       '未派任務診斷.csv','text/csv',key='download_pending_diagnostics')

    override_log_df = load_override_log()
    st.session_state['override_apply_scope'] = '僅本次'

    # ---------- 04 排班結果與覆核：警示案件、已排與未排班次 ----------
    if _module == M_REVIEW:
        subhead('警示案件')
        _render_schedule_notes()
        sa10_reminders = calculate_task_revenue_and_salary(_operational_tasks, config)
        sa10_reminders = sa10_reminders[sa10_reminders['SA10計費提醒'].ne('')]
        if not sa10_reminders.empty:
            st.warning('SA10 單位與原服務時段有不一致資料；未自動延長服務時間，請核對。')
            show_table(sa10_reminders[['任務ID','案家ID','日期','SA10計費提醒']],hide_index=True,width='stretch')
        _issue_go_button('review_go_issues')

        # 全員每日總覽：包含無任務日，清楚區分計算範圍與班表檢視日期。
        if '日期' in df_tasks and not df_tasks.empty:
            _daily_tasks = df_tasks.drop_duplicates('任務ID').copy()
            _daily_tasks['_day'] = pd.to_datetime(_daily_tasks['日期'], errors='coerce').dt.normalize()
            _daily_tasks['_assigned'] = _daily_tasks['任務ID'].astype(str).isin(_assigned_ids).astype(int)
            _period_start = pd.Timestamp(res.get('week_start') or _daily_tasks['_day'].min()).normalize()
            _period_end = pd.Timestamp((res.get('range_end') if res.get('extended_to_end') else res.get('week_end')) or _daily_tasks['_day'].max()).normalize()
            if pd.notna(_period_start) and pd.notna(_period_end):
                _daily_summary = _daily_tasks.groupby('_day').agg(任務數=('任務ID','size'), 已派=('_assigned','sum'))
                _daily_summary = _daily_summary.reindex(pd.date_range(_period_start, _period_end), fill_value=0)
                _daily_summary['未派'] = _daily_summary['任務數'] - _daily_summary['已派']
                _daily_summary.insert(0, '日期', _daily_summary.index.strftime('%m/%d'))
                _daily_summary.insert(1, '星期', [['一','二','三','四','五','六','日'][d.weekday()] for d in _daily_summary.index])
                subhead('已排與未排班次（每日）')
                st.caption(f"已處理 {_period_start:%Y/%m/%d} ～ {_period_end:%Y/%m/%d}，共 {(_period_end-_period_start).days+1} 天；0 筆表示當天沒有匯入的服務任務。")
                show_table(_daily_summary.reset_index(drop=True), hide_index=True, width='stretch')

    # ---------- 02 智慧排班作業：日曆班表與人工調整 ----------
    if _module == M_SCHEDULE:
        # 月曆班表總覽：彙總既有派單結果與居服員資料表（並套用 overrides 顯示目前實際
        # 生效班表），不重呼叫任何排班演算法（見 calendar_view.py）。overrides／on_reassign／
        # on_list_candidates 供「月曆視角一鍵調班」（含空檔優先排序）使用。
        # 未派提示：原因與解法集中在 03，這裡只顯示數量並引導過去（不重複清單）。
        try:
            _eff02 = apply_overrides_to_result(df_result, overrides)
            _done02 = set(_eff02.loc[~_eff02['派單居服員'].astype(str).isin(['', '未指派', 'None', 'nan']), '任務ID'].astype(str))
            _left02 = df_tasks.loc[~df_tasks['任務ID'].astype(str).isin(_done02)]
            _n_task02, _n_client02 = len(_left02), _left02['案家ID'].astype(str).nunique()
        except Exception:
            _n_task02 = _n_client02 = 0
        if _n_task02:
            _g1, _g2 = st.columns([3, 1], vertical_alignment='center')
            _g1.markdown(f"<div class='cf-guide-card cf-guide-warn'><div class='cf-guide-title'>⚠️ 還有 {_n_client02} 個案家、{_n_task02} 筆服務沒排進</div>"
                         "<div class='cf-guide-body'>下一步：到「03 衝突與異常處理」看每一案排不進的原因，系統已備好建議方案（後移時間、換人、拆兩段），點選套用即可。</div></div>",
                         unsafe_allow_html=True)
            _g2.button('前往 03 處理 →', key='go_issues_top', type='primary', width='stretch', on_click=go_to_module, args=(M_ISSUES,))
        else:
            st.markdown("<div class='cf-guide-card cf-guide-ok'><div class='cf-guide-title'>✅ 全部服務都已排進</div>"
                        "<div class='cf-guide-body'>可以直接到「04 排班結果與覆核」檢查班表與下載。</div></div>", unsafe_allow_html=True)
        subhead('日曆班表')
        render_calendar_overview(
            df_result, df_tasks, res.get("df_cg", pd.DataFrame()),
            overrides=overrides,
            on_go_issues=lambda: go_to_module(M_ISSUES),
            on_reassign=_quick_reassign, on_list_candidates=_list_candidates,
            on_move=_calendar_move,
            period_start=res.get('week_start'),
            workflow_store=_workflow_store,
            client_names=_client_names if _show_names else {},
            caregiver_names=_caregiver_names if _show_names else {},
        )

        # 居督人工覆寫（Supervisor Override）：與月曆「一鍵調班」共用同一套覆寫介面與稽覈日誌。
        subhead('人工調整')
        st.caption(
            "當 AI 建議排單不符合實際場域狀況時（例如居服員臨時請假、案家臨時改期），"
            "居督可搜尋／選擇任務後於彈出視窗手動重新指定居服員（與日曆班表「一鍵調班」"
            "共用同一套覆寫互動介面）；每一筆變更皆會記錄原因並寫入稽覈日誌，可在「04 排班結果與覆核」查看。"
        )
        OVERRIDE_TRAVEL_ALERT_THRESHOLD = 3
        travel_reason_count = (
            override_log_df["變更原因"].map(reason_category).eq("交通／時間銜接不足").sum() if not override_log_df.empty else 0
        )
        if travel_reason_count >= OVERRIDE_TRAVEL_ALERT_THRESHOLD:
            st.warning(
                f"📈 稽覈日誌累計已有 {travel_reason_count} 筆修改原因為「交通／時間銜接不足」，"
                "建議到「01 資料匯入與排班設定」提高『車程扣分權重』，讓 AI 派單更優先考量就近指派。"
            )
        st.caption('代班僅套用本次服務；固定人選請假結束後仍由原人選服務。')
        st.caption('🧭 ' + SUBSTITUTE_RULE_TEXT)
        render_task_override_picker(
            df_tasks, df_result, res.get("df_cg", pd.DataFrame()), overrides,
            on_reassign=_quick_reassign, on_clear_override=_clear_override, on_list_candidates=_list_candidates,
        )
        for _notice in st.session_state.pop('consecutive_days_notices', []):
            st.warning('⚠️ ' + _notice)

        _prop_summary = st.session_state.get("last_override_propagation")
        if _prop_summary:
            _updated = int(_prop_summary.get("updated", 0))
            _replanned = int(_prop_summary.get("replanned", 0))
            _unassigned = int(_prop_summary.get("unassigned", 0))
            if _unassigned:
                st.warning(
                    f"固定班延續結果：{_updated} 筆沿用新居服員、"
                    f"{_replanned} 筆由系統重新找人、{_unassigned} 筆暫無可派人選。"
                )
            else:
                st.success(
                    f"固定班延續結果：{_updated} 筆沿用新居服員、"
                    f"{_replanned} 筆由系統重新找人。"
                )
            _prop_details = _prop_summary.get("details", [])
            if _prop_details:
                with st.expander("查看後續週期調整明細", expanded=False):
                    show_table(pd.DataFrame(_prop_details), width="stretch", hide_index=True)
        subhead('個案與居服員媒合、候選人推薦')

    # ---------- 05 人力與效益分析（主管） ----------
    if _module == M_ANALYTICS:
        module_header(M_ANALYTICS)
        st.caption(f'統計範圍：{_scope_label}。本頁只呈現管理數據，不會更動班表；排班與異常處理請回到 01～04。')
        subhead('服務產出與營運效益')
        _hours_total = pd.to_numeric(daily_payroll.get('服務工時', pd.Series(dtype=float)), errors='coerce').sum() if not daily_payroll.empty else 0.0
        _travel_hours = pd.to_numeric(daily_payroll.get('轉場工時', pd.Series(dtype=float)), errors='coerce').sum() if not daily_payroll.empty else 0.0
        _active_people = daily_payroll['居服員ID'].nunique() if not daily_payroll.empty and '居服員ID' in daily_payroll else 0
        _all_people = len(res['df_cg']) if res.get('df_cg') is not None else 0
        e1, e2, e3, e4 = st.columns(4)
        e1.metric("已排服務時數", f"{_hours_total:,.1f} 小時", help="已派服務的服務工時加總（不含轉場）。")
        e2.metric("人力配置", f"{_active_people} ／ {_all_people} 人",
                  help="有排到班的居服員／全部居服員。")
        e3.metric("人均服務時數", f"{(_hours_total / _active_people):,.1f} 小時" if _active_people else "—",
                  help="已排服務時數 ÷ 有排班的居服員人數。")
        e4.metric("轉場工時", f"{_travel_hours:,.1f} 小時", help="前案到下一案的轉場工時加總；首案通勤及末案返家不納入。")
        k4, k5, k6 = st.columns(3)
        k4.metric(
            ("已知營收小計（單價未齊）" if finance_partial else "預估服務總營收"), ("單價待補齊" if pd.isna(total_revenue) else f"{total_revenue:,.0f} 元"),
            help="已派服務的 Service_code 目前給付價格(元) × 服務次數加總；缺單價時僅顯示已知小計。",
        )
        k5.metric(
            '預估居服員總薪資', ('工時待確認' if pd.isna(total_salary) else f'{total_salary:,.0f} 元'),
            help='服務時數×時薪＋轉場次數×轉場工資＋一般工作日加班加給；與服務營收獨立。',
        )
        _margin_ok = not (finance_partial or pd.isna(total_revenue) or pd.isna(total_salary))
        k6.metric('營收扣除薪資', f"{(total_revenue - total_salary):,.0f} 元" if _margin_ok else '資料待補齊',
                  delta=(f"每服務小時 {(total_revenue - total_salary) / _hours_total:,.0f} 元" if _margin_ok and _hours_total else None),
                  delta_color='off', help='預估服務營收 − 預估居服員薪資；未含其他營運成本，僅供管理參考。')
        if finance_partial:
            st.warning('部分已派服務缺單價；營收顯示已知小計。薪資仍依服務工時與有效轉場獨立估算。')

        from dashboard_metrics import build_dashboard
        from dashboard_view import render_dashboard
        try:
            _dashboard = build_dashboard(df_tasks, _effective_for_kpi, _pending_diagnostics, task_payroll, daily_payroll,
                                         monthly_payroll, df_route_result, config.client_primary_map, caregivers=res.get('df_cg'))
            render_dashboard(_dashboard, caregiver_label, client_label)
        except Exception as exc:
            st.warning('儀錶板暫時無法產生，其他班表功能不受影響。原因：' + str(exc))

        subhead('居服員工時與薪資')
        with st.expander('每日與每月工資估算', expanded=False):
            st.caption('服務基本工資＋每次46元轉場工資＋加班加給。轉場與工作緩衝納入工時；首案通勤及末案返家不納入。此為一般工作日估算，休息日／國定假日薪資需另核對。')
            show_table(daily_payroll, hide_index=True, width='stretch')
            show_table(monthly_payroll, hide_index=True, width='stretch')
        subhead('轉場時間')
        st.caption('交通依有效派單的每日服務順序計算；首案由員工家／服務起點出發，後續由前一案出發。交通與緩衝分開顯示；不含末案返家及未匯入的其他行程。')
        with st.expander('每日交通總時間與逐段路線', expanded=False):
            if not df_route_result.empty:
                show_table(daily_travel, hide_index=True, width='stretch')
                show_table(df_route_result[['交通日期','派單居服員','任務ID','交通路段','預估車程(分)','路段緩衝(分)','含緩衝交通時間(分)','交通計算狀態']], hide_index=True, width='stretch')
                download_button_quiet('下載逐段交通明細 CSV', export_frame(df_route_result).to_csv(index=False).encode('utf-8-sig'), '逐段交通明細.csv', 'text/csv')
        module_nav_buttons(M_ANALYTICS)

    if _module != M_SCHEDULE:
        pass
    elif _effective_for_kpi.empty:

        st.info("目前無派單結果。")

    else:

        # 派單明細從有效指派建立，交通未計算也能顯示與匯出。
        df_display = _effective_for_kpi.copy()
        _task_lookup = df_tasks.drop_duplicates('任務ID').set_index('任務ID')
        for _col in ['案家ID', '任務優先級']:
            if _col in _task_lookup:
                _mapped = df_display['任務ID'].map(_task_lookup[_col])
                df_display[_col] = df_display[_col].combine_first(_mapped) if _col in df_display else _mapped
        if {'時間窗_開始', '時間窗_結束'}.issubset(_task_lookup.columns):
            df_display['服務時段'] = df_display['任務ID'].map(_task_lookup['時間窗_開始'].astype(str) + '-' + _task_lookup['時間窗_結束'].astype(str))
        _route_lookup = df_route_result.drop_duplicates('任務ID', keep='last').set_index('任務ID')
        for _col in ['交通路段', '預估車程(分)', '路段緩衝(分)', '含緩衝交通時間(分)']:
            df_display[_col] = df_display['任務ID'].map(_route_lookup[_col])

        has_date_for_display = "日期" in df_tasks.columns

        if has_date_for_display:

            date_map = df_tasks.set_index("任務ID")["日期"]

            df_display["日期"] = df_display["任務ID"].map(date_map)

            if "星期" in df_tasks.columns:

                weekday_map = df_tasks.set_index("任務ID")["星期"]

                df_display["星期"] = df_display["任務ID"].map(weekday_map)

            else:

                df_display["星期"] = df_display["日期"].apply(get_weekday_name)

        display_cols = [

            "任務ID", "案家ID", "派單居服員", "適配分數", "交通路段", "預估車程(分)", "路段緩衝(分)", "含緩衝交通時間(分)", "服務時段", "任務優先級",

            "預估服務營收", "預估居服員薪資", "原首選替換原因",

        ]

        if has_date_for_display:

            display_cols[1:1] = ["日期", "星期"]

        df_display = df_display.reindex(columns=display_cols)
        with st.expander('演算法指派明細（進階）', expanded=False):
            show_table(df_display, width="stretch", hide_index=True)

            csv = export_frame(df_display[display_cols]).to_csv(index=False).encode("utf-8-sig")

            download_button_quiet(

                "⬇️ 匯出指派結果 (assigned_results.csv)",

                csv,

                file_name="assigned_results.csv",

                mime="text/csv",

            )

        with st.expander("進階：單筆配對說明與適配分數", expanded=False):
            # ==========================================

            # Explainable AI：Why this schedule?

            # ==========================================

            st.markdown("##### 單筆配對說明（進階：適配分數）")

            st.caption(

                "查看 AI 為何推薦此居服員。適配分數由專長、照護連續性、"

                "歷史服務品質、交通與工作負荷等因素共同組成。"

            )

            selected_task = st.selectbox(

                "選擇任務查看 AI 推薦原因",

                df_result["任務ID"].astype(str).tolist(),

                key="explain_task_selector",

            )

            selected_row = df_result[

                df_result["任務ID"].astype(str) == str(selected_task)

            ].iloc[0]

            st.markdown(

                f"""

                **AI 推薦居服員：{selected_row['派單居服員']}**  

                **適配分數：{float(selected_row['適配分數']):.1f} 分**  

                **預估車程：{float(selected_row['預估車程(分)']):.1f} 分鐘**

                """

            )

            # ------------------------------

            # 人性化推薦理由

            # ------------------------------

            reasons = []

            if selected_row.get("專長匹配加分", 0) > 0:

                reasons.append("具備符合本案需求的核心照護專長")

            if selected_row.get("是否歷史首選", False):

                reasons.append("為案家歷史首選居服員，有助維持照護連續性")

            if selected_row.get("連續性品質加分", 0) > 0:

                reasons.append("歷史服務滿意度達設定門檻")

            if selected_row.get("交通扣分", 0) <= 10:

                reasons.append("預估轉場車程較短")

            if selected_row.get('本次計畫已排時數',0) > 0:
                reasons.append('已納入本次計畫累計派單時數，協助分配後續任務')
            if selected_row.get('疲勞資料狀態'):
                st.info(selected_row['疲勞資料狀態'])
            if selected_row.get('滿意度資料狀態'):
                st.caption('滿意度：'+selected_row['滿意度資料狀態'])
            if selected_row.get('待確認事項'):
                st.warning(selected_row['待確認事項'])

            if reasons:

                st.markdown("**主要推薦理由**")

                for reason in reasons:

                    st.write(f"✓ {reason}")

            # ------------------------------

            # 分數明細

            # ------------------------------

            explain_df = pd.DataFrame(

                {

                    "評分項目": [

                        "基礎分",

                        "專長匹配",

                        "歷史首選",

                        "連續性品質",

                        "滿意度調整",

                        "交通成本",

                        "已提供工時／本次計畫負荷",

                    ],

                    "分數影響": [

                        float(selected_row.get("基礎分", 0)),

                        float(selected_row.get("專長匹配加分", 0)),

                        float(selected_row.get("歷史首選加分", 0)),

                        float(selected_row.get("連續性品質加分", 0)),

                        float(selected_row.get("滿意度調整", 0)),

                        -float(selected_row.get("交通扣分", 0)),

                        -float(selected_row.get("工作負荷扣分", 0)),

                    ],

                }

            )

            st.markdown("**完整適配分數明細**")

            show_table(

                explain_df,

                width="stretch",

                hide_index=True,

            )

            st.caption(

                "所有候選居服員皆須先通過資格、工時、時段及時空衝突等硬性條件；"

                "上述分數僅用於合法可行候選方案之間的比較。"

            )



    if _module == M_SCHEDULE:
        module_nav_buttons(M_SCHEDULE)

    # ==========================================
    # 區塊④：動態事件處理（03 衝突與異常處理 › 臨時異動）
    # ==========================================

    @st.fragment
    def _changes_section():
        """只重跑這一區：臨時異動的每個操作不再牽動上方的月曆、儀表板與大型表格。"""
        st.session_state.setdefault("dynamic_insertions", [])
        st.session_state.setdefault("dynamic_candidate_result", None)
        st.session_state.setdefault("dynamic_candidate_task", None)
        st.session_state.setdefault("dynamic_adjustment_log", [])

        # 居服員臨時請假事件
        st.session_state.setdefault("leave_event_analysis", None)

        # 案家臨時改期事件
        st.session_state.setdefault("reschedule_event_analysis", None)

        # 前一服務延遲事件
        st.session_state.setdefault("delay_event_context", None)

        st.session_state['_workflow_cg_labels'] = {str(x):caregiver_label(str(x)) for x in res['df_cg']['居服員ID']}
        st.session_state['_workflow_cl_labels'] = {str(k):str(v)+'｜'+str(k) for k,v in (_client_names if _show_names else {}).items()}
        def _leave_owned_tasks(cg_id):
            """原本屬於這位居服員的服務：AI 結果排給他的，加上排班時就因他請假而由別人代班的（固定居服員是他）。"""
            cg_id = str(cg_id)
            owned = {str(t) for t, c in assigned_map.items() if str(c) == cg_id}
            if {'固定居服員ID', '是否請假代班'} <= set(df_result.columns):
                flagged = df_result[(df_result['固定居服員ID'].astype(str) == cg_id)
                                    & df_result['是否請假代班'].astype(str).str.lower().isin(['true', '1', '是'])]
                owned |= set(flagged['任務ID'].astype(str))
            return owned

        def _leave_current_map():
            """目前實際派單（含所有人工改派）；沒有人服務的回傳 None。"""
            _, effective = _workflow_frames()
            current = {}
            for tid, cg in zip(effective.get('任務ID', []), effective.get('派單居服員', [])):
                text = '' if cg is None or (isinstance(cg, float) and pd.isna(cg)) else str(cg)
                current[str(tid)] = None if text in ('', '未指派', 'None', 'nan') else text
            return current

        def _leave_cover_plan(cg_id, day=None, limit=60, top=5):
            """請假登錄後，逐筆列出該居服員當天原排的服務與可代班人選（依代班挑選順序，第一位為推薦），
            由居督確認或改選後才寫入班表。已有人服務的列出由誰代班。"""
            cg_id = str(cg_id)
            tasks_all, _ = _workflow_frames()
            ids_all = tasks_all['任務ID'].astype(str)
            dates = dict(zip(ids_all, pd.to_datetime(tasks_all['日期'], errors='coerce').dt.strftime('%Y-%m-%d')))
            starts = dict(zip(ids_all, tasks_all['時間窗_開始'].astype(str))) if '時間窗_開始' in tasks_all else {}
            mine = [t for t in _leave_owned_tasks(cg_id) if day is None or dates.get(t) == str(day)[:10]]
            mine = sorted(mine, key=lambda t: (dates.get(t) or '', starts.get(t, '')))[:limit]
            base_ids = [str(x) for x in res['df_cg']['居服員ID'] if str(x) != cg_id]
            items, done = [], []
            for tid in mine:
                current = _leave_current_map().get(tid)
                if current == cg_id:
                    continue          # 不在請假時段內，仍由本人服務
                tasks_now, effective_now = _workflow_frames()
                row = tasks_now[tasks_now['任務ID'].astype(str) == tid]
                client = str(row.iloc[0]['案家ID']) if not row.empty else ''
                info = row.iloc[0] if not row.empty else {}
                entry = {'任務ID': tid, '日期': dates.get(tid) or '',
                         '時段': f"{info.get('時間窗_開始', '')}–{info.get('時間窗_結束', '')}" if len(row) else '',
                         '案家': client_label(client), '原居服員': caregiver_label(cg_id)}
                if current is not None:
                    done.append({**entry, '代班居服員': caregiver_label(current), '結果': '已由代班服務'})
                    continue
                team = [str(x) for x in (config.client_backup_map.get(client) or ())]
                ids = [x for x in team if x in base_ids] + [x for x in base_ids if x not in team]
                try:
                    ranked = rank_candidates_by_availability(tid, ids, tasks_now, effective_now, res['df_cg'], config,
                                                             date_column='日期', df_cl=res.get('df_cl'))
                    options = [{'cg': str(r['cg_id']), 'reason': r.get('rank_reason', '')} for r in ranked if r.get('available')][:top]
                    error = '' if options else '找不到可派人選，請到「02 智慧排班作業」的人工調整處理'
                except Exception as exc:
                    options, error = [], f'候選計算失敗：{exc}'
                items.append({'entry': entry, 'options': options, 'error': error})
            return {'cg': cg_id, 'day': day or '', 'items': items, 'done': done}

        def _confirm_leave_plan(plan):
            """依居督在每筆服務選的人寫入班表（逐筆再檢查一次衝突）；回傳明細。"""
            rows = list(plan['done'])
            for item in plan['items']:
                entry, options = item['entry'], item['options']
                tid = entry['任務ID']
                choice = st.session_state.get(f"leave_pick_{tid}")
                if not options or choice in (None, '__skip__'):
                    rows.append({**entry, '代班居服員': '—', '結果': item['error'] or '居督暫不安排', '挑選理由': ''})
                    continue
                recommended = options[0]['cg']
                reason_text = next((o['reason'] for o in options if o['cg'] == choice), '')
                note = '居服員請假或臨時不可｜請假代班（居督確認' + ('推薦人選）' if choice == recommended else '，改選非推薦人選）')
                msg = _quick_reassign(tid, str(choice), note, apply_scope='僅本次')
                rows.append({**entry, '代班居服員': caregiver_label(choice) if msg is None else '—',
                             '結果': ('已代班（推薦人選）' if choice == recommended else '已代班（居督改選）') if msg is None else msg,
                             '挑選理由': _label_ids_plain(reason_text) if msg is None else ''})
            return rows

        _cover_cg = st.session_state.pop('auto_cover_cg', None)
        _cover_day = st.session_state.pop('auto_cover_day', None)
        if _cover_cg:
            with st.spinner('正在為受影響的服務整理可代班人選…'):
                _plan = _leave_cover_plan(_cover_cg, _cover_day)
            st.session_state['leave_cover_plan'] = _plan
            st.session_state.pop('leave_cover_report', None)
            _need = len(_plan['items'])
            st.session_state['workflow_notice'] = (
                f"{caregiver_label(str(_cover_cg))} {_cover_day or ''} 的請假已登錄：需代班 {_need} 筆"
                + (f"、{len(_plan['done'])} 筆已由代班服務" if _plan['done'] else '')
                + ('。請在下方每筆服務確認代班人選（已預選系統推薦，可改選）後按「確認代班」。' if _need else '。'))
        _restore = st.session_state.pop('restore_after_revoke', None)
        if _restore:
            from dispatch_workflow import overlap as _wf_overlap
            _r_cg = str(_restore['cg'])
            _r_tasks, _ = _workflow_frames()
            _r_rows, _restored = [], 0
            _r_tasks = _r_tasks.sort_values([c for c in ('日期', '時間窗_開始') if c in _r_tasks.columns])
            _r_owned, _r_now_map = _leave_owned_tasks(_r_cg), _leave_current_map()
            for _, _t in _r_tasks.iterrows():
                _tid = str(_t['任務ID'])
                if _tid not in _r_owned or not _wf_overlap(_t, _restore):
                    continue
                _now = _r_now_map.get(_tid)
                if _now == _r_cg:
                    continue          # 已經是原居服員
                _entry = {'任務ID': _tid, '日期': str(_t.get('日期', ''))[:10],
                          '時段': f"{_t.get('時間窗_開始', '')}–{_t.get('時間窗_結束', '')}",
                          '案家': client_label(str(_t.get('案家ID', ''))),
                          '代班居服員': caregiver_label(str(_now)) if _now else '待安排'}
                _msg = _quick_reassign(_tid, _r_cg, '居服員請假或臨時不可｜撤銷請假，還給原居服員')
                if _msg is None:
                    _restored += 1
                _r_rows.append({**_entry, '目前居服員': caregiver_label(_r_cg) if _msg is None else _entry['代班居服員'],
                                '結果': '已還給原居服員' if _msg is None else f'未還原：{_msg}'})
            st.session_state['leave_cover_report'] = {
                'title': '撤銷請假還原結果', 'cg': caregiver_label(_r_cg), 'day': str(_restore['start'])[:10], 'rows': _r_rows,
                'empty': '這段時間沒有代班中的服務，不需要還原。'}
            st.session_state['workflow_notice'] = (
                f'已撤銷 {caregiver_label(_r_cg)} {str(_restore["start"])[:10]} 的不可排時段：'
                f'{len(_r_rows)} 筆代班服務中，{_restored} 筆已還給原居服員'
                + (f'；{len(_r_rows) - _restored} 筆因衝突保留代班' if len(_r_rows) > _restored else '')
                + '。明細見下方表格。')
        if st.session_state.get('workflow_notice'):
            st.success(st.session_state.pop('workflow_notice'))
        for _notice in st.session_state.pop('consecutive_days_notices', []):
            st.warning('⚠️ ' + _notice)

        def _task_picker(prefix, tasks, assigned=None, keep_unassigned=False):
            """案家 → 日期 → 服務，逐層縮小範圍；不必在上千筆任務的單一下拉選單裡找。回傳任務ID，沒有可選時回傳 None。"""
            if tasks is None or tasks.empty or '任務ID' not in tasks.columns:
                return None
            keep = [c for c in ('任務ID', '案家ID', '日期', '時間窗_開始', '時間窗_結束') if c in tasks.columns]
            frame = tasks[keep].copy()
            frame['任務ID'] = frame['任務ID'].astype(str)
            if assigned is not None:
                frame['居服員'] = frame['任務ID'].map(assigned)
                if not keep_unassigned:
                    frame = frame[frame['居服員'].notna() & (frame['居服員'].astype(str) != '未指派')]
            if frame.empty:
                return None
            frame['_day'] = pd.to_datetime(frame.get('日期'), errors='coerce').dt.strftime('%Y-%m-%d').fillna('未定')
            frame['_client'] = frame['案家ID'].astype(str) if '案家ID' in frame.columns else ''
            left, middle, right = st.columns([1.3, 1, 1.7])
            client = left.selectbox('案家', sorted(frame['_client'].unique()), format_func=client_label, key=f'{prefix}_pick_client')
            owned = frame[frame['_client'] == client]
            day = middle.selectbox('日期', sorted(owned['_day'].unique()), key=f'{prefix}_pick_day_{client}')
            picked = owned[owned['_day'] == day].sort_values('時間窗_開始') if '時間窗_開始' in owned.columns else owned[owned['_day'] == day]
            labels = {}
            for row in picked.to_dict('records'):
                text = f"{row.get('時間窗_開始', '')}–{row.get('時間窗_結束', '')}"
                if row.get('居服員') is not None and pd.notna(row.get('居服員')) and str(row['居服員']) not in ('', '未指派', 'None'):
                    text += f"｜{caregiver_label(str(row['居服員']))}"
                elif assigned is not None:
                    text += "｜未指派"
                labels[row['任務ID']] = text
            return right.selectbox('服務', list(labels), format_func=labels.get, key=f'{prefix}_pick_task_{client}_{day}')

        _event_choices = ["➕ 臨時新增案件", "🚨 居服員臨時請假", "🕒 案家臨時改期", "⏱️ 前一服務延遲", "🩺 個案照顧需求變更"]
        _dynamic_event_type = big_tabs(_event_choices, "dynamic_event_type_selector", size="medium")

        def _infer_dispatch_year(tasks_df):
            """LINE 訊息只有月／日：以目前班表期間的年份為準。"""
            if "日期" in tasks_df.columns:
                years = pd.to_datetime(tasks_df["日期"], errors="coerce").dropna().dt.year.mode()
                if not years.empty:
                    return int(years.iloc[0])
            return int(pd.Timestamp.today().year)

        def _insertion_frames():
            """臨時新增用的任務表（已併入案家資料）與有效派單；已確認的臨時插單一併納入，避免互相撞班。"""
            tasks = merge_task_client_tables(df_tasks, res.get("df_cl", pd.DataFrame()))
            effective = apply_overrides_to_result(df_result, overrides)
            extra_tasks, extra_rows = [], []
            for item in st.session_state["dynamic_insertions"]:
                extra_tasks.append(item["task"])
                extra_rows.append({"任務ID": item["task"]["任務ID"], "派單居服員": item["cg_id"]})
            if extra_tasks:
                tasks = pd.concat([tasks, pd.DataFrame(extra_tasks)], ignore_index=True, sort=False)
                effective = pd.concat([effective, pd.DataFrame(extra_rows)], ignore_index=True, sort=False)
            return tasks, effective

        def _tc_service_codes():
            master = st.session_state["edit_service_code"]
            return [str(v).strip() for v in master["系統代碼"].dropna().tolist() if str(v).strip()]

        def _tc_parse():
            raw = st.session_state.get("tc_raw", "")
            if not raw.strip():
                return
            parsed = parse_dispatch_message(raw, _tc_service_codes(), _infer_dispatch_year(df_tasks))
            geo = get_geocoder().lookup(parsed["address"]) if parsed["address"] else GeoResult("空白地址", note="訊息中沒有地址。")
            stamp = datetime.now().strftime("%m%d%H%M%S")
            st.session_state["tc"] = {
                "meta": {k: v for k, v in parsed.items() if k != "tasks"},
                "slots": slots_to_frame(parsed["tasks"]),
                "geo": geo.to_dict(), "address": parsed["address"],
                "redacted": redact_message(raw), "prefix": f"TEMP{stamp}", "client_id": f"TEMP_CLIENT_{stamp}",
                "results": None,
            }
            if raw.strip() != str(st.session_state.get("tc_demo_message") or "").strip():
                st.session_state.pop("tc_demo_message", None)   # 貼的是真的派案訊息：不保留原文
            st.session_state["tc_raw"] = ""          # 解析完立即清掉原文，不留在畫面或工作階段（示範訊息另外保留原文顯示）

        def _tc_clear():
            st.session_state.pop("tc", None)
            st.session_state.pop("tc_demo_message", None)
            st.session_state["tc_raw"] = ""

        def _tc_relocate():
            tc = st.session_state.get("tc")
            if not tc:
                return
            address = st.session_state.get(f"tc_addr_{tc['prefix']}", tc["address"])
            tc["address"] = address
            tc["geo"] = get_geocoder().lookup(address).to_dict()
            tc["results"] = None

        def _tc_manual_coord():
            tc = st.session_state.get("tc")
            if not tc:
                return
            raw = st.session_state.get(f"tc_coord_{tc['prefix']}", "")
            found = re.findall(r"-?\d+(?:\.\d+)?", raw)
            if len(found) >= 2:
                lat, lon = float(found[0]), float(found[1])
                if lon < lat:          # 有人會貼成「經度, 緯度」
                    lat, lon = lon, lat
                if 20 <= lat <= 27.5 and 118 <= lon <= 123.5:
                    tc["geo"] = GeoResult("人工座標", lat, lon, 0, "人工輸入的 Google 地圖座標", "manual", "人工輸入座標").to_dict()
                    tc["results"] = None
                    return
            tc["geo_error"] = "座標格式請貼成「25.0123, 121.5012」（從 Google 地圖長按位置即可複製）。"

        def _render_new_case():
            geocoder = get_geocoder()
            service_codes = _tc_service_codes()
            if not geocoder.ready:
                st.warning("找不到離線門牌索引（data/ntpc_doorplate_index.npz）。仍可解析訊息，但需手動輸入座標；"
                           "請依說明執行 build_geo_index.py 建立索引。")

            st.markdown("##### ① 貼上 LINE 派案訊息")
            st.caption("整段貼上即可，姓名、門牌號維持遮蔽（O）沒關係。定位只用政府公開門牌資料，在這台主機上完成，"
                       "不會把地址送到 Google 或任何外部服務；解析完原文會立刻從畫面清掉。")
            _demo_message = st.session_state.get("tc_demo_message")
            if _demo_message:
                from html import escape as _esc
                st.markdown(
                    # 樣式直接寫在元素上（不依賴頁首樣式表），雲端部署版本也一定是藍底對話框；換行改用 <br>，不會被擠成一段。
                    "<div class='cf-line' style='background:linear-gradient(180deg,#9bb6dd,#8aa9d6);border-radius:16px;"
                    "padding:14px 16px 18px;margin:6px 0 10px;'>"
                    "<div class='cf-line-head' style='color:#fff;font-size:13px;font-weight:700;letter-spacing:.04em;margin-bottom:10px;'>"
                    "💬 LINE 派案群組　·　示範訊息（虛構，姓名與門牌已遮蔽）</div>"
                    "<div class='cf-line-bubble' style='display:inline-block;max-width:min(560px,100%);background:#fff;color:#1f2b2a;"
                    "border-radius:18px;padding:12px 16px;font-size:15px;line-height:1.75;box-shadow:0 2px 6px rgba(0,0,0,.12);margin-left:10px;'>"
                    + _esc(_demo_message).replace("\r", "").replace("\n", "<br>") + "</div></div>", unsafe_allow_html=True)
                if st.session_state.get("tc"):
                    st.caption("⬇️ 下方是系統從這段訊息讀出來的內容，請對照確認。")
                else:
                    st.caption("⬇️ 這是收到的原始訊息，也已放進下方輸入框；按「✨ 解析並定位」看系統讀出的資料。")
            st.text_area("派案訊息", height=150, key="tc_raw", label_visibility="collapsed",
                         placeholder="請問貴單位有人力嗎？\n姓名：劉O\n性別/年齡：女性82歲\n地址：新北市中和區大智街106巷O號\n失能等級：3\n失智：X　身障：X\n服務時間及項目：\n10/2(五)08:20~11:20 BA14+BA13\n備註：1.女性居服員…")
            b1, b2, _spacer = st.columns([1.3, 0.8, 4])
            b1.button("✨ 解析並定位", type="primary", key="tc_parse", on_click=_tc_parse, width="stretch")
            b2.button("🧹 清除", key="tc_clear", on_click=_tc_clear, width="stretch")

            tc = st.session_state.get("tc")
            if not tc:
                st.info("貼上訊息後按「解析並定位」，系統會一次帶出案主資料、位置與所有服務時段。")
                return
            meta, geo, prefix = tc["meta"], tc["geo"], tc["prefix"]

            # ---------- ② 確認
            st.markdown("##### ② 確認解析結果")
            with st.container(border=True):
                who = meta.get("name_masked") or "（未提供姓名）"
                facts = [who]
                if meta.get("client_gender") or meta.get("age"):
                    facts.append(f"{meta.get('client_gender', '')} {meta.get('age') or '?'} 歲")
                facts.append(f"失能 {meta['disability_level']} 級" if meta.get("disability_level") else "失能等級未提供")
                facts.append("失智：有" if meta.get("dementia") else "失智：無")
                facts.append("身障：有" if meta.get("disabled") else "身障：無")
                if meta.get("height_cm"):
                    facts.append(f"{meta['height_cm']} cm／{meta['weight_kg']} kg")
                st.markdown("**" + "　·　".join(facts) + "**")
                chips(meta.get("flags", []), "amber")
                if meta.get("splittable"):
                    st.caption("備註寫「可拆單位」：目前系統不會自動拆分；找不到人時，可在下表把一個時段拆成兩列各自找人。")
                if meta.get("summary") and meta["summary"].strip(". …"):
                    st.caption("概況：" + meta["summary"][:120])
                for warning in meta.get("warnings", []):
                    st.warning(warning)

                # 位置
                if geo.get("lat") is None:
                    st.error(f"📍 尚未定位：{geo.get('status')}。{geo.get('note', '')}　請在下方貼上 Google 地圖座標，或修改地址後重新定位。")
                elif geo["status"] == "精確門牌":
                    st.success(f"📍 精確門牌定位完成｜{geo['matched']}")
                elif geo["status"] == "人工座標":
                    st.success(f"📍 使用人工輸入座標｜{geo['lat']:.6f}, {geo['lon']:.6f}")
                else:
                    st.warning(f"📍 {geo['status']}｜{geo['matched']}｜實際位置約在此點 **{geo['radius_m']} 公尺**內。"
                               f"{geo.get('note', '')}")
                with st.expander("調整地址／位置", expanded=geo.get("lat") is None):
                    a1, a2 = st.columns([4, 1])
                    a1.text_input("服務地址", value=tc["address"], key=f"tc_addr_{prefix}")
                    a2.button("📍 重新定位", key=f"tc_reloc_{prefix}", on_click=_tc_relocate, width="stretch")
                    c1, c2 = st.columns([4, 1])
                    c1.text_input("或貼上 Google 地圖座標", key=f"tc_coord_{prefix}", placeholder="25.0123, 121.5012")
                    c2.button("使用座標", key=f"tc_usecoord_{prefix}", on_click=_tc_manual_coord, width="stretch")
                    if tc.get("geo_error"):
                        st.error(tc.pop("geo_error"))
                    st.caption("遮蔽門牌時只能定位到路段／巷弄，已用誤差範圍表示；如果居督知道精確位置，貼上座標會更準。")

            st.markdown("**服務時段**（可直接修改、取消勾選或新增列）")
            slots = st.data_editor(
                tc["slots"], key=f"tc_slots_{prefix}", hide_index=True, num_rows="dynamic", width="stretch",
                column_config={
                    "納入": st.column_config.CheckboxColumn("納入", default=True, width="small"),
                    "日期": st.column_config.DateColumn("日期", format="YYYY/MM/DD"),
                    "開始": st.column_config.TextColumn("開始", help="HH:MM", width="small"),
                    "結束": st.column_config.TextColumn("結束", help="HH:MM", width="small"),
                    "服務碼1": st.column_config.SelectboxColumn("服務碼 1", options=service_codes),
                    "Units1": st.column_config.NumberColumn("Units 1", min_value=0.0, step=1.0, width="small"),
                    "服務碼2": st.column_config.SelectboxColumn("服務碼 2", options=[""] + service_codes),
                    "Units2": st.column_config.NumberColumn("Units 2", min_value=0.0, step=1.0, width="small"),
                    "時段備註": st.column_config.TextColumn("時段備註"),
                },
            )

            with st.expander("照護條件（已自動帶入，可修改）", expanded=False):
                g1, g2, g3 = st.columns(3)
                gender_options = ["男女不拘", "女", "男"]
                gender = g1.selectbox("指定居服員性別", gender_options, index=gender_options.index(meta.get("caregiver_gender", "男女不拘")),
                                      key=f"tc_gender_{prefix}")
                need_options = ["一般照護"]
                if "特殊照護需求" in res.get("df_cl", pd.DataFrame()).columns:
                    need_options += [str(v) for v in res["df_cl"]["特殊照護需求"].dropna().unique() if str(v).strip()]
                if meta.get("special_need") not in need_options:
                    need_options.append(meta["special_need"])
                need = g2.selectbox("特殊照護需求", list(dict.fromkeys(need_options)), index=list(dict.fromkeys(need_options)).index(meta.get("special_need", "一般照護")),
                                    key=f"tc_need_{prefix}")
                heavy = g3.checkbox("需重度移位協助", value=False, key=f"tc_heavy_{prefix}")
                env = st.text_input("案家環境特徵", value="一般居家環境", key=f"tc_env_{prefix}")
                note = st.text_area("備註", value=meta.get("note", ""), key=f"tc_note_{prefix}", height=90)
                if need == "失智引導與精神陪伴" or meta.get("dementia"):
                    st.caption("失智個案只會配對具「失智症照顧專長」訓練的居服員。")

            # ---------- ③ 找人
            st.markdown("##### ③ 一次找出所有時段的人選")
            if st.button("🔎 找人（全部時段）", type="primary", key=f"tc_search_{prefix}", disabled=geo.get("lat") is None):
                case = {"client_id": tc["client_id"], "lat": geo["lat"], "lon": geo["lon"], "gender": gender, "need": need,
                        "heavy": heavy, "env": env, "note": note, "dementia": bool(meta.get("dementia"))}
                temp_tasks, errors = build_temp_tasks(case, slots, st.session_state["edit_service_code"].copy(), prefix,
                                                      "機構三檔" if input_mode == "機構資料（單一Excel／三個工作表）" else "原格式")
                if errors:
                    for message in errors:
                        st.error(message)
                elif temp_tasks.empty:
                    st.warning("沒有勾選任何服務時段。")
                else:
                    with st.spinner(f"正在檢查 {len(res['df_cg'])} 位居服員 × {len(temp_tasks)} 個時段…"):
                        tasks_now, effective_now = _insertion_frames()
                        try:
                            found = search_slots(temp_tasks, tasks_now, effective_now, res["df_cg"], config)
                            tc["results"] = {"slots": found, "case": case}
                        except Exception as exc:
                            tc["results"] = None
                            st.error(f"無法計算候選人：{exc}")

            results = tc.get("results")
            if not results:
                return
            found = results["slots"]
            total_slots = len(found)
            if geo.get("approximate"):
                st.caption(f"ℹ️ 位置為約略定位（±{geo['radius_m']} 公尺），交通時間以該點估算，實際可能略有差異。")

            plan = single_caregiver_plan(found)
            slot_options = []
            for item in found:
                frame = item["frame"]
                ok = frame[frame["可直接插入"] == True] if not frame.empty else frame   # noqa: E712
                slot_options.append(ok)
            slot_label = [f"{pd.Timestamp(i['task']['日期']):%m/%d} {i['task']['時間窗_開始']}-{i['task']['時間窗_結束']}" for i in found]

            summary_rows = []
            for label, ok in zip(slot_label, slot_options):
                summary_rows.append({"時段": label, "可直接插入人數": len(ok),
                                     "最佳人選": caregiver_label(str(ok.iloc[0]["居服員ID"])) if len(ok) else "—"})
            show_table(pd.DataFrame(summary_rows), hide_index=True, width="stretch")
            # 每個時段的推薦名單：依代班挑選順序（熟悉→前一案交通→本月工時→年資），列出從哪裡過來、交通分鐘與理由。
            for _i, (_label, _ok) in enumerate(zip(slot_label, slot_options)):
                if _ok.empty:
                    continue
                with st.expander(f"📋 {_label} 推薦名單（{len(_ok)} 位可排，第一位為系統推薦）", expanded=(_i == 0)):
                    _rec = _ok.head(8).reset_index(drop=True)
                    _view = pd.DataFrame({
                        '順位': ['⭐ 推薦' if n == 0 else str(n + 1) for n in range(len(_rec))],
                        '居服員': _rec['居服員ID'].astype(str).map(caregiver_label),
                        '從哪裡過來': _rec['前一任務（結束）'].map(lambda v: '家裡（當天第一案）' if str(v) in ('無', '', 'nan') else f'前一案 {v}'),
                        '交通(分)': _rec['代班排序理由'].map(lambda t: (re.search(r'過來 (?:約 )?(\d+) 分', str(t)) or [None, '—'])[1]),
                        '插單後當日工時': _rec['插單後排班工時(小時)'],
                        '挑選理由': _rec['代班排序理由'].map(_label_ids_plain),
                    })
                    show_table(_view, hide_index=True, width='stretch',
                               column_config={'挑選理由': st.column_config.TextColumn(width='large')})

            missing = [label for label, ok in zip(slot_label, slot_options) if ok.empty]
            for index, (label, item) in enumerate(zip(slot_label, found)):
                _n_ok = len(slot_options[index])
                _title = f"❌ {label} 找不到人選：為什麼？" if _n_ok == 0 else f"🔍 {label} 只有 {_n_ok} 位可插入，其他居服員為什麼不行？"
                if _n_ok == 0 or _n_ok <= 5:
                    with st.expander(_title, expanded=(_n_ok == 0)):
                        why = explain_slot(pd.Series(item["task"]), item["frame"], res["df_cg"], config)
                        counts = why["類別"].value_counts()
                        st.write("　".join(f"**{name}** {count} 人" for name, count in counts.items()))
                        show_table(why.sort_values("類別"), hide_index=True, width="stretch")
                        st.caption("最常見的原因就是要放寬的地方：例如調整時段、性別條件，或請居服員協助調班。")
            if missing:
                st.error(f"有 {len(missing)} 個時段目前沒有人可直接插入（{'、'.join(missing)}）。可修改上方時段後重新找人。")
                return

            # ---------- ④ 選人與確認
            st.markdown("##### ④ 選擇人選並確認")
            if not plan.empty:
                st.success(f"✅ 有 {len(plan)} 位居服員可以**一個人接下全部 {total_slots} 個時段**，符合「同案固定同一位」。")
                show_table(plan.head(5).rename(columns={"居服員ID": "居服員"}), hide_index=True, width="stretch")
                modes = ["整案交給同一位（建議）", "各時段分別選"]
            else:
                st.warning("沒有任何一位居服員能接下全部時段；請各時段分別選人（同案換人時，系統會把第一位當作固定居服員）。")
                modes = ["各時段分別選"]
            mode = st.radio("安排方式", modes, horizontal=True, key=f"tc_mode_{prefix}")

            chosen = {}
            if mode.startswith("整案"):
                ids = plan["居服員ID"].astype(str).tolist()
                pick = st.selectbox("固定居服員", ids, format_func=lambda v: f"{caregiver_label(v)}｜{_label_ids_plain(plan.set_index('居服員ID').at[v, '挑選理由'] or '')}",
                                    key=f"tc_one_{prefix}")
                chosen = {i: pick for i in range(total_slots)}
            else:
                for i, ok in enumerate(slot_options):
                    ids = ok["居服員ID"].astype(str).tolist()
                    score_col = "適配度分數" if "適配度分數" in ok.columns else ok.columns[3]
                    chosen[i] = st.selectbox(
                        slot_label[i], ids, key=f"tc_pick_{prefix}_{i}",
                        format_func=lambda v, ok=ok: f"{caregiver_label(v)}｜{_label_ids_plain(ok.set_index(ok['居服員ID'].astype(str)).at[v, '代班排序理由'] if '代班排序理由' in ok else '')}")
            reason_text = st.text_input("備註／原因（選填）", key=f"tc_reason_{prefix}", placeholder="例如：案家指定、服務連續性…")

            if st.button(f"📌 確認插入 {total_slots} 個時段", type="primary", key=f"tc_confirm_{prefix}"):
                committed, error = [], None
                for i, item in enumerate(found):
                    task = dict(item["task"])
                    cg_id = chosen[i]
                    tasks_chk, effective_chk = _workflow_frames()
                    for done in committed:           # 同一批前面幾個時段也要算進去
                        if done["task"]["任務ID"] not in set(tasks_chk["任務ID"].astype(str)):
                            tasks_chk = pd.concat([tasks_chk, pd.DataFrame([done["task"]])], ignore_index=True)
                        effective_chk = pd.concat([effective_chk, pd.DataFrame([{"任務ID": done["task"]["任務ID"], "派單居服員": done["cg_id"]}])], ignore_index=True)
                    if str(task["任務ID"]) not in set(tasks_chk["任務ID"].astype(str)):
                        tasks_chk = pd.concat([tasks_chk, pd.DataFrame([task])], ignore_index=True)
                    problem = check_reassignment_conflict(str(task["任務ID"]), cg_id, tasks_chk, effective_chk, res["df_cg"], config,
                                                          date_column="日期", df_cl=res.get("df_cl"))
                    if problem:
                        error = f"{slot_label[i]} 已不能由 {caregiver_label(cg_id)} 接：{problem}"
                        break
                    frame = found[i]["frame"]
                    detail = frame[frame["居服員ID"].astype(str) == cg_id]
                    committed.append({"task": task, "cg_id": cg_id, "detail": detail.iloc[0].to_dict() if not detail.empty else {}, "rank": i})
                if error:
                    st.error(error + "　請重新按「找人」。")
                else:
                    now = pd.Timestamp.now()
                    top_ids = " > ".join(plan["居服員ID"].astype(str).head(3)) if mode.startswith("整案") else ""
                    for done in committed:
                        task, cg_id = done["task"], done["cg_id"]
                        ok = slot_options[done["rank"]]
                        best = str(ok.iloc[0]["居服員ID"]) if len(ok) else cg_id
                        rank = (ok["居服員ID"].astype(str).tolist().index(cg_id) + 1) if cg_id in ok["居服員ID"].astype(str).tolist() else 0
                        st.session_state["dynamic_insertions"].append({
                            "task": task, "cg_id": cg_id, "confirmed_at": now, "candidate_rank": max(rank, 1),
                            "system_top1_id": best, "manual_reason": reason_text.strip(), "candidate_detail": done["detail"],
                        })
                        st.session_state["dynamic_adjustment_log"].append({
                            "時間戳記": now, "事件類型": "臨時新增案件", "任務ID": task["任務ID"], "案家ID": task["案家ID"],
                            "日期": task["日期"], "服務時段": f"{task['時間窗_開始']}-{task['時間窗_結束']}",
                            "Service_Code_1": task.get("Service_Code_1"), "Units_1": task.get("Units_1"),
                            "Service_Code_2": task.get("Service_Code_2"), "Units_2": task.get("Units_2"),
                            "定位模式": ("精確門牌" if geo["status"] == "精確門牌" else f"{geo['status']}（±{geo['radius_m']}m）"),
                            "系統Top3": top_ids or best, "系統第一名": best, "最終選擇居服員": cg_id, "最終候選排名": max(rank, 1),
                            "是否人工改選": "是" if rank > 1 else "否", "是否選擇較高擾動方案": "否", "人工改選原因": reason_text.strip(),
                            "既有班表是否移動": "否", "最低可選擾動成本": 0, "擾動成本": 0,
                            "原始派案文字": tc["redacted"],
                        })
                        _remember_owner(task, cg_id)
                        _workflow_store.remove_pending(str(task["任務ID"]))
                        _tasks_after, _effective_after = _workflow_frames()
                        _notice = _consecutive_days_notice(cg_id, task["日期"], _tasks_after, _effective_after)
                        if _notice:
                            st.session_state.setdefault("consecutive_days_notices", []).append(_notice)
                    # 批次插單完成即寫回持久帳本，供下一次（含跨月）排班／插單回溯「連續上班天數」。
                    _workflow_store.record_workdays((done["cg_id"], done["task"]["日期"]) for done in committed)
                    st.session_state["workflow_notice"] = f"✅ 已插入 {len(committed)} 個時段（{caregiver_label(committed[0]['cg_id'])}"\
                        + ("等" if len({c['cg_id'] for c in committed}) > 1 else "") + "）；既有班表沒有被移動，並已建立稽覈紀錄。"
                    st.session_state.pop("tc", None)
                    st.rerun()

        def _render_inserted():
            if st.session_state["dynamic_insertions"]:
                st.markdown("##### 本次已確認的臨時插單")
                rows = [{"任務ID": i["task"].get("任務ID"), "案家ID": i["task"].get("案家ID"), "日期": i["task"].get("日期"),
                         "服務時段": f"{i['task'].get('時間窗_開始')}-{i['task'].get('時間窗_結束')}",
                         "派單居服員": i["cg_id"], "候選排名": i.get("candidate_rank", 1)} for i in st.session_state["dynamic_insertions"]]
                show_table(pd.DataFrame(rows), width="stretch", hide_index=True)


        # ---------- 🎬 示範案例：從目前班表自動找一個「找得到代班人選」的情境，並預先帶入下方欄位 ----------
        _demo_kinds = {"➕ 臨時新增案件": "new", "🚨 居服員臨時請假": "leave", "🕒 案家臨時改期": "reschedule", "⏱️ 前一服務延遲": "delay",
                       "🩺 個案照顧需求變更": "needchange"}
        _demo_kind = _demo_kinds.get(_dynamic_event_type)
        _demo_buttons = ([('new1', '🎬 範例1：劉O 就醫陪同（單次）'), ('new2', '🎬 範例2：樓O志 失智（三次）')] if _demo_kind == 'new'
                         else [('dementia', '🎬 範例：DEMO0005 新增失智'), ('foot', '🎬 範例：DEMO0004 新增足部照護')] if _demo_kind == 'needchange'
                         else [(_demo_kind, '🎬 載入示範案例')])

        def _demo_context():
            import demo_cases as _dc
            _cfg = replace(config, client_primary_map=dict(config.client_primary_map),
                           client_backup_map=dict(getattr(config, 'client_backup_map', None) or {}),
                           task_support_map=dict(getattr(config, 'task_support_map', None) or {}),
                           primary_profiles=dict(getattr(config, 'primary_profiles', None) or {}),
                           work_day_minutes=dict(config.work_day_minutes), dispatch_cache={})
            _effective = apply_overrides_to_result(df_result, overrides)
            ctx = _dc.Context(df_tasks, _effective, res['df_cg'], res['df_cl'], _cfg)
            ctx.service_master = st.session_state.get('edit_service_code')
            return _dc, ctx

        def _apply_demo(kind, found):
            """把示範情境寫進下方各流程的欄位（元件尚未畫出，可以直接設定）。"""
            if kind == 'leave':
                st.session_state['leave_form_cg'] = found['worker']
                st.session_state['leave_form_day'] = pd.Timestamp(found['day']).date()
                st.session_state['leave_form_mode'] = '整天'
                st.session_state['leave_form_reason'] = '請假'
                picks = '；'.join(f"{p['任務ID']} → {caregiver_label(p['代班'])}（{_label_ids_plain(p['理由'])}；共 {p.get('可選人數', 1)} 位可選）" for p in found['picks'])
                return (f"示範：{caregiver_label(found['worker'])} 在 {found['day']} 全天請假，原排 {len(found['tasks'])} 筆服務。"
                        f"下方已帶入，請按「登錄不可排時間」，系統會列出每筆服務的代班人選讓你確認或改選。推薦人選：{picks}。"
                        "（會登錄一筆真實請假紀錄，測完可在下方「已登錄不可排時段」撤銷並還原。）")
            if kind == 'reschedule':
                client, day, tid = str(found['client']), found['day'], str(found['task'])
                st.session_state['reschedule_pick_client'] = client
                st.session_state[f'reschedule_pick_day_{client}'] = day
                st.session_state[f'reschedule_pick_task_{client}_{day}'] = tid
                st.session_state['reschedule_new_date'] = pd.Timestamp(day).date()
                st.session_state['reschedule_new_start'] = pd.Timestamp('2000-01-01 ' + found['new_start']).time()
                st.session_state['reschedule_new_end'] = pd.Timestamp('2000-01-01 ' + found['new_end']).time()
                st.session_state['reschedule_event_analysis'] = None
                return (f"示範：{client_label(client)} 在 {day} 原本 {found['old']} 的服務，案家要求改到 {found['new_start']}-{found['new_end']}。"
                        f"下方已帶入，請按「🔎 暫時移除原任務並重新找候選人」。預計第一順位：{caregiver_label(found['top'])}"
                        f"（{_label_ids_plain(found['reason'])}）。")
            if kind == 'delay':
                worker, client, day, tid = found['worker'], str(found['client']), found['day'], str(found['task'])
                st.session_state['delay_pick_cg'] = worker
                st.session_state[f'delay_{worker}_pick_client'] = client
                st.session_state[f'delay_{worker}_pick_day_{client}'] = day
                st.session_state[f'delay_{worker}_pick_task_{client}_{day}'] = tid
                st.session_state['delay_event_minutes'] = int(found['delay'])
                st.session_state['delay_event_context'] = None
                return (f"示範：{caregiver_label(worker)} 在 {day} 於 {client_label(client)} 的服務（原 {found['end']} 結束）延遲 {found['delay']} 分鐘，"
                        f"會趕不上下一案 {found['affected']}。下方已帶入，請按「🔎 確認延遲並檢查後續班」。"
                        f"預計第一順位代班：{caregiver_label(found['top'])}（{_label_ids_plain(found['reason'])}）。")
            if kind in ('dementia', 'foot', 'needsplit'):
                client = str(found['client'])
                st.session_state['qual_need'] = found['need']
                st.session_state['qual_client'] = client
                st.session_state['qual_since'] = pd.Timestamp(found['since']).date()
                st.session_state['qual_method'] = 'now'
                st.session_state.pop('qual_avail', None)
                if found.get('split'):
                    _plan = '、'.join(f"{g['星期']}→{caregiver_label(g['居服員'])}" for g in found['split']['groups'])
                    st.session_state['qual_method'] = 'now'
                    return (f"示範：{client_label(client)} 一直由 {caregiver_label(found['worker'])} 服務（本期 {len(found['tasks'])} 次），"
                            f"照專複評後新增{'失智' if found['need'] == 'dementia' else ' BA08 足部照護'}，原居服員沒有訓練。"
                            f"有訓練的居服員裡沒有一位能接下全部 {len(found['tasks'])} 次（時段撞到其他案家或工時），"
                            f"所以系統依星期分給 {found['options']} 位：{_plan}。下方已選「立即改派」（法規只能派給有訓練的人）。")
                _story = ('照專複評後確認失智，但 {w} 沒有失智症照顧服務 20 小時訓練' if kind == 'dementia' else
                          '照專複評後核定新增 BA08 足部照護（例如糖尿病足、趾甲增厚），但 {w} 沒有足部照護訓練')
                return (f"示範：{client_label(client)} 一直由 {caregiver_label(found['worker'])} 服務（本期 {len(found['tasks'])} 筆），"
                        + _story.format(w=caregiver_label(found['worker'])) + "。下方已帶入需求、案家與生效日，"
                        f"法規只能派給有訓練的人，預設「立即改派」；也可加上原居服員陪同交接，或追蹤原居服員補訓後回任。"
                        f"有訓練又能接下全部服務的有 {found['options']} 位，推薦 {caregiver_label(found['top'])}"
                        f"（{_label_ids_plain(found['reason'])}）。")
            if kind in ('new1', 'new2'):
                # 先只把整段 LINE 訊息放進輸入框並顯示原文，讓居督看到「收到什麼」；按「解析並定位」後才出現解析結果。
                st.session_state["tc_raw"] = found['message']
                st.session_state["tc_demo_message"] = found['message']
                st.session_state.pop("tc", None)
                when = '、'.join(f"{pd.Timestamp(d):%m/%d}({'一二三四五六日'[pd.Timestamp(d).weekday()]}) {a}-{b}" for d, a, b in found['slots'])
                tops = '；'.join(f"{pd.Timestamp(d):%m/%d} → {caregiver_label(t)}（{_label_ids_plain(r)}）"
                                for (d, _, _), t, r in zip(found['slots'], found['top'], found['reason']) if t)
                return (f"已載入 LINE 派案範例{found['example']}（日期改為 {when}，其他內容不變）。下方先顯示收到的整段 LINE 訊息；"
                        f"請按「✨ 解析並定位」看系統讀出的資料，再到「③ 一次找出所有時段的人選」按「找人」。"
                        f"各時段可排人數：{'、'.join(str(n) for n in found['per_slot'])} 位；預計第一順位：{tops}。")
            return ''

        if _demo_kind:
            _cols = st.columns([1.2] * len(_demo_buttons) + [3.2 - 1.2 * (len(_demo_buttons) - 1)], vertical_alignment='center')
            _cols[-1].caption('不知道要測哪一筆？按左邊按鈕，系統會從目前班表找一個「一定找得到代班人選」的情境並幫你填好欄位；'
                              '會優先挑第一順位是「從前一案過來」的人，方便確認交通看的是前一案位置。')
            for _col, (_kind, _label) in zip(_cols, _demo_buttons):
                if _col.button(_label, key=f'demo_load_{_kind}', width='stretch'):
                    try:
                        with st.spinner('從目前班表尋找可示範的日期與時段（會用同一套規則試算候選人，約數十秒）…'):
                            _dc, _ctx = _demo_context()
                            _finder = {'leave': _dc.find_leave, 'reschedule': _dc.find_reschedule, 'delay': _dc.find_delay,
                                       'dementia': _dc.find_dementia,
                                       'foot': lambda c: _dc.find_need_change(c, 'foot', prefer='DEMO0004'),
                                       'needsplit': lambda c: (_dc.find_need_change(c, 'dementia', want_split=True, budget=40.0)
                                                               or _dc.find_need_change(c, 'foot', want_split=True, budget=20.0)), 'new1': lambda c: _dc.find_line_example(c, 1), 'new2': lambda c: _dc.find_line_example(c, 2)}[_kind]
                            _found = _finder(_ctx)
                    except Exception as exc:
                        _found = None
                        st.error('示範案例搜尋失敗：' + str(exc))
                    if _kind in ('dementia', 'foot'):
                        st.session_state['qual_need'] = _kind      # 先切到按下的需求（就算找不到範例，畫面也不會停在另一種）
                    if _found:
                        st.session_state.setdefault('demo_notices', {})[_demo_kind] = _apply_demo(_kind, _found)
                        st.rerun()
                    else:
                        st.session_state.get('demo_notices', {}).pop(_demo_kind, None)   # 清掉上一個範例的說明，避免看起來像這次的結果
                        if _kind in ('dementia', 'foot'):
                            st.session_state.pop('qual_client', None)
                        st.warning('目前班表找不到合適的示範情境（例如期間內沒有可接的人，或門牌定位資料不存在）。可以換個排班期間後再試。')
            _notice = st.session_state.get('demo_notices', {}).get(_demo_kind)
            if _notice:
                st.info('🎬 ' + _notice)

        if _dynamic_event_type == "➕ 臨時新增案件":
            _render_new_case()
            _render_inserted()

        elif _dynamic_event_type == "🚨 居服員臨時請假":
            _wf_tasks, _wf_result = _workflow_frames()
            _plan = st.session_state.get('leave_cover_plan')
            if _plan and _plan['items']:
                subhead(f"請確認代班人選｜{caregiver_label(_plan['cg'])} {_plan['day']} 請假")
                st.caption('每筆服務已預選系統推薦人選（⭐，依代班順序排第一）；如果有更適合的人可以改選，按「確認代班」才會寫入班表。')
                for _item in _plan['items']:
                    _e = _item['entry']
                    with st.container(border=True):
                        st.markdown(f"**{_e['日期']} {_e['時段']}**　{_e['案家']}　·　原居服員 {_e['原居服員']}")
                        if not _item['options']:
                            st.warning(_item['error'])
                            continue
                        _opts = [o['cg'] for o in _item['options']] + ['__skip__']
                        _why = {o['cg']: o['reason'] for o in _item['options']}
                        _first = _opts[0]
                        st.selectbox(
                            f"代班人選（{len(_item['options'])} 位可選）", _opts, index=0, key=f"leave_pick_{_e['任務ID']}",
                            format_func=lambda v, _why=_why, _first=_first: ('先不安排（保留待處理）' if v == '__skip__' else
                                (('⭐ 推薦｜' if v == _first else '') + caregiver_label(v) + '　—　' + _label_ids_plain(_why.get(v, '')))))
                        if len(_item['options']) > 1:
                            show_table(pd.DataFrame([{'順位': '⭐ 推薦' if n == 0 else str(n + 1), '居服員': caregiver_label(o['cg']),
                                                      '挑選理由': _label_ids_plain(o['reason'])} for n, o in enumerate(_item['options'])]),
                                       hide_index=True, width='stretch')
                _b1, _b2 = st.columns([1, 1])
                if _b1.button('✅ 確認代班', type='primary', key='leave_plan_confirm', width='stretch'):
                    with st.spinner('寫入代班並逐筆檢查衝突…'):
                        _rows = _confirm_leave_plan(_plan)
                    st.session_state['leave_cover_report'] = {'cg': caregiver_label(_plan['cg']), 'day': _plan['day'], 'rows': _rows}
                    st.session_state.pop('leave_cover_plan', None)
                    st.rerun()
                if _b2.button('稍後再處理', key='leave_plan_later', width='stretch'):
                    st.session_state.pop('leave_cover_plan', None)
                    st.rerun()
            elif _plan is not None and not _plan['items']:
                st.session_state['leave_cover_report'] = {'cg': caregiver_label(_plan['cg']), 'day': _plan['day'], 'rows': _plan['done']}
                st.session_state.pop('leave_cover_plan', None)
            _leave_report = st.session_state.get('leave_cover_report')
            if _leave_report:
                st.markdown(f"**{_leave_report.get('title', '最近一次請假代班結果')}｜{_leave_report['cg']} {_leave_report.get('day', '')}**")
                if _leave_report.get('rows'):
                    show_table(pd.DataFrame(_leave_report['rows']), width='stretch', hide_index=True)
                else:
                    st.info(_leave_report.get('empty', '這位居服員在請假日期沒有原排的服務，不需要代班。'))
                if st.button('關閉代班結果', key='leave_cover_report_close'):
                    st.session_state.pop('leave_cover_report', None)
                    st.rerun()
            render_leave(_workflow_store, res['df_cg'], _wf_tasks, _wf_result, overrides)

        elif _dynamic_event_type == "🕒 案家臨時改期":
            # ==========================================
            st.subheader("🕒 案家臨時改期")
            st.caption(
                "選擇既有任務後，系統會先暫時移除原任務，再用新的日期／時間重新做 insertion。"
                "其他既有任務不移動；若新時段沒有可直接插入的人選，再交由後續局部重排處理。"
            )
            st.caption(
                "工時以實際排班時段計算（例如 08:00–11:00 = 3 小時）；"
                "Service Code Master 分鐘數另作服務碼／申報用途。"
            )

            _effective_result_for_reschedule = apply_overrides_to_result(df_result, overrides)

            _resched_assigned = None
            if not _effective_result_for_reschedule.empty:
                _resched_unique = _effective_result_for_reschedule.drop_duplicates('任務ID')
                _resched_assigned = pd.Series(_resched_unique['派單居服員'].values, index=_resched_unique['任務ID'].astype(str).values)
            _reschedule_task_id = _task_picker('reschedule', df_tasks, assigned=_resched_assigned, keep_unassigned=True)
            _resched_orig_cg = ''
            if _reschedule_task_id is not None and _resched_assigned is not None:
                _resched_orig_cg = _resched_assigned.get(str(_reschedule_task_id))
                _resched_orig_cg = '' if _resched_orig_cg is None or pd.isna(_resched_orig_cg) or str(_resched_orig_cg) in ('未指派', 'None') else str(_resched_orig_cg)
            if _reschedule_task_id is None:
                st.info("目前沒有可供改期的既有任務。")
            else:
                _orig_match = df_tasks[
                    df_tasks["任務ID"].astype(str) == str(_reschedule_task_id)
                ].copy()

                if _orig_match.empty:
                    st.warning("找不到原任務資料。")
                else:
                    _orig_task = _orig_match.iloc[0].copy()
                    _orig_date = pd.to_datetime(_orig_task.get("日期"), errors="coerce")
                    if pd.isna(_orig_date):
                        _orig_date = pd.Timestamp.today()

                    def _reschedule_to_time(value, fallback):
                        if hasattr(value, "hour") and hasattr(value, "minute"):
                            return value
                        try:
                            return pd.to_datetime(str(value)).time()
                        except Exception:
                            return fallback

                    _orig_start = _reschedule_to_time(
                        _orig_task.get("時間窗_開始"),
                        pd.Timestamp("08:00").time(),
                    )
                    _orig_end = _reschedule_to_time(
                        _orig_task.get("時間窗_結束"),
                        pd.Timestamp("09:00").time(),
                    )

                    rc1, rc2, rc3 = st.columns(3)
                    _new_date = rc1.date_input(
                        "新日期",
                        value=_orig_date.date(),
                        key="reschedule_new_date",
                    )
                    _new_start = rc2.time_input(
                        "新開始時間",
                        value=_orig_start,
                        key="reschedule_new_start",
                    )
                    _new_end = rc3.time_input(
                        "新結束時間",
                        value=_orig_end,
                        key="reschedule_new_end",
                    )

                    st.caption(
                        f"原排程：{_orig_date.strftime('%Y-%m-%d')} "
                        f"{_orig_task.get('時間窗_開始', '')}-"
                        f"{_orig_task.get('時間窗_結束', '')}"
                        + (f"｜原居服員 {caregiver_label(_resched_orig_cg)}" if _resched_orig_cg else "｜目前未指派")
                    )

                    def _date_cell(frame, day):
                        """依欄位型別寫入日期：pandas 3 的字串欄位不接受 Timestamp，會丟 TypeError。"""
                        day = pd.Timestamp(day)
                        if "日期" in frame.columns and pd.api.types.is_datetime64_any_dtype(frame["日期"]):
                            return day
                        return day.strftime("%Y-%m-%d")

                    if st.button(
                        "🔎 暫時移除原任務並重新找候選人",
                        type="primary",
                        key="reschedule_analyze",
                    ):
                        if _new_end <= _new_start:
                            st.error("新結束時間必須晚於新開始時間。")
                        else:
                            try:
                                _new_task = _orig_task.copy()
                                _new_task["日期"] = _date_cell(df_tasks, _new_date)
                                _new_task["時間窗_開始"] = _new_start.strftime("%H:%M")
                                _new_task["時間窗_結束"] = _new_end.strftime("%H:%M")

                                # 原任務暫時從既有派單結果移除，避免自己與自己衝突。
                                _base_result = _effective_result_for_reschedule[
                                    _effective_result_for_reschedule["任務ID"].astype(str)
                                    != str(_reschedule_task_id)
                                ].copy()

                                # 任務表保留同一任務ID，但把日期／時間換成新時段。
                                _tasks_changed = df_tasks.copy()
                                _mask = (
                                    _tasks_changed["任務ID"].astype(str)
                                    == str(_reschedule_task_id)
                                )
                                if "日期" in _tasks_changed.columns:
                                    _tasks_changed.loc[_mask, "日期"] = _date_cell(_tasks_changed, _new_date)
                                if "時間窗_開始" in _tasks_changed.columns:
                                    _tasks_changed.loc[_mask, "時間窗_開始"] = _new_start.strftime("%H:%M")
                                if "時間窗_結束" in _tasks_changed.columns:
                                    _tasks_changed.loc[_mask, "時間窗_結束"] = _new_end.strftime("%H:%M")

                                _tasks_changed_merged = merge_task_client_tables(_tasks_changed, res.get("df_cl", pd.DataFrame()))

                                _new_task_merged = _tasks_changed_merged[
                                    _tasks_changed_merged["任務ID"].astype(str)
                                    == str(_reschedule_task_id)
                                ].iloc[0]

                                _candidate_df = find_insertion_candidates(
                                    _new_task_merged,
                                    _tasks_changed_merged,
                                    _base_result,
                                    res.get("df_cg", pd.DataFrame()),
                                    config,
                                    top_n=3,
                                    date_column="日期",
                                )

                                st.session_state["reschedule_event_analysis"] = {
                                    "task_id": str(_reschedule_task_id),
                                    "original_task": _orig_task.to_dict(),
                                    "new_task": _new_task_merged.to_dict(),
                                    "candidates": _candidate_df,
                                }

                            except Exception as e:
                                st.session_state["reschedule_event_analysis"] = None
                                st.error(f"改期候選人計算失敗：{e}")

                    _reschedule_analysis = st.session_state.get("reschedule_event_analysis")

                    if (
                        isinstance(_reschedule_analysis, dict)
                        and str(_reschedule_analysis.get("task_id"))
                        == str(_reschedule_task_id)
                    ):
                        _candidate_df = _reschedule_analysis.get("candidates")
                        _new_task_info = _reschedule_analysis.get("new_task", {})
                        _orig_task_info = _reschedule_analysis.get("original_task", {})

                        if not isinstance(_candidate_df, pd.DataFrame) or _candidate_df.empty:
                            st.warning("目前沒有候選人結果。")
                        else:
                            _available = _candidate_df[
                                _candidate_df["可直接插入"] == True
                            ].copy()

                            if _available.empty:
                                st.warning(
                                    "新時段目前沒有可直接插入的人選。"
                                    "這筆任務之後可進入局部重排。"
                                )
                                show_table(
                                    _candidate_df,
                                    width="stretch",
                                    hide_index=True,
                                )
                            else:
                                _top3 = _available.head(3).copy()
                                if _resched_orig_cg:
                                    _top3.insert(2, "原居服員", _top3["居服員ID"].astype(str).eq(_resched_orig_cg).map({True: "✓", False: ""}))
                                st.success(
                                    f"新時段找到 {len(_top3)} 位可直接插入候選人。"
                                )
                                show_table(
                                    _top3,
                                    width="stretch",
                                    hide_index=True,
                                )

                                st.markdown("**督導確認改期後接手人員**")
                                _top1_id = str(_top3.iloc[0]["居服員ID"])

                                for _rank, (_idx, _cand) in enumerate(
                                    _top3.iterrows(), start=1
                                ):
                                    _new_cg = str(_cand.get("居服員ID", ""))
                                    _score = _cand.get(
                                        "適配度分數",
                                        _cand.get("適配分數", None),
                                    )
                                    _cost = _cand.get("擾動成本", 0)
                                    _prev_slot = str(_cand.get("前一任務（結束）", "無"))
                                    _next_slot = str(_cand.get("下一任務（開始）", "無"))

                                    _left, _right = st.columns([4, 1])
                                    _label = f"候選 {_rank}｜{caregiver_label(_new_cg)}"
                                    if _resched_orig_cg and _new_cg == _resched_orig_cg:
                                        _label += "｜⭐ 原居服員"
                                    if str(_cand.get("代班排序理由", "") or ""):
                                        _label += f"｜{_label_ids_plain(_cand.get('代班排序理由'))}"
                                    _label += f"｜擾動成本 {_cost}"
                                    _left.markdown(_label)
                                    _left.caption(
                                        f"🕒 前一任務 {_prev_slot} → 本次服務 → 下一任務 {_next_slot}"
                                    )

                                    if _right.button(
                                        "✅ 確認改期",
                                        key=f"reschedule_confirm_{_reschedule_task_id}_{_rank}_{_new_cg}",
                                        width="stretch",
                                    ):
                                        _reason = (
                                            "案家要求改時間｜案家臨時改期："
                                            f"{_orig_task_info.get('日期')} "
                                            f"{_orig_task_info.get('時間窗_開始')}-"
                                            f"{_orig_task_info.get('時間窗_結束')} → "
                                            f"{_new_task_info.get('日期')} "
                                            f"{_new_task_info.get('時間窗_開始')}-"
                                            f"{_new_task_info.get('時間窗_結束')}"
                                        )

                                        # 先把任務換成新日期／時段，再做衝突檢查；否則會拿「舊時段」去檢查。
                                        _main_tasks = st.session_state["last_result"]["df_tasks"]
                                        _main_mask = _main_tasks["任務ID"].astype(str) == str(_reschedule_task_id)
                                        _cols = [c for c in ("日期", "時間窗_開始", "時間窗_結束") if c in _main_tasks.columns]
                                        _backup = _main_tasks.loc[_main_mask, _cols].copy()
                                        if "日期" in _cols:
                                            _main_tasks.loc[_main_mask, "日期"] = _date_cell(_main_tasks, _new_task_info.get("日期"))
                                        if "時間窗_開始" in _cols:
                                            _main_tasks.loc[_main_mask, "時間窗_開始"] = _new_task_info.get("時間窗_開始")
                                        if "時間窗_結束" in _cols:
                                            _main_tasks.loc[_main_mask, "時間窗_結束"] = _new_task_info.get("時間窗_結束")

                                        _conflict_msg = _quick_reassign(
                                            _reschedule_task_id,
                                            _new_cg,
                                            _reason,
                                        )

                                        if _conflict_msg:
                                            _main_tasks.loc[_main_mask, _cols] = _backup.values   # 檢查失敗就還原原時段
                                            st.error(_conflict_msg)
                                        else:
                                            _confirmed_at = pd.Timestamp.now()
                                            _top3_ids = " > ".join(
                                                _top3["居服員ID"].astype(str).tolist()
                                            )

                                            st.session_state["dynamic_adjustment_log"].append(
                                                {
                                                    "時間戳記": _confirmed_at,
                                                    "事件類型": "案家臨時改期",
                                                    "任務ID": _reschedule_task_id,
                                                    "案家ID": _new_task_info.get("案家ID"),
                                                    "原日期": _orig_task_info.get("日期"),
                                                    "原時段": (
                                                        f"{_orig_task_info.get('時間窗_開始', '')}-"
                                                        f"{_orig_task_info.get('時間窗_結束', '')}"
                                                    ),
                                                    "新日期": _new_task_info.get("日期"),
                                                    "新時段": (
                                                        f"{_new_task_info.get('時間窗_開始', '')}-"
                                                        f"{_new_task_info.get('時間窗_結束', '')}"
                                                    ),
                                                    "系統Top3": _top3_ids,
                                                    "系統第一名": _top1_id,
                                                    "最終選擇居服員": _new_cg,
                                                    "最終候選排名": _rank,
                                                    "既有其他班表是否移動": "否",
                                                    "擾動成本": _cand.get("擾動成本", 0),
                                                }
                                            )

                                            st.success(
                                                f"✅ {_reschedule_task_id} 已完成改期並指派給 {_new_cg}。"
                                            )
                                            st.session_state["reschedule_event_analysis"] = None
                                            st.rerun()

            # ==========================================

        elif _dynamic_event_type == "⏱️ 前一服務延遲":
            # ==========================================
            st.subheader("⏱️ 前一服務延遲")
            st.caption(
                "例如就醫陪同超時、上一案臨時延長。系統只更新該案的「預計結束時間」，"
                "再往後檢查同一位居服員的班次是否仍能完成交通銜接；"
                "只有真的受影響的後續任務才重新找接手人員。"
            )

            _effective_result_for_delay = apply_overrides_to_result(df_result, overrides)

            def _delay_parse_time(value):
                if hasattr(value, "hour") and hasattr(value, "minute"):
                    return value
                try:
                    return pd.to_datetime(str(value)).time()
                except Exception:
                    return None

            def _delay_same_date(v1, v2):
                try:
                    return pd.to_datetime(v1).date() == pd.to_datetime(v2).date()
                except Exception:
                    return str(v1).split(" ")[0] == str(v2).split(" ")[0]

            def _build_delay_analysis(_ctx):
                """依固定的延遲後結束時間，找出真正受影響的後續任務並重算候選人。"""
                _task_id = str(_ctx["task_id"])
                _cg_id = str(_ctx["cg_id"])
                _event_date = pd.to_datetime(_ctx["date"]).date()
                _expected_end = pd.to_datetime(_ctx["expected_end"]).time()

                _tasks_base = st.session_state.get("last_result", {}).get("df_tasks", df_tasks).copy()
                _clients = res.get("df_cl", pd.DataFrame())
                _caregivers = res.get("df_cg", pd.DataFrame())

                # 任務 + 案家座標／需求
                _tasks_merged = merge_task_client_tables(_tasks_base, _clients)

                _task_lookup = {
                    str(r["任務ID"]): r
                    for _, r in _tasks_merged.iterrows()
                }

                _delayed_row = _task_lookup.get(_task_id)
                if _delayed_row is None:
                    return {"error": "找不到延遲任務資料", "affected": []}

                _cg_row = _caregivers[
                    _caregivers["居服員ID"].astype(str) == _cg_id
                ]
                _transport = "機車"
                if not _cg_row.empty:
                    _transport = str(_cg_row.iloc[0].get("常用交通工具", "機車")).strip()
                    if _transport not in ("機車", "大眾運輸"):
                        _transport = "機車"

                # 同一位居服員、同一天目前仍有效的派單
                _effective_now = apply_overrides_to_result(df_result, st.session_state.get("overrides", {}))
                _cg_tasks = []
                for _, _rr in _effective_now[
                    _effective_now["派單居服員"].astype(str) == _cg_id
                ].iterrows():
                    _tid = str(_rr["任務ID"])
                    _tr = _task_lookup.get(_tid)
                    if _tr is None:
                        continue
                    if "日期" in _tr.index and not _delay_same_date(_tr.get("日期"), _event_date):
                        continue

                    _start = _delay_parse_time(_tr.get("時間窗_開始"))
                    _end = _delay_parse_time(_tr.get("時間窗_結束"))
                    if _start is None or _end is None:
                        continue

                    _cg_tasks.append(
                        {
                            "task_id": _tid,
                            "start": _start,
                            "end": _end,
                            "lat": _tr.get("服務地點_緯度"),
                            "lon": _tr.get("服務地點_經度"),
                            "row": _tr,
                        }
                    )

                _cg_tasks.sort(key=lambda x: x["start"])

                # 延遲任務作為新的「最後一筆保留任務」。
                _prev = {
                    "task_id": _task_id,
                    "start": _delay_parse_time(_delayed_row.get("時間窗_開始")),
                    "end": _expected_end,
                    "lat": _delayed_row.get("服務地點_緯度"),
                    "lon": _delayed_row.get("服務地點_經度"),
                    "row": _delayed_row,
                }

                _affected = []
                _kept_after = []

                for _item in _cg_tasks:
                    if _item["task_id"] == _task_id:
                        continue
                    # 只檢查延遲任務之後的班。
                    if _item["start"] <= _prev["start"]:
                        continue

                    _travel = None
                    if (
                        pd.notna(_prev.get("lat"))
                        and pd.notna(_prev.get("lon"))
                        and pd.notna(_item.get("lat"))
                        and pd.notna(_item.get("lon"))
                    ):
                        _travel = calc_travel_minutes(
                            _prev["lat"],
                            _prev["lon"],
                            _item["lat"],
                            _item["lon"],
                            config,
                            transport_mode=_transport,
                        )

                    _required_gap = float(config.buffer_mins) + (float(_travel) if _travel is not None else 0.0)

                    _prev_end_dt = pd.Timestamp.combine(pd.Timestamp.today().date(), _prev["end"])
                    _next_start_dt = pd.Timestamp.combine(pd.Timestamp.today().date(), _item["start"])
                    _actual_gap = (_next_start_dt - _prev_end_dt).total_seconds() / 60.0

                    if _actual_gap < _required_gap:
                        _affected.append(
                            {
                                "task_id": _item["task_id"],
                                "row": _item["row"],
                                "original_cg": _cg_id,
                                "previous_task": _prev["task_id"],
                                "previous_end": _prev["end"].strftime("%H:%M"),
                                "next_start": _item["start"].strftime("%H:%M"),
                                "travel_min": None if _travel is None else float(_travel),
                                "available_gap": float(_actual_gap),
                                "required_gap": float(_required_gap),
                                "reason": (
                                    f"前一任務 {_prev['task_id']} 預計 {_prev['end'].strftime('%H:%M')} 結束；"
                                    f"本任務 {_item['start'].strftime('%H:%M')} 開始，"
                                    f"可用 {_actual_gap:.0f} 分鐘，但至少需要 {_required_gap:.0f} 分鐘"
                                    f"（交通 + {float(config.buffer_mins):.0f} 分緩衝）。"
                                ),
                            }
                        )
                        # 此任務將被移出原居服員班表，因此不把它當作下一個保留節點。
                        continue

                    # 銜接仍可行，保留在原居服員班表，後續從這一筆再往下檢查。
                    _kept_after.append(_item["task_id"])
                    _prev = _item

                _affected_ids = [x["task_id"] for x in _affected]

                # 受影響任務全部先暫時移除，再逐筆找可直接插入的人。
                _base_result = _effective_now[
                    ~_effective_now["任務ID"].astype(str).isin(_affected_ids)
                ].copy()

                # 原居服員正在處理延遲事件，不列入受影響任務的替補候選。
                _candidate_cg = _caregivers.copy()
                if "居服員ID" in _candidate_cg.columns:
                    _candidate_cg = _candidate_cg[
                        _candidate_cg["居服員ID"].astype(str) != _cg_id
                    ].copy()

                def _delay_cfg_for(_a):
                    """延遲當次視同原固定居服員趕不上這一案（這段時間請假），只在這次計算開放代班，不改動正式設定。"""
                    import json as _json
                    from dataclasses import replace as _replace
                    _row = _a["row"]
                    _client = str(_row.get("案家ID", ""))
                    _owner = config.client_primary_map.get(_client) or _row.get("歷史首選居服員ID")
                    _owner = _cg_id if _owner is None or pd.isna(_owner) or not str(_owner).strip() else str(_owner)
                    _cfg = _replace(config, primary_profiles=dict(config.primary_profiles))
                    _prof = dict(_cfg.primary_profiles.get(_owner) or {})
                    _raw = _prof.get("臨時不可排時段") or "[]"
                    try:
                        _blocks = _json.loads(_raw) if isinstance(_raw, str) else list(_raw)
                    except ValueError:
                        _blocks = []
                    _day = pd.to_datetime(_row.get("日期")).strftime("%Y-%m-%d")
                    _blocks.append({"start": f"{_day}T{str(_row.get('時間窗_開始'))[:5]}:00",
                                    "end": f"{_day}T{str(_row.get('時間窗_結束'))[:5]}:00",
                                    "reason": "請假（前一服務延遲）"})
                    _prof["臨時不可排時段"] = _json.dumps(_blocks, ensure_ascii=False)
                    _cfg.primary_profiles[_owner] = _prof
                    return _cfg

                for _a in _affected:
                    try:
                        _cand = find_insertion_candidates(
                            _a["row"],
                            _tasks_merged,
                            _base_result,
                            _candidate_cg,
                            _delay_cfg_for(_a),
                            top_n=3,
                            date_column="日期",
                        )
                    except Exception as _err:
                        _cand = pd.DataFrame()
                        _a["candidate_error"] = str(_err)
                    _a["candidates"] = _cand

                return {
                    "error": "",
                    "affected": _affected,
                    "kept_after": _kept_after,
                    "transport": _transport,
                }

            # 可選任務：只列目前有實際派單的任務
            _delay_assigned = None
            if not _effective_result_for_delay.empty:
                _delay_unique = _effective_result_for_delay.drop_duplicates('任務ID')
                _delay_assigned = pd.Series(_delay_unique['派單居服員'].values, index=_delay_unique['任務ID'].astype(str).values)
            _delay_tasks_view, _delay_who = df_tasks, '全部'
            if _delay_assigned is not None:
                _delay_people = sorted({str(v) for v in _delay_assigned.dropna()} - {'', '未指派', 'None', 'nan'})
                _delay_who = st.selectbox('居服員（可先選人，再挑案家與服務）', ['全部'] + _delay_people,
                                          format_func=lambda v: '全部居服員' if v == '全部' else caregiver_label(v),
                                          key='delay_pick_cg')
                if _delay_who != '全部':
                    _delay_ids = set(_delay_assigned[_delay_assigned.astype(str) == _delay_who].index)
                    _delay_tasks_view = df_tasks[df_tasks['任務ID'].astype(str).isin(_delay_ids)]
            _delay_task_id = _task_picker(f'delay_{_delay_who}', _delay_tasks_view, assigned=_delay_assigned) if _delay_assigned is not None else None
            if _delay_task_id is None:
                st.info("目前沒有可分析的已派任務。")
            else:
                _delay_task_match = df_tasks[
                    df_tasks["任務ID"].astype(str) == str(_delay_task_id)
                ].copy()
                _delay_assignment = _effective_result_for_delay[
                    _effective_result_for_delay["任務ID"].astype(str) == str(_delay_task_id)
                ].copy()

                if _delay_task_match.empty or _delay_assignment.empty:
                    st.warning("找不到此任務的完整排班資料。")
                else:
                    _delay_task = _delay_task_match.iloc[0]
                    _delay_cg = str(_delay_assignment.iloc[0]["派單居服員"])
                    _delay_date = pd.to_datetime(_delay_task.get("日期"), errors="coerce")

                    _existing_delay_ctx = st.session_state.get("delay_event_context")
                    if (
                        isinstance(_existing_delay_ctx, dict)
                        and str(_existing_delay_ctx.get("task_id")) == str(_delay_task_id)
                        and _existing_delay_ctx.get("original_end")
                    ):
                        _delay_orig_end = _delay_parse_time(
                            _existing_delay_ctx.get("original_end")
                        )
                    else:
                        _delay_orig_end = _delay_parse_time(
                            _delay_task.get("時間窗_結束")
                        )

                    if _delay_orig_end is None:
                        st.warning("此任務沒有有效的結束時間。")
                    else:
                        dc1, dc2 = st.columns(2)
                        _delay_minutes = dc1.number_input(
                            "預計延遲（分鐘）",
                            min_value=5,
                            max_value=240,
                            value=30,
                            step=5,
                            key="delay_event_minutes",
                        )

                        _base_dt = pd.Timestamp.combine(
                            pd.Timestamp.today().date(),
                            _delay_orig_end,
                        )
                        _expected_end_dt = _base_dt + pd.Timedelta(minutes=int(_delay_minutes))
                        dc2.metric(
                            "新的預計結束時間",
                            _expected_end_dt.strftime("%H:%M"),
                        )

                        st.caption(
                            f"原排程：{_delay_cg}｜"
                            f"{_delay_task.get('時間窗_開始', '')}-{_delay_orig_end.strftime('%H:%M')} "
                            f"→ 預計延至 {_expected_end_dt.strftime('%H:%M')}"
                        )

                        if st.button(
                            "🔎 確認延遲並檢查後續班",
                            type="primary",
                            key="delay_event_analyze",
                        ):
                            _ctx = {
                                "task_id": str(_delay_task_id),
                                "cg_id": _delay_cg,
                                "date": _delay_task.get("日期"),
                                "original_end": _delay_orig_end.strftime("%H:%M"),
                                "expected_end": _expected_end_dt,
                                "delay_minutes": int(_delay_minutes),
                            }

                            # 不修改原始班表的「時間窗_結束」。
                            # 延遲只存於事件 context，因此每次都從原排程結束時間重新計算。
                            st.session_state["delay_event_context"] = _ctx

                            st.session_state["dynamic_adjustment_log"].append(
                                {
                                    "時間戳記": pd.Timestamp.now(),
                                    "事件類型": "前一服務延遲",
                                    "任務ID": str(_delay_task_id),
                                    "案家ID": _delay_task.get("案家ID"),
                                    "日期": _delay_task.get("日期"),
                                    "居服員": _delay_cg,
                                    "原預計結束時間": _delay_orig_end.strftime("%H:%M"),
                                    "新預計結束時間": _expected_end_dt.strftime("%H:%M"),
                                    "延遲分鐘": int(_delay_minutes),
                                    "處理方式": "只檢查並重排受影響後續區段",
                                }
                            )
                            st.rerun()

                _delay_ctx = st.session_state.get("delay_event_context")
                if (
                    isinstance(_delay_ctx, dict)
                    and str(_delay_ctx.get("task_id")) == str(_delay_task_id)
                ):
                    _delay_analysis = _build_delay_analysis(_delay_ctx)

                    st.info(
                        f"目前延遲事件：{_delay_ctx['cg_id']} 的 {_delay_ctx['task_id']} "
                        f"預計延至 {pd.to_datetime(_delay_ctx['expected_end']).strftime('%H:%M')} 結束。"
                    )

                    if _delay_analysis.get("error"):
                        st.error(_delay_analysis["error"])
                    else:
                        _affected = _delay_analysis.get("affected", [])

                        if not _affected:
                            st.success(
                                "✅ 延遲後的交通與緩衝時間仍足夠，後續班表不需要調整。"
                            )
                        else:
                            st.warning(
                                f"⚠️ 找到 {len(_affected)} 筆真正受到延遲影響的後續任務。"
                                "只有以下任務需要重新安排。"
                            )

                            _affected_summary = []
                            for _a in _affected:
                                _affected_summary.append(
                                    {
                                        "受影響任務": _a["task_id"],
                                        "原居服員": _a["original_cg"],
                                        "前一任務": _a["previous_task"],
                                        "前一任務預計結束": _a["previous_end"],
                                        "本任務開始": _a["next_start"],
                                        "可用間隔(分)": round(_a["available_gap"], 1),
                                        "至少需要(分)": round(_a["required_gap"], 1),
                                        "原因": _a["reason"],
                                    }
                                )
                            show_table(
                                pd.DataFrame(_affected_summary),
                                width="stretch",
                                hide_index=True,
                            )

                            for _i, _a in enumerate(_affected, start=1):
                                _task_id = str(_a["task_id"])
                                _cand_df = _a.get("candidates", pd.DataFrame())

                                with st.expander(
                                    f"受影響任務 {_i}｜{_task_id}｜{_a['next_start']} 開始",
                                    expanded=True,
                                ):
                                    st.caption(_a["reason"])

                                    if _a.get("candidate_error"):
                                        st.error(
                                            f"候選人計算失敗：{_a['candidate_error']}"
                                        )
                                        continue

                                    if not isinstance(_cand_df, pd.DataFrame) or _cand_df.empty:
                                        st.warning("目前沒有可用候選人結果。")
                                        continue

                                    _available = _cand_df[
                                        _cand_df["可直接插入"] == True
                                    ].copy()

                                    if _available.empty:
                                        st.warning(
                                            "目前沒有可直接接手的人選；此任務需要進一步做局部重排。"
                                        )
                                        show_table(
                                            _cand_df,
                                            width="stretch",
                                            hide_index=True,
                                        )
                                        continue

                                    _top3 = _available.head(3).copy()
                                    show_table(
                                        _top3,
                                        width="stretch",
                                        hide_index=True,
                                    )

                                    for _rank, (_idx, _cand) in enumerate(
                                        _top3.iterrows(), start=1
                                    ):
                                        _new_cg = str(_cand.get("居服員ID", ""))
                                        _score = _cand.get(
                                            "適配度分數",
                                            _cand.get("適配分數", None),
                                        )
                                        _cost = _cand.get("擾動成本", 0)
                                        _prev_slot = str(
                                            _cand.get("前一任務（結束）", "無")
                                        )
                                        _next_slot = str(
                                            _cand.get("下一任務（開始）", "無")
                                        )

                                        _left, _right = st.columns([4, 1])
                                        _label = f"候選 {_rank}｜{caregiver_label(_new_cg)}"
                                        if str(_cand.get("代班排序理由", "") or ""):
                                            _label += f"｜{_label_ids_plain(_cand.get('代班排序理由'))}"
                                        _label += f"｜擾動成本 {_cost}"
                                        _left.markdown(_label)
                                        _left.caption(
                                            f"🕒 {_prev_slot} → 本次 {_task_id} → {_next_slot}"
                                        )

                                        if _right.button(
                                            "✅ 改派此任務",
                                            key=(
                                                f"delay_reassign_{_delay_ctx['task_id']}_"
                                                f"{_task_id}_{_rank}_{_new_cg}"
                                            ),
                                            width="stretch",
                                        ):
                                            _reason = (
                                                f"交通／時間銜接不足｜前一服務延遲：{_delay_ctx['task_id']} "
                                                f"延至 {pd.to_datetime(_delay_ctx['expected_end']).strftime('%H:%M')}；"
                                                f"{_task_id} 無法維持原銜接"
                                            )

                                            # 原居服員趕不上這一案：登錄只涵蓋這一案時段的「請假（前一服務延遲）」，
                                            # 才能依規定開放代班，且重新整理後改派不會被退回。
                                            _d_row = _a["row"]
                                            _d_day = pd.to_datetime(_d_row.get("日期")).strftime("%Y-%m-%d")
                                            _d_span = (str(_d_row.get("時間窗_開始"))[:5], str(_d_row.get("時間窗_結束"))[:5])
                                            _d_cg = str(_a.get("original_cg") or _delay_ctx["cg_id"])
                                            _d_before = {b["id"] for b in _workflow_store.blocks()}
                                            _workflow_store.add_blocks(_d_cg, _d_day, [_d_span], "請假（前一服務延遲）")
                                            res["df_cg"] = attach_blocks(res["df_cg"], _workflow_store)
                                            config.primary_profiles = {str(r["居服員ID"]): r.to_dict() for _, r in res["df_cg"].iterrows()}
                                            _conflict_msg = _quick_reassign(
                                                _task_id,
                                                _new_cg,
                                                _reason,
                                            )
                                            if _conflict_msg:
                                                for _b in _workflow_store.blocks():
                                                    if _b["id"] not in _d_before:
                                                        _workflow_store.remove_block(_b["id"])
                                                res["df_cg"] = attach_blocks(res["df_cg"], _workflow_store)
                                                config.primary_profiles = {str(r["居服員ID"]): r.to_dict() for _, r in res["df_cg"].iterrows()}

                                            if _conflict_msg:
                                                st.error(_conflict_msg)
                                            else:
                                                st.session_state[
                                                    "dynamic_adjustment_log"
                                                ].append(
                                                    {
                                                        "時間戳記": pd.Timestamp.now(),
                                                        "事件類型": "前一服務延遲－後續改派",
                                                        "延遲任務ID": _delay_ctx["task_id"],
                                                        "受影響任務ID": _task_id,
                                                        "原居服員": _delay_ctx["cg_id"],
                                                        "最終選擇居服員": _new_cg,
                                                        "最終候選排名": _rank,
                                                        "擾動成本": _cost,
                                                        "原因": _reason,
                                                    }
                                                )
                                                st.success(
                                                    f"✅ {_task_id} 已改派給 {_new_cg}。"
                                                )
                                                # rerun 後會以最新 override 自動重算「剩餘受影響區段」。
                                                st.rerun()

                    if st.button(
                        "✅ 結束本次延遲事件處理",
                        key="delay_event_finish",
                    ):
                        st.session_state["delay_event_context"] = None
                        st.rerun()

            # ==========================================

        elif _dynamic_event_type == "🩺 個案照顧需求變更":
            import qualification_recheck as _qr
            subhead('🩺 個案照顧需求變更：重新檢核目前的居服員資格')
            st.caption('照專複評後，案家新增「失智」或新增「BA08 足部照護」等需要特殊訓練的照顧需求，'
                       '而原本服務的居服員沒有這項訓練時使用。系統會找出受影響的服務，全部改派給有訓練的居服員，'
                       '並提供降低換人衝擊的作法（熟面孔陪同交接、原居服員補訓後回任）。')
            st.info('⚖️ ' + _qr.RULE_TEXT)
            # 需求與處理方式用大按鈕切換；範例按鈕會改 qual_need／qual_method，這裡同步到按鈕狀態。
            _q_need_labels = {'dementia': '🧠 新增失智', 'foot': '🦶 新增足部照護（BA08）'}
            _q_need_keys = {v: k for k, v in _q_need_labels.items()}
            _q_want = st.session_state.get('qual_need') if st.session_state.get('qual_need') in _q_need_labels else 'dementia'
            if st.session_state.get('_qual_need_seen') != _q_want:
                st.session_state['qual_need_tab'] = _q_need_labels[_q_want]
            st.markdown('##### ① 新增了什麼需求')
            _q_need = _q_need_keys[big_tabs(list(_q_need_labels.values()), 'qual_need_tab', size='medium')]
            st.session_state['qual_need'] = st.session_state['_qual_need_seen'] = _q_need
            _q_info = _qr.NEEDS[_q_need]
            if _q_need == 'dementia':
                st.caption('📌 個案確認失智後，服務須由完成失智症照顧服務 20 小時訓練並登錄的居服員提供。')
            else:
                st.caption('📌 BA08 足部照護須由完成足部照護訓練（10 小時）並登錄的居服員提供。若不想整案換人，也可以原居服員繼續原本的服務、'
                           '另排一位有訓練的人單獨做 BA08（週服務計畫另加一列，用「新增服務」流程）。')
            _q_tasks, _q_effective = _workflow_frames()
            _q_current = {str(t): (None if str(c) in _qr.EMPTY else str(c))
                          for t, c in zip(_q_effective.get('任務ID', []), _q_effective.get('派單居服員', []))}
            _q_clients = sorted({str(c) for c in _q_tasks['案家ID'].astype(str)
                                 if any(_q_current.get(str(t)) for t in _q_tasks.loc[_q_tasks['案家ID'].astype(str) == c, '任務ID'])})
            st.markdown('##### ② 哪一個案家、從哪天開始')
            _qc1, _qc2 = st.columns([2, 1])
            _q_client = _qc1.selectbox('案家', _q_clients, index=None, format_func=client_label, key='qual_client',
                                       placeholder='選擇照專複評後新增需求的案家')
            _q_days = pd.to_datetime(_q_tasks['日期'], errors='coerce').dropna()
            if 'qual_since' not in st.session_state and len(_q_days):
                st.session_state['qual_since'] = _q_days.min().date()
            _q_since = _qc2.date_input('新需求生效日', key='qual_since')
            if _q_client:
                _q_aff = _qr.affected_tasks(_q_client, _q_tasks, _q_current, res['df_cg'], _q_since, need=_q_need)
                if not _q_aff:
                    st.success(f'✅ {client_label(_q_client)} 從 {_q_since} 起的服務，目前的居服員都有{_q_info["course"]}，不需要調整。')
                else:
                    _q_owner = _q_aff[0]['居服員']
                    _q_owners = sorted({a['居服員'] for a in _q_aff})
                    _q_key = (_q_need, str(_q_client), str(_q_since), tuple(a['任務ID'] for a in _q_aff))
                    _q_cache = st.session_state.get('qual_avail')
                    if not _q_cache or _q_cache.get('key') != _q_key:
                        with st.spinner('用同一套規則試算有訓練的居服員能不能接（含前一案交通、工時、休息）…'):
                            _q_cfg = replace(config, client_primary_map=dict(config.client_primary_map),
                                             client_backup_map=dict(getattr(config, 'client_backup_map', None) or {}),
                                             task_support_map=dict(getattr(config, 'task_support_map', None) or {}),
                                             work_day_minutes=dict(config.work_day_minutes), dispatch_cache={})
                            _q_avail = _qr.availability([a['任務ID'] for a in _q_aff], sorted(_qr.certified(res['df_cg'], _q_need)),
                                                        _q_tasks, _q_effective, res['df_cg'], _q_cfg, res.get('df_cl'), task_locations)
                        st.session_state['qual_avail'] = {'key': _q_key, 'avail': _q_avail}
                    _q_avail = st.session_state['qual_avail']['avail']
                    _q_all_ids = [a['任務ID'] for a in _q_aff]
                    _q_takers_all = _qr.covering(_q_avail, _q_all_ids)
                    _q_days_text = '、'.join(sorted({'週' + '一二三四五六日'[pd.Timestamp(a['日期']).weekday()] for a in _q_aff},
                                                    key='週一週二週三週四週五週六週日'.index))
                    st.markdown(
                        "<div class='cf-guide-card cf-guide-warn'><div class='cf-guide-title'>⚠️ 重新檢核："
                        f"{client_label(_q_client)} 有 {len(_q_aff)} 次服務需要{_q_info['course']}</div>"
                        f"<div class='cf-guide-body'>目前由 {'、'.join(caregiver_label(o) for o in _q_owners)} 服務（{_q_days_text}），"
                        "沒有這項訓練，依法規不能再服務這些時段，需改由有訓練的居服員。</div></div>", unsafe_allow_html=True)
                    _qm1, _qm2, _qm3 = st.columns(3)
                    _qm1.metric('受影響服務', f'{len(_q_aff)} 次', help='從生效日起，由沒有這項訓練的人服務的次數')
                    _qm2.metric('原居服員', caregiver_label(_q_owner).split('｜')[0])
                    _qm3.metric('有訓練、能全部接手', f'{len(_q_takers_all)} 位' if _q_takers_all else '需分給多位',
                                help='能接下全部受影響服務的有訓練居服員人數（已檢查時段、交通、工時、休息）')
                    with st.expander(f'受影響的服務明細（{len(_q_aff)} 次）'):
                        show_table(pd.DataFrame([{'日期': a['日期'], '時段': f"{a['開始']}–{a['結束']}", '目前居服員': caregiver_label(a['居服員'])}
                                                 for a in _q_aff]), hide_index=True, width='stretch')

                    st.markdown('##### ③ 選擇處理方式')
                    _q_method_labels = {'now': '① 立即改派', 'handover': '② 熟面孔陪同交接', 'grace': '③ 補訓後回任'}
                    _q_method_keys = {v: k for k, v in _q_method_labels.items()}
                    _q_mwant = st.session_state.get('qual_method') if st.session_state.get('qual_method') in _q_method_labels else _qr.DEFAULT_METHOD
                    if st.session_state.get('_qual_method_seen') != _q_mwant:
                        st.session_state['qual_method_tab'] = _q_method_labels[_q_mwant]
                    _q_method = _q_method_keys[big_tabs(list(_q_method_labels.values()), 'qual_method_tab', size='medium')]
                    st.session_state['qual_method'] = st.session_state['_qual_method_seen'] = _q_method
                    st.markdown("<div class='cf-guide-card'><div class='cf-guide-title'>"
                                + _qr.method_label(_q_method, _q_need)[2:] + "</div><div class='cf-guide-body'>💡 "
                                + _qr.method_help(_q_method, _q_need) + "</div></div>", unsafe_allow_html=True)
                    _q_deadline, _q_n = None, _qr.HANDOVER_VISITS
                    if _q_method == 'grace':
                        _q_deadline = st.date_input('原居服員補訓期限（預設生效日後 60 天）', value=_qr.default_deadline(_q_since),
                                                    key=f'qual_deadline_{_q_client}')
                    elif _q_method == 'handover':
                        _q_n = int(st.number_input('陪同交接次數（新居服員服務、原居服員陪同）', 1, max(1, len(_q_aff)),
                                                   min(_qr.HANDOVER_VISITS, max(1, len(_q_aff))), key='qual_handover_n'))
                    _q_parts = _qr.split(_q_method, _q_aff, _q_deadline, _q_n)
                    _q_cover = _qr.covering(_q_avail, _q_parts['move'], _q_parts['accompany'])
                    _q_need_new = bool(_q_parts['move'] or _q_parts['accompany'])
                    _q_new = None
                    _q_assign, _q_split = {}, None
                    if _q_need_new:
                        if not _q_cover:
                            # 沒有一位能全包：依星期分給多位有訓練的人（同一個星期固定同一位）
                            _q_dayof = {a['任務ID']: a['日期'] for a in _q_aff}
                            _q_split = _qr.split_cover(_q_avail, _q_parts['move'], _q_parts['accompany'], _q_dayof)
                            _q_assign = dict(_q_split['assign'])
                            if not _q_assign:
                                st.error('有訓練的居服員在這些時段都排不進。依法規不能讓原居服員繼續服務；請到「未派案家」用整案處理'
                                         '與案家協調調整時間，或回報照專／轉介其他單位。')
                            else:
                                _q_people = sorted(set(_q_assign.values()), key=lambda w: -list(_q_assign.values()).count(w))
                                st.warning(f"沒有一位有訓練的居服員能接下全部 {len(_q_parts['move']) + len(_q_parts['accompany'])} 次服務，"
                                           f"系統改成依星期分給 {len(_q_people)} 位（同一個星期固定同一位，長輩比較好適應）：")
                                show_table(pd.DataFrame([{'星期': g['星期'], '接手居服員': caregiver_label(g['居服員']), '次數': g['筆數'],
                                                          '挑選理由': _label_ids_plain(g['理由'])} for g in _q_split['groups']]),
                                           hide_index=True, width='stretch')
                                if _q_split['missing']:
                                    st.error(f"還有 {len(_q_split['missing'])} 次找不到有訓練又排得進的人。套用後這幾次會列在失敗清單，"
                                                     "原居服員不符資格不能服務，請調整時間、回報照專或轉介。")
                        else:
                            _q_ids = [c['cg'] for c in _q_cover]
                            _q_why = {c['cg']: c['reason'] for c in _q_cover}
                            _q_new = st.selectbox(f'接手的居服員（有{_q_info["course"]}，{len(_q_ids)} 位可選）', _q_ids, index=0,
                                                  key=f'qual_new_{_q_client}_{_q_method}',
                                                  format_func=lambda v: ('⭐ 推薦｜' if v == _q_ids[0] else '') + caregiver_label(v)
                                                  + '　—　' + _label_ids_plain(_q_why.get(v, '')))
                            _q_assign = {t: _q_new for t in [*_q_parts['move'], *_q_parts['accompany']]}
                    _q_label = {t: f"{a['日期']} {a['開始']}–{a['結束']}" for a in _q_aff for t in [a['任務ID']]}
                    _q_preview = []
                    for a in _q_aff:
                        t = a['任務ID']
                        _who = _q_assign.get(t)
                        if not _who:
                            how = '⚠️ 缺人：找不到有訓練又排得進的人（原居服員不符資格，不可服務）'
                        elif t in _q_parts.get('escort', []):
                            how = f"改由 {caregiver_label(_who)} 服務，{caregiver_label(a['居服員'])} 陪同交接（不提供服務）"
                        else:
                            how = f"改由 {caregiver_label(_who)} 服務"
                        _q_preview.append({'服務': _q_label[t], '處理後': how})
                    st.markdown('##### ④ 套用後的樣子')
                    show_table(pd.DataFrame(_q_preview), hide_index=True, width='stretch')
                    if _q_method == 'grace':
                        st.caption('原居服員的補訓期限會記入追蹤清單（04 也看得到）；取得資格並登錄後，可到固定人選評估改回主責。')
                    if _q_parts.get('escort') and _q_assign:
                        st.caption('陪同交接不另計一筆服務、不申報；原居服員本來就排在這些時段，請另外通知兩位居服員與家屬。')
                    _q_blocked = _q_need_new and not _q_assign
                    if st.button('✅ 套用處理方式', type='primary', key='qual_apply', disabled=_q_blocked):
                        _q_now = pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
                        _q_reason = compose_reason('資格證照不符', f'個案{_q_info["label"]}：{_qr.METHOD_SHORT[_q_method]}')
                        _q_takers = sorted(set(_q_assign.values()), key=lambda w: (-list(_q_assign.values()).count(w), w))
                        if _q_takers:
                            _prim, _backs = _qr.team_after(
                                _q_method, _q_owner, list((config.client_backup_map or {}).get(str(_q_client)) or []), _q_takers[0])
                            if len(_q_takers) > 1:       # 分給多位：其他接手的人也加入固定人選團隊
                                _backs = list(dict.fromkeys([*_q_takers[1:], *_backs]))[:max(2, len(_q_takers) - 1)]
                            config.client_primary_map = {**config.client_primary_map, str(_q_client): _prim}
                            config.client_backup_map = {**config.client_backup_map, str(_q_client): _backs}
                            for _store in (res, res.get('phase2') or {}):
                                _store['client_primary_map'] = dict(config.client_primary_map)
                                _store['client_backup_map'] = dict(config.client_backup_map)
                        _q_fail = []
                        for t in _q_parts['move']:
                            if not _q_assign.get(t):
                                _q_fail.append(f"{_q_label[t]}：找不到有訓練又排得進的人")
                                continue
                            _msg = _quick_reassign(t, _q_assign[t], _q_reason, apply_scope='僅本次')
                            if _msg:
                                _q_fail.append(f"{_q_label[t]}：{_label_ids_plain(_msg)}")
                        st.session_state['dynamic_adjustment_log'].append({
                            '時間戳記': _q_now, '事件類型': '個案' + _q_info['label'], '案家ID': _q_client, '日期': str(_q_since),
                            '處理方式': _qr.METHOD_SHORT[_q_method], '原居服員': _q_owner, '最終選擇居服員': '、'.join(_q_takers),
                            '改派筆數': len(_q_parts['move']) - len(_q_fail), '原居服員陪同筆數': len(_q_parts.get('escort', [])),
                            '缺人筆數': len(_q_fail), '原因': _q_reason})
                        _follow = _qr.followup(_q_method, client_label(_q_client), caregiver_label(_q_owner),
                                               '、'.join(caregiver_label(w) for w in _q_takers), _q_parts, _q_deadline, _q_now, need=_q_need)
                        st.session_state.setdefault('qualification_followups', []).append(_follow)
                        st.session_state['qual_notice'] = (
                            f"✅ {client_label(_q_client)} 已套用「{_qr.METHOD_SHORT[_q_method]}」：" + _follow['待辦'] + '。'
                            + (f" 有 {len(_q_fail)} 筆沒有改成功：{'；'.join(_q_fail)}" if _q_fail else '')
                            + ' ' + _follow['提醒'] + '。')
                        st.session_state.pop('qual_avail', None)
                        st.rerun()
            if st.session_state.get('qual_notice'):
                st.success(st.session_state.pop('qual_notice'))
            if st.session_state.get('qualification_followups'):
                st.markdown('**資格追蹤清單**（補訓期限、陪同交接進度）')
                show_table(pd.DataFrame(st.session_state['qualification_followups']), hide_index=True, width='stretch')

        if st.session_state.get("dynamic_adjustment_log"):
            with st.expander("🧾 本次臨時異動紀錄", expanded=False):
                _dynamic_log_df = pd.DataFrame(st.session_state["dynamic_adjustment_log"])
                show_table(_dynamic_log_df, width="stretch", hide_index=True)
                st.caption(
                    "稽覈紀錄供動態調整追蹤使用；匯出後請依機構資料治理規範保存。"
                )
                _dynamic_log_csv = export_frame(_dynamic_log_df).to_csv(index=False).encode("utf-8-sig")
                download_button_quiet(
                    "⬇️ 匯出 Dynamic_Adjustment_Log.csv",
                    _dynamic_log_csv,
                    file_name="Dynamic_Adjustment_Log.csv",
                    mime="text/csv",
                    key="download_dynamic_adjustment_log",
                )

    if _module == M_ISSUES and _issue_view == '臨時異動':
        subhead('臨時異動')
        st.caption(
            "遇到臨時狀況時先選事件類型（新增服務、請假、改期、延遲）；系統只展開需要的流程，"
            "並盡量保留其他既有班表不動，再提供可替代的人選。"
        )
        st.info('🧭 ' + SUBSTITUTE_RULE_TEXT)
        _changes_section()
    if _module == M_ISSUES:
        module_nav_buttons(M_ISSUES)


    # 班表整表（含交通、薪資欄位）只在指派／任務／交通／人工調整／工作流程狀態改變時才重建；
    # 單純開關視窗、切換日期或拖曳取消時直接沿用上次結果。
    try:
        _workflow_digest = hashlib.sha256(repr(sorted(_workflow_store.all_records().items())).encode()).hexdigest()
    except Exception:
        _workflow_digest = str(__import__('time').time_ns())  # 讀不到工作流程狀態時不使用快取
    _final_key = hashlib.sha256(repr((
        route_signature, _route_digest, _frame_digest(df_tasks), _frame_digest(df_result),
        sorted((str(k), repr(v)) for k, v in overrides.items()),
        repr(st.session_state.get('dynamic_insertions', [])), _workflow_digest, input_mode, _show_names,
    )).encode()).hexdigest()
    _final_cached = res.get('_final_table_cache')
    if _final_cached and _final_cached[0] == _final_key:
        df_final = _final_cached[1].copy()
    else:
        has_date_for_final = "日期" in df_tasks.columns

        has_weekday_col_for_final = "星期" in df_tasks.columns

        final_rows = []
        _ai_lookup = (df_result.drop_duplicates('任務ID', keep='first').set_index('任務ID').to_dict('index')
                      if not df_result.empty else {})
        for _, task_row in df_tasks.iterrows():

            t_id = task_row["任務ID"]

            cl_id = task_row["案家ID"]

            ai_row = _ai_lookup.get(t_id)

            ai_cg = ai_row["派單居服員"] if ai_row is not None else None

            override = overrides.get(t_id)

            if override:

                final_cg = override["cg_id"]

                change_note = f"居督覆寫：{override['reason']}"

            else:

                final_cg = ai_cg

                change_note = ai_row["原首選替換原因"] if ai_row is not None else ""

            row_dict = {"任務ID": t_id, "案家ID": cl_id,
                        '結果用途':'待確認排班草案' if input_mode=='機構資料（單一Excel／三個工作表）' else '派單建議',
                        '待確認事項': task_row.get('任務待確認事項',''),
                        '疲勞資料狀態': ai_row.get('疲勞資料狀態','未計算') if ai_row is not None else '未計算',
                        '本次計畫已排時數': ai_row.get('本次計畫已排時數') if ai_row is not None else None}

            if has_date_for_final:

                date_val = task_row.get("日期")

                row_dict["日期"] = date_val

                row_dict["星期"] = (

                    task_row.get("星期")

                    if has_weekday_col_for_final and pd.notna(task_row.get("星期"))

                    else get_weekday_name(date_val)

                )

            row_dict.update({

                "AI建議居服員": ai_cg if ai_cg is not None else "未指派",

                "最終派單居服員": final_cg if final_cg is not None else "未指派",

                "適配分數": ai_row["適配分數"] if ai_row is not None else None,

                "預估車程(分)": ai_row["預估車程(分)"] if ai_row is not None else None,

                "服務時段": (str(task_row.get("時間窗_開始", "")) + "-" + str(task_row.get("時間窗_結束", ""))) if pd.notna(task_row.get("時間窗_開始")) and pd.notna(task_row.get("時間窗_結束")) else (ai_row["服務時段"] if ai_row is not None else ""),

                "任務優先級": ai_row["任務優先級"] if ai_row is not None else task_row.get("任務優先級", ""),

                "預估服務營收": ai_row["預估服務營收"] if ai_row is not None else None,

                "預估居服員薪資": ai_row["預估居服員薪資"] if ai_row is not None else None,

                "備註": change_note,

                # 系統挑代班時的理由（熟悉／前一案交通／本月工時／年資）；居督覆寫後不再適用。
                "代班排序理由": (ai_row.get("代班排序理由", "") if (ai_row is not None and not override) else ""),

            })

            final_rows.append(row_dict)

        # 將已確認的臨時插單一起併入最終派單結果（不改動原既有任務）。
        for _item in st.session_state.get("dynamic_insertions", []):
            _t = _item["task"]
            _dynamic_row = {
                "任務ID": _t.get("任務ID"),
                "案家ID": _t.get("案家ID"),
            }
            if has_date_for_final:
                _date_val = _t.get("日期")
                _dynamic_row["日期"] = _date_val
                _dynamic_row["星期"] = _t.get("星期") or get_weekday_name(_date_val)
            _dynamic_row.update(
                {
                    "AI建議居服員": (
                        _item.get("candidate_detail", {}).get("居服員ID")
                        if _item.get("candidate_rank", 1) == 1
                        else _item.get("system_top1_id", _item["cg_id"])
                    ),
                    "最終派單居服員": _item["cg_id"],
                    "適配分數": _item.get("candidate_detail", {}).get(
                        "適配度分數",
                        _item.get("candidate_detail", {}).get("適配分數"),
                    ),
                    "預估車程(分)": _item.get("candidate_detail", {}).get(
                        "前段交通時間",
                        _item.get("candidate_detail", {}).get("前段交通時間(分)"),
                    ),
                    "服務時段": f"{_t.get('時間窗_開始')}-{_t.get('時間窗_結束')}",
                    "任務優先級": _t.get("任務優先級", "高"),
                    "預估服務營收": None,
                    "預估居服員薪資": None,
                    "備註": "臨時新增案件｜直接插入（既有班表未移動）",
                }
            )
            final_rows.append(_dynamic_row)

        df_final = build_schedule_output(travel_tasks, _operational_assignments, res['df_cg'],
            df_route_result, metadata=pd.DataFrame(final_rows), buffer_mins=config.buffer_mins)
        financial_output = calculate_task_revenue_and_salary(travel_tasks, config).set_index('任務ID')
        for column in ['費用類型','預估服務營收','預估自費營收','預估BA服務營收','財務估算狀態','SA10計費提醒']:
            df_final[column] = df_final['任務ID'].map(financial_output[column])
        if not task_payroll.empty:
            payroll_output=task_payroll.set_index('任務ID')
            for column in ['固定居服員ID','是否請假代班','代班原因','服務基本工資','轉場次數','轉場工資','加班加給','預估居服員薪資']:
                df_final[column] = df_final['任務ID'].map(payroll_output[column])
        else:
            df_final['預估居服員薪資'] = float('nan')
        res['_final_table_cache'] = (_final_key, df_final.copy())
        # 最終有效班表（含居督覆寫與臨時插單）重建完成即寫回持久帳本，確保跨月回溯「連續上班天數」
        # 用的是當下真正生效的指派，不是尚未套用覆寫前的 AI 原始建議。
        if '日期' in df_final.columns and '最終派單居服員' in df_final.columns:
            _final_assigned = df_final[df_final['最終派單居服員'].astype(str) != '未指派']
            _workflow_store.record_workdays(zip(_final_assigned['最終派單居服員'], _final_assigned['日期']))
    if _module == M_REVIEW:
        subhead('最終班表（目前有效派單）')
        st.caption('沿用週服務計畫欄位，加入派單居服員與逐段交通。交通時間為估計分鐘，原檔的0不當成實測；空白表示尚未計算。姓名依側邊欄「個資顯示設定」呈現。')

        @st.fragment
        def _schedule_view():
            """篩選與表格只重跑這一區。"""
            _filter_date, _filter_cg, _filter_case = st.columns([1, 1, 2])
            _date_options = ['全部日期'] + sorted(df_final['日期'].dropna().astype(str).unique().tolist())
            _cg_options = ['全部居服員'] + sorted(df_final['派單居服員'].dropna().astype(str).unique().tolist())
            _selected_date = _filter_date.selectbox('日期', _date_options, key='output_filter_date')
            _selected_cg = _filter_cg.selectbox('居服員', _cg_options,
                format_func=lambda value: caregiver_label(value) if value != '全部居服員' else value, key='output_filter_cg')
            _search_case = _filter_case.text_input('搜尋案號、案家姓名或任務ID', key='output_filter_case').strip()
            _visible = df_final.copy()
            if _selected_date != '全部日期':
                _visible = _visible.loc[_visible['日期'].astype(str).eq(_selected_date)]
            if _selected_cg != '全部居服員':
                _visible = _visible.loc[_visible['派單居服員'].astype(str).eq(_selected_cg)]
            if _search_case:
                _case_text = _visible['案號'].astype(str) + ' ' + _visible['任務ID'].astype(str)
                if _show_names:
                    _case_text += ' ' + _visible['案號'].map(_client_names).fillna('')
                _visible = _visible.loc[_case_text.str.contains(_search_case, case=False, regex=False)]
            _full_columns = st.checkbox('顯示週服務計畫與交通完整欄位', value=True, key='output_full_columns')
            _full_view_columns = MAIN_COLUMNS + [col for col in ['狀態', '服務時數', '費用類型', '頻率', '生效日', '到期日',
                '固定居服員ID', '是否請假代班', '代班原因', '代班排序理由', '預估服務營收', '服務基本工資', '轉場工資', '加班加給', '預估居服員薪資', 'SA10計費提醒', '交通估算來源', '交通查詢方式', '交通查詢時間', '轉場時段檢查', '備註', '待確認事項'] if col in _visible and col not in MAIN_COLUMNS]
            show_table(_visible[_full_view_columns if _full_columns else MAIN_COLUMNS], width='stretch', hide_index=True,
                column_config={
                    '案號': st.column_config.TextColumn('案家（案號／姓名）', width='medium'),
                    '派單居服員': st.column_config.TextColumn('居服員（編號／姓名）', width='medium'),
                    '交通時間': st.column_config.NumberColumn('交通時間（分）', format='%.2f', help='估計交通，不含轉場緩衝；空白表示未計算。'),
                    '出發地': st.column_config.TextColumn('出發地', width='large'),
                    '服務項目': st.column_config.TextColumn('服務項目', width='medium'),
                })
            st.caption(f'目前顯示 {len(_visible)} 筆／共 {len(df_final)} 筆；下載包含全部班表，交通細節與座標放在「逐段交通」工作表。')

        _schedule_view()

        subhead('匯出')
        # 班表 Excel（含逐段交通）產生很費時：改成按需產生，並依班表內容快取；平常調整班表不再為它等待。
        _excel_key = hashlib.sha256(repr((_final_key, _export_names, sorted(config.client_primary_map.items()))).encode()).hexdigest()

        @st.fragment
        def _excel_download():
            cached = res.get('_excel_cache')
            if cached and cached[0] == _excel_key:
                download_button_quiet('⬇️ 下載服務班表 Excel（含逐段交通）', cached[1], '服務班表_含交通工資.xlsx',
                    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key='download_schedule_excel')
                return
            st.caption('Excel 含逐段交通與薪資工作表，檔案較大；需要時再產生（約十數秒），產生後可重複下載。')
            if st.button('📦 產生服務班表 Excel', key='build_schedule_excel'):
                with st.spinner('正在產生 Excel…'):
                    res['_excel_cache'] = (_excel_key, schedule_excel_bytes(
                        export_frame(df_final), export_frame(daily_payroll), export_frame(monthly_payroll), config.client_primary_map,
                        pd.DataFrame([d for d in df_matches.attrs.get('matching_diagnostics', []) if d['任務ID'] not in set(_operational_assignments['任務ID'])])))
                try:
                    st.rerun(scope='fragment')
                except TypeError:
                    st.rerun()

        _excel_download()

        csv_final = export_frame(df_final).to_csv(index=False).encode("utf-8-sig")
        download_button_quiet(
            "⬇️ 匯出最終派單結果（含覆寫） (final_assignment_with_overrides.csv)",
            csv_final,
            file_name="final_assignment_with_overrides.csv",
            mime="text/csv",
        )

        # ---- 更新後的機構 Excel：把固定人選、備援與整列時段調整直接寫回原檔，不必照紀錄手動回填 ----
        subhead('更新機構 Excel')
        st.caption('把本次確定的「歷史首選（固定人選）」「歷史備援」與「週服務計畫的時段調整」寫回你上傳的 Excel，其他工作表、公式與格式保持不變。'
                   '只影響某一天的改派、請假代班與臨時插單不會改動週計畫，它們已自動存檔，並列在上方服務班表 Excel 裡。')

        @st.fragment
        def _source_update_download():
            from schedule_assist import merged_backups, weekly_changes_from_shifts
            from source_workbook import update_source_workbook
            _orig = sheets['Client_Profiles']
            _col = next((c for c in ('歷史首選', '歷史首選居服員ID') if c in _orig), None)
            _was = {str(k).strip(): ('' if pd.isna(v) else str(v).strip())
                    for k, v in zip(_orig['案家ID'], _orig[_col])} if _col else {}
            _owners = {str(c): str(w) for c, w in config.client_primary_map.items()
                       if str(w).strip() and _was.get(str(c), '') != str(w).strip()}
            _backups = merged_backups({'client_backup_map': config.client_backup_map}, _orig,
                                      set(res['df_cg']['居服員ID'].astype(str)))
            _weekly = dict(res.get('weekly_plan_changes') or {})
            _shift_changes, _shift_skipped = weekly_changes_from_shifts((res.get('phase2') or {}).get('time_shifts'),
                                                                        res.get('df_tasks_all', df_tasks))
            _weekly.update(_shift_changes)
            _u1, _u2, _u3 = st.columns(3)
            _u1.metric('固定人選變更', f'{len(_owners)} 案')
            _u2.metric('寫入歷史備援', f'{len(_backups)} 案')
            _u3.metric('週計畫時段調整', f'{len(_weekly)} 列')
            if not (_owners or _backups or _weekly):
                st.info('目前沒有需要寫回 Excel 的固定人選、備援或週計畫時段變更。')
                return
            if _owners:
                with st.expander(f'查看 {len(_owners)} 個固定人選變更'):
                    show_table(pd.DataFrame([{'案家ID': c, '原歷史首選': _was.get(c, '') or '（空白）', '新歷史首選': w}
                                             for c, w in sorted(_owners.items())]), hide_index=True, width='stretch')
            if _shift_skipped is not None and not _shift_skipped.empty:
                st.caption(f'另有 {len(_shift_skipped)} 列時段調整只影響部分日期，不寫回週計畫（需居督決定）。')
            _key = session_store.change_token((sorted(_owners.items()), sorted(_backups.items()), sorted(_weekly.items()), _file_hash))
            _cached = st.session_state.get('_source_update_cache')
            if _cached and _cached[0] == _key:
                download_button_quiet('⬇️ 下載更新後的機構 Excel', _cached[1],
                                      pathlib.Path(_source_name).stem + '_已更新.xlsx',
                                      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key='download_source_update')
                st.caption('下載後請以這份檔案作為下次排班的來源 Excel。原檔不會被覆蓋。')
                return
            if st.button('📦 產生更新後的機構 Excel', key='build_source_update'):
                try:
                    with st.spinner('寫回固定人選、備援與時段調整…'):
                        st.session_state['_source_update_cache'] = (_key, update_source_workbook(
                            _agency_bytes, _owners, _weekly or None, _backups or None))
                except Exception as exc:
                    st.error('產生失敗（請確認 Excel 含「個案資料」與「週服務計畫」工作表）：' + str(exc))
                    return
                try:
                    st.rerun(scope='fragment')
                except TypeError:
                    st.rerun()

        _source_update_download()

        subhead('人工覆寫紀錄')
        _session_overrides = [{'任務ID': k, '目前居服員': caregiver_label(v.get('cg_id')) if v.get('cg_id') else '待安排', '原因': v.get('reason', '')}
                              for k, v in overrides.items()]
        st.caption(f"本次作業共有 {len(_session_overrides)} 筆服務經人工覆寫或標記待安排；"
                   f"稽覈日誌累計 {len(override_log_df)} 筆。")
        _log_tabs = st.tabs(['修改原因統計', '本次人工覆寫', '居督覆寫稽覈日誌', '臨時異動紀錄', '整案調整紀錄', '資格追蹤'])
        with _log_tabs[0]:
            _reasons = override_log_df['變更原因'] if (not override_log_df.empty and '變更原因' in override_log_df) else pd.Series(dtype=object)
            if _reasons.empty:
                st.info('尚無修改紀錄。修改班表時選擇的原因會累計在這裡。')
            else:
                _cats = _reasons.map(reason_category)
                _stats = (_cats.value_counts().reindex([c if c != CHANGE_REASONS[-1] else '其他' for c in CHANGE_REASONS], fill_value=0)
                          .rename_axis('原因類別').reset_index(name='次數'))
                _stats['占比'] = (_stats['次數'] / max(1, int(_stats['次數'].sum())) * 100).round(1).astype(str) + '%'
                st.caption(f'依稽覈日誌累計 {len(_reasons)} 筆修改；「其他」含補充說明及舊版自由輸入的原因。')
                _s1, _s2 = st.columns([1, 1.3])
                with _s1:
                    show_table(_stats, hide_index=True, width='stretch')
                with _s2:
                    st.bar_chart(_stats.set_index('原因類別')[['次數']], height=260)
        _log_tabs = _log_tabs[1:]
        with _log_tabs[0]:
            if _session_overrides:
                show_table(pd.DataFrame(_session_overrides), width='stretch', hide_index=True)
            else:
                st.info('本次作業尚無人工覆寫。')
        with _log_tabs[1]:
            if override_log_df.empty:
                st.info("尚無居督覆寫紀錄。")
            else:
                show_table(override_log_df, width="stretch", hide_index=True)
                download_button_quiet('下載稽覈日誌 CSV', export_frame(override_log_df).to_csv(index=False).encode('utf-8-sig'),
                                      'supervisor_override_log.csv', 'text/csv', key='download_override_log')
        with _log_tabs[2]:
            _dyn_log = st.session_state.get('dynamic_adjustment_log') or []
            if _dyn_log:
                show_table(pd.DataFrame(_dyn_log), width='stretch', hide_index=True)
            else:
                st.info('本次作業尚無臨時異動紀錄（請假、改期、延遲、臨時新增）。')
        with _log_tabs[3]:
            _whole_log = res.get('whole_client_log') or []
            if _whole_log:
                show_table(pd.DataFrame(_whole_log), width='stretch', hide_index=True)
                st.caption('原始 Excel 不會被自動改寫；若這些調整要長期保留，請照這份紀錄修改 Excel 後重新匯入。')
            else:
                st.info('本次作業尚無整案調整。')
        with _log_tabs[4]:
            _follow_rows = st.session_state.get('qualification_followups') or []
            if _follow_rows:
                show_table(pd.DataFrame(_follow_rows), width='stretch', hide_index=True)
                st.caption('補訓期限到期前請確認原居服員已完成訓練並更新員工資料「取得資格」；取得資格後才可改回主責，未完成者維持由有訓練的居服員服務。')
            else:
                st.info('目前沒有資格追蹤項目（案家新增失智、足部照護等資格重新檢核會記在這裡）。')
        module_nav_buttons(M_REVIEW)

    # 自動存檔：作業內容有變動（排班、改派、插單、整案調整、改期…）就寫入本機資料庫。
    try:
        _key_cols = [c for c in ['任務ID', '日期', '時間窗_開始', '時間窗_結束'] if c in res['df_tasks']]
        _token = session_store.change_token((
            res.get('input_signature'), sorted((str(k), repr(v)) for k, v in overrides.items()),
            [(str(i['task'].get('任務ID')), str(i.get('cg_id'))) for i in st.session_state.get('dynamic_insertions', [])],
            len(st.session_state.get('dynamic_adjustment_log') or []), len(res.get('whole_client_log') or []),
            _frame_digest(res['df_tasks'][_key_cols]), sorted((str(k), str(v)) for k, v in config.client_primary_map.items()),
            sorted((str(k), repr(v)) for k, v in (config.client_backup_map or {}).items()),
            repr(sorted((st.session_state.get('pending_edits') or {}).items()))))
        if st.session_state.get('_autosave_token') != _token:
            _changes = sum(1 for v in overrides.values()) + len(st.session_state.get('dynamic_insertions', []))
            _stamp, _size = session_store.save(
                institution_key, _file_hash, st.session_state.get('_data_source_key', ''), st.session_state,
                file_name=_source_name, file_bytes=_agency_bytes,
                summary=f'期間 {_scope_label}　·　已排 {assigned_count:,}／{total_tasks:,} 筆　·　人工修改 {_changes} 筆')
            st.session_state['_autosave_token'] = _token
            st.session_state['_autosave_at'] = _stamp
    except Exception as exc:
        st.session_state['_autosave_at'] = None
        st.session_state['_autosave_error'] = str(exc)[:160]
    with _sidebar_status:
        if st.session_state.get('_autosave_at'):
            st.caption(f"💾 已自動存檔 {str(st.session_state['_autosave_at'])[11:16]}，重新整理可還原")
        elif st.session_state.get('_autosave_error'):
            st.caption('⚠️ 自動存檔失敗：' + st.session_state['_autosave_error'])

    # 整頁畫完後才啟動背景分析（見「03 衝突與異常處理」的未派案家面板）
    try:
        _bg_submit(_bg_pending)
    except NameError:
        pass

else:

    if _module == M_SCHEDULE:
        st.info("確認資料後，按上方「開始排班 · 全期間」產生班表。")
        module_nav_buttons(M_SCHEDULE)
    elif _module in (M_ISSUES, M_REVIEW, M_ANALYTICS):
        module_header(_module)
        empty_module(_module, '尚未產生班表。請先到「02 智慧排班作業」按「開始排班」，完成後這裡會顯示結果。')
