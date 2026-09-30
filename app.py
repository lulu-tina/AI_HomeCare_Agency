"""長照居家照顧 AI 派單系統 - 網頁儀表板



執行方式：

    streamlit run app.py



工作流程：

    1. 側邊欄調整核心派單政策參數（含車程、優先級、照護連續性、財務拆帳比例）。

    2. 區塊①：即時 BA 服務代碼併報法規防呆健檢（Phase 0），資料一載入／編輯

       即顯示，不需按鈕即可檢視。

    3. 區塊②：上傳／檢視／直接編輯居服員、案家、任務三張表格（可新增列與欄位）。

       任務工作表可為單日的「Today_Pending_Tasks」或含日期／星期欄位的月批次

       「Monthly_Pending_Tasks」，兩者皆可直接使用。

    4. 區塊③：按下「執行 AI 最佳化派單」後，依序呼叫 caregiver_engine.py 的

       Phase 1～3（若任務含「日期」欄位，可另選單日或全月一鍵批次排程），並顯示

       KPI（含財務試算）、四合一分析儀表板與指派明細表。



核心運算邏輯（適配度評分、OR-Tools 最佳化、財務點數試算、BA 代碼防呆、DiD 效益

回溯）皆定義於 caregiver_engine.py，本檔僅負責資料編輯介面與結果呈現，不重複

實作演算法。

"""



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



from calendar_view import apply_overrides_to_result, render_calendar_overview, render_task_override_picker

from caregiver_engine import (

    DEFAULT_EXCEL_PATH,

    PipelineConfig,

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



# ==========================================

# 跨平台中文字型處理（避免 st.pyplot 圖表出現亂碼）

# ==========================================

# 直接隨 repo 附帶一份開源中文字型（Noto Sans TC, SIL Open Font License），

# 因為系統套件（packages.txt: fonts-noto-cjk）是否真的裝進 Streamlit Cloud 容器、

# 以及字型註冊名稱是否符合猜測，皆不受我們控制；自帶字型檔可確保任何部署環境

# 都一定找得到，不必依賴系統/雲端環境是否裝好中文字型。

_BUNDLED_FONT_PATH = pathlib.Path(__file__).parent / "assets" / "fonts" / "NotoSansTC-Regular.ttf"





def setup_chinese_font():

    # 依平台猜測系統字型名稱不可靠：猜錯字型 matplotlib 不會報錯，只會靜默退回

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



st.set_page_config(page_title="長照居家照顧 AI 派單系統", page_icon="🏠", layout="wide")

st.title("🏠 長照居家照顧 AI 派單系統")

st.caption("居督與營運團隊可上傳／編輯排班資料、調整派單政策參數，並一鍵執行 AI 最佳化派單")



# ==========================================

# 側邊欄：核心派單政策參數

# ==========================================

defaults = PipelineConfig()



with st.sidebar:

    st.header("⚙️ 核心派單政策參數")



    if st.button("🔄 還原預設值", width="stretch"):

        for key in [

            "cfg_buffer_mins", "cfg_travel_penalty_weight", "cfg_urgent_priority_bonus",

            "cfg_continuity_bonus", "cfg_skill_bonus", "cfg_salary_rate_pct",

        ]:

            if key in st.session_state:

                del st.session_state[key]

        st.rerun()



    buffer_mins = st.slider(

        "任務間轉場緩衝時間（分鐘）", 0.0, 60.0, defaults.buffer_mins, 1.0,

        key="cfg_buffer_mins",

        help="兩個服務任務之間，除實際交通時間外額外預留的緩衝時間，"

        "用於停車、步行、上下樓、門禁、服務紀錄及臨時延誤。",

    )

    travel_penalty_weight = st.slider(

        "車程扣分權重（分／分鐘）", 0.0, 5.0, defaults.travel_penalty_weight, 0.1,

        key="cfg_travel_penalty_weight",

        help="每一分鐘預估車程對適配度分數的扣分幅度。"

        "（目前車程扣分最高為 25 分，避免距離因素過度凌駕照護連續性。）",

    )

    urgent_priority_bonus = st.slider(

        "緊急任務派單優先權（分）", 0.0, 100.0, defaults.urgent_priority_bonus, 5.0,

        key="cfg_urgent_priority_bonus",

        help= "用於整體排班最佳化。當人力或時段不足、無法完成所有任務時，"

        "緊急任務會取得較高的派單優先度；不直接改變居服員本身的適配度。",

    )

    continuity_bonus = st.slider(

        "照護連續性加分（歷史首選居服員）", 0.0, 100.0, defaults.preferred_caregiver_bonus, 1.0,

        key="cfg_continuity_bonus",

        help="若候選居服員等於案家的「歷史首選居服員ID」，即獲得此加分。"

        "目前預設為 +60 分；若該居服員歷史滿意度 ≥ 4.3，系統另再加 +15 分。",

    )

    skill_bonus = st.slider(

        "核心專長匹配加分（分）", 0.0, 50.0, defaults.cert_bonus_dementia, 1.0,

        key="cfg_skill_bonus",

        help="居服員持有之核心專長證照符合案家特殊照護需求時的加分。",

    )



    st.divider()

    st.subheader("💰 財務試算設定")



    salary_rate_pct = st.slider(

        "居服員拆帳比例（%）", 0, 100, int(defaults.caregiver_salary_rate_per_point * 100), 5,

        key="cfg_salary_rate_pct",

        help="僅用於估算居服員拆帳薪資，不作為居服員適配度的加減分依據。"

        "正式導入時可依各機構實際薪資與拆帳制度設定。",

    )



config = PipelineConfig(

    buffer_mins=buffer_mins,

    travel_penalty_weight=travel_penalty_weight,

    urgent_priority_bonus=urgent_priority_bonus,

    preferred_caregiver_bonus=continuity_bonus,

    cert_bonus_dementia=skill_bonus,

    cert_bonus_other=skill_bonus,

    caregiver_salary_rate_per_point=salary_rate_pct / 100.0,

)



# ==========================================

# 資料載入（先於區塊①②，供 Phase 0 健檢與編輯區共用）

# ==========================================

uploaded_file = st.file_uploader("上傳派單資料庫（.xlsx，未上傳則使用預設檔案）", type=["xlsx"])





@st.cache_data(show_spinner="讀取 Excel 資料中...")

def _read_raw_sheets(file_or_path):

    xl = pd.ExcelFile(file_or_path)

    missing_sheets = [s for s in FIXED_REQUIRED_SHEETS if s not in xl.sheet_names]

    if missing_sheets:

        raise ValueError(f"檔案缺少必要工作表：{'、'.join(missing_sheets)}")



    tasks_sheet_name = next((s for s in TASKS_SHEET_CANDIDATES if s in xl.sheet_names), None)

    if tasks_sheet_name is None:

        raise ValueError(f"檔案缺少任務工作表，需具備「{'」或「'.join(TASKS_SHEET_CANDIDATES)}」其中之一。")



    sheets = {name: pd.read_excel(xl, sheet_name=name) for name in FIXED_REQUIRED_SHEETS}

    sheets["Tasks"] = pd.read_excel(xl, sheet_name=tasks_sheet_name)



    column_checks = dict(REQUIRED_COLUMNS)

    column_checks["Tasks"] = REQUIRED_TASKS_COLUMNS

    for name, required_cols in column_checks.items():

        missing_cols = [c for c in required_cols if c not in sheets[name].columns]

        if missing_cols:

            label = tasks_sheet_name if name == "Tasks" else name

            raise ValueError(f"工作表「{label}」缺少必要欄位：{'、'.join(missing_cols)}")



    return sheets, tasks_sheet_name





try:

    if uploaded_file is not None:

        source_key = f"upload::{uploaded_file.name}::{uploaded_file.size}"

        sheets, tasks_sheet_name = _read_raw_sheets(uploaded_file)

    else:

        import os



        if not os.path.exists(DEFAULT_EXCEL_PATH):

            st.error(f"找不到預設檔案：{DEFAULT_EXCEL_PATH}，請改用上方上傳功能。")

            st.stop()

        source_key = f"default::{os.path.getmtime(DEFAULT_EXCEL_PATH)}"

        sheets, tasks_sheet_name = _read_raw_sheets(DEFAULT_EXCEL_PATH)

except ValueError as e:

    st.error(f"⚠️ 資料格式錯誤：{e}")

    st.stop()

except Exception as e:

    st.error(f"⚠️ 無法讀取上傳的檔案，請確認上傳的是有效的 Excel（.xlsx）檔案。錯誤訊息：{e}")

    st.stop()



# 僅在資料來源變更時（首次載入／換檔案／原檔被覆寫）重設編輯區，避免使用者的編輯內容被覆蓋

if st.session_state.get("_data_source_key") != source_key:

    st.session_state["_data_source_key"] = source_key

    st.session_state["edit_cg"] = sheets["Caregiver_Profiles"].copy()

    st.session_state["edit_cl"] = sheets["Client_Profiles"].copy()

    st.session_state["edit_tasks"] = sheets["Tasks"].copy()

    st.session_state["data_hist"] = sheets["Historical_Service_Logs"].copy()

    st.session_state["edit_service_code"] = (

    sheets["Service_Code"].copy()

    )

    st.session_state["tasks_sheet_name"] = tasks_sheet_name



# ==========================================

# 區塊①：Phase 0 申報法規防呆健檢

# ==========================================

st.header("① Phase 0：申報法規防呆健檢")

st.caption(

    "依目前編輯中的任務與案家資料，即時檢核 BA 服務代碼併報合規性（依長照給付支付基準併報規則）。"

    "本健檢僅檢核與提示、不會排除任何任務，亦不影響下方③區塊的派單運算。"

)

try:

    _merged_preview = st.session_state["edit_tasks"].merge(

        st.session_state["edit_cl"], on="案家ID", how="left"

    )

    _validated_preview = validate_ba_codes(_merged_preview)

    _total_tasks_preview = len(_validated_preview)

    _violation_count = int(_validated_preview["含違規代碼"].sum())

    _qualified_count = _total_tasks_preview - _violation_count



    h1, h2, h3 = st.columns(3)

    h1.metric("總任務數", f"{_total_tasks_preview}")

    h2.metric("合格任務數", f"{_qualified_count}")

    h3.metric("⚠️ 偵測到違規申報數", f"{_violation_count}")



    if _violation_count > 0:

        st.warning(

            f"偵測到 {_violation_count} 筆任務的 BA 服務代碼併報疑似違反長照給付支付基準規則，"

            "建議於下方②區塊修正服務代碼，或由居督複核後仍可略過警告繼續派單。"

        )

        st.dataframe(

            _validated_preview.loc[_validated_preview["含違規代碼"], ["任務ID", "案家ID", "BA代碼檢核異常"]],

            width="stretch",

            hide_index=True,

        )

    else:

        st.success("目前資料未偵測到已知的 BA 服務代碼併報違規。")

except Exception as e:

    st.info(f"暫無法執行 BA 代碼健檢（請確認任務與案家資料的「案家ID」欄位可正常對應）：{e}")



# ==========================================

# 區塊②：Excel 載入與現況檢視／動態編輯

# ==========================================

st.header("② 資料載入與編輯")





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

    edited = st.data_editor(

        df,

        num_rows="dynamic",

        width="stretch",

        key=f"editor_{session_key}_{len(df.columns)}",

    )

    st.session_state[session_key] = edited





tab_cg, tab_cl, tab_tasks, tab_service_code = st.tabs([

    "居服員資料 (Caregiver_Profiles)",

    "案家資料 (Client_Profiles)",

    f"任務資料 ({st.session_state['tasks_sheet_name']})",

     "服務碼主檔 (Service_Code)",

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

    st.caption(

        "修改「CareFlow排班分鐘(暫定)」後，"

        "下一次執行派單即會重新計算全部任務的服務歷時。"

    )

    render_editable_sheet("edit_service_code")

# ==========================================

# 區塊③：一鍵執行與成果儀表板

# ==========================================

st.header("③ 執行最佳化派單與成果儀表板")



_tasks_preview_df = st.session_state["edit_tasks"]

has_date_column = "日期" in _tasks_preview_df.columns



schedule_mode = "單日排程"

selected_schedule_date = None

if has_date_column:

    schedule_mode = st.radio(

        "排程範圍",

        ["單日排程", "全月一鍵排程"],

        horizontal=True,

        key="schedule_mode",

        help="偵測到任務資料含「日期」欄位：可選擇僅排單一天的班表，或一次排完整個月"

        "（依日期逐日各自求解 Phase 1 + Phase 2 後彙整為全期間派單總表）。",

    )

    if schedule_mode == "單日排程":

        available_dates = sorted(_tasks_preview_df["日期"].dropna().unique())

        if available_dates:

            selected_schedule_date = st.selectbox("選擇排程日期", available_dates, key="schedule_selected_date")



if st.button("🚀 執行 AI 最佳化派單", type="primary"):

    df_cg = st.session_state["edit_cg"].copy()

    df_cl = st.session_state["edit_cl"].copy()

    df_hist = st.session_state["data_hist"].copy()



    df_tasks = st.session_state["edit_tasks"].copy()

    df_service_code = (

    st.session_state["edit_service_code"].copy()

    )



    if has_date_column and schedule_mode == "單日排程" and selected_schedule_date is not None:

        df_tasks = df_tasks[df_tasks["日期"] == selected_schedule_date].copy()



    try:

        # 先依Service Code Master重算服務歷時

        df_tasks = apply_service_duration(

            df_tasks,

            df_service_code,

        )

        # 再合併案家資料並進入派單

        tasks = df_tasks.merge(

            df_cl,

            on="案家ID",

            how="left",

        )



        df_matches = run_phase1_matching(

            tasks,

            df_cg,

            config,

        )



        df_matches = run_phase1_matching(tasks, df_cg, config)



        if has_date_column and schedule_mode == "全月一鍵排程":

            batch = run_monthly_batch_dispatch(tasks, df_cg, config, date_column="日期")

            failed_dates = [

                d for d, r in batch["daily_results"].items() if r["status"] != pywraplp.Solver.OPTIMAL

            ]

            phase2 = {

                "df_valid": pd.DataFrame(),

                "status": pywraplp.Solver.OPTIMAL if not failed_dates else None,

                "df_result": batch["df_result_all"],

                "assigned_count": batch["total_assigned_count"],

            }

        else:

            phase2 = run_phase2_optimization(df_matches, tasks, df_cg, config)

            failed_dates = []



        did = run_phase3_did(df_hist)

    except Exception as e:

        st.error(

            "⚠️ 派單運算過程發生錯誤，請確認表格中的 ID 對應是否存在、數值欄位是否填寫正確"

            f"（例如經緯度、工時、時間格式「HH:MM」）。錯誤訊息：{e}"

        )

        st.stop()



    st.session_state["last_result"] = {

        "df_tasks": df_tasks,

        "df_cg": df_cg,

        "df_cl": df_cl,

        "df_matches": df_matches,

        "phase2": phase2,

        "did": did,

        "schedule_failed_dates": failed_dates,

    }

    # 每次重新執行最佳化後，AI 建議已改變，先前的居督覆寫不再對應同一份建議，故一併清空。

    st.session_state["overrides"] = {}



if "last_result" in st.session_state:

    res = st.session_state["last_result"]

    df_tasks = res["df_tasks"]

    df_matches = res["df_matches"]

    phase2 = res["phase2"]

    did = res["did"]

    schedule_failed_dates = res.get("schedule_failed_dates") or []



    df_result = phase2["df_result"]

    assigned_count = phase2["assigned_count"]

    solver_ok = phase2["status"] == pywraplp.Solver.OPTIMAL



    # overrides／assigned_map 提前到此處初始化（原僅存在於④區塊），因為月曆視角

    # 「一鍵調班」與④區塊「居督人工覆寫」共用同一份 overrides 狀態與同一套

    # save_override_log 稽核紀錄，月曆總覽（KPI 之後即會渲染）需要在此之前就能

    # 讀寫這份狀態，兩處才不會各自維護一份互不同步的覆寫紀錄。

    st.session_state.setdefault("overrides", {})

    overrides = st.session_state["overrides"]

    assigned_map = dict(zip(df_result["任務ID"], df_result["派單居服員"])) if not df_result.empty else {}

    task_locations = (

        df_matches.drop_duplicates("任務ID").set_index("任務ID")[["地點緯度", "地點經度"]].apply(tuple, axis=1).to_dict()

        if not df_matches.empty

        else {}

    )



    def _quick_reassign(task_id, new_cg_id, reason):

        """統一的居督覆寫寫入入口：即時衝突檢查通過後，寫入 overrides 狀態與稽核

        日誌；回傳 None 表示成功，否則回傳供呼叫端 Modal 顯示的錯誤訊息。



        供月曆「居服員 x 日期」改派 Modal、月曆「未派單案件」處置，以及④區塊

        「居督人工覆寫」任務搜尋器共用（calendar_view._task_override_dialog／

        _quick_reassign_dialog）：三者差異只在 task_id 原本是否已有指派、觸發

        來源不同，apply_overrides_to_result 本就支援替未指派任務新增覆寫列，

        不需要另外實作一套指派邏輯；reason 由呼叫端的原因選擇器收集（必填）。

        """

        df_result_effective = apply_overrides_to_result(df_result, overrides)

        conflict_msg = check_reassignment_conflict(

            task_id, new_cg_id, df_tasks, df_result_effective, res.get("df_cg", pd.DataFrame()),

            config, task_locations=task_locations, date_column="日期",

        )

        if conflict_msg:

            return conflict_msg



        task_rows = df_tasks[df_tasks["任務ID"] == task_id]

        cl_id = task_rows.iloc[0]["案家ID"] if not task_rows.empty else ""

        ai_cg = assigned_map.get(task_id)

        ai_label = str(ai_cg) if ai_cg is not None else "未指派"

        save_override_log(task_id, cl_id, ai_label, new_cg_id if new_cg_id is not None else "未指派", reason)

        overrides[task_id] = {"cg_id": new_cg_id, "reason": reason}

        return None



    def _clear_override(task_id):

        """清除單一任務的居督覆寫，還原為 AI 建議；與原④區塊「清除覆寫」行為

        相同，不寫入稽核日誌（稽核日誌只記錄實際發生過的覆寫變更）。"""

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

        df_result_effective = apply_overrides_to_result(df_result, overrides)

        return rank_candidates_by_availability(

            task_id, candidate_cg_ids, df_tasks, df_result_effective, res.get("df_cg", pd.DataFrame()),

            config, task_locations=task_locations, date_column="日期",

            df_cl=res.get("df_cl", pd.DataFrame()),

        )



    if not solver_ok:

        if schedule_failed_dates:

            st.warning(

                f"以下 {len(schedule_failed_dates)} 個日期無法求得最佳解（其餘日期仍正常派單）："

                f"{', '.join(str(d) for d in schedule_failed_dates)}"

            )

        else:

            st.warning("OR-Tools 無法在目前資料與參數下找到最佳解，請檢查資料是否有效或調整參數。")



    total_tasks = len(df_tasks)

    assign_rate = (assigned_count / total_tasks * 100) if total_tasks else 0.0

    avg_transition = (

        (df_result["預估車程(分)"] + config.buffer_mins).mean() if not df_result.empty else 0.0

    )

    avg_score = df_result["適配分數"].mean() if not df_result.empty else 0.0

    total_revenue = df_result["預估長照申報點數(營收)"].sum() if not df_result.empty else 0.0

    total_salary = df_result["預估居服員拆帳薪資"].sum() if not df_result.empty else 0.0



    st.subheader("📌 KPI 指標")

    k1, k2, k3 = st.columns(3)

    k1.metric("派單成功率", f"{assign_rate:.1f}%", help=f"{assigned_count} / {total_tasks} 筆任務成功指派")

    k2.metric("平均轉場時間", f"{avg_transition:.1f} 分", help="已派單任務的平均車程時間＋轉場緩衝時間")

    k3.metric("平均適配得分", f"{avg_score:.1f} 分", help="已派單配對的平均適配度分數")



    k4, k5 = st.columns(2)

    k4.metric(

        "預估長照申報總點數（營收）", f"{total_revenue:,.0f} 點",

        help="已派單任務的長照申報點數總和，依 BA 服務代碼點值試算（無代碼者以服務歷時概算）。",

    )

    k5.metric(

        "預估居服員拆帳總薪資", f"{total_salary:,.0f} 元",

        help=f"申報點數 × 側邊欄設定的拆帳比例（目前 {salary_rate_pct}%）加總。",

    )



    # 月曆班表總覽：彙總既有派單結果與居服員資料表（並套用 overrides 顯示目前實際

    # 生效班表），不重呼叫任何排班演算法（見 calendar_view.py）；作為整個結果區的

    # 「首頁總覽」，置於 KPI 之後、派單分析儀表板之前。overrides／on_reassign／

    # on_list_candidates 供「月曆視角一鍵調班」（含空檔優先排序）使用。

    render_calendar_overview(

        df_result, df_tasks, res.get("df_cg", pd.DataFrame()),

        overrides=overrides,

        on_reassign=_quick_reassign, on_list_candidates=_list_candidates,

    )



    st.subheader("📊 派單分析儀表板")

    if df_matches.empty:

        st.info("目前無候選配對可供分析（可能所有配對皆被硬性條件過濾）。")

    else:

        sns.set_style("whitegrid")

        setup_chinese_font()  # 須在 sns.set_style 之後呼叫，否則字型會被 seaborn 預設值覆蓋

        fig, axes = plt.subplots(2, 2, figsize=(13, 9))



        # (1) 適配度分數分布

        ax = axes[0, 0]

        ax.hist(df_matches["適配度分數"], bins=20, color=BLUE, edgecolor="white")

        mean_score = df_matches["適配度分數"].mean()

        ax.axvline(mean_score, color=MUTED, linestyle="--", label=f"平均 {mean_score:.1f}")

        ax.set_title("適配度分數分布（全部候選配對）")

        ax.set_xlabel("適配度分數")

        ax.set_ylabel("候選配對數")

        ax.legend()



        # (2) 各居服員派單量

        ax = axes[0, 1]

        if not df_result.empty:

            load_counts = df_result["派單居服員"].value_counts().sort_values(ascending=True)

            ax.barh(load_counts.index.astype(str), load_counts.values, color=BLUE)

            ax.set_xlabel("派單任務數")

        else:

            ax.text(0.5, 0.5, "無派單結果", ha="center", va="center", transform=ax.transAxes)

        ax.set_title("各居服員派單量")



        # (3) 派單狀態占比

        ax = axes[1, 0]

        unassigned = total_tasks - assigned_count

        if total_tasks > 0:

            ax.pie(

                [assigned_count, unassigned],

                labels=["已派單", "未派單"],

                colors=[BLUE, MUTED],

                autopct="%1.1f%%",

                startangle=90,

            )

        ax.set_title("任務派單狀態占比")



        # (4) 目前派單結果中車程較長的任務（不使用不存在的歷史 AI vs 人工基準）

        ax = axes[1, 1]

        if (
            not df_result.empty
            and "預估車程(分)" in df_result.columns
            and "任務ID" in df_result.columns
        ):
            _travel_plot = df_result[["任務ID", "預估車程(分)"]].copy()
            _travel_plot["預估車程(分)"] = pd.to_numeric(
                _travel_plot["預估車程(分)"], errors="coerce"
            )
            _travel_plot = (
                _travel_plot.dropna()
                .sort_values("預估車程(分)", ascending=False)
                .head(5)
                .sort_values("預估車程(分)", ascending=True)
            )

            if not _travel_plot.empty:
                ax.barh(
                    _travel_plot["任務ID"].astype(str),
                    _travel_plot["預估車程(分)"].values,
                    color=ORANGE,
                )
                ax.set_xlabel("預估車程（分鐘）")
                ax.set_title("車程最長的 5 筆已派任務")
            else:
                ax.text(
                    0.5, 0.5, "暫無車程資料",
                    ha="center", va="center", transform=ax.transAxes
                )
                ax.set_title("車程最長的 5 筆已派任務")
        else:
            ax.text(
                0.5, 0.5, "暫無車程資料",
                ha="center", va="center", transform=ax.transAxes
            )
            ax.set_title("車程最長的 5 筆已派任務")


        fig.tight_layout()

        st.pyplot(fig)



    st.subheader("📋 指派明細表")

    if df_result.empty:

        st.info("目前無派單結果。")

    else:

        df_display = df_result.copy()

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

            "任務ID", "案家ID", "派單居服員", "適配分數", "預估車程(分)", "服務時段", "任務優先級",

            "預估長照申報點數(營收)", "預估居服員拆帳薪資", "原首選替換原因",

        ]

        if has_date_for_display:

            display_cols[1:1] = ["日期", "星期"]

        st.dataframe(df_display[display_cols], width="stretch", hide_index=True)



        csv = df_display[display_cols].to_csv(index=False).encode("utf-8-sig")

        st.download_button(

            "⬇️ 匯出指派結果 (assigned_results.csv)",

            csv,

            file_name="assigned_results.csv",

            mime="text/csv",

        )



        # ==========================================

        # Explainable AI：Why this schedule?

        # ==========================================

        st.subheader("🔎 Why this schedule?")

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



        if selected_row.get("工作負荷扣分", 0) <= 5:

            reasons.append("目前工作負荷相對可接受")



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

                    "疲勞／工作負荷",

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



        with st.expander("📊 查看完整適配分數明細"):

            st.dataframe(

                explain_df,

                width="stretch",

                hide_index=True,

            )



            st.caption(

                "所有候選居服員皆須先通過資格、工時、時段及時空衝突等硬性條件；"

                "上述分數僅用於合法可行候選方案之間的比較。"

            )






    # ==========================================
    # 臨時派案文字解析工具（LINE / 訊息貼上）
    # ==========================================
    def _normalize_dispatch_text(raw_text: str) -> str:
        """統一常見符號，讓 LINE / 複製文字更容易解析。"""
        return (
            str(raw_text or "")
            .replace("\\~", "~")
            .replace("～", "~")
            .replace("﹣", "-")
            .replace("－", "-")
            .replace("—", "-")
            .replace("–", "-")
        )


    def _normalize_service_code(raw_code: str, valid_codes: set[str]):
        """回傳 (標準化代碼, 是否為已知誤植修正)。未知代碼不自行猜測。"""
        code = str(raw_code or "").strip().upper()

        # 王處長提供範例中已確認 BA133 為 BA13 的誤植。
        known_aliases = {
            "BA133": "BA13",
        }
        if code in known_aliases and known_aliases[code] in valid_codes:
            return known_aliases[code], True

        if code in valid_codes:
            return code, False

        return None, False


    def _parse_service_items(service_text: str, valid_codes: list[str]):
        """
        解析：BA14+BA13、BA13×2、GA09x1、BA13*2。
        回傳 items=[(code, units), ...] 與 warnings。
        """
        valid_set = set(valid_codes)
        items = []
        warnings = []

        # 代碼格式兼容 BA05-1 / BA16-2 / GA09 / SC09 / AA07 等。
        pattern = re.compile(
            r"(?<![A-Z0-9-])([A-Z]{2,3}\d+(?:-\d+)?)\s*(?:[×xX*]\s*(\d+(?:\.\d+)?))?",
            flags=re.IGNORECASE,
        )

        for m in pattern.finditer(str(service_text or "").upper()):
            raw_code = m.group(1)
            raw_units = m.group(2)
            code, corrected = _normalize_service_code(raw_code, valid_set)

            if code is None:
                warnings.append(f"無法辨識服務碼 {raw_code}，請人工確認。")
                continue

            units = float(raw_units) if raw_units else 1.0
            if units <= 0:
                warnings.append(f"{code} 的 Units 必須大於 0，請人工確認。")
                continue

            if corrected:
                warnings.append(f"已依確認規則將 {raw_code} 修正為 {code}（Units={units:g}）。")

            # 同一行若重複出現同一代碼，合併 Units。
            existing_idx = next((i for i, (c, _) in enumerate(items) if c == code), None)
            if existing_idx is not None:
                old_code, old_units = items[existing_idx]
                items[existing_idx] = (old_code, old_units + units)
            else:
                items.append((code, units))

        return items, warnings


    def _infer_dispatch_year(tasks_df: pd.DataFrame) -> int:
        """只有月/日的 LINE 訊息，以目前任務資料的年份優先推定。"""
        if "日期" in tasks_df.columns:
            parsed_dates = pd.to_datetime(tasks_df["日期"], errors="coerce").dropna()
            if not parsed_dates.empty:
                year_mode = parsed_dates.dt.year.mode()
                if not year_mode.empty:
                    return int(year_mode.iloc[0])
        return int(pd.Timestamp.today().year)


    def _parse_dispatch_message(raw_text: str, valid_codes: list[str], default_year: int):
        """
        將一段 LINE 派案文字解析成 metadata + 1~N 筆服務任務。
        只抽取排班會用到的資訊；原文仍保留供人工複核。
        """
        text_norm = _normalize_dispatch_text(raw_text)
        warnings = []

        address_match = re.search(r"(?:^|\n)\s*地址\s*[:：]\s*([^\n]+)", text_norm)
        address = address_match.group(1).strip() if address_match else ""

        # 指定「居服員性別」只從備註/明示條件抓，不把個案本人性別誤當成派班限制。
        if re.search(r"女性\s*居服員|女\s*居服員", text_norm):
            caregiver_gender = "女"
        elif re.search(r"男性\s*居服員|男\s*居服員", text_norm):
            caregiver_gender = "男"
        else:
            caregiver_gender = "男女不拘"

        note_match = re.search(r"(?:^|\n)\s*備註\s*[:：]\s*(.*)$", text_norm, flags=re.DOTALL)
        note = note_match.group(1).strip() if note_match else ""

        # 例：10/2(五)08:20~11:20 BA14+BA13
        # 也支援 2026/10/2 08:20-11:20 GA09×1
        task_pattern = re.compile(
            r"(?:(?P<year>\d{4})[/-])?"
            r"(?P<month>\d{1,2})[/-](?P<day>\d{1,2})"
            r"(?:\s*\([^\n)]*\))?\s*"
            r"(?P<start_h>\d{1,2}):(?P<start_m>\d{2})\s*[~-]\s*"
            r"(?P<end_h>\d{1,2}):(?P<end_m>\d{2})\s*"
            r"(?P<service>[^\n]*)",
            flags=re.IGNORECASE,
        )

        parsed_tasks = []
        for task_idx, m in enumerate(task_pattern.finditer(text_norm), start=1):
            year = int(m.group("year")) if m.group("year") else int(default_year)
            month = int(m.group("month"))
            day = int(m.group("day"))
            try:
                task_date = pd.Timestamp(year=year, month=month, day=day)
            except ValueError:
                warnings.append(f"無法解析日期：{m.group(0).split()[0]}")
                continue

            start_text = f'{int(m.group("start_h")):02d}:{int(m.group("start_m")):02d}'
            end_text = f'{int(m.group("end_h")):02d}:{int(m.group("end_m")):02d}'
            service_text = m.group("service").strip()
            items, item_warnings = _parse_service_items(service_text, valid_codes)
            warnings.extend(item_warnings)

            if not items:
                warnings.append(
                    f"{task_date.strftime('%m/%d')} {start_text}-{end_text} 未抓到有效服務碼，請人工確認。"
                )

            if len(items) > 2:
                warnings.append(
                    f"{task_date.strftime('%m/%d')} 一次抓到 {len(items)} 個服務碼；目前 CareFlow 單筆任務先支援 2 個，僅帶入前 2 個。"
                )

            first = items[0] if len(items) >= 1 else (None, 0.0)
            second = items[1] if len(items) >= 2 else (None, 0.0)

            parsed_tasks.append({
                "序號": task_idx,
                "日期": task_date,
                "開始時間": start_text,
                "結束時間": end_text,
                "Service_Code_1": first[0],
                "Units_1": first[1],
                "Service_Code_2": second[0],
                "Units_2": second[1],
                "原始服務字串": service_text,
            })

        # 特殊需求目前只做保守辨識；不從長篇病況自行推論重度移位等硬性條件。
        special_need = "一般照護"
        if re.search(r"失智\s*[:：]\s*[Vv✓√]", text_norm):
            special_need = "失智引導與精神陪伴"

        return {
            "address": address,
            "caregiver_gender": caregiver_gender,
            "special_need": special_need,
            "note": note,
            "tasks": parsed_tasks,
            "warnings": list(dict.fromkeys(warnings)),
            "raw_text": raw_text,
        }


    # ==========================================
    # 本機門牌定位：直接讀官方門牌 CSV（沿用 integrated.html / addressConvert.js 的資料來源）
    # ==========================================
    _DEFAULT_GOV_ADDRESS_CSV = (
        r"C:\Users\呂亞庭\OneDrive\桌面\中化比賽\合併去識別化門牌轉換\新北市門牌位置數值資料.csv"
    )

    _NEW_TAIPEI_AREA = {
        "65000010": "板橋區", "65000020": "三重區", "65000030": "中和區", "65000040": "永和區",
        "65000050": "新莊區", "65000060": "新店區", "65000070": "樹林區", "65000080": "鶯歌區",
        "65000090": "三峽區", "65000100": "淡水區", "65000110": "汐止區", "65000120": "瑞芳區",
        "65000130": "土城區", "65000140": "蘆洲區", "65000150": "五股區", "65000160": "泰山區",
        "65000170": "林口區", "65000180": "深坑區", "65000190": "石碇區", "65000200": "坪林區",
        "65000210": "三芝區", "65000220": "石門區", "65000230": "八里區", "65000240": "平溪區",
        "65000250": "雙溪區", "65000260": "貢寮區", "65000270": "金山區", "65000280": "萬里區",
        "65000290": "烏來區",
    }

    def _normalize_str_js_style(value) -> str:
        if value is None:
            return ""
        s = str(value).strip()

        # 全形英數 -> 半形
        chars = []
        for ch in s:
            code = ord(ch)
            if 0xFF10 <= code <= 0xFF19 or 0xFF21 <= code <= 0xFF3A or 0xFF41 <= code <= 0xFF5A:
                chars.append(chr(code - 0xFEE0))
            else:
                chars.append(ch)
        s = "".join(chars)

        s = s.replace("　", "")
        s = re.sub(r"\s+", "", s)
        s = s.replace("臺", "台")

        zh_nums = {
            "一": "1", "二": "2", "三": "3", "四": "4", "五": "5",
            "六": "6", "七": "7", "八": "8", "九": "9", "十": "10",
        }
        s = re.sub(
            r"([一二三四五六七八九十]+)段",
            lambda m: zh_nums.get(m.group(1), m.group(1)) + "段",
            s,
        )
        return s


    def _clean_address_js_style(raw) -> str:
        """比照 addressConvert.js 的 cleanAddress()。"""
        s = _normalize_str_js_style(raw)
        s = re.sub(
            r"^\d{3,6}(?=(台|新北|桃園|新竹|苗栗|台中|彰化|南投|雲林|嘉義|高雄|屏東|宜蘭|花蓮|台東|澎湖|金門|連江))",
            "",
            s,
        )
        s = re.sub(r"(\d+)[\-－](\d+)號", r"\1之\2號", s)
        m = re.match(r"^(.*?\d+(?:之\d+)?號)(?:地下.*|B\d+.*|\d+樓.*|\d+F.*|樓.*|.*室.*)?$", s, flags=re.I)
        return m.group(1) if m else s


    def _detect_text_encoding(path: str) -> str:
        """
        新北官方 CSV 常見 UTF-8 / UTF-8-SIG / CP950(Big5)。
        逐一嘗試讀首段。
        """
        for enc in ("utf-8-sig", "utf-8", "cp950", "big5"):
            try:
                with open(path, "r", encoding=enc, newline="") as f:
                    f.read(65536)
                return enc
            except UnicodeDecodeError:
                continue
        return "utf-8-sig"


    def _normalize_header(v) -> str:
        return str(v or "").replace("\ufeff", "").replace('"', "").replace("'", "").strip().upper()


    def _first_header_index(headers, names):
        h = [_normalize_header(x) for x in headers]
        for name in names:
            name_u = _normalize_header(name)
            if name_u in h:
                return h.index(name_u)
        return -1


    def _detect_gov_header(headers) -> dict:
        return {
            "x": _first_header_index(headers, ["X_3826", "X3826", "TWD97_X", "TWD97X", "X座標", "橫座標", "經度", "LONGITUDE", "LON", "LNG"]),
            "y": _first_header_index(headers, ["Y_3826", "Y3826", "TWD97_Y", "TWD97Y", "Y座標", "縱座標", "緯度", "LATITUDE", "LAT"]),
            "full": _first_header_index(headers, ["FULLADDR", "FULL_ADDR", "完整地址", "門牌地址", "地址"]),
            "county_code": _first_header_index(headers, ["COUNTYCODE", "COUNTY_CODE", "縣市代碼", "省市縣市代碼"]),
            "area_code": _first_header_index(headers, ["AREACODE", "AREA_CODE", "鄉鎮市區代碼"]),
            "county": _first_header_index(headers, ["COUNTY", "CITY", "縣市", "縣市名稱"]),
            "town": _first_header_index(headers, ["TOWN", "DISTRICT", "鄉鎮市區", "市區", "鄉鎮", "行政區"]),
            "road": _first_header_index(headers, ["STREET、ROAD、SECTION", "STREET,ROAD,SECTION", "STREET/ROAD/SECTION", "ROAD", "STREET", "街、路段", "街路段", "路街"]),
            "area": _first_header_index(headers, ["AREA", "地區"]),
            "lane": _first_header_index(headers, ["LANE", "巷"]),
            "alley": _first_header_index(headers, ["ALLEY", "弄"]),
            "num": _first_header_index(headers, ["NUMBER", "號", "門牌號碼"]),
        }


    def _field(row, idx: int) -> str:
        if idx is None or idx < 0 or idx >= len(row):
            return ""
        return str(row[idx] or "").strip()


    def _build_gov_address(row, H) -> str:
        if H["full"] >= 0 and _field(row, H["full"]):
            return _clean_address_js_style(_field(row, H["full"]))

        county = _field(row, H["county"])
        town = _field(row, H["town"])
        cc = _field(row, H["county_code"])
        ac = re.sub(r"\D", "", _field(row, H["area_code"]))

        if (not county or county.isdigit()) and cc == "65000":
            county = "新北市"

        # 新北有些資料是 7 碼行政區代碼，補 0 與原 JS/PS 相容
        if len(ac) == 7:
            ac = ac + "0"
        if (not town or town.isdigit()) and ac in _NEW_TAIPEI_AREA:
            town = _NEW_TAIPEI_AREA[ac]

        road = _field(row, H["road"])
        area = _field(row, H["area"])
        lane = _field(row, H["lane"])
        alley = _field(row, H["alley"])
        num = _field(row, H["num"])

        return _clean_address_js_style(county + town + road + area + lane + alley + num)


    def _twd97_to_wgs84(x: float, y: float):
        """EPSG:3826 -> WGS84，與舊工具 proj4 用途相同。"""
        a = 6378137.0
        b = 6356752.314245
        lon0 = math.radians(121.0)
        k0 = 0.9999
        dx = 250000.0
        e = math.sqrt(1.0 - (b * b) / (a * a))
        xx = x - dx
        M = y / k0
        mu = M / (a * (1 - e * e / 4 - 3 * e**4 / 64 - 5 * e**6 / 256))
        e1 = (1 - math.sqrt(1 - e * e)) / (1 + math.sqrt(1 - e * e))
        J1 = 3 * e1 / 2 - 27 * e1**3 / 32
        J2 = 21 * e1**2 / 16 - 55 * e1**4 / 32
        J3 = 151 * e1**3 / 96
        J4 = 1097 * e1**4 / 512
        fp = mu + J1 * math.sin(2 * mu) + J2 * math.sin(4 * mu) + J3 * math.sin(6 * mu) + J4 * math.sin(8 * mu)
        e2 = (e * a / b) ** 2
        C1 = e2 * math.cos(fp) ** 2
        T1 = math.tan(fp) ** 2
        R1 = a * (1 - e * e) / (1 - e * e * math.sin(fp) ** 2) ** 1.5
        N1 = a / math.sqrt(1 - e * e * math.sin(fp) ** 2)
        D = xx / (N1 * k0)
        Q1 = N1 * math.tan(fp) / R1
        Q2 = D * D / 2
        Q3 = (5 + 3 * T1 + 10 * C1 - 4 * C1 * C1 - 9 * e2) * D**4 / 24
        Q4 = (61 + 90 * T1 + 298 * C1 + 45 * T1 * T1 - 252 * e2 - 3 * C1 * C1) * D**6 / 720
        lat = fp - Q1 * (Q2 - Q3 + Q4)
        Q5 = D
        Q6 = (1 + 2 * T1 + C1) * D**3 / 6
        Q7 = (5 - 2 * C1 + 28 * T1 - 3 * C1 * C1 + 8 * e2 + 24 * T1 * T1) * D**5 / 120
        lon = lon0 + (Q5 - Q6 + Q7) / math.cos(fp)
        return math.degrees(lon), math.degrees(lat)


    def _coord_to_wgs84(x: float, y: float):
        if 118 <= x <= 123.5 and 20 <= y <= 27.5:
            return x, y
        if 100000 <= x <= 400000 and 2300000 <= y <= 2900000:
            return _twd97_to_wgs84(x, y)
        return None


    def _get_official_gov_csv_path() -> pathlib.Path | None:
        env_path = os.getenv("CAREFLOW_GOV_ADDRESS_CSV", "").strip()
        if env_path and pathlib.Path(env_path).is_file():
            return pathlib.Path(env_path)

        default = pathlib.Path(_DEFAULT_GOV_ADDRESS_CSV)
        if default.is_file():
            return default

        # 若之後把門牌檔複製到專案內，也能自動找到
        base = pathlib.Path(__file__).resolve().parent
        for candidate in [
            base / "新北市門牌位置數值資料.csv",
            base / "data" / "新北市門牌位置數值資料.csv",
        ]:
            if candidate.is_file():
                return candidate
        return None


    @st.cache_data(show_spinner=False)
    def _scan_official_gov_csv(csv_path: str, raw_address: str) -> dict:
        """
        直接掃描 integrated.html 原本使用的官方門牌 CSV。
        完整地址：精確匹配。
        O號/X號：以相同路/街/巷前綴的門牌點中位數作 Demo 近似定位。
        """
        target = _clean_address_js_style(raw_address)
        masked = bool(re.search(r"(?:[OoＯ○〇XxＸ＊*]+)\s*號", str(raw_address or "")))

        if masked:
            prefix = re.sub(r"(?:[OoＯ○〇XxＸ＊*]+)\s*號.*$", "", target).strip()
            if not re.search(r"(路|街|巷|弄)", prefix):
                return {
                    "status": "Demo近似定位失敗", "lat": None, "lon": None,
                    "matched": "", "source": pathlib.Path(csv_path).name,
                    "method": "遮蔽地址近似定位",
                    "note": "地址資訊不足，至少需包含路／街／巷／弄。",
                    "demo_only": True,
                }
        else:
            prefix = ""

        encoding = _detect_text_encoding(csv_path)
        exact_hits = []
        approx_hits = []

        with open(csv_path, "r", encoding=encoding, newline="") as f:
            reader = csv_module.reader(f)
            try:
                headers = next(reader)
            except StopIteration:
                return {
                    "status": "門牌檔為空", "lat": None, "lon": None,
                    "matched": "", "source": pathlib.Path(csv_path).name,
                    "method": "官方門牌 CSV", "note": "",
                    "demo_only": masked,
                }

            H = _detect_gov_header(headers)
            if H["x"] < 0 or H["y"] < 0:
                return {
                    "status": "門牌檔欄位不相容", "lat": None, "lon": None,
                    "matched": "", "source": pathlib.Path(csv_path).name,
                    "method": "官方門牌 CSV",
                    "note": "找不到 X/Y 座標欄位。",
                    "demo_only": masked,
                }

            for row in reader:
                if not row:
                    continue

                gov_addr = _build_gov_address(row, H)
                if not gov_addr:
                    continue

                if masked:
                    hit = gov_addr.startswith(prefix)
                else:
                    hit = gov_addr == target

                if not hit:
                    continue

                try:
                    x = float(_field(row, H["x"]))
                    y = float(_field(row, H["y"]))
                except (TypeError, ValueError):
                    continue

                coord = _coord_to_wgs84(x, y)
                if not coord:
                    continue
                lon, lat = coord

                rec = {"lon": float(lon), "lat": float(lat), "matched": gov_addr}
                if masked:
                    approx_hits.append(rec)
                else:
                    exact_hits.append(rec)

        source = pathlib.Path(csv_path).name

        if not masked:
            if not exact_hits:
                return {
                    "status": "找不到", "lat": None, "lon": None,
                    "matched": "", "source": source,
                    "method": "完整地址精確比對",
                    "note": "",
                    "demo_only": False,
                }

            coords = {(round(x["lon"], 7), round(x["lat"], 7)) for x in exact_hits}
            if len(coords) == 1:
                c = exact_hits[0]
                return {
                    "status": "精確門牌",
                    "lat": c["lat"], "lon": c["lon"],
                    "matched": c["matched"], "source": source,
                    "method": "完整地址精確比對",
                    "note": "",
                    "demo_only": False,
                }

            return {
                "status": "需人工確認", "lat": None, "lon": None,
                "matched": " | ".join(x["matched"] for x in exact_hits[:3]),
                "source": source, "method": "完整地址多筆不同座標",
                "note": "未自動選擇座標。",
                "demo_only": False,
            }

        if not approx_hits:
            return {
                "status": "Demo近似定位失敗", "lat": None, "lon": None,
                "matched": "", "source": source,
                "method": "區+路/街/巷近似中心",
                "note": f"官方門牌檔找不到「{prefix}」的可用門牌點。",
                "demo_only": True,
            }

        approx_lat = float(statistics.median(x["lat"] for x in approx_hits))
        approx_lon = float(statistics.median(x["lon"] for x in approx_hits))
        return {
            "status": "Demo近似定位",
            "lat": approx_lat,
            "lon": approx_lon,
            "matched": f"{prefix}（近似中心，{len(approx_hits)} 個門牌點）",
            "source": source,
            "method": "區+路/街/巷近似中心",
            "note": "⚠️ 僅供 Demo／功能測試，不可用於正式派單或實際到宅導航。",
            "demo_only": True,
        }


    def _lookup_local_address(address: str) -> dict:
        raw = str(address or "").strip()
        if not raw:
            return {
                "status": "空白地址", "lat": None, "lon": None,
                "matched": "", "source": "", "method": "", "note": "",
                "db_path": "", "demo_only": False,
            }

        csv_path = _get_official_gov_csv_path()
        if csv_path is None:
            return {
                "status": "找不到官方門牌 CSV", "lat": None, "lon": None,
                "matched": "", "source": "", "method": "未定位",
                "note": (
                    "找不到「新北市門牌位置數值資料.csv」。"
                    "目前預設位置為："
                    + _DEFAULT_GOV_ADDRESS_CSV
                ),
                "db_path": "",
                "demo_only": False,
            }

        try:
            result = _scan_official_gov_csv(str(csv_path), raw)
            result["db_path"] = str(csv_path)
            return result
        except Exception as exc:
            return {
                "status": "定位失敗", "lat": None, "lon": None,
                "matched": "", "source": pathlib.Path(csv_path).name,
                "method": "官方門牌 CSV",
                "note": str(exc),
                "db_path": str(csv_path),
                "demo_only": False,
            }


    def _apply_local_geocode_to_form(address: str) -> dict:
        result = _lookup_local_address(address)

        st.session_state["dynamic_geocode_status"] = result.get("status", "")
        st.session_state["dynamic_geocode_matched"] = result.get("matched", "")
        st.session_state["dynamic_geocode_note"] = result.get("note", "")
        st.session_state["dynamic_geocode_db"] = result.get("db_path", "")
        st.session_state["dynamic_geocode_demo_only"] = bool(result.get("demo_only", False))

        if result.get("lat") is not None and result.get("lon") is not None:
            st.session_state["dynamic_lat"] = float(result["lat"])
            st.session_state["dynamic_lon"] = float(result["lon"])
        else:
            st.session_state["dynamic_lat"] = 0.0
            st.session_state["dynamic_lon"] = 0.0
        return result


    def _apply_parsed_task_to_form(parsed_payload: dict, task_index: int, service_codes: list[str]):
        """把解析結果寫進 Streamlit session_state，之後 rerun 讓表單自動帶入。"""
        tasks = parsed_payload.get("tasks") or []
        if not tasks:
            return

        task_index = max(0, min(int(task_index), len(tasks) - 1))
        task = tasks[task_index]

        st.session_state["dynamic_date"] = pd.Timestamp(task["日期"]).date()
        st.session_state["dynamic_start"] = pd.Timestamp(task["開始時間"]).time()
        st.session_state["dynamic_end"] = pd.Timestamp(task["結束時間"]).time()
        st.session_state["dynamic_task_id"] = f"TEMP_{task_index + 1:03d}"
        st.session_state["dynamic_client_id"] = "TEMP_CLIENT_001"
        st.session_state["dynamic_gender"] = parsed_payload.get("caregiver_gender", "男女不拘")
        st.session_state["dynamic_address"] = parsed_payload.get("address", "")

        # 解析出地址後立即用院內／本機門牌資料庫定位；不把地址送到外部。
        _apply_local_geocode_to_form(st.session_state["dynamic_address"])

        parsed_need = parsed_payload.get("special_need", "一般照護")
        st.session_state["dynamic_need"] = parsed_need

        code1 = task.get("Service_Code_1")
        code2 = task.get("Service_Code_2")
        if code1 in service_codes:
            st.session_state["dynamic_code1"] = code1
            st.session_state["dynamic_units1"] = float(task.get("Units_1") or 1.0)
        if code2 in service_codes:
            st.session_state["dynamic_code2"] = code2
            st.session_state["dynamic_units2"] = float(task.get("Units_2") or 0.0)
        else:
            st.session_state["dynamic_code2"] = ""
            st.session_state["dynamic_units2"] = 0.0

        # 備註只自動帶入「備註段落」；原始全文另外保存，不塞爆表單。
        st.session_state["dynamic_note"] = parsed_payload.get("note", "")

    # ==========================================
    # ==========================================
    # 區塊④：動態事件處理
    # ==========================================
    st.header("④ ⚡ 動態事件處理")
    st.caption(
        "遇到臨時狀況時先選事件類型；系統只展開需要的流程，"
        "並盡量保留其他既有班表不動。"
    )

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


    _dynamic_event_type = st.radio(
        "選擇事件類型",
        ["➕ 臨時新增案件", "🚨 居服員臨時請假", "🕒 案家臨時改期", "⏱️ 前一服務延遲"],
        horizontal=True,
        key="dynamic_event_type_selector",
    )

    if _dynamic_event_type == "➕ 臨時新增案件":
        with st.expander("➕ 建立臨時新增案件", expanded=False):
            _service_master = st.session_state["edit_service_code"].copy()
            _service_codes = [
                str(v).strip() for v in _service_master["系統代碼"].dropna().tolist()
                if str(v).strip()
            ]

            st.markdown("#### ✨ 先貼 LINE / 派案文字，自動帶入")
            st.caption(
                "例如 `BA13×2` 會自動解析成 Service Code = BA13、Units = 2；"
                "`BA14+BA13` 則會帶成兩個服務碼、各 1 Units。解析後仍可手動修改。"
            )
            dynamic_raw_message = st.text_area(
                "貼上完整派案訊息",
                height=220,
                placeholder=(
                    "請問貴單位有人力嗎？\n"
                    "地址：新北市中和區...\n"
                    "服務時間及項目：\n"
                    "10/2(五)08:20~11:20 BA14+BA13\n"
                    "備註：\n1.女性居服員..."
                ),
                key="dynamic_raw_message",
            )

            parse_col1, parse_col2 = st.columns([1, 3])
            if parse_col1.button("✨ 自動解析文字", type="secondary", key="dynamic_parse_text"):
                if not dynamic_raw_message.strip():
                    st.warning("請先貼上派案文字。")
                else:
                    _parsed = _parse_dispatch_message(
                        dynamic_raw_message,
                        _service_codes,
                        _infer_dispatch_year(df_tasks),
                    )
                    st.session_state["dynamic_parsed_payload"] = _parsed
                    st.session_state["dynamic_parsed_task_index"] = 0
                    if _parsed.get("tasks"):
                        _apply_parsed_task_to_form(_parsed, 0, _service_codes)
                        st.rerun()

            _parsed_payload = st.session_state.get("dynamic_parsed_payload")
            if isinstance(_parsed_payload, dict):
                _parsed_tasks = _parsed_payload.get("tasks") or []
                if _parsed_tasks:
                    st.success(f"已解析出 {len(_parsed_tasks)} 筆服務時段，可先確認後再手動修正。")
                    _preview = pd.DataFrame(_parsed_tasks).copy()
                    if "日期" in _preview.columns:
                        _preview["日期"] = pd.to_datetime(_preview["日期"]).dt.strftime("%Y/%m/%d")
                    st.dataframe(
                        _preview[
                            [
                                "序號", "日期", "開始時間", "結束時間",
                                "Service_Code_1", "Units_1", "Service_Code_2", "Units_2",
                                "原始服務字串",
                            ]
                        ],
                        width="stretch",
                        hide_index=True,
                    )

                    if len(_parsed_tasks) > 1:
                        _task_labels = [
                            f"第 {i+1} 筆｜{pd.Timestamp(t['日期']).strftime('%m/%d')} "
                            f"{t['開始時間']}-{t['結束時間']}"
                            for i, t in enumerate(_parsed_tasks)
                        ]
                        _selected_idx = st.selectbox(
                            "選擇要載入表單的服務時段",
                            range(len(_task_labels)),
                            format_func=lambda i: _task_labels[i],
                            key="dynamic_parsed_task_index",
                        )
                        if st.button("⬇️ 套用選取時段到下方欄位", key="dynamic_apply_parsed_task"):
                            _apply_parsed_task_to_form(
                                _parsed_payload, _selected_idx, _service_codes
                            )
                            st.rerun()

                    if _parsed_payload.get("warnings"):
                        for _warning in _parsed_payload["warnings"]:
                            st.warning(_warning)

                    if _parsed_payload.get("address"):
                        _geo_status = st.session_state.get("dynamic_geocode_status", "")
                        _geo_matched = st.session_state.get("dynamic_geocode_matched", "")
                        _geo_note = st.session_state.get("dynamic_geocode_note", "")
                        if _geo_status == "精確門牌":
                            st.success(
                                f"📍 已在本機完成定位：{_parsed_payload['address']} → "
                                f"{st.session_state.get('dynamic_lat', 0):.6f}, "
                                f"{st.session_state.get('dynamic_lon', 0):.6f}"
                            )
                            if _geo_matched:
                                st.caption(f"本機匹配門牌：{_geo_matched}")
                        elif _geo_status == "Demo近似定位":
                            st.warning(
                                f"🧪 已使用遮蔽地址做 Demo 近似定位："
                                f"{st.session_state.get('dynamic_lat', 0):.6f}, "
                                f"{st.session_state.get('dynamic_lon', 0):.6f}\n\n"
                                "⚠️ 僅供 Demo／功能測試，不可正式派單或作為到宅導航地址。"
                            )
                            if _geo_matched:
                                st.caption(f"近似範圍：{_geo_matched}")
                        elif _geo_status == "需人工確認" and st.session_state.get("dynamic_lat"):
                            st.warning(
                                f"📍 已找到一組可能座標，但需要人工確認：{_geo_matched or _parsed_payload['address']}"
                            )
                        else:
                            st.warning(
                                f"📍 地址已抓到，但尚未完成定位：{_geo_status or '未定位'}。"
                                f"{(' ' + _geo_note) if _geo_note else ''}"
                            )
                else:
                    st.warning("這段文字沒有抓到『日期＋時間＋服務碼』的服務時段，請手動輸入或調整文字格式。")

            st.divider()
            st.markdown("#### 📝 解析後可在這裡手動確認 / 修正")

            d1, d2, d3 = st.columns(3)

            if "日期" in df_tasks.columns and not df_tasks["日期"].dropna().empty:
                _default_dynamic_date = pd.to_datetime(df_tasks["日期"].dropna().iloc[0]).date()
            else:
                _default_dynamic_date = pd.Timestamp.today().date()

            dynamic_date = d1.date_input(
                "日期", value=_default_dynamic_date, key="dynamic_date"
            )
            dynamic_start = d2.time_input(
                "開始時間", value=pd.Timestamp("09:00").time(), key="dynamic_start"
            )
            dynamic_end = d3.time_input(
                "結束時間", value=pd.Timestamp("10:00").time(), key="dynamic_end"
            )

            c1, c2 = st.columns(2)
            dynamic_task_id = c1.text_input(
                "臨時任務ID", value="TEMP_001", key="dynamic_task_id"
            )
            dynamic_client_id = c2.text_input(
                "案家ID", value="TEMP_CLIENT_001", key="dynamic_client_id"
            )

            addr_col, geo_col = st.columns([4, 1])
            dynamic_address = addr_col.text_input(
                "服務地址（由派案文字自動抓取，可手動修改）",
                value="",
                key="dynamic_address",
            )
            if geo_col.button("📍 本機重新定位", key="dynamic_regeocode", width="stretch"):
                _geo = _apply_local_geocode_to_form(dynamic_address)
                if _geo.get("status") == "精確門牌":
                    st.success("本機定位完成。")
                elif _geo.get("lat") is not None and _geo.get("lon") is not None:
                    st.warning("找到可能座標，但請人工確認匹配門牌。")
                else:
                    st.warning(f"尚未定位：{_geo.get('status')}。{_geo.get('note', '')}")
                st.rerun()

            p1, p2 = st.columns(2)
            dynamic_lat = p1.number_input(
                "服務地點緯度", value=0.000000, format="%.6f", key="dynamic_lat"
            )
            dynamic_lon = p2.number_input(
                "服務地點經度", value=0.000000, format="%.6f", key="dynamic_lon"
            )

            _geo_status = st.session_state.get("dynamic_geocode_status", "")
            _geo_matched = st.session_state.get("dynamic_geocode_matched", "")
            _geo_note = st.session_state.get("dynamic_geocode_note", "")
            _geo_db = st.session_state.get("dynamic_geocode_db", "")

            if _geo_status == "精確門牌":
                st.success(f"🔒 本機定位完成｜{_geo_matched or dynamic_address}")
            elif _geo_status == "Demo近似定位":
                st.error(
                    "🧪 Demo 近似定位｜"
                    f"{_geo_matched or dynamic_address}\n\n"
                    "⚠️ 此座標僅供系統展示與演算法測試，不可用於正式派單或實際到宅導航。"
                )
            elif _geo_status == "需人工確認" and dynamic_lat and dynamic_lon:
                st.warning(f"⚠️ 本機定位需人工確認｜{_geo_matched or dynamic_address}")
            elif _geo_status:
                st.info(f"定位狀態：{_geo_status}" + (f"｜{_geo_note}" if _geo_note else ""))

            if _geo_db:
                st.caption(f"本機官方門牌檔：{_geo_db}")
            else:
                st.caption(
                    "地址定位只使用本機 address_db.tsv，不會把真實地址送到外部 geocoding API。"
                    "若尚未建立資料庫，可先手動輸入經緯度。"
                )

            if st.session_state.get("dynamic_geocode_demo_only", False):
                st.caption("🧪 DEMO ONLY：遮蔽門牌座標為同一路／街／巷門牌點的近似中心。")

            f1, f2, f3 = st.columns(3)
            dynamic_gender = f1.selectbox(
                "指定居服員性別",
                ["男女不拘", "女", "男"],
                key="dynamic_gender",
            )

            _care_need_options = ["一般照護"]
            if "特殊照護需求" in res.get("df_cl", pd.DataFrame()).columns:
                _existing_needs = [
                    str(v) for v in res["df_cl"]["特殊照護需求"].dropna().unique()
                    if str(v).strip()
                ]
                _care_need_options = list(dict.fromkeys(_care_need_options + _existing_needs))

            # 若文字解析抓到一個新選項，但原始案家主檔沒有，仍先讓畫面能顯示並人工確認。
            _parsed_need = st.session_state.get("dynamic_need", "一般照護")
            if _parsed_need not in _care_need_options:
                _care_need_options.append(_parsed_need)

            dynamic_need = f2.selectbox(
                "特殊照護需求", _care_need_options, key="dynamic_need"
            )
            dynamic_heavy = f3.checkbox(
                "需重度移位協助", value=False, key="dynamic_heavy"
            )

            dynamic_env = st.text_input(
                "案家環境特徵",
                value="一般居家環境",
                key="dynamic_env",
            )

            s1, s2 = st.columns(2)
            dynamic_code1 = s1.selectbox(
                "Service Code 1", _service_codes, key="dynamic_code1"
            )
            dynamic_units1 = s2.number_input(
                "Units 1", min_value=0.0, value=1.0, step=1.0, key="dynamic_units1"
            )

            s3, s4 = st.columns(2)
            dynamic_code2 = s3.selectbox(
                "Service Code 2（可不選）",
                [""] + _service_codes,
                key="dynamic_code2",
            )
            dynamic_units2 = s4.number_input(
                "Units 2", min_value=0.0, value=0.0, step=1.0, key="dynamic_units2"
            )

            dynamic_note = st.text_area(
                "備註（解析後仍可手動修改）",
                placeholder="例如：女性居服員；外出就醫需注意安全；可拆單位…",
                key="dynamic_note",
            )

            if st.button("🔎 尋找可直接插入的 Top 3", type="primary", key="dynamic_find"):
                if dynamic_end <= dynamic_start:
                    st.error("結束時間必須晚於開始時間。")
                elif not (20.0 <= float(dynamic_lat) <= 27.5 and 118.0 <= float(dynamic_lon) <= 123.5):
                    st.error(
                        "尚未取得有效的台灣經緯度。請先按「📍 本機重新定位」，"
                        "或人工確認後輸入正確座標，再進行插單。"
                    )
                else:
                    if st.session_state.get("dynamic_geocode_demo_only", False):
                        st.warning(
                            "🧪 本次候選排序使用的是 Demo 近似座標；結果只用來展示插單流程，"
                            "不可視為正式派單依據。"
                        )
                    _temp_task = pd.DataFrame(
                        [{
                            "任務ID": dynamic_task_id.strip(),
                            "案家ID": dynamic_client_id.strip(),
                            "日期": pd.Timestamp(dynamic_date),
                            "星期": get_weekday_name(pd.Timestamp(dynamic_date)),
                            "時間窗_開始": dynamic_start.strftime("%H:%M"),
                            "時間窗_結束": dynamic_end.strftime("%H:%M"),
                            "任務優先級": "高",
                            "Service_Code_1": dynamic_code1,
                            "Units_1": dynamic_units1,
                            "Service_Code_2": dynamic_code2 if dynamic_code2 else None,
                            "Units_2": dynamic_units2 if dynamic_code2 else None,
                        }]
                    )

                    try:
                        _temp_task = apply_service_duration(
                            _temp_task,
                            st.session_state["edit_service_code"].copy(),
                        )

                        _temp_merged = _temp_task.copy()
                        _temp_merged["指定居服員性別"] = dynamic_gender
                        _temp_merged["服務地點_緯度"] = float(dynamic_lat)
                        _temp_merged["服務地點_經度"] = float(dynamic_lon)
                        _temp_merged["特殊照護需求"] = dynamic_need
                        _temp_merged["案家環境特徵"] = dynamic_env
                        _temp_merged["需重度移位協助(0/1)"] = int(dynamic_heavy)
                        _temp_merged["歷史首選居服員ID"] = ""
                        _temp_merged["臨時事件備註"] = dynamic_note

                        _current_tasks_for_insert = df_tasks.merge(
                            res.get("df_cl", pd.DataFrame()),
                            on="案家ID",
                            how="left",
                        )

                        _effective_result_for_insert = apply_overrides_to_result(
                            df_result, overrides
                        )

                        # 已確認的臨時插單也要納入下一次候選檢查，避免後續插單互相撞班。
                        if st.session_state["dynamic_insertions"]:
                            _extra_task_rows = []
                            _extra_result_rows = []
                            for _item in st.session_state["dynamic_insertions"]:
                                _extra_task_rows.append(_item["task"])
                                _extra_result_rows.append(
                                    {
                                        "任務ID": _item["task"]["任務ID"],
                                        "派單居服員": _item["cg_id"],
                                    }
                                )
                            if _extra_task_rows:
                                _current_tasks_for_insert = pd.concat(
                                    [
                                        _current_tasks_for_insert,
                                        pd.DataFrame(_extra_task_rows),
                                    ],
                                    ignore_index=True,
                                    sort=False,
                                )
                                _effective_result_for_insert = pd.concat(
                                    [
                                        _effective_result_for_insert,
                                        pd.DataFrame(_extra_result_rows),
                                    ],
                                    ignore_index=True,
                                    sort=False,
                                )

                        _candidate_df = find_insertion_candidates(
                            _temp_merged.iloc[0],
                            _current_tasks_for_insert,
                            _effective_result_for_insert,
                            res.get("df_cg", pd.DataFrame()),
                            config,
                            top_n=3,
                            date_column="日期",
                        )

                        st.session_state["dynamic_candidate_result"] = _candidate_df
                        st.session_state["dynamic_candidate_task"] = _temp_merged.iloc[0].to_dict()
                        st.session_state["dynamic_candidate_raw_message"] = dynamic_raw_message
                    except Exception as e:
                        st.session_state["dynamic_candidate_result"] = None
                        st.session_state["dynamic_candidate_task"] = None
                        st.error(f"無法計算臨時插單候選人：{e}")

        _candidate_df = st.session_state.get("dynamic_candidate_result")
        _candidate_task = st.session_state.get("dynamic_candidate_task")

        if isinstance(_candidate_df, pd.DataFrame) and not _candidate_df.empty:
            _available_candidates = _candidate_df[_candidate_df["可直接插入"] == True].copy()

            if _available_candidates.empty:
                st.warning(
                    "目前找不到不移動既有班表即可直接插入的人選。"
                    "下方顯示最接近的候選與阻擋原因；下一階段可再做「局部重排」。"
                )
                st.dataframe(_candidate_df, width="stretch", hide_index=True)
            else:
                st.success(
                    f"找到 {len(_available_candidates)} 位可直接插入的候選居服員。"
                    "請由督導選擇最終人選。"
                )

                # 保存當次 Top 3 排序，用於之後的稽核紀錄。
                _top3_snapshot = _available_candidates.head(3).copy()
                st.dataframe(_top3_snapshot, width="stretch", hide_index=True)

                st.markdown("##### 👤 督導確認人選")
                _top1_id = str(_top3_snapshot.iloc[0]["居服員ID"]) if not _top3_snapshot.empty else ""

                # 每位候選人各自一個按鈕，避免還要另外下拉選擇。
                for _rank, (_idx, _cand) in enumerate(_top3_snapshot.iterrows(), start=1):
                    _cg_id = str(_cand.get("居服員ID", ""))
                    _score = _cand.get("適配度分數", _cand.get("適配分數", None))
                    _prev_travel = _cand.get("前段交通時間", _cand.get("前段交通時間(分)", None))
                    _next_travel = _cand.get("後段交通時間", _cand.get("後段交通時間(分)", None))
                    _reason = _cand.get("推薦原因", "")
                    _cost = _cand.get("擾動成本", 0)
                    _prev_slot = str(_cand.get("前一任務（結束）", "無"))
                    _next_slot = str(_cand.get("下一任務（開始）", "無"))

                    with st.container(border=True):
                        _a, _b = st.columns([4, 1])
                        _summary_parts = [f"**候選 {_rank}｜{_cg_id}**"]
                        if pd.notna(_score):
                            _summary_parts.append(f"適配分數：{float(_score):.1f}")
                        if pd.notna(_prev_travel):
                            _summary_parts.append(f"前段交通：{float(_prev_travel):.0f} 分")
                        if pd.notna(_next_travel):
                            _summary_parts.append(f"後段交通：{float(_next_travel):.0f} 分")
                        _summary_parts.append(f"擾動成本：{_cost}")
                        _a.markdown("　｜　".join(_summary_parts))
                        _a.caption(
                            f"🕒 排班銜接：前一任務 {_prev_slot} → 本次服務 → 下一任務 {_next_slot}"
                        )
                        if str(_reason).strip():
                            _a.caption(str(_reason))

                        if _b.button(
                            "✅ 選擇此人",
                            key=f"dynamic_choose_{_candidate_task.get('任務ID','TEMP')}_{_rank}_{_cg_id}",
                            width="stretch",
                        ):
                            st.session_state["dynamic_pending_cg"] = _cg_id
                            st.session_state["dynamic_pending_rank"] = _rank
                            st.session_state["dynamic_pending_candidate"] = _cand.to_dict()
                            st.rerun()

                _pending_cg = st.session_state.get("dynamic_pending_cg")
                if _pending_cg:
                    _pending_rank = int(st.session_state.get("dynamic_pending_rank", 1))
                    _selected_candidate = st.session_state.get("dynamic_pending_candidate", {})

                    def _safe_cost(v):
                        try:
                            return float(v)
                        except (TypeError, ValueError):
                            return 0.0

                    _selected_cost = _safe_cost(_selected_candidate.get("擾動成本", 0))
                    _min_cost = min(
                        _safe_cost(v)
                        for v in _top3_snapshot["擾動成本"].tolist()
                    ) if "擾動成本" in _top3_snapshot.columns and not _top3_snapshot.empty else 0.0

                    _requires_reason = _selected_cost > (_min_cost + 1e-9)

                    st.info(
                        f"目前選擇：{_pending_cg}（系統候選第 {_pending_rank} 名｜"
                        f"擾動成本 {_selected_cost:g}）"
                    )

                    _manual_reason = ""
                    if _requires_reason:
                        st.warning(
                            f"目前方案擾動成本為 {_selected_cost:g}，高於可選方案最低擾動成本 "
                            f"{_min_cost:g}。若仍要採用此方案，請留下原因以供後續稽核。"
                        )
                        _manual_reason = st.text_input(
                            "改選原因（必填）",
                            placeholder="例如：案家偏好、服務連續性、長者接受度、督導已掌握的現場資訊…",
                            key="dynamic_manual_choice_reason",
                        )
                    elif _pending_rank > 1:
                        st.caption(
                            "此人雖不是適配分數第 1 名，但與最佳候選的擾動成本相同，"
                            "可直接由督導選擇，不需填寫改選原因。"
                        )

                    _confirm_disabled = _requires_reason and not str(_manual_reason).strip()

                    if st.button(
                        "📌 確認插入當日班表",
                        type="primary",
                        key="dynamic_confirm_final",
                        disabled=_confirm_disabled,
                    ):
                        _task_id = str(_candidate_task["任務ID"])
                        _already_exists = any(
                            str(item["task"]["任務ID"]) == _task_id
                            for item in st.session_state["dynamic_insertions"]
                        )

                        if _already_exists:
                            st.warning(f"臨時任務 {_task_id} 已加入本次班表。")
                        else:
                            _selected_candidate = st.session_state.get(
                                "dynamic_pending_candidate", {}
                            )
                            _raw_message = st.session_state.get(
                                "dynamic_candidate_raw_message", ""
                            )

                            # Top 3 只保留可讀欄位字串，方便 CSV 稽核。
                            _top3_ids = " > ".join(
                                _top3_snapshot["居服員ID"].astype(str).tolist()
                            )

                            _confirmed_at = pd.Timestamp.now()
                            _is_manual_override = (_pending_rank > 1)
                            _is_higher_disruption_override = bool(_requires_reason)

                            st.session_state["dynamic_insertions"].append(
                                {
                                    "task": dict(_candidate_task),
                                    "cg_id": _pending_cg,
                                    "confirmed_at": _confirmed_at,
                                    "candidate_rank": _pending_rank,
                                    "system_top1_id": _top1_id,
                                    "manual_reason": str(_manual_reason).strip(),
                                    "candidate_detail": dict(_selected_candidate),
                                }
                            )

                            st.session_state["dynamic_adjustment_log"].append(
                                {
                                    "時間戳記": _confirmed_at,
                                    "事件類型": "臨時新增案件",
                                    "任務ID": _task_id,
                                    "案家ID": _candidate_task.get("案家ID"),
                                    "日期": _candidate_task.get("日期"),
                                    "服務時段": (
                                        f"{_candidate_task.get('時間窗_開始')}-"
                                        f"{_candidate_task.get('時間窗_結束')}"
                                    ),
                                    "Service_Code_1": _candidate_task.get("Service_Code_1"),
                                    "Units_1": _candidate_task.get("Units_1"),
                                    "Service_Code_2": _candidate_task.get("Service_Code_2"),
                                    "Units_2": _candidate_task.get("Units_2"),
                                    "定位模式": (
                                        "Demo近似定位"
                                        if st.session_state.get("dynamic_geocode_demo_only", False)
                                        else "正式座標"
                                    ),
                                    "系統Top3": _top3_ids,
                                    "系統第一名": _top1_id,
                                    "最終選擇居服員": _pending_cg,
                                    "最終候選排名": _pending_rank,
                                    "是否人工改選": "是" if _is_manual_override else "否",
                                    "是否選擇較高擾動方案": "是" if _is_higher_disruption_override else "否",
                                    "人工改選原因": str(_manual_reason).strip(),
                                    "既有班表是否移動": "否",
                                    "最低可選擾動成本": _min_cost,
                                    "擾動成本": _selected_candidate.get("擾動成本", 0),
                                    "原始派案文字": _raw_message,
                                }
                            )

                            st.success(
                                f"✅ 已將 {_task_id} 插入 {_pending_cg} 的班表；"
                                "既有任務未被移動，並已建立稽核紀錄。"
                            )

                            # 清掉待確認狀態，避免 rerun 後重複確認。
                            for _k in [
                                "dynamic_candidate_result",
                                "dynamic_candidate_task",
                                "dynamic_pending_cg",
                                "dynamic_pending_rank",
                                "dynamic_pending_candidate",
                                "dynamic_manual_choice_reason",
                            ]:
                                st.session_state.pop(_k, None)
                            st.rerun()

        if st.session_state["dynamic_insertions"]:
            st.markdown("##### 本次已確認的臨時插單")
            _insert_rows = []
            for _item in st.session_state["dynamic_insertions"]:
                _t = _item["task"]
                _insert_rows.append(
                    {
                        "任務ID": _t.get("任務ID"),
                        "案家ID": _t.get("案家ID"),
                        "日期": _t.get("日期"),
                        "服務時段": f"{_t.get('時間窗_開始')}-{_t.get('時間窗_結束')}",
                        "派單居服員": _item["cg_id"],
                        "候選排名": _item.get("candidate_rank", 1),
                        "人工改選原因": _item.get("manual_reason", ""),
                        "服務歷時(分鐘)": _t.get("服務歷時(分鐘)"),
                        "備註": "臨時新增案件｜直接插入",
                    }
                )
            st.dataframe(pd.DataFrame(_insert_rows), width="stretch", hide_index=True)


    elif _dynamic_event_type == "🚨 居服員臨時請假":
        # ==========================================
        st.subheader("🚨 居服員臨時請假")
        st.caption(
            "指定請假居服員與日期後，系統會找出該日受影響任務，"
            "把這些任務暫時視為未派單，再逐筆重新尋找可直接插入的 Top 3 候選人。"
            "其他既有班表不移動。"
        )

        _effective_result_for_leave = apply_overrides_to_result(df_result, overrides)

        # 日期只取目前任務資料中實際存在的日期
        _leave_date_options = []
        if "日期" in df_tasks.columns:
            _leave_date_series = pd.to_datetime(df_tasks["日期"], errors="coerce").dropna()
            _leave_date_options = sorted({d.date() for d in _leave_date_series})

        if not _leave_date_options:
            st.info("目前班表沒有可用的日期資料，請先執行含日期的排程。")
        else:
            _leave_date = st.selectbox(
                "請假日期",
                _leave_date_options,
                format_func=lambda d: pd.Timestamp(d).strftime("%Y-%m-%d"),
                key="leave_event_date",
            )

            # 找出該日實際有被派班的居服員，讓督導不用從全部員工中找
            _task_date_map = (
                df_tasks[["任務ID", "日期"]]
                .assign(_date_norm=lambda x: pd.to_datetime(x["日期"], errors="coerce").dt.date)
            )
            _result_with_date = _effective_result_for_leave.merge(
                _task_date_map[["任務ID", "_date_norm"]],
                on="任務ID",
                how="left",
            )
            _assigned_on_leave_date = _result_with_date[
                _result_with_date["_date_norm"] == _leave_date
            ].copy()

            _leave_cg_options = sorted(
                {
                    str(v)
                    for v in _assigned_on_leave_date.get("派單居服員", pd.Series(dtype=object)).dropna().tolist()
                    if str(v).strip() and str(v).strip() != "未指派"
                }
            )

            if not _leave_cg_options:
                st.info("這一天目前沒有已指派的居服員。")
            else:
                _leave_cg = st.selectbox(
                    "請假的居服員",
                    _leave_cg_options,
                    key="leave_event_caregiver",
                )

                if st.button(
                    "🔎 找出受影響任務並重新找候選人",
                    type="primary",
                    key="leave_event_analyze",
                ):
                    try:
                        # 找出該居服員在該日目前實際負責的任務
                        _affected_result = _assigned_on_leave_date[
                            _assigned_on_leave_date["派單居服員"].astype(str) == str(_leave_cg)
                        ].copy()
                        _affected_task_ids = _affected_result["任務ID"].astype(str).tolist()

                        if not _affected_task_ids:
                            st.session_state["leave_event_analysis"] = {
                                "date": _leave_date,
                                "absent_cg": _leave_cg,
                                "items": [],
                            }
                        else:
                            # 完整任務資料：任務 + 案家條件
                            _current_tasks_for_leave = df_tasks.merge(
                                res.get("df_cl", pd.DataFrame()),
                                on="案家ID",
                                how="left",
                            )

                            # 請假者當日受影響任務先全部從「既有指派」移除，
                            # 這樣每一筆才能以真正待補班的狀態重新做 insertion。
                            _base_result_without_affected = _effective_result_for_leave[
                                ~_effective_result_for_leave["任務ID"].astype(str).isin(_affected_task_ids)
                            ].copy()

                            # 候選池直接排除請假本人
                            _df_cg_leave = res.get("df_cg", pd.DataFrame()).copy()
                            if "居服員ID" in _df_cg_leave.columns:
                                _df_cg_leave = _df_cg_leave[
                                    _df_cg_leave["居服員ID"].astype(str) != str(_leave_cg)
                                ].copy()

                            _leave_items = []

                            for _task_id in _affected_task_ids:
                                _task_match = _current_tasks_for_leave[
                                    _current_tasks_for_leave["任務ID"].astype(str) == str(_task_id)
                                ]
                                if _task_match.empty:
                                    _leave_items.append({
                                        "task_id": _task_id,
                                        "task": {},
                                        "candidates": pd.DataFrame(),
                                        "error": "找不到原任務資料",
                                    })
                                    continue

                                _task_row = _task_match.iloc[0].copy()

                                try:
                                    _cand = find_insertion_candidates(
                                        _task_row,
                                        _current_tasks_for_leave,
                                        _base_result_without_affected,
                                        _df_cg_leave,
                                        config,
                                        top_n=3,
                                        date_column="日期",
                                    )
                                    _leave_items.append({
                                        "task_id": _task_id,
                                        "task": _task_row.to_dict(),
                                        "candidates": _cand,
                                        "error": "",
                                    })
                                except Exception as _leave_err:
                                    _leave_items.append({
                                        "task_id": _task_id,
                                        "task": _task_row.to_dict(),
                                        "candidates": pd.DataFrame(),
                                        "error": str(_leave_err),
                                    })

                            st.session_state["leave_event_analysis"] = {
                                "date": _leave_date,
                                "absent_cg": _leave_cg,
                                "items": _leave_items,
                            }

                    except Exception as e:
                        st.session_state["leave_event_analysis"] = None
                        st.error(f"無法分析居服員請假事件：{e}")

                _leave_analysis = st.session_state.get("leave_event_analysis")

                if (
                    isinstance(_leave_analysis, dict)
                    and _leave_analysis.get("date") == _leave_date
                    and str(_leave_analysis.get("absent_cg")) == str(_leave_cg)
                ):
                    _leave_items = _leave_analysis.get("items", [])

                    if not _leave_items:
                        st.warning("找不到這位居服員在所選日期的受影響任務。")
                    else:
                        st.success(
                            f"找到 {len(_leave_items)} 筆受影響任務。"
                            "以下每一筆都已獨立重新計算可直接插入的候選人。"
                        )

                        _affected_summary = []
                        for _item in _leave_items:
                            _t = _item.get("task", {})
                            _affected_summary.append({
                                "任務ID": _item.get("task_id"),
                                "案家ID": _t.get("案家ID"),
                                "日期": _t.get("日期"),
                                "服務時段": f"{_t.get('時間窗_開始', '')}-{_t.get('時間窗_結束', '')}",
                                "原派居服員": _leave_cg,
                                "服務歷時(分鐘)": _t.get("服務歷時(分鐘)"),
                            })
                        st.dataframe(
                            pd.DataFrame(_affected_summary),
                            width="stretch",
                            hide_index=True,
                        )

                        for _item_index, _item in enumerate(_leave_items, start=1):
                            _task_id = str(_item.get("task_id"))
                            _task = _item.get("task", {})
                            _cand_df = _item.get("candidates")
                            _err = _item.get("error", "")

                            with st.expander(
                                f"受影響任務 {_item_index}｜{_task_id}｜"
                                f"{_task.get('時間窗_開始', '')}-{_task.get('時間窗_結束', '')}",
                                expanded=True,
                            ):
                                if _err:
                                    st.error(f"候選人計算失敗：{_err}")
                                    continue

                                if not isinstance(_cand_df, pd.DataFrame) or _cand_df.empty:
                                    st.warning("目前沒有候選人結果。")
                                    continue

                                _available = _cand_df[
                                    _cand_df["可直接插入"] == True
                                ].copy()

                                if _available.empty:
                                    st.warning(
                                        "目前沒有可直接接手的人選。"
                                        "此任務之後可進入局部重排流程。"
                                    )
                                    st.dataframe(_cand_df, width="stretch", hide_index=True)
                                    continue

                                _top3 = _available.head(3).copy()
                                st.dataframe(_top3, width="stretch", hide_index=True)

                                st.markdown("**督導選擇接手人員**")
                                _top1 = str(_top3.iloc[0]["居服員ID"])

                                for _rank, (_idx, _cand) in enumerate(_top3.iterrows(), start=1):
                                    _new_cg = str(_cand.get("居服員ID", ""))
                                    _score = _cand.get(
                                        "適配度分數",
                                        _cand.get("適配分數", None),
                                    )
                                    _cost = _cand.get("擾動成本", 0)
                                    _prev_slot = str(_cand.get("前一任務（結束）", "無"))
                                    _next_slot = str(_cand.get("下一任務（開始）", "無"))

                                    _left, _right = st.columns([4, 1])
                                    _label = f"候選 {_rank}｜{_new_cg}"
                                    if pd.notna(_score):
                                        _label += f"｜適配分數 {float(_score):.1f}"
                                    _label += f"｜擾動成本 {_cost}"
                                    _left.markdown(_label)
                                    _left.caption(
                                        f"🕒 前一任務 {_prev_slot} → 本次服務 → 下一任務 {_next_slot}"
                                    )

                                    if _right.button(
                                        "✅ 改派給此人",
                                        key=f"leave_assign_{_leave_date}_{_leave_cg}_{_task_id}_{_rank}_{_new_cg}",
                                        width="stretch",
                                    ):
                                        _reason = f"居服員臨時請假：{_leave_cg}"
                                        _conflict_msg = _quick_reassign(
                                            _task_id,
                                            _new_cg,
                                            _reason,
                                        )

                                        if _conflict_msg:
                                            st.error(_conflict_msg)
                                        else:
                                            _confirmed_at = pd.Timestamp.now()
                                            _top3_ids = " > ".join(
                                                _top3["居服員ID"].astype(str).tolist()
                                            )

                                            st.session_state["dynamic_adjustment_log"].append(
                                                {
                                                    "時間戳記": _confirmed_at,
                                                    "事件類型": "居服員臨時請假",
                                                    "任務ID": _task_id,
                                                    "案家ID": _task.get("案家ID"),
                                                    "日期": _task.get("日期"),
                                                    "服務時段": (
                                                        f"{_task.get('時間窗_開始', '')}-"
                                                        f"{_task.get('時間窗_結束', '')}"
                                                    ),
                                                    "請假居服員": _leave_cg,
                                                    "系統Top3": _top3_ids,
                                                    "系統第一名": _top1,
                                                    "最終選擇居服員": _new_cg,
                                                    "最終候選排名": _rank,
                                                    "是否人工改選": "是" if _rank > 1 else "否",
                                                    "既有班表是否移動": "否",
                                                    "擾動成本": _cand.get("擾動成本", 0),
                                                    "事件備註": "請假任務逐筆重新插入",
                                                }
                                            )

                                            st.success(
                                                f"✅ {_task_id} 已由 {_leave_cg} 改派給 {_new_cg}。"
                                            )
                                            # 改派後原候選結果可能已過期，要求重新分析剩餘任務。
                                            st.session_state["leave_event_analysis"] = None
                                            st.rerun()

        # ==========================================

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

        _reschedule_options = []
        if not df_tasks.empty and "任務ID" in df_tasks.columns:
            for _, _r in df_tasks.iterrows():
                _tid = str(_r.get("任務ID", ""))
                if not _tid:
                    continue
                _client = str(_r.get("案家ID", ""))
                _date_val = pd.to_datetime(_r.get("日期"), errors="coerce")
                _date_txt = (
                    _date_val.strftime("%Y-%m-%d")
                    if pd.notna(_date_val)
                    else ""
                )
                _start_txt = str(_r.get("時間窗_開始", ""))
                _end_txt = str(_r.get("時間窗_結束", ""))
                _label = f"{_tid}｜{_client}｜{_date_txt}｜{_start_txt}-{_end_txt}"
                _reschedule_options.append((_tid, _label))

        if not _reschedule_options:
            st.info("目前沒有可供改期的既有任務。")
        else:
            _reschedule_label_map = dict(_reschedule_options)
            _reschedule_task_id = st.selectbox(
                "選擇要改期的任務",
                [x[0] for x in _reschedule_options],
                format_func=lambda tid: _reschedule_label_map.get(tid, tid),
                key="reschedule_task_id",
            )

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
                )

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
                            _new_task["日期"] = pd.Timestamp(_new_date)
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
                                _tasks_changed.loc[_mask, "日期"] = pd.Timestamp(_new_date)
                            if "時間窗_開始" in _tasks_changed.columns:
                                _tasks_changed.loc[_mask, "時間窗_開始"] = _new_start.strftime("%H:%M")
                            if "時間窗_結束" in _tasks_changed.columns:
                                _tasks_changed.loc[_mask, "時間窗_結束"] = _new_end.strftime("%H:%M")

                            _tasks_changed_merged = _tasks_changed.merge(
                                res.get("df_cl", pd.DataFrame()),
                                on="案家ID",
                                how="left",
                            )

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
                            st.dataframe(
                                _candidate_df,
                                width="stretch",
                                hide_index=True,
                            )
                        else:
                            _top3 = _available.head(3).copy()
                            st.success(
                                f"新時段找到 {len(_top3)} 位可直接插入候選人。"
                            )
                            st.dataframe(
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
                                _label = f"候選 {_rank}｜{_new_cg}"
                                if pd.notna(_score):
                                    _label += f"｜適配分數 {float(_score):.1f}"
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
                                        "案家臨時改期："
                                        f"{_orig_task_info.get('日期')} "
                                        f"{_orig_task_info.get('時間窗_開始')}-"
                                        f"{_orig_task_info.get('時間窗_結束')} → "
                                        f"{_new_task_info.get('日期')} "
                                        f"{_new_task_info.get('時間窗_開始')}-"
                                        f"{_new_task_info.get('時間窗_結束')}"
                                    )

                                    _conflict_msg = _quick_reassign(
                                        _reschedule_task_id,
                                        _new_cg,
                                        _reason,
                                    )

                                    if _conflict_msg:
                                        st.error(_conflict_msg)
                                    else:
                                        # 真正更新目前 session 中的任務日期／時間
                                        _main_mask = (
                                            st.session_state["last_result"]["df_tasks"]["任務ID"]
                                            .astype(str)
                                            == str(_reschedule_task_id)
                                        )
                                        if "日期" in st.session_state["last_result"]["df_tasks"].columns:
                                            st.session_state["last_result"]["df_tasks"].loc[
                                                _main_mask, "日期"
                                            ] = pd.Timestamp(_new_task_info.get("日期"))
                                        if "時間窗_開始" in st.session_state["last_result"]["df_tasks"].columns:
                                            st.session_state["last_result"]["df_tasks"].loc[
                                                _main_mask, "時間窗_開始"
                                            ] = _new_task_info.get("時間窗_開始")
                                        if "時間窗_結束" in st.session_state["last_result"]["df_tasks"].columns:
                                            st.session_state["last_result"]["df_tasks"].loc[
                                                _main_mask, "時間窗_結束"
                                            ] = _new_task_info.get("時間窗_結束")

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
            _tasks_merged = _tasks_base.merge(
                _clients,
                on="案家ID",
                how="left",
            )

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
                if _transport not in ("機車", "大眾運輸", "汽車"):
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

            for _a in _affected:
                try:
                    _cand = find_insertion_candidates(
                        _a["row"],
                        _tasks_merged,
                        _base_result,
                        _candidate_cg,
                        config,
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
        _delay_options = []
        if not _effective_result_for_delay.empty:
            _delay_join = _effective_result_for_delay.merge(
                df_tasks[
                    [c for c in ["任務ID", "案家ID", "日期", "時間窗_開始", "時間窗_結束"] if c in df_tasks.columns]
                ],
                on="任務ID",
                how="left",
            )
            for _, _r in _delay_join.iterrows():
                _tid = str(_r.get("任務ID", ""))
                _cg = str(_r.get("派單居服員", ""))
                if not _tid or not _cg or _cg == "未指派":
                    continue
                _date = pd.to_datetime(_r.get("日期"), errors="coerce")
                _date_txt = _date.strftime("%Y-%m-%d") if pd.notna(_date) else ""
                _delay_options.append(
                    (
                        _tid,
                        f"{_tid}｜{_cg}｜{_date_txt}｜"
                        f"{_r.get('時間窗_開始', '')}-{_r.get('時間窗_結束', '')}",
                    )
                )

        if not _delay_options:
            st.info("目前沒有可分析的已派任務。")
        else:
            _delay_label_map = dict(_delay_options)
            _delay_task_id = st.selectbox(
                "哪一筆服務發生延遲？",
                [x[0] for x in _delay_options],
                format_func=lambda tid: _delay_label_map.get(tid, tid),
                key="delay_event_task_id",
            )

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
                        st.dataframe(
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
                                    st.dataframe(
                                        _cand_df,
                                        width="stretch",
                                        hide_index=True,
                                    )
                                    continue

                                _top3 = _available.head(3).copy()
                                st.dataframe(
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
                                    _label = f"候選 {_rank}｜{_new_cg}"
                                    if pd.notna(_score):
                                        _label += f"｜適配分數 {float(_score):.1f}"
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
                                            f"前一服務延遲：{_delay_ctx['task_id']} "
                                            f"延至 {pd.to_datetime(_delay_ctx['expected_end']).strftime('%H:%M')}；"
                                            f"{_task_id} 無法維持原銜接"
                                        )

                                        _conflict_msg = _quick_reassign(
                                            _task_id,
                                            _new_cg,
                                            _reason,
                                        )

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


    if st.session_state.get("dynamic_adjustment_log"):
        with st.expander("🧾 Dynamic Adjustment Log", expanded=False):
            _dynamic_log_df = pd.DataFrame(st.session_state["dynamic_adjustment_log"])
            st.dataframe(_dynamic_log_df, width="stretch", hide_index=True)
            st.caption(
                "稽核紀錄供動態調整追蹤使用；匯出後請依機構資料治理規範保存。"
            )
            _dynamic_log_csv = _dynamic_log_df.to_csv(index=False).encode("utf-8-sig")
            st.download_button(
                "⬇️ 匯出 Dynamic_Adjustment_Log.csv",
                _dynamic_log_csv,
                file_name="Dynamic_Adjustment_Log.csv",
                mime="text/csv",
                key="download_dynamic_adjustment_log",
            )

    # 區塊⑤：居督人工覆寫（Supervisor Override）與稽核日誌

    # ==========================================

    st.header("⑤ 居督人工覆寫（Supervisor Override）")

    st.caption(

        "當 AI 建議排單不符合實際場域狀況時（例如居服員臨時請假、案家臨時改期），"

        "居督可搜尋／選擇任務後於彈出視窗手動重新指定居服員（與月曆視角「一鍵調班」"

        "共用同一套覆寫互動介面）；每一筆變更皆會記錄原因並寫入稽核日誌，供後續演算法迭代分析。"

    )



    # overrides 已於本區塊之前（月曆總覽渲染前）初始化，此處沿用同一份 session_state。



    OVERRIDE_TRAVEL_ALERT_THRESHOLD = 3



    override_log_df = load_override_log()

    travel_reason_count = (

        (override_log_df["變更原因"] == "車程／交通因素").sum() if not override_log_df.empty else 0

    )

    if travel_reason_count >= OVERRIDE_TRAVEL_ALERT_THRESHOLD:

        st.warning(

            f"📈 稽核日誌累計已有 {travel_reason_count} 筆覆寫原因為「車程／交通因素」，"

            "建議提高側邊欄的『車程扣分權重』，讓 AI 派單更優先考量就近指派。"

        )



    # assigned_map 已於本區塊之前初始化並供月曆快速改派共用，此處沿用同一份。

    render_task_override_picker(

        df_tasks, df_result, res.get("df_cg", pd.DataFrame()), overrides,

        on_reassign=_quick_reassign, on_clear_override=_clear_override, on_list_candidates=_list_candidates,

    )



    st.subheader("📄 最終派單結果（含居督覆寫）")

    has_date_for_final = "日期" in df_tasks.columns

    has_weekday_col_for_final = "星期" in df_tasks.columns

    final_rows = []

    for _, task_row in df_tasks.iterrows():

        t_id = task_row["任務ID"]

        cl_id = task_row["案家ID"]

        base = df_result[df_result["任務ID"] == t_id] if not df_result.empty else pd.DataFrame()

        ai_row = base.iloc[0] if not base.empty else None

        ai_cg = ai_row["派單居服員"] if ai_row is not None else None



        override = overrides.get(t_id)

        if override:

            final_cg = override["cg_id"]

            change_note = f"居督覆寫：{override['reason']}"

        else:

            final_cg = ai_cg

            change_note = ai_row["原首選替換原因"] if ai_row is not None else ""



        row_dict = {"任務ID": t_id, "案家ID": cl_id}

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

            "服務時段": ai_row["服務時段"] if ai_row is not None else "",

            "任務優先級": ai_row["任務優先級"] if ai_row is not None else task_row.get("任務優先級", ""),

            "預估長照申報點數(營收)": ai_row["預估長照申報點數(營收)"] if ai_row is not None else None,

            "預估居服員拆帳薪資": ai_row["預估居服員拆帳薪資"] if ai_row is not None else None,

            "備註": change_note,

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
                "預估長照申報點數(營收)": None,
                "預估居服員拆帳薪資": None,
                "備註": "臨時新增案件｜直接插入（既有班表未移動）",
            }
        )
        final_rows.append(_dynamic_row)

    df_final = pd.DataFrame(final_rows)

    st.dataframe(df_final, width="stretch", hide_index=True)



    csv_final = df_final.to_csv(index=False).encode("utf-8-sig")

    st.download_button(

        "⬇️ 匯出最終派單結果（含覆寫） (final_assignment_with_overrides.csv)",

        csv_final,

        file_name="final_assignment_with_overrides.csv",

        mime="text/csv",

    )



    with st.expander("🗂️ 檢視完整稽核日誌 (supervisor_override_log.csv)"):

        if override_log_df.empty:

            st.info("尚無居督覆寫紀錄。")

        else:

            st.dataframe(override_log_df, width="stretch", hide_index=True)

else:

    st.info("請先於上方確認／編輯資料，再按下「🚀 執行 AI 最佳化派單」開始運算。")
