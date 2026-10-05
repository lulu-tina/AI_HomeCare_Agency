"""Presentation-only helpers. No scheduling or external requests."""
import streamlit as st

CSS = """
<style>
:root { --cf-ink:#183d3d; --cf-teal:#147d73; --cf-muted:#627777; }
.stApp { background:#f5f8f7; }
[data-testid="stAppViewContainer"] .main .block-container,
[data-testid="stMainBlockContainer"] { max-width:1440px; padding-top:2rem; padding-bottom:4rem; }
[data-testid="stSidebar"] { background:#edf4f1; border-right:1px solid #dce8e3; }
h1,h2,h3 { color:var(--cf-ink); letter-spacing:-.025em; }
[data-testid="stCaptionContainer"] { color:var(--cf-muted); }
.cf-hero { background:linear-gradient(120deg,#123f3c,#1a685d); border-radius:22px; padding:32px 36px; color:#fff; margin-bottom:18px; }
.cf-eyebrow { font-size:12px; font-weight:700; letter-spacing:.18em; color:#b5dcd0; }
.cf-hero h1 { color:white; font-size:34px; font-weight:750; margin:9px 0; padding:0; }
.cf-hero p { color:#d7ebe4; margin:0; font-size:15px; line-height:1.8; }
.cf-steps { display:flex; gap:12px; flex-wrap:wrap; margin-top:24px; }
.cf-steps span { background:#ffffff12; border:1px solid #ffffff29; border-radius:10px; padding:10px 16px; font-size:14px; }
.cf-section { display:flex; align-items:center; gap:14px; margin:28px 0 14px; scroll-margin-top:80px; }
.cf-number { background:#dfefE8; color:#147d73; border-radius:12px; padding:10px 13px; font-size:15px; font-weight:750; }
.cf-section h2 { font-size:23px; padding:0; margin:0 0 4px; }
.cf-section p { font-size:14px; color:#627777; margin:0; }
.cf-nav { display:grid; gap:6px; }
.cf-nav a { padding:11px 14px; border-radius:10px; color:#28574e !important; text-decoration:none; font-size:14px; font-weight:600; }
.cf-nav a:hover { background:#dcece4; }
.cf-empty { border:1px dashed #b9d4c8; background:white; border-radius:18px; text-align:center; padding:35px 20px; margin:12px 0; }
.cf-empty span { font-size:30px; }.cf-empty h3 { font-size:21px; padding:12px 0 8px; }.cf-empty p { color:#627777; font-size:14px; }
[data-testid="stFileUploader"] { background:white; padding:18px; border:1px solid #dce8e3; border-radius:16px; }
[data-testid="stFileUploaderDropzone"] { background:#f4f9f6; border-radius:12px; }
[data-testid="stMetric"] { background:white; border:1px solid #dce8e3; border-radius:16px; padding:18px 20px; }
[data-testid="stMetricValue"] { color:#147d73; font-weight:700; }
[data-testid="stExpander"] { background:white; border:1px solid #dce8e3; border-radius:14px; }
[data-testid="stExpander"] summary { padding:14px 16px; }
.stButton button,.stDownloadButton button { border-radius:10px; min-height:44px; font-weight:600; }
.stButton button[kind="primary"] { background:#147d73; border-color:#147d73; box-shadow:0 4px 12px #147d7320; }
.stButton button[kind="primary"]:hover { background:#0e665d; border-color:#0e665d; }
[data-testid="stTabs"] [role="tablist"] { gap:10px; border-bottom:1px solid #dce8e3; }
[data-testid="stTabs"] [role="tab"] { border-radius:8px 8px 0 0; padding:10px 18px; }
[data-testid="stTabs"] [aria-selected="true"] { color:#147d73; background:#edf6f1; }
[data-testid="stAlert"] { border-radius:12px; }
@media(max-width:640px) { .cf-hero { padding:24px 20px; }.cf-hero h1 {font-size:27px;}.cf-steps {gap:6px;}.cf-steps span {font-size:12px; padding:8px 10px;}.cf-section h2 {font-size:20px;} }
</style>
"""

def render_brand():
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown('''<div class="cf-hero"><div class="cf-eyebrow">CAREFLOW · 居服排班工作台</div><h1>把排班變得更簡單。</h1><p>整合服務資料，產生排班草案，讓每一次照顧安排更清楚。</p><div class="cf-steps"><span>01　匯入 Excel</span><span>02　確認資料</span><span>03　產生班表</span></div></div>''', unsafe_allow_html=True)
    st.caption('🔒 排班使用 ID，姓名以遮罩方式顯示；可在「個資顯示設定」調整。')

def section_title(anchor, number, title, description):
    from html import escape
    st.markdown(f'<div id="{escape(anchor)}" class="cf-section"><span class="cf-number">{escape(number)}</span><div><h2>{escape(title)}</h2><p>{escape(description)}</p></div></div>', unsafe_allow_html=True)

def render_summary(caregivers, clients, tasks):
    cols = st.columns(3)
    for col, label, value in zip(cols, ['居服員', '服務個案', '待排服務'], [caregivers, clients, tasks]):
        col.metric(label, f'{value:,}')
