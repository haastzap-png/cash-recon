import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from pathlib import Path
import re
import html
from datetime import timedelta
from io import BytesIO

def _theme():
    return {
        "config": {
            "background": "rgba(0,0,0,0)",
            "view": {"stroke": "transparent"},
            "axis": {
                "labelColor": "#3b3b35",
                "titleColor": "#3b3b35",
                "gridColor": "#e6dfd3",
                "labelFont": "Noto Sans TC",
                "titleFont": "Noto Sans TC",
            },
            "title": {
                "color": "#1d1b16",
                "font": "Noto Serif TC",
                "fontSize": 18,
            },
            "range": {
                "category": [
                    "#2b7a78",
                    "#f2b134",
                    "#8f5e3b",
                    "#9bc1bc",
                    "#d95d39",
                    "#5f6c37",
                    "#7a4e2d",
                ]
            },
        }
    }

alt.themes.register("custom_theme", _theme)
alt.themes.enable("custom_theme")

st.set_page_config(page_title="顧客關係經營分析", layout="wide")

st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@600;700&family=Noto+Sans+TC:wght@400;500;600&display=swap');
:root {
  --ink: #1d1b16;
  --muted: #6b6b63;
  --accent: #2b7a78;
  --accent-2: #f2b134;
  --card: #ffffff;
  --line: #e6dfd3;
  --shadow: 0 6px 18px rgba(33, 28, 20, 0.08);
}
html, body, [class*="css"] {
  font-family: "Noto Sans TC", "PingFang TC", "Microsoft JhengHei", sans-serif;
}
.stApp {
  background:
    radial-gradient(1200px 600px at 10% -10%, #f6efe3 0%, transparent 60%),
    radial-gradient(1000px 500px at 110% 10%, #eaf3f2 0%, transparent 55%),
    linear-gradient(180deg, #f7f4ee 0%, #f3f0ea 35%, #f9f7f3 100%);
}
h1, h2, h3, h4 {
  font-family: "Noto Serif TC", "Noto Sans TC", serif;
  letter-spacing: 0.4px;
}
h1 {
  font-size: 2.2rem;
  color: var(--ink);
}
div[data-testid="stAppViewContainer"] section.main,
div[data-testid="stAppViewContainer"] .main {
  color: var(--ink) !important;
  --text-color: var(--ink);
  --secondary-text-color: var(--muted);
}
div[data-testid="stAppViewContainer"] .main h1,
div[data-testid="stAppViewContainer"] .main h2,
div[data-testid="stAppViewContainer"] .main h3,
div[data-testid="stAppViewContainer"] .main h4,
div[data-testid="stAppViewContainer"] .main h5,
div[data-testid="stAppViewContainer"] .main h6,
div[data-testid="stAppViewContainer"] .main p,
div[data-testid="stAppViewContainer"] .main li,
div[data-testid="stAppViewContainer"] .main label {
  color: var(--ink) !important;
  -webkit-text-fill-color: var(--ink) !important;
  opacity: 1 !important;
}
div[data-testid="stAppViewContainer"] .main div[data-testid="stMarkdownContainer"] p,
div[data-testid="stAppViewContainer"] .main div[data-testid="stMarkdownContainer"] li,
div[data-testid="stAppViewContainer"] .main div[data-testid="stMarkdownContainer"] span {
  color: var(--ink) !important;
  -webkit-text-fill-color: var(--ink) !important;
  opacity: 1 !important;
}
div[data-testid="stAppViewContainer"] .main .stCaption,
div[data-testid="stAppViewContainer"] .main .stCaption *,
div[data-testid="stAppViewContainer"] .main div[data-testid="stCaptionContainer"] * {
  color: var(--muted) !important;
  -webkit-text-fill-color: var(--muted) !important;
  opacity: 1 !important;
}
div[data-testid="stAppViewContainer"] .main div[data-testid="stMetricValue"] {
  color: var(--ink) !important;
  -webkit-text-fill-color: var(--ink) !important;
  opacity: 1 !important;
}
div[data-testid="stAppViewContainer"] .main div[data-testid="stMetricLabel"],
div[data-testid="stAppViewContainer"] .main div[data-testid="stMetricDelta"] {
  color: var(--muted) !important;
  -webkit-text-fill-color: var(--muted) !important;
  opacity: 1 !important;
}
div[data-testid="stAppViewContainer"] .main button,
div[data-testid="stAppViewContainer"] .main button * {
  color: inherit !important;
  -webkit-text-fill-color: inherit !important;
}
div[data-testid="stSidebar"] {
  background: #faf7f2;
  border-right: 1px solid var(--line);
}
div[data-testid="stMetric"] {
  background: var(--card);
  border: 1px solid var(--line);
  border-radius: 16px;
  padding: 10px 12px;
  box-shadow: var(--shadow);
}
div[data-testid="stMetric"] label {
  color: var(--muted);
  font-weight: 600;
  opacity: 1 !important;
}
.section-note {
  padding: 8px 12px;
  background: rgba(43, 122, 120, 0.08);
  border: 1px solid rgba(43, 122, 120, 0.2);
  border-radius: 12px;
  color: #1f314f !important;
  font-size: 0.95rem;
}
.stCaption {
  color: var(--muted);
}
.stDataFrame, .stTable {
  border: 1px solid var(--line);
  border-radius: 12px;
}
.metric-card {
  position: relative;
  background: var(--card);
  border: 1px solid var(--line);
  border-radius: 16px;
  padding: 14px 16px;
  box-shadow: var(--shadow);
  min-height: 96px;
  margin-bottom: 8px;
}
.metric-card.metric-flat {
  background: transparent;
  border: 0;
  box-shadow: none;
  margin-bottom: 0;
  padding: 8px 6px 4px 6px;
  min-height: 84px;
}
.metric-card.metric-horizontal {
  display: grid;
  grid-template-columns: 160px 1fr auto auto;
  align-items: center;
  gap: 12px;
  min-height: 110px;
}
.metric-card.metric-horizontal .metric-title {
  margin-bottom: 0;
  min-width: 140px;
}
.metric-card.metric-horizontal .metric-value {
  font-size: 1.8rem;
  justify-self: center;
}
.metric-card.metric-horizontal .metric-sub {
  margin-top: 0;
  display: inline-flex;
  align-items: center;
  gap: 8px;
  justify-self: end;
}
.metric-title {
  font-size: 0.95rem;
  color: var(--muted);
  font-weight: 600;
  margin-bottom: 6px;
}
.metric-value {
  font-size: 2rem;
  color: var(--ink);
  font-weight: 700;
  line-height: 1.1;
}
.metric-suffix {
  font-size: 0.95rem;
  font-weight: 700;
  color: var(--muted);
  margin-left: 10px;
  vertical-align: middle;
}
.metric-meta {
  margin-top: 4px;
  font-size: 0.85rem;
  color: var(--muted);
  font-weight: 600;
}
.metric-sub {
  margin-top: 8px;
  font-size: 0.85rem;
  color: var(--muted);
}
.metric-tag {
  display: inline-block;
  padding: 2px 8px;
  border-radius: 999px;
  font-size: 0.8rem;
  font-weight: 600;
  margin-left: 6px;
  min-width: 56px;
  text-align: center;
  justify-content: center;
  display: inline-flex;
  align-items: center;
}
.metric-help {
  position: absolute;
  top: 10px;
  right: 10px;
  width: 22px;
  height: 22px;
  border-radius: 50%;
  border: 1px solid #c7c1b7;
  color: #6b6b63;
  background: #ffffff;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 12px;
  cursor: help;
}
.metric-help::after {
  content: attr(data-tooltip);
  position: absolute;
  top: 28px;
  right: 0;
  width: 240px;
  padding: 8px 10px;
  background: #1d1b16;
  color: #ffffff;
  border-radius: 8px;
  font-size: 12px;
  line-height: 1.4;
  opacity: 0;
  pointer-events: none;
  transform: translateY(-4px);
  transition: opacity 0.15s ease, transform 0.15s ease;
  z-index: 10;
}
.metric-help:hover::after {
  opacity: 1;
  transform: translateY(0);
}
.section-gap {
  height: 18px;
}
div[data-testid="stVerticalBlockBorderWrapper"] {
  border: 1px solid #ddd3c4;
  border-radius: 16px;
  background: linear-gradient(180deg, #fffdf9 0%, #f7f2e9 100%);
  box-shadow: 0 10px 24px rgba(33, 28, 20, 0.10);
  padding: 4px 8px 8px 8px;
  margin-bottom: 10px;
}
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stExpander"] {
  margin-top: 2px;
  margin-bottom: 0;
}
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stExpander"] details {
  border: 0;
  border-top: 1px dashed #d8ccb9;
  border-radius: 0 0 12px 12px;
  background: rgba(255, 255, 255, 0.58);
  box-shadow: none;
  overflow: hidden;
}
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stExpander"] details > summary {
  padding: 6px 10px;
  font-size: 0.86rem;
  color: var(--muted);
}
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stExpander"] details[open] > summary {
  border-bottom: 0;
  color: var(--ink);
}
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stExpander"] details > summary p {
  margin: 0;
  line-height: 1.2;
}
div[data-testid="stVerticalBlockBorderWrapper"] div[data-testid="stExpander"] details > div {
  padding-top: 6px;
}
</style>
""",
    unsafe_allow_html=True,
)

st.title("顧客關係經營分析")
st.markdown('<div class="section-note">圖表優先，表格放在最後。指標固定：出勤狀態、新客流失、熟客化、熟客維持、空窗率、業績穩定度。</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("輸入")
    member_file = st.file_uploader("上傳 會員名單.xlsx（選填）", type=["xlsx"], key="member")
    bill_files = st.file_uploader("上傳 帳單紀錄.xlsx（可多檔）", type=["xlsx"], key="bill", accept_multiple_files=True)
    st.divider()
    include_types = st.multiselect(
        "計算包含結帳類型",
        ["服務", "票券", "商品", "儲值金"],
        default=["服務", "票券"],
    )
    chart_top_n = st.number_input("圖表顯示前 N 名（0=全部）", min_value=0, max_value=100, value=0, step=1)
    min_repeat_base = st.number_input("回指率最低樣本數（低於則不顯示）", min_value=1, max_value=100, value=5, step=1)
    store_chart_type = st.selectbox("分店比較圖表", ["群組直條圖", "熱度圖", "堆疊條圖"])

st.write("""
本工具會：
- 以全品牌帳單歷史找出新客，並計算 60 天內是否回店（同分店）
- 依分店/師傅呈現出勤狀態（每月平均有單天數、近 3 月有單月份數）、新客流失、熟客化與熟客維持、空窗率與業績穩定度
- 提供圖表與排行榜，快速看出差異
熟客定義：同分店同師傅，180 天內消費 ≥5 次
熟客維持：熟客達成後 180 天內回訪 ≥3 次
回指判定：T2=30 天、T3=60 天（固定，作為輔助指標）
空窗率計算：依項目分鐘估算時長（1～30=0.5；31～60=1；61～90=1.5，以此類推），月上限 168 小時
""")

if not bill_files:
    st.info("請先上傳帳單檔。")
    st.stop()

CHURN_DAYS = 60
T2_DAYS = 30
T3_DAYS = 60
REGULAR_DAYS = 180
REGULAR_VISITS = 5
RETENTION_DAYS = 180
RETENTION_VISITS = 3
RELATIONSHIP_COHORT_MONTHS = 12

@st.cache_data(show_spinner=False)
def load_member(file):
    return pd.read_excel(file, sheet_name="會員名單")

@st.cache_data(show_spinner=False)
def load_bill(file, sheets):
    xls = pd.ExcelFile(file)
    available = [s for s in sheets if s in xls.sheet_names]
    frames = []
    for s in available:
        frames.append(pd.read_excel(file, sheet_name=s))
    if not frames:
        return pd.DataFrame(), available
    return pd.concat(frames, ignore_index=True), available

def infer_store_name(filename):
    stem = Path(filename).stem
    name = stem
    name = re.sub(r"帳單紀錄", "", name)
    name = re.sub(r"帳單|紀錄", "", name)
    name = re.sub(r"\d{4}-\d{2}-\d{2}", "", name)
    name = re.sub(r"\d{8,}", "", name)
    name = re.sub(r"[_\-]+", " ", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name if name else stem

def load_bills(files, sheets):
    frames = []
    used = set()
    for f in files:
        df, used_sheets = load_bill(f, sheets)
        if not df.empty:
            store_name = infer_store_name(getattr(f, "name", "未命名分店"))
            df["來源檔案"] = getattr(f, "name", "未命名檔案")
            if "分店" not in df.columns:
                df["分店"] = store_name
            else:
                df["分店"] = df["分店"].fillna(store_name)
            frames.append(df)
        used.update(used_sheets)
    if not frames:
        return pd.DataFrame(), list(used)
    return pd.concat(frames, ignore_index=True), list(used)

bills, used_sheets = load_bills(bill_files, include_types)

if bills.empty:
    st.error("帳單檔中找不到選擇的工作表。")
    st.stop()

# Normalize phone numbers

def norm_digits(x):
    if pd.isna(x):
        return None
    if isinstance(x, (int, np.integer)):
        s = str(int(x))
    elif isinstance(x, (float, np.floating)):
        if np.isnan(x):
            return None
        s = str(int(x)) if float(x).is_integer() else str(x)
    else:
        s = str(x)
    digits = re.sub(r"\D+", "", s)
    return digits or None

def norm_yes_no(x):
    if pd.isna(x):
        return None
    s = str(x).strip().upper()
    if s in ("Y", "YES", "TRUE", "1", "指定", "是"):
        return True
    if s in ("N", "NO", "FALSE", "0", "非指定", "否"):
        return False
    return None

def mask_last3(x):
    if pd.isna(x):
        return ""
    digits = re.sub(r"\D+", "", str(x))
    if not digits:
        return ""
    return digits[-3:] if len(digits) >= 3 else digits

def extract_minutes(item_text):
    if pd.isna(item_text):
        return 0
    mins = re.findall(r"(\d+)\s*分鐘", str(item_text))
    if not mins:
        return 0
    return sum(int(m) for m in mins)

def minutes_to_hours(mins):
    if mins <= 0:
        return 0.0
    # 1~30=0.5, 31~60=1, 61~90=1.5, ...
    units = int((mins - 1) // 30 + 1)
    return units * 0.5

def render_bar_chart(df, category_col, value_col, title, color="#2b7a78", top_n=0, value_format="percent", orient="vertical", ascending=False):
    if df.empty:
        st.info("沒有可顯示的資料。")
        return 0, 0
    chart_df = df[[category_col, value_col]].dropna().copy()
    chart_df = chart_df.sort_values(value_col, ascending=ascending)
    if top_n and len(chart_df) > top_n:
        chart_df = chart_df.head(int(top_n))
    chart_df[value_col] = chart_df[value_col].astype(float)
    height = 320 if orient == "vertical" else max(320, 26 * len(chart_df) + 40)
    if value_format == "percent":
        axis_format = "%"
        label_format = ".1%"
        tooltip = [category_col, alt.Tooltip(value_col, format=".2%")]
    elif value_format == "number1":
        axis_format = ",.1f"
        label_format = ".1f"
        tooltip = [category_col, alt.Tooltip(value_col, format=",.1f")]
    else:
        axis_format = ",.0f"
        label_format = ".0f"
        tooltip = [category_col, alt.Tooltip(value_col, format=",.0f")]

    if orient == "vertical":
        base = alt.Chart(chart_df).encode(
            x=alt.X(
                f"{category_col}:N",
                sort=alt.SortField(field=value_col, order="ascending" if ascending else "descending"),
                title="",
                axis=alt.Axis(labelAngle=-30),
            ),
            y=alt.Y(f"{value_col}:Q", axis=alt.Axis(format=axis_format, title=title)),
            tooltip=tooltip,
        )
        bars = base.mark_bar(color=color)
        labels = base.mark_text(align="center", dy=-6, color="#333").encode(
            text=alt.Text(f"{value_col}:Q", format=label_format)
        )
    else:
        base = alt.Chart(chart_df).encode(
            y=alt.Y(
                f"{category_col}:N",
                sort=alt.SortField(field=value_col, order="ascending" if ascending else "descending"),
                title="",
                axis=alt.Axis(labelLimit=0, labelOverlap=False, labelPadding=6, labelFontSize=12),
            ),
            x=alt.X(f"{value_col}:Q", axis=alt.Axis(format=axis_format, title=title)),
            tooltip=tooltip,
        )
        bars = base.mark_bar(color=color)
        labels = base.mark_text(align="left", dx=4, color="#333").encode(
            text=alt.Text(f"{value_col}:Q", format=label_format)
        )
    st.altair_chart((bars + labels).properties(height=height), use_container_width=True)
    return len(chart_df), len(df)

def render_rank_bar(df, name_col, value_col, title, ascending, value_format, color, top_n=6):
    if df.empty:
        st.info("沒有可顯示的資料。")
        return
    chart_df = df[[name_col, value_col]].dropna().copy()
    chart_df = chart_df.sort_values(value_col, ascending=ascending).head(top_n)
    chart_df[value_col] = chart_df[value_col].astype(float)

    if value_format == "percent":
        axis_format = "%"
        label_format = ".1%"
        tooltip = [name_col, alt.Tooltip(value_col, format=".2%")]
    elif value_format == "number1":
        axis_format = ",.1f"
        label_format = ".1f"
        tooltip = [name_col, alt.Tooltip(value_col, format=",.1f")]
    else:
        axis_format = ",.0f"
        label_format = ".0f"
        tooltip = [name_col, alt.Tooltip(value_col, format=",.0f")]

    base = alt.Chart(chart_df).encode(
        y=alt.Y(
            f"{name_col}:N",
            sort=alt.SortField(field=value_col, order="ascending" if ascending else "descending"),
            title="",
            axis=alt.Axis(labelLimit=0, labelOverlap=False, labelPadding=6, labelFontSize=12),
        ),
        x=alt.X(f"{value_col}:Q", axis=alt.Axis(format=axis_format, title=title)),
        tooltip=tooltip,
    )
    bars = base.mark_bar(color=color)
    labels = base.mark_text(align="left", dx=4, color="#333").encode(
        text=alt.Text(f"{value_col}:Q", format=label_format)
    )
    height = max(260, 26 * len(chart_df) + 40)
    st.altair_chart((bars + labels).properties(height=height), use_container_width=True)

def metric_card(label, value, help_text, subtext=None, tag_text=None, tag_bg=None, tag_color=None, value_suffix=None, meta_text=None, horizontal=False, flat=False):
    safe_label = html.escape(str(label))
    safe_value = html.escape(str(value))
    suffix_html = ""
    if value_suffix:
        safe_suffix = html.escape(str(value_suffix))
        suffix_html = f'<span class="metric-suffix">{safe_suffix}</span>'
    meta_html = ""
    if meta_text:
        safe_meta = html.escape(str(meta_text))
        meta_html = f'<div class="metric-meta">{safe_meta}</div>'
    safe_help = html.escape(str(help_text), quote=True)
    sub_html = ""
    if subtext is not None:
        safe_sub = html.escape(str(subtext))
        tag_html = ""
        if tag_text:
            safe_tag = html.escape(str(tag_text))
            style = ""
            if tag_bg or tag_color:
                style = f' style="background:{tag_bg or "#eee"};color:{tag_color or "#333"};"'
            tag_html = f'<span class="metric-tag"{style}>{safe_tag}</span>'
        sub_html = f'<div class="metric-sub">{safe_sub}{tag_html}</div>'
    extra_cls = ""
    if horizontal:
        extra_cls += " metric-horizontal"
    if flat:
        extra_cls += " metric-flat"
    st.markdown(
        f"""
        <div class="metric-card{extra_cls}">
          <div class="metric-title">{safe_label}</div>
          <div class="metric-value">{safe_value}{suffix_html}</div>
          {meta_html}
          {sub_html}
          <div class="metric-help" data-tooltip="{safe_help}">!</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

def section_gap():
    st.markdown('<div class="section-gap"></div>', unsafe_allow_html=True)


for df, col in [(bills, "國碼"), (bills, "電話號碼")]:
    if col in df.columns:
        bills[col] = bills[col].apply(norm_digits)

if "國碼" not in bills.columns or "電話號碼" not in bills.columns:
    st.error("帳單檔缺少 '國碼' 或 '電話號碼' 欄位。")
    st.stop()

bills["phone_key"] = bills["國碼"].fillna("") + "-" + bills["電話號碼"].fillna("")

valid_bills = bills[bills["phone_key"].str.contains("-") & (bills["phone_key"] != "-")].copy()

valid_bills["結帳操作時間"] = pd.to_datetime(valid_bills["結帳操作時間"], errors="coerce")
if "指定" in valid_bills.columns:
    valid_bills["is_requested"] = valid_bills["指定"].apply(norm_yes_no)
else:
    valid_bills["is_requested"] = pd.NA

# Merge member data (optional)
merged = valid_bills.copy()
if member_file:
    members = load_member(member_file)
    for df, col in [(members, "國碼"), (members, "手機號碼")]:
        if col in df.columns:
            df[col] = df[col].apply(norm_digits)
    if "國碼" not in members.columns or "手機號碼" not in members.columns:
        st.error("會員名單缺少 '國碼' 或 '手機號碼' 欄位。")
        st.stop()
    members["phone_key"] = members["國碼"].fillna("") + "-" + members["手機號碼"].fillna("")
    valid_members = members[members["phone_key"].str.contains("-") & (members["phone_key"] != "-")].copy()
    member_cols = ["phone_key", "來店次數", "會員姓名"]
    member_cols = [c for c in member_cols if c in valid_members.columns]
    merged = valid_bills.merge(valid_members[member_cols], on="phone_key", how="left")

# First checkout per phone
merged_sorted = merged.sort_values("結帳操作時間")
first_checkout = merged_sorted.groupby("phone_key", as_index=False).first()

# New customer definition (全品牌首次結帳)
new_first = first_checkout.copy()

end_date_all = merged["結帳操作時間"].max()

# Choose member name column
name_col = "會員姓名"
if "會員姓名" not in new_first.columns:
    # fallback if merge suffixes exist
    if "會員姓名_y" in new_first.columns:
        new_first[name_col] = new_first["會員姓名_y"]
    elif "會員姓名_x" in new_first.columns:
        new_first[name_col] = new_first["會員姓名_x"]

# Designer column
if "設計師" not in new_first.columns:
    st.error("帳單檔缺少 '設計師' 欄位，無法分師傅。")
    st.stop()

# Store column (optional)
has_store = "分店" in new_first.columns
if not has_store:
    st.warning("帳單檔缺少 '分店' 欄位，將無法依分店分組。")

# Sidebar filters
with st.sidebar:
    st.header("篩選")
    if has_store:
        store_options = sorted([s for s in merged["分店"].dropna().unique()])
        store_filter = st.multiselect("分店", store_options, default=store_options)
    else:
        store_filter = None
    if has_store and store_filter is not None:
        base_for_designers = merged[merged["分店"].isin(store_filter)]
    else:
        base_for_designers = merged
    designer_options = sorted([s for s in base_for_designers["設計師"].dropna().unique()])
    exclude_designers = st.multiselect("排除師傅", designer_options, default=[])
    include_options = [d for d in designer_options if d not in set(exclude_designers)]
    designer_filter = st.multiselect("師傅", include_options, default=include_options)

merged_store = merged.copy()
if has_store and store_filter is not None:
    merged_store = merged_store[merged_store["分店"].isin(store_filter)]
merged_store = merged_store.sort_values("結帳操作時間")
merged_designer = merged_store[merged_store["設計師"].isin(designer_filter)].copy()

# 新客以全品牌判定，請確保已上傳全品牌資料
st.caption("新客口徑為全品牌歷史首次結帳；若未上傳全品牌帳單，可能高估新客與流失率。")

end_date = merged_store["結帳操作時間"].max()
if pd.isna(end_date):
    st.error("篩選後沒有可用資料。")
    st.stop()

# 新客（全品牌首購）→ 同分店回店
new_first_store = new_first.copy()
if has_store and store_filter is not None:
    new_first_store = new_first_store[new_first_store["分店"].isin(store_filter)]

checkouts_by_phone_store = merged_store.groupby(["phone_key", "分店"])["結帳操作時間"].apply(list)

def distinct_visit_dates(times):
    dates = set()
    for t in times:
        if pd.isna(t):
            continue
        ts = pd.to_datetime(t, errors="coerce")
        if pd.isna(ts):
            continue
        dates.add(ts.date())
    return sorted(dates)

churn_flags = []
return_days_store = []
for _, row in new_first_store.iterrows():
    pk = row["phone_key"]
    store = row.get("分店")
    first_time = row["結帳操作時間"]
    if pd.isna(first_time):
        return_days_store.append(np.nan)
        churn_flags.append(True)
        continue
    first_date = pd.to_datetime(first_time).date()
    times = checkouts_by_phone_store.get((pk, store), [])
    visit_dates = distinct_visit_dates(times)
    next_dates = [d for d in visit_dates if d > first_date]
    next_date = next_dates[0] if next_dates else None
    if next_date is not None:
        return_days = (next_date - first_date).days
        return_days_store.append(return_days)
        churn_flags.append(return_days > CHURN_DAYS)
    else:
        return_days_store.append(np.nan)
        churn_flags.append(True)
new_first_store["return_days_store"] = return_days_store
new_first_store["churn"] = churn_flags
new_first_store["matured"] = new_first_store["結帳操作時間"] + pd.Timedelta(days=CHURN_DAYS) <= end_date

filtered_new_first = new_first_store[new_first_store["設計師"].isin(designer_filter)].copy()

def add_repeat_flags(df, list_map, key_cols, first_col, t2, t3):
    days2 = []
    days3 = []
    repeat2 = []
    repeat3 = []
    for _, row in df.iterrows():
        key = tuple(row[c] for c in key_cols)
        first_time = row[first_col]
        if pd.isna(first_time):
            days2.append(np.nan)
            days3.append(np.nan)
            repeat2.append(False)
            repeat3.append(False)
            continue
        first_date = pd.to_datetime(first_time).date()
        times = list_map.get(key, [])
        visit_dates = distinct_visit_dates(times)
        after = [d for d in visit_dates if d > first_date]
        d2 = (after[0] - first_date).days if len(after) >= 1 else np.nan
        d3 = (after[1] - first_date).days if len(after) >= 2 else np.nan
        days2.append(d2)
        days3.append(d3)
        repeat2.append(pd.notna(d2) and d2 <= t2)
        repeat3.append(pd.notna(d3) and d3 <= t3)
    df["days_to_2nd"] = days2
    df["days_to_3rd"] = days3
    df["repeat2"] = repeat2
    df["repeat3"] = repeat3
    return df

def zscore_series(series):
    s = pd.to_numeric(series, errors="coerce")
    mean = s.mean()
    std = s.std()
    if std == 0 or pd.isna(std):
        return pd.Series(np.nan, index=s.index)
    return (s - mean) / std

def reliability_factor(n, n0=30):
    n = pd.to_numeric(n, errors="coerce")
    return np.where((n > 0) & pd.notna(n), np.minimum(1, np.sqrt(n / n0)), np.nan)

def quantile_default(df, col, q, fallback):
    if col not in df.columns:
        return fallback
    s = pd.to_numeric(df[col], errors="coerce").dropna()
    if s.empty:
        return fallback
    return float(s.quantile(q))

def col_series(df, col):
    if col in df.columns:
        return df[col]
    return pd.Series(np.nan, index=df.index)

def goal_score_high(series, target, floor=0.0):
    v = pd.to_numeric(series, errors="coerce").astype(float)
    if pd.isna(target):
        return pd.Series(np.nan, index=v.index)
    denom = float(target) - float(floor)
    if denom <= 0:
        return pd.Series(np.nan, index=v.index)
    score = (v - float(floor)) / denom * 100.0
    return score.clip(lower=0, upper=100)

def goal_score_low(series, target, ceiling=1.0):
    v = pd.to_numeric(series, errors="coerce").astype(float)
    if pd.isna(target):
        return pd.Series(np.nan, index=v.index)
    denom = float(ceiling) - float(target)
    if denom <= 0:
        return pd.Series(np.nan, index=v.index)
    score = (float(ceiling) - v) / denom * 100.0
    return score.clip(lower=0, upper=100)

def shrink_to_neutral(score, rel, neutral=50.0):
    s = pd.to_numeric(score, errors="coerce").astype(float)
    r = pd.to_numeric(rel, errors="coerce").astype(float)
    return neutral + (s - neutral) * r

def score_insight(df, score_col, value, tag_mode="generic"):
    if pd.isna(value):
        return None, None, None, None, None
    s = pd.to_numeric(df[score_col], errors="coerce").dropna()
    if s.empty:
        return None, None, None, None, None
    pct = (s <= value).sum() / len(s) * 100
    rank = int((s > value).sum() + 1)
    total = int(len(s))
    rank_text = f"{rank}/{total}"
    pct_value = int(round(pct))
    pct_value_text = f"{pct_value}/100"
    pct_thresholds = [90, 75, 60, 40, 25]
    if pct >= pct_thresholds[0]:
        tier = 0
    elif pct >= pct_thresholds[1]:
        tier = 1
    elif pct >= pct_thresholds[2]:
        tier = 2
    elif pct >= pct_thresholds[3]:
        tier = 3
    elif pct >= pct_thresholds[4]:
        tier = 4
    else:
        tier = 5

    generic_labels = ["頂尖", "領先", "略高", "中等", "落後", "脫隊"]
    acq_labels = ["最多", "許多", "略多", "中等", "偏少", "極少"]
    colors = [
        ("#e6f4ea", "#1b7f3b"),
        ("#e8f5e9", "#2ca02c"),
        ("#e8f1fb", "#4e79a7"),
        ("#eef2f7", "#6b7280"),
        ("#fff4e5", "#f2b134"),
        ("#fdecea", "#d62728"),
    ]
    labels = acq_labels if tag_mode == "acq" else generic_labels
    bg, color = colors[tier]
    return rank_text, pct_value_text, labels[tier], bg, color

def median_suffix(df, col, fmt):
    if col not in df.columns:
        return None
    s = pd.to_numeric(df[col], errors="coerce").dropna()
    if s.empty:
        return None
    med = s.median()
    if fmt == "percent":
        formatted = f"{med:.1%}"
    elif fmt == "number1":
        formatted = f"{med:.1f}"
    else:
        formatted = f"{med:.0f}"
    return f"中位數 {formatted}"

def _fmt_alt(fmt_key, delta=False):
    if fmt_key == "percent":
        return "+.1%" if delta else ".1%"
    if fmt_key == "number1":
        return "+.1f" if delta else ".1f"
    return "+.0f" if delta else ".0f"

def build_delta_map(df, selected_name, x_col, y_col, x_label, y_label, x_fmt, y_fmt, x_reverse=False, y_reverse=False):
    view = df[["設計師", x_col, y_col]].dropna().copy()
    if view.empty:
        return None
    selected_row = view[view["設計師"] == selected_name]
    if selected_row.empty:
        return None
    base_x = selected_row.iloc[0][x_col]
    base_y = selected_row.iloc[0][y_col]
    view["x_delta"] = (base_x - view[x_col]) if x_reverse else (view[x_col] - base_x)
    view["y_delta"] = (base_y - view[y_col]) if y_reverse else (view[y_col] - base_y)
    view["is_selected"] = view["設計師"] == selected_name

    x_axis = alt.Axis(format=_fmt_alt(x_fmt, delta=True), title=x_label)
    y_axis = alt.Axis(format=_fmt_alt(y_fmt, delta=True), title=y_label)

    base = alt.Chart(view).mark_circle(size=70, color="#bdbdbd").encode(
        x=alt.X("x_delta:Q", axis=x_axis),
        y=alt.Y("y_delta:Q", axis=y_axis),
        tooltip=[
            "設計師",
            alt.Tooltip(x_col, title="X 原始值", format=_fmt_alt(x_fmt)),
            alt.Tooltip(y_col, title="Y 原始值", format=_fmt_alt(y_fmt)),
            alt.Tooltip("x_delta:Q", title="X 差距", format=_fmt_alt(x_fmt, delta=True)),
            alt.Tooltip("y_delta:Q", title="Y 差距", format=_fmt_alt(y_fmt, delta=True)),
        ],
    )
    highlight = alt.Chart(view[view["is_selected"]]).mark_circle(size=160, color="#e15759").encode(
        x="x_delta:Q",
        y="y_delta:Q",
        tooltip=["設計師"],
    )
    rule_x = alt.Chart(pd.DataFrame({"x": [0]})).mark_rule(color="#d6c9b8").encode(x="x:Q")
    rule_y = alt.Chart(pd.DataFrame({"y": [0]})).mark_rule(color="#d6c9b8").encode(y="y:Q")
    return (rule_x + rule_y + base + highlight).properties(height=320)

def add_regular_metrics(df, list_map, key_cols, baseline_col):
    regular_counts = []
    regular_achieved = []
    regular_dates = []
    post_regular_visits = []
    retention_achieved = []
    for _, row in df.iterrows():
        key = tuple(row[c] for c in key_cols)
        baseline_time = row[baseline_col]
        times = list_map.get(key, [])
        times = [t for t in times if pd.notna(t)]
        times.sort()
        after = [t for t in times if t >= baseline_time]
        window_end = baseline_time + timedelta(days=REGULAR_DAYS)
        within = [t for t in after if t <= window_end]
        count_within = len(within)
        regular_counts.append(count_within)
        achieved = count_within >= REGULAR_VISITS
        regular_achieved.append(achieved)
        if achieved:
            reg_date = within[REGULAR_VISITS - 1]
        else:
            reg_date = pd.NaT
        regular_dates.append(reg_date)
        if achieved:
            retention_end = reg_date + timedelta(days=RETENTION_DAYS)
            post = [t for t in after if t > reg_date and t <= retention_end]
            post_count = len(post)
            post_regular_visits.append(post_count)
            retention_achieved.append(post_count >= RETENTION_VISITS)
        else:
            post_regular_visits.append(np.nan)
            retention_achieved.append(np.nan)
    df["regular_count_180"] = regular_counts
    df["regular_achieved"] = regular_achieved
    df["regular_date"] = regular_dates
    df["post_regular_visits_180"] = post_regular_visits
    df["retention_achieved"] = retention_achieved
    return df

checkouts_by_phone_store_designer = (
    merged_store.groupby(["phone_key", "分店", "設計師"])["結帳操作時間"]
    .apply(list)
)

# 建立師傅關係起點（同分店同師傅第一次）
relationship_first = (
    merged_store.sort_values("結帳操作時間")
    .groupby(["phone_key", "分店", "設計師"], as_index=False)
    .first()
    .rename(columns={"結帳操作時間": "baseline_time"})
)
if not relationship_first.empty:
    relationship_first = add_regular_metrics(
        relationship_first,
        checkouts_by_phone_store_designer,
        ["phone_key", "分店", "設計師"],
        "baseline_time",
    )
    relationship_first["regular_matured_180"] = (
        relationship_first["baseline_time"] + pd.Timedelta(days=REGULAR_DAYS) <= end_date
    )
    relationship_first["retention_matured_180"] = (
        relationship_first["regular_date"] + pd.Timedelta(days=RETENTION_DAYS) <= end_date
    )

new_first_store["matured_2"] = new_first_store["結帳操作時間"] + pd.Timedelta(days=T2_DAYS) <= end_date
new_first_store["matured_3"] = new_first_store["結帳操作時間"] + pd.Timedelta(days=T3_DAYS) <= end_date
new_first_store = add_repeat_flags(
    new_first_store,
    checkouts_by_phone_store_designer,
    ["phone_key", "分店", "設計師"],
    "結帳操作時間",
    T2_DAYS,
    T3_DAYS,
)

end_ts = pd.to_datetime(end_date)
start_ts_3m = end_ts - pd.DateOffset(months=3)
start_ts_6m = end_ts - pd.DateOffset(months=6)
start_ts_12m = end_ts - pd.DateOffset(months=RELATIONSHIP_COHORT_MONTHS)

# 新客經營力（近 3 個月）
new_recent = new_first_store[new_first_store["結帳操作時間"] >= start_ts_3m].copy()
new_recent_churn = new_recent[new_recent["matured"]].copy()
new_churn_by_designer = (
    new_recent_churn.groupby("設計師")
    .agg(
        new_customers_3m=("phone_key", "count"),
        new_churned_3m=("churn", "sum"),
    )
    .reset_index()
)
new_churn_by_designer["new_retained_3m"] = (
    new_churn_by_designer["new_customers_3m"] - new_churn_by_designer["new_churned_3m"]
)
new_churn_by_designer["new_churn_rate_3m"] = np.where(
    new_churn_by_designer["new_customers_3m"] > 0,
    new_churn_by_designer["new_churned_3m"] / new_churn_by_designer["new_customers_3m"],
    np.nan,
)

new_recent2 = new_recent[new_recent["matured_2"]].copy()
new_by_designer = (
    new_recent2.groupby("設計師")
    .agg(
        new_repeat_base_3m=("phone_key", "count"),
        new_repeat2_count=("repeat2", "sum"),
    )
    .reset_index()
)
new_by_designer["new_repeat_rate_3m"] = np.where(
    new_by_designer["new_repeat_base_3m"] > 0,
    new_by_designer["new_repeat2_count"] / new_by_designer["new_repeat_base_3m"],
    np.nan,
)
new_by_designer.loc[
    new_by_designer["new_repeat_base_3m"] < min_repeat_base, "new_repeat_rate_3m"
] = np.nan

new_recent3 = new_recent[new_recent["matured_3"]].copy()
new_deep = (
    new_recent3.groupby("設計師")
    .agg(new_deep_rate_3m=("repeat3", "mean"), new_deep_n=("repeat3", "count"))
    .reset_index()
)
new_deep.loc[new_deep["new_deep_n"] < min_repeat_base, "new_deep_rate_3m"] = np.nan

# 熟客經營力（近 3 個月，同分店熟客）
history = merged_store[merged_store["結帳操作時間"] < start_ts_3m]
familiar_keys = (
    history.groupby(["phone_key", "分店"])
    .size()
    .reset_index(name="visits_before")
)
familiar_keys = familiar_keys[familiar_keys["visits_before"] >= 2][["phone_key", "分店"]]

if familiar_keys.empty:
    familiar_first = pd.DataFrame(columns=["phone_key", "分店", "設計師", "baseline_time"])
else:
    familiar_visits = merged_store[merged_store["結帳操作時間"] >= start_ts_3m]
    familiar_visits = familiar_visits.merge(familiar_keys, on=["phone_key", "分店"], how="inner")
    familiar_first = (
        familiar_visits.sort_values("結帳操作時間")
        .groupby(["phone_key", "分店", "設計師"], as_index=False)
        .first()
        .rename(columns={"結帳操作時間": "baseline_time"})
    )

if not familiar_first.empty:
    familiar_first["matured_2"] = familiar_first["baseline_time"] + pd.Timedelta(days=T2_DAYS) <= end_date
    familiar_first["matured_3"] = familiar_first["baseline_time"] + pd.Timedelta(days=T3_DAYS) <= end_date
    familiar_first = add_repeat_flags(
        familiar_first,
        checkouts_by_phone_store_designer,
        ["phone_key", "分店", "設計師"],
        "baseline_time",
        T2_DAYS,
        T3_DAYS,
    )

if familiar_first.empty:
    fam_recent2 = familiar_first.copy()
    fam_recent3 = familiar_first.copy()
else:
    fam_recent2 = familiar_first[familiar_first["matured_2"]].copy()
    fam_recent3 = familiar_first[familiar_first["matured_3"]].copy()
fam_by_designer = (
    fam_recent2.groupby("設計師")
    .agg(
        familiar_customers_3m=("phone_key", "count"),
        familiar_repeat2_count=("repeat2", "sum"),
    )
    .reset_index()
)
fam_by_designer["familiar_repeat_rate_3m"] = np.where(
    fam_by_designer["familiar_customers_3m"] > 0,
    fam_by_designer["familiar_repeat2_count"] / fam_by_designer["familiar_customers_3m"],
    np.nan,
)
fam_by_designer.loc[
    fam_by_designer["familiar_customers_3m"] < min_repeat_base, "familiar_repeat_rate_3m"
] = np.nan

fam_deep = (
    fam_recent3.groupby("設計師")
    .agg(familiar_deep_rate_3m=("repeat3", "mean"), familiar_deep_n=("repeat3", "count"))
    .reset_index()
)
fam_deep.loc[fam_deep["familiar_deep_n"] < min_repeat_base, "familiar_deep_rate_3m"] = np.nan

# 合併師傅指標
designer_metrics = (
    new_churn_by_designer
    .merge(new_by_designer, on="設計師", how="outer")
    .merge(new_deep, on="設計師", how="outer")
    .merge(fam_by_designer, on="設計師", how="outer")
    .merge(fam_deep, on="設計師", how="outer")
)

# 近 3 個月總單量
orders_recent = merged_store[merged_store["結帳操作時間"] >= start_ts_3m].copy()
orders_summary = (
    orders_recent.groupby("設計師")
    .size()
    .reset_index(name="total_orders_3m")
)
designer_metrics = designer_metrics.merge(orders_summary, on="設計師", how="left")
designer_metrics["new_share_3m"] = np.where(
    pd.to_numeric(designer_metrics["total_orders_3m"], errors="coerce") > 0,
    pd.to_numeric(designer_metrics["new_customers_3m"], errors="coerce")
    / pd.to_numeric(designer_metrics["total_orders_3m"], errors="coerce"),
    np.nan,
)

# 出勤狀態（以月份計）
orders_monthly = merged_store.copy()
orders_monthly["month_period"] = orders_monthly["結帳操作時間"].dt.to_period("M")
orders_monthly["date"] = orders_monthly["結帳操作時間"].dt.date
end_month = end_ts.to_period("M")
start_month_3m = end_month - 2
recent_months = orders_monthly[orders_monthly["month_period"] >= start_month_3m]
active_months_3m = (
    recent_months.groupby("設計師")["month_period"]
    .nunique()
    .reset_index(name="active_months_3m")
)
active_days_monthly = (
    orders_monthly.groupby(["設計師", "month_period"])["date"]
    .nunique()
    .reset_index(name="active_days")
)
active_days_recent = active_days_monthly[active_days_monthly["month_period"] >= start_month_3m]
active_days_3m = (
    active_days_recent.groupby("設計師")["active_days"]
    .sum()
    .reset_index(name="active_days_3m")
)
avg_active_days_3m = (
    active_days_recent.groupby("設計師")["active_days"]
    .sum()
    .div(3)
    .reset_index(name="avg_active_days_3m")
)
prev_month = end_month - 1
orders_prev = orders_monthly[orders_monthly["month_period"] == prev_month]
has_order_prev = (
    orders_prev.groupby("設計師")
    .size()
    .reset_index(name="orders_prev_month")
)
has_order_prev["has_order_prev_month"] = has_order_prev["orders_prev_month"] > 0
last_month = (
    orders_monthly.groupby("設計師")["month_period"]
    .max()
    .reset_index(name="last_order_month")
)
last_month["months_since_last"] = last_month["last_order_month"].apply(lambda p: end_month.ordinal - p.ordinal)

attendance_summary = (
    active_months_3m
    .merge(active_days_3m, on="設計師", how="left")
    .merge(avg_active_days_3m, on="設計師", how="left")
    .merge(has_order_prev[["設計師", "has_order_prev_month"]], on="設計師", how="left")
    .merge(last_month[["設計師", "months_since_last"]], on="設計師", how="left")
)
designer_metrics = designer_metrics.merge(attendance_summary, on="設計師", how="left")
designer_metrics["new_per_active_day_3m"] = np.where(
    pd.to_numeric(designer_metrics.get("active_days_3m"), errors="coerce") > 0,
    pd.to_numeric(designer_metrics.get("new_customers_3m"), errors="coerce")
    / pd.to_numeric(designer_metrics.get("active_days_3m"), errors="coerce"),
    np.nan,
)
designer_metrics["new_retention_rate_3m"] = np.where(
    pd.notna(designer_metrics.get("new_churn_rate_3m")),
    1 - pd.to_numeric(designer_metrics.get("new_churn_rate_3m"), errors="coerce"),
    np.nan,
)

# 熟客化/熟客維持（近 12 個月關係起點）
regular_summary = None
retention_summary = None
if not relationship_first.empty:
    rel_recent = relationship_first[relationship_first["baseline_time"] >= start_ts_12m].copy()
    regular_base = rel_recent[rel_recent["regular_matured_180"]].copy()
    if not regular_base.empty:
        regular_base["regular_achieved_int"] = regular_base["regular_achieved"].fillna(False).astype(int)
        regular_base["regular_days"] = (regular_base["regular_date"] - regular_base["baseline_time"]).dt.days
        regular_summary = (
            regular_base.groupby("設計師")
            .agg(
                regular_base_180=("phone_key", "count"),
                regular_achieved_180=("regular_achieved_int", "sum"),
                regular_days_avg_180=("regular_days", "mean"),
            )
            .reset_index()
        )
        regular_summary["regular_rate_180"] = np.where(
            regular_summary["regular_base_180"] > 0,
            regular_summary["regular_achieved_180"] / regular_summary["regular_base_180"],
            np.nan,
        )
        designer_metrics = designer_metrics.merge(regular_summary, on="設計師", how="left")

    retention_base = rel_recent[
        (rel_recent["regular_achieved"] == True) & (rel_recent["retention_matured_180"])
    ].copy()
    if not retention_base.empty:
        retention_base["retention_achieved_int"] = retention_base["retention_achieved"].fillna(False).astype(int)
        retention_summary = (
            retention_base.groupby("設計師")
            .agg(
                retention_base_180=("phone_key", "count"),
                retention_achieved_180=("retention_achieved_int", "sum"),
                post_regular_visits_avg_180=("post_regular_visits_180", "mean"),
            )
            .reset_index()
        )
        retention_summary["post_regular_visits_monthly_avg_180"] = (
            retention_summary["post_regular_visits_avg_180"] / 6
        )
        retention_summary["retention_rate_180"] = np.where(
            retention_summary["retention_base_180"] > 0,
            retention_summary["retention_achieved_180"] / retention_summary["retention_base_180"],
            np.nan,
        )
        designer_metrics = designer_metrics.merge(retention_summary, on="設計師", how="left")

    for col in [
        "regular_rate_180",
        "regular_base_180",
        "regular_achieved_180",
        "regular_days_avg_180",
        "retention_rate_180",
        "retention_base_180",
        "retention_achieved_180",
        "post_regular_visits_avg_180",
        "post_regular_visits_monthly_avg_180",
    ]:
        if col not in designer_metrics.columns:
            designer_metrics[col] = np.nan

# 經營中熟客：已達熟客(180天達5次)但後180天維持觀察期尚未滿
designer_metrics["in_service_regular_180"] = (
    pd.to_numeric(designer_metrics["regular_achieved_180"], errors="coerce")
    - pd.to_numeric(designer_metrics["retention_base_180"], errors="coerce")
).clip(lower=0)

# 指定率（近 3 個月）
if "is_requested" in merged_store.columns and merged_store["is_requested"].notna().any():
    request_recent = merged_store[merged_store["結帳操作時間"] >= start_ts_3m].copy()
    request_recent["is_requested_num"] = request_recent["is_requested"].map({True: 1, False: 0})
    request_summary = (
        request_recent.groupby("設計師")["is_requested_num"]
        .agg(request_yes_3m="sum", request_total_3m="count")
        .reset_index()
    )
    request_summary["request_rate_3m"] = np.where(
        request_summary["request_total_3m"] > 0,
        request_summary["request_yes_3m"] / request_summary["request_total_3m"],
        np.nan,
    )
    designer_metrics = designer_metrics.merge(request_summary, on="設計師", how="left")
else:
    designer_metrics["request_rate_3m"] = np.nan
    designer_metrics["request_yes_3m"] = np.nan
    designer_metrics["request_total_3m"] = np.nan
    st.warning("帳單檔缺少「指定」欄位或無有效值，指定率將不計算。")

# Store monthly summary (average per month)
store_monthly_avg = None
if has_store:
    store_month = new_first_store[new_first_store["matured"]].copy()
    store_month["month"] = store_month["結帳操作時間"].dt.to_period("M").astype(str)
    month_summary = (
        store_month.groupby(["分店", "month"])
        .agg(matured_new_customers=("phone_key", "count"), churned=("churn", "sum"))
        .reset_index()
    )
    month_summary["retained"] = month_summary["matured_new_customers"] - month_summary["churned"]
    month_summary["repeat_rate"] = np.where(
        month_summary["matured_new_customers"] > 0,
        month_summary["retained"] / month_summary["matured_new_customers"],
        np.nan,
    )
    store_monthly_avg = (
        month_summary.groupby("分店")
        .agg(
            月平均新客數=("matured_new_customers", "mean"),
            月平均流失數=("churned", "mean"),
            月平均留住數=("retained", "mean"),
            平均回店率=("repeat_rate", "mean"),
            月份數=("month", "nunique"),
        )
        .reset_index()
    )

summary_by_store = None
summary_by_store_designer = None
if has_store:
    matured_store = new_first_store[new_first_store["matured"]].copy()
    summary_by_store = (
        matured_store.groupby("分店")
        .agg(
            matured_new_customers=("phone_key", "count"),
            churned=("churn", "sum"),
        )
        .reset_index()
    )
    summary_by_store["churn_rate"] = summary_by_store["churned"] / summary_by_store["matured_new_customers"]
    summary_by_store["repeat_rate"] = 1 - summary_by_store["churn_rate"]

    summary_by_store_designer = (
        matured_store[matured_store["設計師"].isin(designer_filter)]
        .groupby(["分店", "設計師"])
        .agg(
            matured_new_customers=("phone_key", "count"),
            churned=("churn", "sum"),
        )
        .reset_index()
    )
    summary_by_store_designer["churn_rate"] = (
        summary_by_store_designer["churned"] / summary_by_store_designer["matured_new_customers"]
    )
    summary_by_store_designer["repeat_rate"] = 1 - summary_by_store_designer["churn_rate"]

# Vacancy metrics (monthly, 168h cap)
vacancy_monthly = None
vacancy_recent = None
if "項目" in merged_store.columns:
    time_df = merged_store.copy()
    time_df["duration_minutes"] = time_df["項目"].apply(extract_minutes)
    time_df["duration_hours"] = time_df["duration_minutes"].apply(minutes_to_hours)
    time_df["month"] = time_df["結帳操作時間"].dt.to_period("M").astype(str)
    if has_store:
        group_cols = ["分店", "設計師", "month"]
    else:
        group_cols = ["設計師", "month"]
    vacancy_monthly = (
        time_df.groupby(group_cols)["duration_hours"]
        .sum()
        .reset_index()
    )
    vacancy_monthly["vacancy_rate"] = (1 - vacancy_monthly["duration_hours"] / 168.0).clip(lower=0, upper=1)
    vm = vacancy_monthly.copy()
    vm["month_start"] = pd.to_datetime(vm["month"] + "-01")
    vm = vm[vm["month_start"] >= start_ts_3m]
    vacancy_recent = (
        vm.groupby("設計師")["vacancy_rate"]
        .mean()
        .reset_index()
        .rename(columns={"vacancy_rate": "vacancy_rate_3m"})
    )
    designer_metrics = designer_metrics.merge(vacancy_recent, on="設計師", how="left")

# 業績穩定度（近 6 個月）
stability_by_designer = None
stability_df = merged_store.copy()
stability_df["month"] = stability_df["結帳操作時間"].dt.to_period("M").astype(str)
stability_df["date"] = stability_df["結帳操作時間"].dt.date
active_days = (
    stability_df.groupby(["設計師", "month"])["date"]
    .nunique()
    .reset_index(name="active_days")
)
active_days["month_start"] = pd.to_datetime(active_days["month"] + "-01")
active_days_6m = active_days[active_days["month_start"] >= start_ts_6m]

stability_by_designer = (
    active_days_6m.groupby("設計師")
    .agg(
        active_days_avg_6m=("active_days", "mean"),
        active_days_cv_6m=("active_days", lambda s: s.std() / s.mean() if s.mean() else np.nan),
        active_months_6m=("month", "nunique"),
    )
    .reset_index()
)

if vacancy_monthly is not None:
    hours_df = vacancy_monthly.copy()
    hours_df = hours_df.rename(columns={"duration_hours": "service_hours"})
    hours_df["month_start"] = pd.to_datetime(hours_df["month"] + "-01")
    hours_6m = hours_df[hours_df["month_start"] >= start_ts_6m]
    hours_summary = (
        hours_6m.groupby("設計師")["service_hours"]
        .agg(service_hours_avg_6m="mean", service_hours_cv_6m=lambda s: s.std() / s.mean() if s.mean() else np.nan)
        .reset_index()
    )
    stability_by_designer = stability_by_designer.merge(hours_summary, on="設計師", how="left")

last_tx = (
    merged_store.groupby("設計師")["結帳操作時間"]
    .max()
    .reset_index()
    .rename(columns={"結帳操作時間": "last_tx"})
)
last_tx["days_since_last_tx"] = (end_ts - last_tx["last_tx"]).dt.days
stability_by_designer = stability_by_designer.merge(last_tx, on="設計師", how="left")
designer_metrics = designer_metrics.merge(stability_by_designer, on="設計師", how="left")

# 四大區塊戰力指標（Z-score + 樣本數修正）
basic_z = pd.DataFrame({
    "z_active_days": zscore_series(designer_metrics.get("avg_active_days_3m")),
    "z_active_days_3m": zscore_series(designer_metrics.get("active_days_3m")),
    "z_total_orders": zscore_series(designer_metrics.get("total_orders_3m")),
    "z_vacancy": -zscore_series(designer_metrics.get("vacancy_rate_3m")),
})
designer_metrics["basic_z"] = basic_z.mean(axis=1, skipna=True)
designer_metrics["basic_rel"] = reliability_factor(designer_metrics.get("total_orders_3m"))
designer_metrics["basic_score"] = designer_metrics["basic_z"] * designer_metrics["basic_rel"]
designer_metrics["basic_score_0100"] = np.clip(50 + 10 * designer_metrics["basic_score"], 0, 100)

new_acq_z = pd.DataFrame({
    "z_new_share": zscore_series(designer_metrics.get("new_share_3m")),
    "z_new_per_day": zscore_series(designer_metrics.get("new_per_active_day_3m")),
})
designer_metrics["new_acq_z"] = new_acq_z.mean(axis=1, skipna=True)
designer_metrics["new_acq_rel"] = reliability_factor(designer_metrics.get("new_customers_3m"))
designer_metrics["new_acq_score"] = designer_metrics["new_acq_z"] * designer_metrics["new_acq_rel"]
designer_metrics["new_acq_score_0100"] = np.clip(50 + 10 * designer_metrics["new_acq_score"], 0, 100)

new_ret_z = pd.DataFrame({
    "z_new_retention": zscore_series(designer_metrics.get("new_retention_rate_3m")),
})
designer_metrics["new_ret_z"] = new_ret_z.mean(axis=1, skipna=True)
designer_metrics["new_ret_rel"] = reliability_factor(designer_metrics.get("new_customers_3m"))
designer_metrics["new_ret_score"] = designer_metrics["new_ret_z"] * designer_metrics["new_ret_rel"]
designer_metrics["new_ret_score_0100"] = np.clip(50 + 10 * designer_metrics["new_ret_score"], 0, 100)

convert_z = pd.DataFrame({
    "z_regular_rate": zscore_series(designer_metrics.get("regular_rate_180")),
    "z_regular_days": -zscore_series(designer_metrics.get("regular_days_avg_180")),
})
designer_metrics["convert_z"] = convert_z.mean(axis=1, skipna=True)
designer_metrics["convert_rel"] = reliability_factor(designer_metrics.get("regular_base_180"))
designer_metrics["convert_score"] = designer_metrics["convert_z"] * designer_metrics["convert_rel"]
designer_metrics["convert_score_0100"] = np.clip(50 + 10 * designer_metrics["convert_score"], 0, 100)

retain_z = pd.DataFrame({
    "z_retention_rate": zscore_series(designer_metrics.get("retention_rate_180")),
    "z_post_visits": zscore_series(designer_metrics.get("post_regular_visits_avg_180")),
})
designer_metrics["retain_z"] = retain_z.mean(axis=1, skipna=True)
designer_metrics["retain_rel"] = reliability_factor(designer_metrics.get("retention_base_180"))
designer_metrics["retain_score"] = designer_metrics["retain_z"] * designer_metrics["retain_rel"]
designer_metrics["retain_score_0100"] = np.clip(50 + 10 * designer_metrics["retain_score"], 0, 100)

stability_cv = designer_metrics["service_hours_cv_6m"] if "service_hours_cv_6m" in designer_metrics.columns else pd.Series(np.nan, index=designer_metrics.index)
stability_cv = np.where(pd.notna(stability_cv), stability_cv, designer_metrics.get("active_days_cv_6m"))
designer_metrics["stability_cv"] = stability_cv
designer_metrics["stability_z"] = -zscore_series(designer_metrics.get("stability_cv"))
designer_metrics["stability_rel"] = reliability_factor(designer_metrics.get("active_months_6m"), n0=6)
designer_metrics["stability_score"] = designer_metrics["stability_z"] * designer_metrics["stability_rel"]
designer_metrics["stability_score_0100"] = np.clip(50 + 10 * designer_metrics["stability_score"], 0, 100)

block_scores = designer_metrics[
    ["basic_score", "new_acq_score", "new_ret_score", "convert_score", "retain_score", "stability_score"]
]
weights = np.array([1/6, 1/6, 1/6, 1/6, 1/6, 1/6])
valid_mask = block_scores.notna().values
weighted = block_scores.fillna(0).values * weights
weight_sum = (valid_mask * weights).sum(axis=1)
overall_z = np.where(weight_sum > 0, weighted.sum(axis=1) / weight_sum, np.nan)
designer_metrics["overall_score_z"] = overall_z
designer_metrics["overall_score"] = np.clip(50 + 10 * overall_z, 0, 100)

# 目標達成分（0-100，達標=100；並依樣本數修正）
goal_base = designer_metrics[designer_metrics["設計師"].isin(designer_filter)].copy()

default_target_avg_active_days = 16.5
default_target_active_days_3m = 50.0
default_target_total_orders_3m = 188.0
default_target_vacancy_rate_3m = 0.25

default_target_new_share_3m = 0.15
default_target_new_per_active_day_3m = 0.35
default_target_new_retention_rate_3m = 0.60

default_target_regular_rate_180 = 0.18
default_target_regular_days_avg_180 = 60.0
default_target_retention_rate_180 = 0.80
default_target_post_regular_visits_monthly_avg_180 = 1.50

stability_ceiling = 1.0
default_target_stability_cv = 0.45

with st.sidebar:
    with st.expander("目標達成分設定（可選）", expanded=False):
        st.caption("達標=100 分；未達標依比例扣分。低樣本會往 50 分收斂。")
        target_avg_active_days = st.number_input(
            "每月平均有單天數目標(近3月)",
            min_value=0.0,
            max_value=31.0,
            value=float(round(default_target_avg_active_days, 1)),
            step=0.5,
        )
        target_active_days_3m = st.number_input(
            "近3月有單天數目標",
            min_value=0.0,
            max_value=93.0,
            value=float(round(default_target_active_days_3m, 0)),
            step=1.0,
        )
        target_total_orders_3m = st.number_input(
            "總單量目標(3M)",
            min_value=0.0,
            value=float(round(default_target_total_orders_3m, 0)),
            step=10.0,
        )
        target_vacancy_rate_3m_pct = st.slider(
            "空窗率目標(3M, 越低越好)",
            min_value=0.0,
            max_value=100.0,
            value=float(round(default_target_vacancy_rate_3m * 100, 1)),
            step=0.5,
        )
        st.divider()
        target_new_share_3m_pct = st.slider(
            "新客占比目標(3M)",
            min_value=0.0,
            max_value=100.0,
            value=float(round(default_target_new_share_3m * 100, 1)),
            step=0.5,
        )
        target_new_per_active_day_3m = st.number_input(
            "新客/有單天數目標(3M)",
            min_value=0.0,
            value=float(round(default_target_new_per_active_day_3m, 2)),
            step=0.05,
        )
        target_new_retention_rate_3m_pct = st.slider(
            "新客留存率目標(60天, 同分店)",
            min_value=0.0,
            max_value=100.0,
            value=float(round(default_target_new_retention_rate_3m * 100, 1)),
            step=0.5,
        )
        st.divider()
        target_regular_rate_180_pct = st.slider(
            "熟客化率目標(180天達≥5次)",
            min_value=0.0,
            max_value=100.0,
            value=float(round(default_target_regular_rate_180 * 100, 1)),
            step=0.5,
        )
        target_regular_days_avg_180 = st.number_input(
            "平均達標天數目標(越低越好)",
            min_value=1.0,
            max_value=float(REGULAR_DAYS),
            value=float(round(default_target_regular_days_avg_180, 0)),
            step=5.0,
        )
        target_retention_rate_180_pct = st.slider(
            "熟客維持率目標(後180天≥3次)",
            min_value=0.0,
            max_value=100.0,
            value=float(round(default_target_retention_rate_180 * 100, 1)),
            step=0.5,
        )
        target_post_regular_visits_monthly_avg_180 = st.number_input(
            "熟客月均回訪次數目標(後180天)",
            min_value=0.0,
            value=float(round(default_target_post_regular_visits_monthly_avg_180, 2)),
            step=0.1,
        )
        st.divider()
        target_stability_cv = st.number_input(
            "業績穩定度目標(CV, 越低越好)",
            min_value=0.0,
            max_value=float(stability_ceiling),
            value=float(round(default_target_stability_cv, 2)),
            step=0.05,
        )

target_vacancy_rate_3m = float(target_vacancy_rate_3m_pct) / 100.0
target_new_share_3m = float(target_new_share_3m_pct) / 100.0
target_new_retention_rate_3m = float(target_new_retention_rate_3m_pct) / 100.0
target_regular_rate_180 = float(target_regular_rate_180_pct) / 100.0
target_retention_rate_180 = float(target_retention_rate_180_pct) / 100.0

basic_goal_components = pd.DataFrame({
    "g_avg_active_days": goal_score_high(col_series(designer_metrics, "avg_active_days_3m"), target_avg_active_days, floor=0.0),
    "g_active_days_3m": goal_score_high(col_series(designer_metrics, "active_days_3m"), target_active_days_3m, floor=0.0),
    "g_total_orders": goal_score_high(col_series(designer_metrics, "total_orders_3m"), target_total_orders_3m, floor=0.0),
    "g_vacancy": goal_score_low(col_series(designer_metrics, "vacancy_rate_3m"), target_vacancy_rate_3m, ceiling=1.0),
})
designer_metrics["basic_goal_raw"] = basic_goal_components.mean(axis=1, skipna=True)
designer_metrics["basic_goal_0100"] = shrink_to_neutral(designer_metrics["basic_goal_raw"], designer_metrics.get("basic_rel"))

new_acq_goal_components = pd.DataFrame({
    "g_new_share": goal_score_high(col_series(designer_metrics, "new_share_3m"), target_new_share_3m, floor=0.0),
    "g_new_per_day": goal_score_high(col_series(designer_metrics, "new_per_active_day_3m"), target_new_per_active_day_3m, floor=0.0),
})
designer_metrics["new_acq_goal_raw"] = new_acq_goal_components.mean(axis=1, skipna=True)
designer_metrics["new_acq_goal_0100"] = shrink_to_neutral(designer_metrics["new_acq_goal_raw"], designer_metrics.get("new_acq_rel"))

designer_metrics["new_ret_goal_raw"] = goal_score_high(col_series(designer_metrics, "new_retention_rate_3m"), target_new_retention_rate_3m, floor=0.0)
designer_metrics["new_ret_goal_0100"] = shrink_to_neutral(designer_metrics["new_ret_goal_raw"], designer_metrics.get("new_ret_rel"))

convert_goal_components = pd.DataFrame({
    "g_regular_rate": goal_score_high(col_series(designer_metrics, "regular_rate_180"), target_regular_rate_180, floor=0.0),
    "g_regular_days": goal_score_low(col_series(designer_metrics, "regular_days_avg_180"), target_regular_days_avg_180, ceiling=float(REGULAR_DAYS)),
})
designer_metrics["convert_goal_raw"] = convert_goal_components.mean(axis=1, skipna=True)
designer_metrics["convert_goal_0100"] = shrink_to_neutral(designer_metrics["convert_goal_raw"], designer_metrics.get("convert_rel"))

retain_goal_components = pd.DataFrame({
    "g_retention_rate": goal_score_high(col_series(designer_metrics, "retention_rate_180"), target_retention_rate_180, floor=0.0),
    "g_post_visits": goal_score_high(col_series(designer_metrics, "post_regular_visits_monthly_avg_180"), target_post_regular_visits_monthly_avg_180, floor=0.0),
})
designer_metrics["retain_goal_raw"] = retain_goal_components.mean(axis=1, skipna=True)
designer_metrics["retain_goal_0100"] = shrink_to_neutral(designer_metrics["retain_goal_raw"], designer_metrics.get("retain_rel"))

designer_metrics["stability_goal_raw"] = goal_score_low(col_series(designer_metrics, "stability_cv"), target_stability_cv, ceiling=float(stability_ceiling))
designer_metrics["stability_goal_0100"] = shrink_to_neutral(designer_metrics["stability_goal_raw"], designer_metrics.get("stability_rel"))

goal_blocks = designer_metrics[
    ["basic_goal_0100", "new_acq_goal_0100", "new_ret_goal_0100", "convert_goal_0100", "retain_goal_0100", "stability_goal_0100"]
]
# 將戰力指標權重調整：暫時不計入「新客獲取量」以避免偏差
goal_weights = np.array([1/5, 1/5, 0, 1/5, 1/5, 1/5])
goal_valid = goal_blocks.notna().values
goal_weighted = goal_blocks.fillna(0).values * goal_weights
goal_weight_sum = (goal_valid * goal_weights).sum(axis=1)
designer_metrics["overall_goal_0100"] = np.where(goal_weight_sum > 0, goal_weighted.sum(axis=1) / goal_weight_sum, np.nan)

designer_metrics_filtered = designer_metrics[designer_metrics["設計師"].isin(designer_filter)].copy()

overall = {
    "matured_new_customers": int(len(new_first_store[new_first_store["matured"]])),
    "churned_matured": int(new_first_store.loc[new_first_store["matured"], "churn"].sum()),
}
overall["retained_matured"] = overall["matured_new_customers"] - overall["churned_matured"]
overall["churn_rate_matured"] = (
    overall["churned_matured"] / overall["matured_new_customers"]
    if overall["matured_new_customers"]
    else None
)
overall["repeat_rate_matured"] = (
    1 - overall["churn_rate_matured"]
    if overall["churn_rate_matured"] is not None
    else None
)

st.subheader("總覽摘要")
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("滿60天新客數", overall["matured_new_customers"])
col2.metric("流失人數", overall["churned_matured"])
col3.metric("留住人數", overall["retained_matured"])
col4.metric("流失率", f"{overall['churn_rate_matured']:.2%}" if overall["churn_rate_matured"] is not None else "-")
col5.metric("回店率(同分店)", f"{overall['repeat_rate_matured']:.2%}" if overall["repeat_rate_matured"] is not None else "-")

st.caption(f"資料截止時間：{end_date}")

if has_store and store_monthly_avg is not None and not store_monthly_avg.empty:
    st.subheader("各分店摘要（每月平均）")
    st.caption("僅計入滿 60 天的新客 cohort（最嚴謹）。")
    for _, row in store_monthly_avg.iterrows():
        store_name = row["分店"]
        avg_new = row["月平均新客數"]
        avg_churn = row["月平均流失數"]
        avg_retained = row["月平均留住數"]
        avg_repeat = row["平均回店率"]
        repeat_text = f"{avg_repeat:.2%}" if pd.notna(avg_repeat) else "-"
        st.write(
            f"- {store_name}：每月平均新客 {avg_new:.1f} 人、流失 {avg_churn:.1f} 人、留住 {avg_retained:.1f} 人、回店率 {repeat_text}"
        )

    st.subheader("分店指標對比（每月平均）")
    metric_base = store_monthly_avg.copy()
    if chart_top_n and chart_top_n > 0:
        top_stores = (
            metric_base.sort_values("月平均新客數", ascending=False)
            .head(int(chart_top_n))["分店"]
            .tolist()
        )
        metric_base = metric_base[metric_base["分店"].isin(top_stores)]

    metric_long = metric_base.melt(
        id_vars=["分店"],
        value_vars=["月平均新客數", "月平均流失數", "月平均留住數"],
        var_name="指標",
        value_name="人數",
    )
    metric_long["指標"] = metric_long["指標"].replace(
        {
            "月平均新客數": "新客(人)",
            "月平均流失數": "流失(人)",
            "月平均留住數": "留住(人)",
        }
    )

    if store_chart_type == "群組直條圖":
        group_chart = alt.Chart(metric_long).mark_bar().encode(
            x=alt.X("分店:N", sort="-y", title="", axis=alt.Axis(labelAngle=-30)),
            y=alt.Y("人數:Q", title="人數(每月平均)"),
            color=alt.Color("指標:N", title="指標"),
            xOffset="指標:N",
            tooltip=["分店", "指標", alt.Tooltip("人數:Q", format=",.1f")],
        ).properties(height=320)
        st.altair_chart(group_chart, use_container_width=True)
    elif store_chart_type == "熱度圖":
        heat = alt.Chart(metric_long).mark_rect().encode(
            x=alt.X("指標:N", title=""),
            y=alt.Y("分店:N", title=""),
            color=alt.Color("人數:Q", title="人數(每月平均)"),
            tooltip=["分店", "指標", alt.Tooltip("人數:Q", format=",.1f")],
        ).properties(height=320)
        text = alt.Chart(metric_long).mark_text(color="white").encode(
            x="指標:N",
            y="分店:N",
            text=alt.Text("人數:Q", format=".1f"),
        )
        st.altair_chart(heat + text, use_container_width=True)
    else:
        stack_long = metric_long[metric_long["指標"].isin(["流失(人)", "留住(人)"])].copy()
        stack = alt.Chart(stack_long).mark_bar().encode(
            x=alt.X("分店:N", sort="-y", title="", axis=alt.Axis(labelAngle=-30)),
            y=alt.Y("人數:Q", stack=True, title="每月平均新客(人)"),
            color=alt.Color("指標:N", title="構成"),
            tooltip=["分店", "指標", alt.Tooltip("人數:Q", format=",.1f")],
        ).properties(height=320)
        totals = store_monthly_avg.copy()
        totals["總人"] = totals["月平均流失數"] + totals["月平均留住數"]
        total_text = alt.Chart(totals).mark_text(dy=-6, color="#333").encode(
            x=alt.X("分店:N", sort="-y"),
            y=alt.Y("總人:Q"),
            text=alt.Text("總人:Q", format=".1f"),
        )
        st.altair_chart(stack + total_text, use_container_width=True)

    if chart_top_n and chart_top_n > 0:
        st.caption(f"分店比較圖僅顯示前 {int(chart_top_n)} 名（依每月平均新客數排序）。")

    # 移除 Small Multiples 依需求

st.subheader("師傅指標比較")
st.caption("新客流失/空窗/總單量以近 3 個月計；熟客化/維持以近 12 個月 cohort 且滿 180 天計。")
st.caption(f"回指率(30/60天)僅顯示樣本數 ≥ {int(min_repeat_base)}。")
metric_df = designer_metrics_filtered.copy()
service_cv = metric_df["service_hours_cv_6m"] if "service_hours_cv_6m" in metric_df.columns else pd.Series(np.nan, index=metric_df.index)
active_cv = metric_df["active_days_cv_6m"] if "active_days_cv_6m" in metric_df.columns else pd.Series(np.nan, index=metric_df.index)
metric_df["業績穩定度(CV)"] = np.where(service_cv.notna(), service_cv, active_cv)

metric_options = {
    "戰力指標(0-100, 高越好)": ("overall_goal_0100", "number1", False),
    "基本狀態(0-100, 高越好)": ("basic_goal_0100", "number1", False),
    "新客獲取量(0-100, 高越好)": ("new_acq_goal_0100", "number1", False),
    "新客留存力(0-100, 高越好)": ("new_ret_goal_0100", "number1", False),
    "熟客轉化力(0-100, 高越好)": ("convert_goal_0100", "number1", False),
    "熟客經營力(0-100, 高越好)": ("retain_goal_0100", "number1", False),
    "業績穩定度(0-100, 高越好)": ("stability_goal_0100", "number1", False),
    "新客留存率(60天，高越好)": ("new_retention_rate_3m", "percent", False),
    "新客流失率(60天，低越好)": ("new_churn_rate_3m", "percent", True),
    "新客數(3M,滿60天，高越好)": ("new_customers_3m", "number0", False),
    "新客留住人數(3M，高越好)": ("new_retained_3m", "number0", False),
    "新客/有單天數(3M,高越好)": ("new_per_active_day_3m", "number1", False),
    "新客占比(新客/總單量,高越好)": ("new_share_3m", "percent", False),
    "新客回指率(30天,3M,高越好)": ("new_repeat_rate_3m", "percent", False),
    "新客深度回指率(60天,3M,高越好)": ("new_deep_rate_3m", "percent", False),
    "熟客回指率(30天,3M,高越好)": ("familiar_repeat_rate_3m", "percent", False),
    "熟客深度回指率(60天,3M,高越好)": ("familiar_deep_rate_3m", "percent", False),
    "熟客化率(180天達5次，高越好)": ("regular_rate_180", "percent", False),
    "熟客維持率(後180天≥3次，高越好)": ("retention_rate_180", "percent", False),
    "熟客月均回訪次數(後180天,高越好)": ("post_regular_visits_monthly_avg_180", "number1", False),
    "總單量(3M，高越好)": ("total_orders_3m", "number0", False),
    "指定率(3M，高越好)": ("request_rate_3m", "percent", False),
    "空窗率(3M，低越好)": ("vacancy_rate_3m", "percent", True),
    "業績穩定度(CV，低越好)": ("業績穩定度(CV)", "number1", True),
    "每月平均有單天數(近3月，高越好)": ("avg_active_days_3m", "number1", False),
}

metric_choice = st.selectbox("選擇指標", list(metric_options.keys()))
metric_col, metric_fmt, metric_asc = metric_options[metric_choice]
metric_view = metric_df[["設計師", metric_col]].dropna().copy()
render_bar_chart(
    metric_view,
    "設計師",
    metric_col,
    metric_choice,
    color="#4e79a7",
    top_n=chart_top_n,
    value_format=metric_fmt,
    orient="horizontal",
    ascending=metric_asc,
)

st.subheader("個別師傅狀態")
st.caption("新客＝全品牌首次；回店口徑＝同分店；熟客＝180 天內同分店同師傅消費 ≥5 次。")

if not designer_filter:
    st.info("目前沒有可顯示的師傅。")
else:
    designer_select = st.selectbox("選擇師傅", designer_filter)
    selected_row = designer_metrics_filtered[designer_metrics_filtered["設計師"] == designer_select].copy()

    if selected_row.empty:
        st.info("此師傅在最近 3 個月內沒有足夠資料。")
    else:
        r = selected_row.iloc[0]
        detail_recent = new_recent_churn[new_recent_churn["設計師"] == designer_select].copy()
        rel_all = relationship_first[relationship_first["設計師"] == designer_select].copy()
        st.markdown("**戰力指標**")
        row1 = st.columns(1)
        with row1[0]:
            rank_txt, pct_value_text, tag, bg, color = score_insight(designer_metrics_filtered, "overall_goal_0100", r.get("overall_goal_0100"))
            with st.container(border=True):
                metric_card(
                    "戰力指標",
                    f"{r['overall_goal_0100']:.0f}分" if pd.notna(r.get("overall_goal_0100")) else "-",
                    "整體分數：把六項數值先各自換成同一把尺，再合在一起。數字越高，代表整體狀況越好、越有競爭力（用來跟同批人比較）。",
                    subtext=pct_value_text if pct_value_text else "",
                    tag_text=tag,
                    tag_bg=bg,
                    tag_color=color,
                    value_suffix=None,
                    meta_text=None,
                    horizontal=True,
                    flat=True,
                )
                with st.expander("構成細項", expanded=False):
                    cc1, cc2, cc3 = st.columns(3)
                    with cc1:
                        metric_card("新客留存力", f"{r['new_ret_goal_0100']:.0f}分" if pd.notna(r.get("new_ret_goal_0100")) else "-", "戰力指標子分數。")
                        metric_card("合作穩定度", f"{r['basic_goal_0100']:.0f}分" if pd.notna(r.get("basic_goal_0100")) else "-", "戰力指標子分數。")
                    with cc2:
                        metric_card("熟客轉化力", f"{r['convert_goal_0100']:.0f}分" if pd.notna(r.get("convert_goal_0100")) else "-", "戰力指標子分數。")
                        metric_card("業績波動度", f"{r['stability_goal_0100']:.0f}分" if pd.notna(r.get("stability_goal_0100")) else "-", "戰力指標子分數。")
                    with cc3:
                        metric_card("熟客經營力", f"{r['retain_goal_0100']:.0f}分" if pd.notna(r.get("retain_goal_0100")) else "-", "戰力指標子分數。")
        row2 = st.columns(1)
        with row2[0]:
            rank_txt, pct_value_text, tag, bg, color = score_insight(designer_metrics_filtered, "new_ret_goal_0100", r.get("new_ret_goal_0100"))
            with st.container(border=True):
                metric_card(
                    "新客留存力",
                    f"{r['new_ret_goal_0100']:.0f}分" if pd.notna(r.get("new_ret_goal_0100")) else "-",
                    "看「新客回不回來」：新客在 60 天內是否回到同分店的比例（留存率＝1−流失率）。數字越高，代表新客更容易在短期內回來。（注意：是同分店回店，不限定同師傅。）",
                    subtext=pct_value_text if pct_value_text else "",
                    tag_text=tag,
                    tag_bg=bg,
                    tag_color=color,
                    value_suffix=None,
                    meta_text=None,
                    horizontal=True,
                    flat=True,
                )
                with st.expander("構成細項", expanded=False):
                    c1, c2, c3, c4 = st.columns(4)
                    with c1:
                        metric_card(
                            "新客留存率(60天)",
                            f"{r['new_retention_rate_3m']:.2%}" if pd.notna(r.get("new_retention_rate_3m")) else "-",
                            "滿 60 天新客中，60 天內有回店（同分店）的比例。",
                            value_suffix=median_suffix(designer_metrics_filtered, "new_retention_rate_3m", "percent"),
                        )
                    with c2:
                        metric_card(
                            "新客流失率(60天)",
                            f"{r['new_churn_rate_3m']:.2%}" if pd.notna(r.get("new_churn_rate_3m")) else "-",
                            "滿 60 天新客中，60 天內未回店（同分店）的比例。",
                            value_suffix=median_suffix(designer_metrics_filtered, "new_churn_rate_3m", "percent"),
                        )
                    with c3:
                        metric_card(
                            "流失人數(3M)",
                            f"{int(r['new_churned_3m'])}" if pd.notna(r.get("new_churned_3m")) else "-",
                            "上述滿 60 天新客中的流失人數。",
                            value_suffix=median_suffix(designer_metrics_filtered, "new_churned_3m", "number0"),
                        )
                    with c4:
                        metric_card(
                            "留住人數(3M)",
                            f"{int(r['new_retained_3m'])}" if pd.notna(r.get("new_retained_3m")) else "-",
                            "上述滿 60 天新客中的留住人數。",
                            value_suffix=median_suffix(designer_metrics_filtered, "new_retained_3m", "number0"),
                        )

                    if not detail_recent.empty:
                        detail_recent["末三碼"] = detail_recent["phone_key"].apply(mask_last3)
                        detail_recent["首單時間"] = detail_recent["結帳操作時間"]
                        base_cols = ["末三碼", "分店", "首單時間"]
                        if name_col in detail_recent.columns:
                            base_cols.insert(1, name_col)

                        churn_list = detail_recent[detail_recent["churn"]].copy()
                        retained_list = detail_recent[~detail_recent["churn"]].copy()
                        if "return_days_store" in retained_list.columns:
                            retained_list = retained_list.rename(columns={"return_days_store": "回店天數"})
                            retained_cols = base_cols + ["回店天數"]
                        else:
                            retained_cols = base_cols

                        with st.expander("查看流失名單（近3月、滿60天）", expanded=False):
                            if churn_list.empty:
                                st.info("目前沒有流失名單。")
                            else:
                                show_cols = [c for c in base_cols if c in churn_list.columns]
                                st.dataframe(churn_list[show_cols].sort_values("首單時間"), use_container_width=True)

                        with st.expander("查看留住名單（近3月、滿60天）", expanded=False):
                            if retained_list.empty:
                                st.info("目前沒有留住名單。")
                            else:
                                show_cols = [c for c in retained_cols if c in retained_list.columns]
                                st.dataframe(retained_list[show_cols].sort_values("首單時間"), use_container_width=True)
        row3 = st.columns(2)
        with row3[0]:
            rank_txt, pct_value_text, tag, bg, color = score_insight(designer_metrics_filtered, "convert_goal_0100", r.get("convert_goal_0100"))
            with st.container(border=True):
                metric_card(
                    "熟客轉化力",
                    f"{r['convert_goal_0100']:.0f}分" if pd.notna(r.get("convert_goal_0100")) else "-",
                    "看「把客人養成熟客的能力/速度」：同分店同師傅，180 天內是否能累積到 ≥5 次，以及平均多久達到第 5 次。數字越高，代表更容易、也更快把客人養成穩定熟客。",
                    subtext=pct_value_text if pct_value_text else "",
                    tag_text=tag,
                    tag_bg=bg,
                    tag_color=color,
                    value_suffix=None,
                    meta_text=None,
                    horizontal=True,
                    flat=True,
                )
                with st.expander("構成細項", expanded=False):
                    c1, c2, c3, c4 = st.columns(4)
                    with c1:
                        metric_card(
                            "熟客化率",
                            f"{r['regular_rate_180']:.2%}" if pd.notna(r.get("regular_rate_180")) else "-",
                            "同分店同師傅，180 天內消費 ≥5 次的比例。",
                            value_suffix=median_suffix(designer_metrics_filtered, "regular_rate_180", "percent"),
                        )
                    with c2:
                        metric_card(
                            "熟客化樣本數",
                            f"{int(r['regular_base_180'])}" if pd.notna(r.get("regular_base_180")) else "-",
                            "關係起點已滿 180 天的樣本數。",
                            value_suffix=median_suffix(designer_metrics_filtered, "regular_base_180", "number0"),
                        )
                    with c3:
                        metric_card(
                            "熟客達標人數",
                            f"{int(r['regular_achieved_180'])}" if pd.notna(r.get("regular_achieved_180")) else "-",
                            "在 180 天內達成 ≥5 次的人數。",
                            value_suffix=median_suffix(designer_metrics_filtered, "regular_achieved_180", "number0"),
                        )
                    with c4:
                        metric_card(
                            "平均達標天數",
                            f"{r['regular_days_avg_180']:.0f}" if pd.notna(r.get("regular_days_avg_180")) else "-",
                            "達成第 5 次消費的平均天數。",
                            value_suffix=median_suffix(designer_metrics_filtered, "regular_days_avg_180", "number0"),
                        )

                    if not rel_all.empty:
                        rel_all_conv = rel_all[rel_all["regular_matured_180"]].copy()
                        rel_all_conv["regular_achieved"] = rel_all_conv["regular_achieved"].fillna(False)
                        regular_list = rel_all_conv[rel_all_conv["regular_achieved"] == True].copy()
                        in_service_list = rel_all_conv[
                            (rel_all_conv["regular_achieved"] == True)
                            & (rel_all_conv["retention_matured_180"] != True)
                        ].copy()
                        if not regular_list.empty:
                            regular_list["末三碼"] = regular_list["phone_key"].apply(mask_last3)
                            regular_list["關係起點"] = regular_list["baseline_time"]
                            regular_list["熟客達標日"] = regular_list["regular_date"]
                            regular_list["達標天數"] = (regular_list["regular_date"] - regular_list["baseline_time"]).dt.days
                        if not in_service_list.empty:
                            in_service_list["末三碼"] = in_service_list["phone_key"].apply(mask_last3)
                            in_service_list["關係起點"] = in_service_list["baseline_time"]
                            in_service_list["熟客達標日"] = in_service_list["regular_date"]
                            in_service_list["觀察到期日"] = in_service_list["regular_date"] + pd.Timedelta(days=RETENTION_DAYS)
                            in_service_list["剩餘天數"] = (in_service_list["觀察到期日"] - end_date).dt.days.clip(lower=0)
                        base_cols = ["末三碼", "分店", "關係起點", "熟客達標日", "達標天數"]
                        if name_col in rel_all_conv.columns:
                            base_cols.insert(1, name_col)
                        in_service_cols = ["末三碼", "分店", "關係起點", "熟客達標日", "觀察到期日", "剩餘天數"]
                        if name_col in rel_all_conv.columns:
                            in_service_cols.insert(1, name_col)
                        label_count = int(r["regular_achieved_180"]) if pd.notna(r.get("regular_achieved_180")) else 0
                        in_service_count = int(r["in_service_regular_180"]) if pd.notna(r.get("in_service_regular_180")) else 0
                        with st.expander(f"查看熟客名單（熟客達標人數：{label_count}）", expanded=False):
                            if regular_list.empty:
                                st.info("目前沒有熟客達標名單。")
                            else:
                                show_cols = [c for c in base_cols if c in regular_list.columns]
                                st.dataframe(regular_list[show_cols].sort_values("熟客達標日"), use_container_width=True)
                        with st.expander(f"查看經營中熟客名單（經營中熟客：{in_service_count}）", expanded=False):
                            if in_service_list.empty:
                                st.info("目前沒有經營中熟客名單。")
                            else:
                                show_cols = [c for c in in_service_cols if c in in_service_list.columns]
                                st.dataframe(in_service_list[show_cols].sort_values("熟客達標日"), use_container_width=True)
        with row3[1]:
            rank_txt, pct_value_text, tag, bg, color = score_insight(designer_metrics_filtered, "retain_goal_0100", r.get("retain_goal_0100"))
            with st.container(border=True):
                metric_card(
                    "熟客經營力",
                    f"{r['retain_goal_0100']:.0f}分" if pd.notna(r.get("retain_goal_0100")) else "-",
                    "看「熟客養成後能不能維持」：成為熟客後的下一個 180 天內，是否仍有 ≥3 次回訪，以及後 180 天的平均回訪頻率。數字越高，代表熟客更有黏著度、更常回來。",
                    subtext=pct_value_text if pct_value_text else "",
                    tag_text=tag,
                    tag_bg=bg,
                    tag_color=color,
                    value_suffix=None,
                    meta_text=None,
                    horizontal=True,
                    flat=True,
                )
                with st.expander("構成細項", expanded=False):
                    c1, c2, c3, c4 = st.columns(4)
                    with c1:
                        metric_card(
                            "熟客維持率",
                            f"{r['retention_rate_180']:.2%}" if pd.notna(r.get("retention_rate_180")) else "-",
                            "熟客達成後 180 天內回訪 ≥3 次的比例。",
                            value_suffix=median_suffix(designer_metrics_filtered, "retention_rate_180", "percent"),
                        )
                    with c2:
                        metric_card(
                            "熟客維持樣本數",
                            f"{int(r['retention_base_180'])}" if pd.notna(r.get("retention_base_180")) else "-",
                            "熟客達成且後 180 天已滿期的樣本數。",
                            value_suffix=median_suffix(designer_metrics_filtered, "retention_base_180", "number0"),
                        )
                    with c3:
                        metric_card(
                            "熟客維持達標人數",
                            f"{int(r['retention_achieved_180'])}" if pd.notna(r.get("retention_achieved_180")) else "-",
                            "後 180 天回訪 ≥3 次的人數。",
                            value_suffix=median_suffix(designer_metrics_filtered, "retention_achieved_180", "number0"),
                        )
                    with c4:
                        metric_card(
                            "熟客月均回訪次數",
                            f"{r['post_regular_visits_monthly_avg_180']:.2f}" if pd.notna(r.get("post_regular_visits_monthly_avg_180")) else "-",
                            "熟客達成後 180 天內平均回訪次數 / 6 個月。",
                            value_suffix=median_suffix(designer_metrics_filtered, "post_regular_visits_monthly_avg_180", "number1"),
                        )

                    if not rel_all.empty:
                        retention_base = rel_all[(rel_all["regular_achieved"] == True) & (rel_all["retention_matured_180"])].copy()
                        retention_base["retention_achieved"] = retention_base["retention_achieved"].fillna(False)
                        deep_list = retention_base[retention_base["retention_achieved"] == True].copy()
                        churned_deep_list = retention_base[retention_base["retention_achieved"] == False].copy()
                        if not deep_list.empty:
                            deep_list["末三碼"] = deep_list["phone_key"].apply(mask_last3)
                            deep_list["關係起點"] = deep_list["baseline_time"]
                            deep_list["熟客達標日"] = deep_list["regular_date"]
                            deep_list["後180天回訪次數"] = deep_list["post_regular_visits_180"]
                        if not churned_deep_list.empty:
                            churned_deep_list["末三碼"] = churned_deep_list["phone_key"].apply(mask_last3)
                            churned_deep_list["關係起點"] = churned_deep_list["baseline_time"]
                            churned_deep_list["熟客達標日"] = churned_deep_list["regular_date"]
                            churned_deep_list["後180天回訪次數"] = churned_deep_list["post_regular_visits_180"]
                        deep_cols = ["末三碼", "分店", "關係起點", "熟客達標日", "後180天回訪次數"]
                        if name_col in retention_base.columns:
                            deep_cols.insert(1, name_col)
                        label_count = int(r["retention_achieved_180"]) if pd.notna(r.get("retention_achieved_180")) else 0
                        churned_label_count = int(r["retention_base_180"] - r["retention_achieved_180"]) if pd.notna(r.get("retention_base_180")) and pd.notna(r.get("retention_achieved_180")) else 0
                        with st.expander(f"查看深度熟客名單（熟客維持達標人數：{label_count}）", expanded=False):
                            if deep_list.empty:
                                st.info("目前沒有深度熟客名單。")
                            else:
                                show_cols = [c for c in deep_cols if c in deep_list.columns]
                                st.dataframe(deep_list[show_cols].sort_values("熟客達標日"), use_container_width=True)
                        with st.expander(f"查看流失熟客名單（熟客維持未達標人數：{churned_label_count}）", expanded=False):
                            if churned_deep_list.empty:
                                st.info("目前沒有流失熟客名單。")
                            else:
                                show_cols = [c for c in deep_cols if c in churned_deep_list.columns]
                                st.dataframe(churned_deep_list[show_cols].sort_values("熟客達標日"), use_container_width=True)
        row4 = st.columns(2)
        with row4[0]:
            rank_txt, pct_value_text, tag, bg, color = score_insight(designer_metrics_filtered, "basic_goal_0100", r.get("basic_goal_0100"))
            with st.container(border=True):
                metric_card(
                    "合作穩定度",
                    f"{r['basic_goal_0100']:.0f}分" if pd.notna(r.get("basic_goal_0100")) else "-",
                    "看這位師傅最近是否「正常有在上班、接單、排班/客量是否穩定」：近 3 個月的有單天數、總單量、空窗率等組合。數字越高，代表近期更穩定。",
                    subtext=pct_value_text if pct_value_text else "",
                    tag_text=tag,
                    tag_bg=bg,
                    tag_color=color,
                    value_suffix=None,
                    meta_text=None,
                    horizontal=True,
                    flat=True,
                )
                with st.expander("構成細項", expanded=False):
                    c1, c2, c3, c4 = st.columns(4)
                    with c1:
                        metric_card(
                            "每月平均有單天數(近3月)",
                            f"{r['avg_active_days_3m']:.1f}" if pd.notna(r.get("avg_active_days_3m")) else "-",
                            "近 3 個月每月平均有單的天數。",
                            value_suffix=median_suffix(designer_metrics_filtered, "avg_active_days_3m", "number1"),
                        )
                    with c2:
                        metric_card(
                            "近3月有單天數",
                            f"{int(r['active_days_3m'])}" if pd.notna(r.get("active_days_3m")) else "-",
                            "近 3 個月內所有有單天數的加總。",
                            value_suffix=median_suffix(designer_metrics_filtered, "active_days_3m", "number0"),
                        )
                    with c3:
                        metric_card(
                            "總單量(3M)",
                            f"{int(r['total_orders_3m'])}" if pd.notna(r.get("total_orders_3m")) else "-",
                            "近 3 個月的訂單總數。",
                            value_suffix=median_suffix(designer_metrics_filtered, "total_orders_3m", "number0"),
                        )
                    with c4:
                        metric_card(
                            "空窗率(3M)",
                            f"{r['vacancy_rate_3m']:.2%}" if pd.notna(r.get("vacancy_rate_3m")) else "-",
                            "近 3 個月平均空窗率（1 - 服務時數/168 小時）。",
                            value_suffix=median_suffix(designer_metrics_filtered, "vacancy_rate_3m", "percent"),
                        )
        with row4[1]:
            rank_txt, pct_value_text, tag, bg, color = score_insight(designer_metrics_filtered, "stability_goal_0100", r.get("stability_goal_0100"))
            with st.container(border=True):
                metric_card(
                    "業績波動度",
                    f"{r['stability_goal_0100']:.0f}分" if pd.notna(r.get("stability_goal_0100")) else "-",
                    "看「工作量起伏大不大」：近 6 個月每月工時（或有單天數）的波動程度，越穩定分數越高。數字越高，代表月與月之間更穩定。",
                    subtext=pct_value_text if pct_value_text else "",
                    tag_text=tag,
                    tag_bg=bg,
                    tag_color=color,
                    value_suffix=None,
                    meta_text=None,
                    horizontal=True,
                    flat=True,
                )
                with st.expander("構成細項", expanded=False):
                    c1, _, _, _ = st.columns(4)
                    if pd.notna(r.get("service_hours_cv_6m")):
                        with c1:
                            metric_card(
                                "業績穩定度(工時CV)",
                                f"{r['service_hours_cv_6m']:.2f}",
                                "近 6 個月服務時數的變異係數（CV），越低越穩定。",
                                value_suffix=median_suffix(designer_metrics_filtered, "service_hours_cv_6m", "number1"),
                            )
                    else:
                        with c1:
                            metric_card(
                                "業績穩定度(出勤CV)",
                                f"{r['active_days_cv_6m']:.2f}" if pd.notna(r.get("active_days_cv_6m")) else "-",
                                "近 6 個月出勤天數的變異係數（CV），越低越穩定。",
                                value_suffix=median_suffix(designer_metrics_filtered, "active_days_cv_6m", "number1"),
                            )
        st.caption("熟客化/熟客維持皆以「同分店、同師傅」計算；業績穩定度 CV 越低代表越穩定。")

        st.subheader("差距地圖（以選定師傅為中心）")
        st.caption("中心點為選定師傅 (0,0)；其餘點表示與他在 X/Y 指標上的差距。")

        map_defs = [
            {
                "name": "穩定性地圖",
                "x": "basic_goal_0100",
                "y": "stability_goal_0100",
                "x_label": "基本狀態差距(分數)",
                "y_label": "業績穩定度差距(分數)",
                "x_fmt": "number1",
                "y_fmt": "number1",
            },
            {
                "name": "新客成長地圖",
                "x": "new_per_active_day_3m",
                "y": "new_retention_rate_3m",
                "x_label": "新客/有單天數差距",
                "y_label": "新客留存率差距(%)",
                "x_fmt": "number1",
                "y_fmt": "percent",
            },
            {
                "name": "新客結構地圖",
                "x": "new_share_3m",
                "y": "new_retention_rate_3m",
                "x_label": "新客占比差距(%)",
                "y_label": "新客留存率差距(%)",
                "x_fmt": "percent",
                "y_fmt": "percent",
            },
            {
                "name": "熟客策略地圖",
                "x": "regular_rate_180",
                "y": "retention_rate_180",
                "x_label": "熟客化率差距(%)",
                "y_label": "熟客維持率差距(%)",
                "x_fmt": "percent",
                "y_fmt": "percent",
            },
            {
                "name": "熟客深度地圖",
                "x": "retention_rate_180",
                "y": "post_regular_visits_monthly_avg_180",
                "x_label": "熟客維持率差距(%)",
                "y_label": "熟客月均回訪次數差距",
                "x_fmt": "percent",
                "y_fmt": "number1",
            },
            {
                "name": "熟客速度地圖",
                "x": "regular_rate_180",
                "y": "regular_days_avg_180",
                "x_label": "熟客化率差距(%)",
                "y_label": "達標速度差距(天，+代表更快)",
                "x_fmt": "percent",
                "y_fmt": "number0",
                "y_reverse": True,
            },
        ]

        map_tabs = st.tabs([m["name"] for m in map_defs])
        for tab, cfg in zip(map_tabs, map_defs):
            with tab:
                chart = build_delta_map(
                    designer_metrics_filtered,
                    designer_select,
                    cfg["x"],
                    cfg["y"],
                    cfg["x_label"],
                    cfg["y_label"],
                    cfg["x_fmt"],
                    cfg["y_fmt"],
                    x_reverse=cfg.get("x_reverse", False),
                    y_reverse=cfg.get("y_reverse", False),
                )
                if chart is None:
                    st.info("此地圖目前沒有足夠資料。")
                else:
                    st.altair_chart(chart, use_container_width=True)

        # 成長趨勢（多指標）
        st.subheader("成長趨勢")
        new_cohort = new_first_store[new_first_store["設計師"] == designer_select].copy()
        new_cohort["cohort_month"] = new_cohort["結帳操作時間"].dt.to_period("M").dt.to_timestamp()
        new_cohort_matured = new_cohort[new_cohort["matured"]].copy()
        new_churn_trend = (
            new_cohort_matured.groupby("cohort_month")
            .agg(rate=("churn", "mean"), n=("churn", "count"))
            .reset_index()
        )
        new_churn_trend["指標"] = "新客流失率(60天)"

        rel_cohort = relationship_first[relationship_first["設計師"] == designer_select].copy()
        if not rel_cohort.empty:
            rel_cohort = rel_cohort[rel_cohort["baseline_time"] >= start_ts_12m]
        if not rel_cohort.empty:
            rel_cohort["cohort_month"] = rel_cohort["baseline_time"].dt.to_period("M").dt.to_timestamp()
            regular_trend = (
                rel_cohort[rel_cohort["regular_matured_180"]]
                .groupby("cohort_month")
                .agg(rate=("regular_achieved", "mean"), n=("regular_achieved", "count"))
                .reset_index()
            )
            regular_trend["指標"] = "熟客化率(180天達5次)"

            retention_trend = (
                rel_cohort[(rel_cohort["regular_achieved"] == True) & (rel_cohort["retention_matured_180"])]
                .groupby("cohort_month")
                .agg(rate=("retention_achieved", "mean"), n=("retention_achieved", "count"))
                .reset_index()
            )
            retention_trend["指標"] = "熟客維持率(後180天≥3次)"

            trend_long = pd.concat([new_churn_trend, regular_trend, retention_trend], ignore_index=True)
        else:
            trend_long = new_churn_trend.copy()

        # 業績穩定度(工時CV)加入成長趨勢（Z-score）
        cv_trend = None
        if vacancy_monthly is not None:
            cv_monthly = vacancy_monthly.copy()
            cv_monthly = cv_monthly.rename(columns={"duration_hours": "service_hours"})
            cv_monthly["month_start"] = pd.to_datetime(cv_monthly["month"] + "-01")
            cv_monthly = cv_monthly[cv_monthly["month_start"] >= start_ts_6m]
            cv_monthly = cv_monthly[cv_monthly["設計師"] == designer_select]
            if not cv_monthly.empty:
                cv_monthly = cv_monthly.sort_values("month_start")
                cv_series = cv_monthly["service_hours"].rolling(window=6, min_periods=2).apply(
                    lambda s: s.std() / s.mean() if s.mean() else np.nan,
                    raw=False,
                )
                cv_monthly["cv_6m"] = cv_series.values
                cv_trend = cv_monthly[["month_start", "cv_6m"]].dropna().rename(
                    columns={"month_start": "cohort_month", "cv_6m": "rate"}
                )
                cv_trend["指標"] = "業績穩定度(工時CV)"

        trend_for_z = trend_long.copy()
        if cv_trend is not None and not cv_trend.empty:
            trend_for_z = pd.concat([trend_for_z, cv_trend], ignore_index=True)

        if not trend_for_z.empty:
            trend_for_z["z"] = trend_for_z.groupby("指標")["rate"].transform(
                lambda s: (s - s.mean()) / s.std() if s.std() and not np.isnan(s.std()) else np.nan
            )
            z_chart = trend_for_z.dropna(subset=["z"])
            if not z_chart.empty:
                line_z = alt.Chart(z_chart).mark_line(point=True).encode(
                    x=alt.X("cohort_month:T", title="月份"),
                    y=alt.Y("z:Q", title="Z-score"),
                    color=alt.Color("指標:N", title=""),
                    tooltip=["cohort_month:T", "指標", alt.Tooltip("z:Q", format=".2f")],
                ).properties(height=280)
                st.caption("成長趨勢（Z-score 標準化後合併）")
                st.altair_chart(line_z, use_container_width=True)
            else:
                st.info("目前沒有足夠的趨勢資料可視覺化。")
        else:
            st.info("目前沒有足夠的趨勢資料可視覺化。")

st.subheader("師傅排行榜（Top 6）")
st.caption("新客流失/空窗/總單量以近 3 個月計；熟客化/維持以近 12 個月 cohort 且滿 180 天計。")
col_good, col_watch = st.columns(2)
has_hours_cv = "service_hours_cv_6m" in designer_metrics_filtered.columns and designer_metrics_filtered["service_hours_cv_6m"].notna().any()

with col_good:
    st.markdown("<span style='color:#2ca02c;font-weight:700;'>表現較佳</span>", unsafe_allow_html=True)
    render_rank_bar(
        designer_metrics_filtered.dropna(subset=["new_churn_rate_3m"]),
        "設計師",
        "new_churn_rate_3m",
        "新客流失率最低(60天)",
        ascending=True,
        value_format="percent",
        color="#2ca02c",
    )
    render_rank_bar(
        designer_metrics_filtered.dropna(subset=["regular_rate_180"]),
        "設計師",
        "regular_rate_180",
        "熟客化率最高(180天達5次)",
        ascending=False,
        value_format="percent",
        color="#2ca02c",
    )
    render_rank_bar(
        designer_metrics_filtered.dropna(subset=["retention_rate_180"]),
        "設計師",
        "retention_rate_180",
        "熟客維持率最高(後180天≥3次)",
        ascending=False,
        value_format="percent",
        color="#2ca02c",
    )
    if "request_rate_3m" in designer_metrics_filtered.columns:
        render_rank_bar(
            designer_metrics_filtered.dropna(subset=["request_rate_3m"]),
            "設計師",
            "request_rate_3m",
            "指定率最高(3M)",
            ascending=False,
            value_format="percent",
            color="#2ca02c",
        )
    if vacancy_recent is None:
        st.info("無法計算空窗率（缺少項目分鐘）。")
    else:
        render_rank_bar(
            designer_metrics_filtered.dropna(subset=["vacancy_rate_3m"]),
            "設計師",
            "vacancy_rate_3m",
            "空窗率最低(3M)",
            ascending=True,
            value_format="percent",
            color="#2ca02c",
        )
    if has_hours_cv:
        render_rank_bar(
            designer_metrics_filtered.dropna(subset=["service_hours_cv_6m"]),
            "設計師",
            "service_hours_cv_6m",
            "業績穩定度最佳(工時CV最低)",
            ascending=True,
            value_format="number1",
            color="#2ca02c",
        )
    else:
        render_rank_bar(
            designer_metrics_filtered.dropna(subset=["active_days_cv_6m"]),
            "設計師",
            "active_days_cv_6m",
            "業績穩定度最佳(出勤CV最低)",
            ascending=True,
            value_format="number1",
            color="#2ca02c",
        )

with col_watch:
    st.markdown("<span style='color:#d62728;font-weight:700;'>需關注</span>", unsafe_allow_html=True)
    render_rank_bar(
        designer_metrics_filtered.dropna(subset=["new_churn_rate_3m"]),
        "設計師",
        "new_churn_rate_3m",
        "新客流失率最高(60天)",
        ascending=False,
        value_format="percent",
        color="#d62728",
    )
    render_rank_bar(
        designer_metrics_filtered.dropna(subset=["regular_rate_180"]),
        "設計師",
        "regular_rate_180",
        "熟客化率最低(180天達5次)",
        ascending=True,
        value_format="percent",
        color="#d62728",
    )
    render_rank_bar(
        designer_metrics_filtered.dropna(subset=["retention_rate_180"]),
        "設計師",
        "retention_rate_180",
        "熟客維持率最低(後180天≥3次)",
        ascending=True,
        value_format="percent",
        color="#d62728",
    )
    if "request_rate_3m" in designer_metrics_filtered.columns:
        render_rank_bar(
            designer_metrics_filtered.dropna(subset=["request_rate_3m"]),
            "設計師",
            "request_rate_3m",
            "指定率最低(3M)",
            ascending=True,
            value_format="percent",
            color="#d62728",
        )
    if vacancy_recent is None:
        st.info("無法計算空窗率（缺少項目分鐘）。")
    else:
        render_rank_bar(
            designer_metrics_filtered.dropna(subset=["vacancy_rate_3m"]),
            "設計師",
            "vacancy_rate_3m",
            "空窗率最高(3M)",
            ascending=False,
            value_format="percent",
            color="#d62728",
        )
    if has_hours_cv:
        render_rank_bar(
            designer_metrics_filtered.dropna(subset=["service_hours_cv_6m"]),
            "設計師",
            "service_hours_cv_6m",
            "業績穩定度較弱(工時CV最高)",
            ascending=False,
            value_format="number1",
            color="#d62728",
        )
    else:
        render_rank_bar(
            designer_metrics_filtered.dropna(subset=["active_days_cv_6m"]),
            "設計師",
            "active_days_cv_6m",
            "業績穩定度較弱(出勤CV最高)",
            ascending=False,
            value_format="number1",
            color="#d62728",
        )

st.subheader("回店時間分布（同分店）")
ret_series = filtered_new_first["return_days_store"].dropna() if "return_days_store" in filtered_new_first.columns else pd.Series([], dtype=float)
if ret_series.empty:
    st.info("目前沒有可計算的回訪資料。")
else:
    p50 = ret_series.quantile(0.5)
    p70 = ret_series.quantile(0.7)
    p80 = ret_series.quantile(0.8)
    p90 = ret_series.quantile(0.9)

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("P50(天)", f"{p50:.0f}")
    c2.metric("P70(天)", f"{p70:.0f}")
    c3.metric("P80(天)", f"{p80:.0f}")
    c4.metric("P90(天)", f"{p90:.0f}")

    hist_df = pd.DataFrame({"回店天數": ret_series})
    hist = alt.Chart(hist_df).mark_bar(color="#4e79a7").encode(
        x=alt.X("回店天數:Q", bin=alt.Bin(step=7), title="回店天數(天)"),
        y=alt.Y("count():Q", title="人數"),
        tooltip=[alt.Tooltip("count()", title="人數")],
    ).properties(height=260)

    p_df = pd.DataFrame({"x": [p70, p80], "label": ["P70", "P80"]})
    rules = alt.Chart(p_df).mark_rule(color="#f28e2b").encode(x="x:Q")
    labels = alt.Chart(p_df).mark_text(align="left", dx=4, dy=-4, color="#f28e2b").encode(
        x="x:Q", y=alt.value(0), text="label:N"
    )
    st.altair_chart(hist + rules + labels, use_container_width=True)

st.subheader("詳細表格")
st.caption("依照你的篩選條件，以下是完整明細表格。")
store_table = None
store_designer_table = None
display_vacancy = None

# 師傅彙總
st.markdown("**師傅彙總（新客近 3 月 / 熟客化近 12 月）**")
if designer_metrics_filtered.empty:
    st.info("目前沒有足夠的近 3 月資料。")
    designer_table = None
else:
    designer_table = designer_metrics_filtered.copy()
    service_cv = designer_table["service_hours_cv_6m"] if "service_hours_cv_6m" in designer_table.columns else pd.Series(np.nan, index=designer_table.index)
    active_cv = designer_table["active_days_cv_6m"] if "active_days_cv_6m" in designer_table.columns else pd.Series(np.nan, index=designer_table.index)
    designer_table["業績穩定度(CV)"] = np.where(service_cv.notna(), service_cv, active_cv)
    designer_table = designer_table.rename(
        columns={
            "設計師": "師傅",
            "overall_goal_0100": "戰力指標(0-100)",
            "basic_goal_0100": "基本狀態(0-100)",
            "new_acq_goal_0100": "新客獲取量(0-100)",
            "new_ret_goal_0100": "新客留存力(0-100)",
            "convert_goal_0100": "熟客轉化力(0-100)",
            "retain_goal_0100": "熟客經營力(0-100)",
            "stability_goal_0100": "業績穩定度(0-100)",
            "new_share_3m": "新客占比(新客/總單量)",
            "new_per_active_day_3m": "新客/有單天數(3M)",
            "new_customers_3m": "新客數(3M,滿60天)",
            "new_churn_rate_3m": "新客流失率(60天)",
            "new_retention_rate_3m": "新客留存率(60天)",
            "total_orders_3m": "總單量(3M)",
            "new_churned_3m": "流失人數(3M)",
            "new_retained_3m": "留住人數(3M)",
            "new_repeat_rate_3m": "新客回指率(30天)",
            "new_repeat_base_3m": "新客回指樣本數(30天)",
            "new_deep_rate_3m": "新客深度回指率(60天)",
            "new_deep_n": "新客深度回指樣本數(60天)",
            "familiar_repeat_rate_3m": "熟客回指率(30天)",
            "familiar_customers_3m": "熟客回指樣本數(30天)",
            "familiar_deep_rate_3m": "熟客深度回指率(60天)",
            "familiar_deep_n": "熟客深度回指樣本數(60天)",
            "avg_active_days_3m": "每月平均有單天數(近3月)",
            "active_days_3m": "近3月有單天數",
            "regular_rate_180": "熟客化率(180天達5次)",
            "regular_base_180": "熟客化樣本數(滿180天)",
            "regular_achieved_180": "熟客化達標人數",
            "regular_days_avg_180": "平均達標天數",
            "retention_rate_180": "熟客維持率(後180天≥3次)",
            "retention_base_180": "熟客維持樣本數(滿後180天)",
            "retention_achieved_180": "熟客維持達標人數",
            "post_regular_visits_avg_180": "後180天平均回訪次數",
            "post_regular_visits_monthly_avg_180": "熟客月均回訪次數(後180天)",
            "request_rate_3m": "指定率(3M)",
            "request_yes_3m": "指定單數(3M)",
            "request_total_3m": "總單數(3M)",
            "vacancy_rate_3m": "空窗率(3M)",
            "days_since_last_tx": "最近有單距今(天)",
        }
    )
    cols = [
        "師傅",
        "戰力指標(0-100)",
        "基本狀態(0-100)",
        "新客獲取量(0-100)",
        "新客留存力(0-100)",
        "熟客轉化力(0-100)",
        "熟客經營力(0-100)",
        "業績穩定度(0-100)",
        "新客占比(新客/總單量)",
        "新客/有單天數(3M)",
        "新客數(3M,滿60天)",
        "新客流失率(60天)",
        "新客留存率(60天)",
        "流失人數(3M)",
        "留住人數(3M)",
        "新客回指率(30天)",
        "新客回指樣本數(30天)",
        "新客深度回指率(60天)",
        "新客深度回指樣本數(60天)",
        "熟客回指率(30天)",
        "熟客回指樣本數(30天)",
        "熟客深度回指率(60天)",
        "熟客深度回指樣本數(60天)",
        "總單量(3M)",
        "每月平均有單天數(近3月)",
        "近3月有單天數",
        "熟客化率(180天達5次)",
        "熟客化樣本數(滿180天)",
        "熟客化達標人數",
        "平均達標天數",
        "熟客維持率(後180天≥3次)",
        "熟客維持樣本數(滿後180天)",
        "熟客維持達標人數",
        "後180天平均回訪次數",
        "熟客月均回訪次數(後180天)",
        "指定率(3M)",
        "指定單數(3M)",
        "總單數(3M)",
        "空窗率(3M)",
        "業績穩定度(CV)",
    ]
    cols = [c for c in cols if c in designer_table.columns]
    st.dataframe(designer_table[cols].sort_values("新客流失率(60天)", ascending=True), use_container_width=True)
    st.caption("業績穩定度：若缺少項目分鐘，將以出勤 CV 代替工時 CV。")

# 分店彙總
if has_store:
    st.markdown("**分店彙總**")
    store_table = summary_by_store.rename(
        columns={
            "分店": "分店",
            "matured_new_customers": "滿60天新客數",
            "churned": "流失數",
            "churn_rate": "流失率",
            "repeat_rate": "回店率",
        }
    )
    st.dataframe(store_table.sort_values("流失率", ascending=False), use_container_width=True)

    st.markdown("**分店 x 師傅彙總**")
    store_designer_table = summary_by_store_designer.rename(
        columns={
            "分店": "分店",
            "設計師": "師傅",
            "matured_new_customers": "滿60天新客數",
            "churned": "流失數",
            "churn_rate": "流失率",
            "repeat_rate": "回店率",
        }
    )
    st.dataframe(store_designer_table.sort_values("流失率", ascending=False), use_container_width=True)

# 空窗率（月，168 小時上限）
st.markdown("**空窗率（月，168 小時上限）**")
if vacancy_monthly is None:
    st.warning("帳單檔缺少 '項目' 欄位，無法估算服務時數與空窗率。")
else:
    display_vacancy = vacancy_monthly.copy()
    if has_store and store_filter is not None:
        display_vacancy = display_vacancy[display_vacancy["分店"].isin(store_filter)]
    display_vacancy = display_vacancy[display_vacancy["設計師"].isin(designer_filter)]
    display_vacancy = display_vacancy.rename(
        columns={
            "分店": "分店",
            "設計師": "師傅",
            "month": "月份",
            "duration_hours": "服務時數",
            "vacancy_rate": "空窗率",
        }
    )
    st.dataframe(display_vacancy.sort_values(["月份", "師傅"]), use_container_width=True)

# 流失名單
st.markdown("**流失名單**")
show_churned_only = st.checkbox("只看流失者", value=True)

if show_churned_only:
    detail = filtered_new_first[filtered_new_first["churn"]].copy()
else:
    detail = filtered_new_first.copy()

if has_store and store_filter is not None:
    detail = detail[detail["分店"].isin(store_filter)]
detail = detail[detail["設計師"].isin(designer_filter)]

detail_display = detail.copy()
detail_display["是否滿期"] = detail_display["matured"].map({True: "是", False: "否"})
detail_display["是否流失"] = detail_display["churn"].map({True: "是", False: "否"})
detail_display["電話"] = detail_display["phone_key"]
detail_display = detail_display.rename(columns={"設計師": "師傅", "結帳操作時間": "首單時間"})

display_cols = ["電話", name_col, "分店", "師傅", "首單時間", "是否滿期", "是否流失"]
display_cols = [c for c in display_cols if c in detail_display.columns]
st.dataframe(detail_display[display_cols].sort_values("首單時間"), use_container_width=True)

# Download Excel
st.subheader("下載報表")

output = BytesIO()
overall_df = pd.DataFrame([{
    "滿60天新客數": overall["matured_new_customers"],
    "流失人數": overall["churned_matured"],
    "留住人數": overall["retained_matured"],
    "流失率": overall["churn_rate_matured"],
    "回店率": overall["repeat_rate_matured"],
}])
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    overall_df.to_excel(writer, index=False, sheet_name="總覽")
    if designer_table is not None:
        designer_table.to_excel(writer, index=False, sheet_name="師傅彙總(3M)")
    if store_table is not None:
        store_table.to_excel(writer, index=False, sheet_name="分店彙總")
    if store_designer_table is not None:
        store_designer_table.to_excel(writer, index=False, sheet_name="分店師傅彙總")
    detail_display[display_cols].to_excel(writer, index=False, sheet_name="流失名單")
    if display_vacancy is not None:
        display_vacancy.to_excel(writer, index=False, sheet_name="空窗率(月)")

st.download_button(
    label="下載 Excel",
    data=output.getvalue(),
    file_name="customer_relationship_analysis.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
