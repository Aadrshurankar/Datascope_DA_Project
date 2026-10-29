from __future__ import annotations

from io import BytesIO
from pathlib import Path
import html
import re

import pandas as pd
import plotly.express as px
import streamlit as st

APP_TITLE = "DataScope | Data Analytics Workspace"
SUPPORTED_FILE_TYPES = ["csv", "xlsx", "xls"]

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ---------- Visual theme ----------
st.markdown(
    """
    <style>
      :root {
        --ink:#101b3a; --muted:#61708f; --line:#e0e5f7;
        --purple:#7c3aed; --pink:#ec4899; --blue:#3b82f6;
        --green:#16a765; --gold:#f59e0b;
      }
      html, body, [class*="css"] {
        font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      }
      .stApp {
        color:var(--ink);
        background:
          radial-gradient(ellipse at 14% 3%, rgba(255,199,233,.68), transparent 32%),
          radial-gradient(ellipse at 88% 4%, rgba(164,203,255,.75), transparent 31%),
          linear-gradient(140deg,#fff8fc 0%,#f5f3ff 54%,#f1f7ff 100%);
      }
      [data-testid="stAppViewContainer"] > .main { background:transparent; }
      .block-container { max-width:1540px; padding-top:1.15rem; padding-bottom:2.5rem; }
      header[data-testid="stHeader"] { background:transparent; }
      #MainMenu, footer { visibility:hidden; }
      [data-testid="stSidebar"] { display:none; }

      .topbar { display:flex; align-items:center; justify-content:space-between; gap:1rem; margin:0 0 1.15rem; }
      .brand { display:flex; align-items:center; gap:.8rem; min-width:0; }
      .brand-icon { display:flex; align-items:center; justify-content:center; width:48px; height:48px; flex-shrink:0; border-radius:16px; color:white; background:linear-gradient(145deg,#d946ef 0%,#8b5cf6 48%,#38bdf8 100%); box-shadow:0 8px 22px rgba(124,58,237,.22); font-size:1.05rem; font-weight:950; letter-spacing:-.06em; }
      .brand-name { margin:0; color:#101b3a; font-size:1.8rem; line-height:1.02; font-weight:950; letter-spacing:-.065em; }
      .brand-name span { background:linear-gradient(90deg,#7c3aed,#d946ef); -webkit-background-clip:text; background-clip:text; color:transparent; }
      .brand-sub { margin:.28rem 0 0; color:#586887; font-size:.86rem; font-weight:550; }
      .top-pill { border:1px solid rgba(255,255,255,.9); background:rgba(255,255,255,.58); padding:.7rem 1rem; border-radius:999px; color:#2e3b5c; font-size:.84rem; font-weight:700; white-space:nowrap; box-shadow:0 6px 22px rgba(82,99,150,.06); }
      .top-pill span { color:#c026d3; margin-right:.4rem; }

      [data-testid="stVerticalBlockBorderWrapper"] { border:1px solid rgba(211,219,244,.96) !important; border-radius:22px !important; background:rgba(255,255,255,.78) !important; box-shadow:0 12px 35px rgba(51,64,120,.07); }
      .card-heading { display:flex; align-items:center; gap:.75rem; margin:0 0 .25rem; }
      .heading-icon { display:flex; align-items:center; justify-content:center; flex-shrink:0; width:58px; height:58px; border-radius:17px; font-size:1.55rem; }
      .heading-icon.pink { color:#c026d3; background:linear-gradient(140deg,#fce7f3,#f5d0fe); border:1px solid #f5b8ec; }
      .heading-icon.blue { color:#2563eb; background:linear-gradient(140deg,#e0f2fe,#dbeafe); border:1px solid #b6dcff; }
      .card-title { margin:0; color:#101b3a; font-size:1.25rem; line-height:1.2; font-weight:900; letter-spacing:-.035em; }
      .card-title .accent { color:#7c3aed; }
      .card-copy { color:#61708f; font-size:.86rem; line-height:1.55; margin:.26rem 0 0; }
      .upload-zone { margin:.9rem 0 .45rem; border:1.5px dashed #b79aff; border-radius:17px; padding:.45rem .55rem; background:linear-gradient(135deg,rgba(255,249,255,.94),rgba(246,244,255,.94)); }
      div[data-testid="stFileUploader"] { background:linear-gradient(135deg,rgba(255,249,255,.94),rgba(246,244,255,.94)); border:1.5px dashed #b79aff; border-radius:17px; padding:.45rem; margin:.9rem 0 .45rem; }
      div[data-testid="stFileUploader"] section { border:0; background:rgba(255,255,255,.78); border-radius:12px; }
      div[data-testid="stFileUploader"] section > div { padding:.55rem .55rem; }
      div[data-testid="stFileUploader"] button { border-radius:10px; border:1px solid #d7c8ff; color:#5b21b6; background:#fff; font-weight:800; }
      .upload-note { color:#61708f; font-size:.77rem; line-height:1.5; margin-top:.45rem; }
      .file-pill { display:inline-flex; align-items:center; gap:.5rem; margin:.8rem 0 0; border:1px solid #e0e7ff; background:#fff; border-radius:12px; padding:.55rem .7rem; color:#182544; font-size:.8rem; font-weight:750; max-width:100%; }
      .file-pill small { display:block; color:#74819d; font-weight:550; margin-top:.1rem; }
      .file-check { width:22px; height:22px; display:inline-flex; align-items:center; justify-content:center; border-radius:50%; background:#dcfce7; color:#16a765; font-size:.8rem; }

      .overview-head { display:flex; align-items:center; justify-content:space-between; gap:.8rem; margin-bottom:.85rem; }
      .overview-title { color:#101b3a; font-size:1.25rem; font-weight:900; letter-spacing:-.035em; margin:0; }
      .overview-copy { color:#61708f; font-size:.82rem; margin:.2rem 0 0; }
      .summary-grid { display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); gap:.5rem; margin:.7rem 0 .85rem; }
      .summary-card { min-width:0; border:1px solid rgba(226,232,250,.82); border-radius:15px; padding:.7rem .68rem .62rem; background:rgba(255,255,255,.74); }
      .summary-card .summary-label { color:#33415f; font-size:.73rem; line-height:1.25; min-height:1.8rem; }
      .summary-card .summary-value { color:#0f1b38; font-size:1.42rem; font-weight:950; letter-spacing:-.045em; margin-top:.12rem; }
      .summary-card .summary-icon { display:flex; align-items:center; justify-content:center; width:37px; height:37px; border-radius:11px; font-size:1.05rem; margin-bottom:.48rem; }
      .tone-blue { background:linear-gradient(140deg,#eff6ff,#e0f2fe); }
      .tone-pink { background:linear-gradient(140deg,#fff1f7,#fce7f3); }
      .tone-gold { background:linear-gradient(140deg,#fffbeb,#fef3c7); }
      .tone-green { background:linear-gradient(140deg,#f0fdf4,#dcfce7); }
      .tone-purple { background:linear-gradient(140deg,#faf5ff,#f3e8ff); }
      .icon-blue { color:#1685e5; background:#dbeafe; }
      .icon-pink { color:#db2777; background:#fce7f3; }
      .icon-gold { color:#d97706; background:#fef3c7; }
      .icon-green { color:#16a765; background:#dcfce7; }
      .icon-purple { color:#9333ea; background:#f3e8ff; }
      .basic-title { display:flex; align-items:center; gap:.8rem; color:#15213f; font-weight:850; font-size:.91rem; margin:.75rem 0 .6rem; }
      .basic-title:after { content:""; height:1px; flex:1; background:#e0e5f7; }
      .info-grid { display:grid; grid-template-columns:repeat(3,minmax(0,1fr)); gap:.7rem .8rem; }
      .info-item { display:flex; align-items:center; gap:.55rem; min-width:0; }
      .info-icon { width:38px; height:38px; border-radius:12px; display:flex; align-items:center; justify-content:center; flex-shrink:0; color:#2563eb; background:linear-gradient(140deg,#e0f2fe,#dbeafe); font-size:1rem; }
      .info-label { color:#52617e; font-size:.73rem; margin-bottom:.12rem; }
      .info-value { color:#182544; font-size:.82rem; font-weight:750; overflow-wrap:anywhere; }
      .empty-overview { min-height:300px; display:flex; flex-direction:column; justify-content:center; }
      .empty-overview .empty-symbol { display:flex; width:58px; height:58px; border-radius:18px; align-items:center; justify-content:center; margin-bottom:1rem; background:linear-gradient(135deg,#dbeafe,#f5d0fe); color:#7c3aed; font-size:1.6rem; }
      .empty-overview h3 { color:#101b3a; margin:0 0 .35rem; font-size:1.2rem; font-weight:900; }
      .empty-overview p { color:#61708f; font-size:.88rem; line-height:1.65; margin:0 0 .5rem; }
      .empty-overview .empty-tip { color:#7c3aed; font-size:.8rem; font-weight:750; margin-top:.4rem; }
      .loaded-banner { border:1px solid #bdebd0; background:linear-gradient(90deg,#f0fdf4,#eff6ff); color:#168052; border-radius:13px; padding:.7rem .9rem; margin:1.1rem 0 .75rem; font-weight:800; font-size:.85rem; }

      div[data-testid="stTabs"] { margin-top:1.05rem; }
      div[data-testid="stTabs"] > div[role="tablist"] { gap:.15rem; padding:.25rem; border:1px solid rgba(255,255,255,.9); border-radius:16px; background:rgba(255,255,255,.65); box-shadow:0 7px 24px rgba(51,64,120,.05); }
      div[data-testid="stTabs"] button[role="tab"] { border-radius:11px; font-weight:800; color:#415174; padding:.75rem .8rem; }
      div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] { color:white; background:linear-gradient(100deg,#d946ef,#635bff); }
      div[data-testid="stTabs"] button[role="tab"] p { font-size:.82rem; }
      div[data-testid="stTabContent"] { border:1px solid rgba(211,219,244,.96); border-radius:20px; background:rgba(255,255,255,.82); box-shadow:0 12px 35px rgba(51,64,120,.05); padding:1.2rem 1.15rem; margin-top:.55rem; }
      .section-title { color:#101b3a; font-size:1.2rem; line-height:1.2; font-weight:900; letter-spacing:-.035em; margin:0; }
      .section-copy { color:#61708f; font-size:.83rem; margin:.24rem 0 .9rem; }
      div[data-testid="stDataFrame"] { border:1px solid #e0e5f7; border-radius:13px; overflow:hidden; }
      div[data-testid="stMetric"] { background:linear-gradient(145deg,#fff 0%,#fbfaff 100%); border:1px solid #e0e5f7; border-radius:15px; padding:.75rem; box-shadow:0 4px 15px rgba(51,64,120,.035); }
      div[data-testid="stMetricLabel"] { color:#61708f; font-size:.78rem; font-weight:750; }
      div[data-testid="stMetricValue"] { color:#101b3a; font-weight:950; letter-spacing:-.04em; }
      div[data-testid="stExpander"] { border:1px solid #e0e5f7; border-radius:13px; background:rgba(255,255,255,.8); }
      .footer { text-align:center; color:#8290ad; border-top:1px solid rgba(216,222,244,.9); margin-top:2rem; padding-top:1rem; font-size:.77rem; }
      .stButton button,.stDownloadButton button { border-radius:11px; font-weight:800; }
      .stDownloadButton button { border:1px solid #d8c9ff; color:#5b21b6; background:#fff; }
      @media (max-width:1050px) { .summary-grid { grid-template-columns:repeat(3,minmax(0,1fr)); } .info-grid { grid-template-columns:repeat(2,minmax(0,1fr)); } }
      @media (max-width:700px) { .block-container { padding-left:.8rem; padding-right:.8rem; } .topbar { align-items:flex-start; flex-direction:column; } .brand-name { font-size:1.5rem; } .top-pill { white-space:normal; } .summary-grid { grid-template-columns:repeat(2,minmax(0,1fr)); } .info-grid { grid-template-columns:1fr; } }

      /* Professional business dashboard theme */
      .stApp {
        color:#172b4d !important;
        background:linear-gradient(180deg,#f8fafc 0%,#f1f5f9 100%) !important;
      }
      :root { --ink:#172b4d; --muted:#64748b; --line:#dbe3ed; --purple:#2563eb; --pink:#2563eb; --blue:#2563eb; }
      .topbar {
        position:relative; margin:0 0 1.4rem; padding:1rem 1.25rem;
        border:1px solid #e2e8f0; border-radius:16px;
        background:rgba(255,255,255,.96); box-shadow:0 4px 16px rgba(15,23,42,.045);
      }
      .brand-icon { background:linear-gradient(145deg,#2563eb,#1d4ed8) !important; box-shadow:0 5px 14px rgba(37,99,235,.16) !important; border-radius:12px !important; }
      .brand-name { color:#0f172a !important; }
      .brand-name span { color:#2563eb !important; background:none !important; -webkit-text-fill-color:#2563eb !important; }
      .brand-sub { color:#64748b !important; }
      [data-testid="stVerticalBlockBorderWrapper"], div[data-testid="stTabContent"] {
        border-color:#e2e8f0 !important; background:#fff !important;
        box-shadow:0 5px 18px rgba(15,23,42,.035) !important; border-radius:16px !important;
      }
      .heading-icon.pink { color:#2563eb !important; background:#eff6ff !important; border-color:#dbeafe !important; }
      .heading-icon.blue { color:#2563eb !important; background:#eff6ff !important; border-color:#dbeafe !important; }
      .card-title .accent { color:#2563eb !important; }
      div[data-testid="stFileUploader"] {
        border-color:#93c5fd !important; background:#f8fbff !important; border-radius:14px !important;
      }
      div[data-testid="stFileUploader"] section { background:#fff !important; }
      div[data-testid="stFileUploader"] button { color:#1d4ed8 !important; border-color:#bfdbfe !important; }
      .summary-card { background:#fff !important; border-color:#e2e8f0 !important; border-radius:12px !important; }
      .tone-blue,.tone-pink,.tone-gold,.tone-green,.tone-purple { background:#fff !important; }
      .icon-blue,.icon-pink,.icon-gold,.icon-green,.icon-purple { color:#2563eb !important; background:#eff6ff !important; }
      .loaded-banner { border-color:#bfdbfe !important; background:#eff6ff !important; color:#1d4ed8 !important; }
      div[data-testid="stTabs"] > div[role="tablist"] { background:#fff !important; border-color:#e2e8f0 !important; }
      div[data-testid="stTabs"] button[role="tab"] { color:#475569 !important; }
      div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] { color:#fff !important; background:#2563eb !important; }
      .empty-overview .empty-symbol { background:#eff6ff !important; color:#2563eb !important; }
      .empty-overview .empty-tip { color:#2563eb !important; }
      .stDownloadButton button { color:#1d4ed8 !important; border-color:#bfdbfe !important; }
      .footer { color:#64748b !important; border-color:#e2e8f0 !important; }
      @media (max-width:700px) { .topbar { padding:.85rem; } }


      /* Match the upload and overview panels in width and minimum height */
      .block-container > div > div[data-testid="stHorizontalBlock"] > div[data-testid="column"] > div[data-testid="stVerticalBlock"] > div[data-testid="stVerticalBlockBorderWrapper"] {
        min-height: 430px !important;
        height: 100%;
        box-sizing: border-box;
      }
      .empty-overview { min-height: 330px; }
\n    </style>
    <div class="topbar">
      <div class="brand">
        <div class="brand-icon">▥</div>
        <div><div class="brand-name">Data<span>Scope</span></div><div class="brand-sub">From raw data to clear insights.</div></div>
      </div>
    </div>
    """,
    unsafe_allow_html=True,
)


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    """Read CSV or Excel data without changing the original uploaded file."""
    suffix = Path(uploaded_file.name).suffix.lower()
    uploaded_file.seek(0)
    if suffix == ".csv":
        data = uploaded_file.getvalue()
        last_error = None
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                return pd.read_csv(BytesIO(data), encoding=encoding, low_memory=False)
            except UnicodeDecodeError as exc:
                last_error = exc
        if last_error:
            raise last_error
    if suffix in (".xlsx", ".xls"):
        return pd.read_excel(uploaded_file)
    raise ValueError("Unsupported format. Upload a CSV, XLSX, or XLS file.")


def guess_encoding(uploaded_file) -> str:
    suffix = Path(uploaded_file.name).suffix.lower()
    if suffix in (".xlsx", ".xls"):
        return "Excel workbook"
    data = uploaded_file.getvalue()
    try:
        data.decode("utf-8-sig")
        return "UTF-8"
    except UnicodeDecodeError:
        try:
            data.decode("cp1252")
            return "Windows-1252"
        except UnicodeDecodeError:
            return "Other / legacy"


def blank_mask(series: pd.Series) -> pd.Series:
    if pd.api.types.is_object_dtype(series) or pd.api.types.is_string_dtype(series):
        return (series.astype("string").str.strip() == "").fillna(False)
    return pd.Series(False, index=series.index, dtype=bool)


def text_null_mask(series: pd.Series) -> pd.Series:
    if pd.api.types.is_object_dtype(series) or pd.api.types.is_string_dtype(series):
        return series.astype("string").str.strip().str.lower().isin(["null", "none", "nan", "n/a", "na", "<na>"]).fillna(False)
    return pd.Series(False, index=series.index, dtype=bool)


def get_column_quality(df: pd.DataFrame) -> pd.DataFrame:
    records = []
    total_rows = len(df)
    for column in df.columns:
        series = df[column]
        missing_count = int(series.isna().sum())
        blank_count = int(blank_mask(series).sum())
        marker_count = int(text_null_mask(series).sum())
        records.append({
            "Column": str(column),
            "Data Type": str(series.dtype),
            "Missing Values": missing_count,
            "Missing %": round(missing_count / total_rows * 100, 2) if total_rows else 0.0,
            "Blank Cells": blank_count,
            "Text NULL Markers": marker_count,
            "Unique Values": int(series.nunique(dropna=True)),
        })
    return pd.DataFrame(records)


def human_size(size: int) -> str:
    if size < 1024:
        return f"{size} B"
    if size < 1024 * 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size / (1024 * 1024):.2f} MB"


def escaped(value) -> str:
    return html.escape(str(value))


def render_snapshot(df: pd.DataFrame, source_label: str, source_size: int, encoding: str) -> None:
    rows, columns = len(df), len(df.columns)
    missing = int(df.isna().sum().sum())
    blanks = sum(int(blank_mask(df[col]).sum()) for col in df.columns)
    duplicates = int(df.duplicated().sum()) if rows else 0
    file_type = Path(source_label).suffix.upper().replace(".", "")
    memory = int(df.memory_usage(index=True, deep=True).sum())
    file_size_label = human_size(source_size)
    metric_items = [
        ("♧", "Rows", f"{rows:,}", "tone-blue", "icon-blue"),
        ("▦", "Columns", f"{columns:,}", "tone-pink", "icon-pink"),
        ("◴", "Missing Values", f"{missing:,}", "tone-gold", "icon-gold"),
        ("↧", "Blank Cells", f"{blanks:,}", "tone-green", "icon-green"),
        ("▤", "Duplicate Rows", f"{duplicates:,}", "tone-purple", "icon-purple"),
    ]
    metric_html = "".join(
        f'<div class="summary-card {tone}"><div class="summary-icon {icon}">{ico}</div><div class="summary-label">{label}</div><div class="summary-value">{value}</div></div>'
        for ico, label, value, tone, icon in metric_items
    )
    basic_items = [
        ("▣", "File Name", source_label),
        ("▤", "File Type", file_type),
        ("▤", "Data Shape", f"{rows:,} × {columns:,}"),
        ("◷", "File Size", file_size_label),
        ("◉", "Encoding", encoding),
        ("⬡", "Memory Usage", human_size(memory)),
    ]
    basic_html = "".join(
        f'<div class="info-item"><div class="info-icon">{ico}</div><div><div class="info-label">{label}</div><div class="info-value">{escaped(value)}</div></div></div>'
        for ico, label, value in basic_items
    )
    safe_name = escaped(source_label)
    st.markdown(
        f'<div class="overview-head"><div style="display:flex;align-items:center;gap:.75rem"><div class="heading-icon blue">▥</div><div><div class="overview-title">Dataset Overview</div><div class="overview-copy">Quick summary of your uploaded dataset</div></div></div><div class="file-pill"><span style="font-size:1.2rem;color:#2684ee">▤</span><span>{safe_name}<small>{file_size_label}</small></span><span class="file-check">✓</span></div></div>'
        f'<div class="summary-grid">{metric_html}</div><div class="basic-title">Basic Information</div><div class="info-grid">{basic_html}</div>',
        unsafe_allow_html=True,
    )
    if missing or blanks or duplicates or int(df.apply(lambda s: text_null_mask(s).sum()).sum()):
        st.caption("Data quality issues are summarized here. Open the Data Quality tab below to inspect them by column.")
    else:
        st.caption("Initial quality checks look clear.")


def chart_layout(fig, height: int = 330):
    fig.update_layout(
        height=height,
        margin=dict(l=12, r=12, t=45, b=12),
        paper_bgcolor="rgba(255,255,255,0)",
        plot_bgcolor="rgba(255,255,255,0)",
        font=dict(family="Inter, Arial, sans-serif", color="#415174", size=12),
        title=dict(font=dict(size=16, color="#182544"), x=0.02, xanchor="left"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
    )
    fig.update_xaxes(showgrid=False, linecolor="#e0e5f7")
    fig.update_yaxes(gridcolor="#edf0fb", zerolinecolor="#e0e5f7")
    return fig


def render_dataset(df: pd.DataFrame, source_label: str) -> None:
    if len(df.columns) == 0:
        st.error("No columns were found in this dataset. Please check the file.")
        return
    if df.empty:
        st.warning("This file contains column headers but no data rows.")

    rows = len(df)
    duplicate_count = int(df.duplicated().sum()) if rows else 0
    missing_cells = int(df.isna().sum().sum())
    blank_cells = sum(int(blank_mask(df[col]).sum()) for col in df.columns)
    text_null_markers = sum(int(text_null_mask(df[col]).sum()) for col in df.columns)
    numeric_cols = list(df.select_dtypes(include="number").columns)
    quality_df = get_column_quality(df)
    category_cols = [c for c in df.columns if (pd.api.types.is_object_dtype(df[c]) or pd.api.types.is_string_dtype(df[c]) or isinstance(df[c].dtype, pd.CategoricalDtype)) and 0 < df[c].nunique(dropna=True) <= 100]

    st.markdown(f'<div class="loaded-banner">✓ Dataset loaded successfully &nbsp;·&nbsp; {escaped(source_label)}</div>', unsafe_allow_html=True)

    preview_tab, quality_tab, quick_tab, visuals_tab, corr_tab, download_tab = st.tabs([
        "▦  Data Preview", "⬟  Data Quality", "▥  Quick Analysis", "◔  Visualizations", "▦  Correlation", "⇩  Download"
    ])

    with preview_tab:
        left, right = st.columns([2.4, 1], vertical_alignment="bottom")
        with left:
            st.markdown('<div class="section-title">Data Preview</div><div class="section-copy">Here’s a preview of your dataset. Search to find specific values.</div>', unsafe_allow_html=True)
        with right:
            search = st.text_input("Search in data", placeholder="⌕  Search in data...", label_visibility="collapsed", key="preview_search")
        limit_col, info_col = st.columns([1, 4], vertical_alignment="center")
        with limit_col:
            row_limit = st.selectbox("Show rows", [10, 25, 50, 100], index=0, key="preview_row_limit")
        with info_col:
            st.caption(f"{len(df):,} total rows · {len(df.columns):,} columns")
        filtered = df
        if search.strip():
            needle = re.escape(search.strip())
            mask = df.astype("string").fillna("").apply(lambda col: col.str.contains(needle, case=False, regex=True, na=False)).any(axis=1)
            filtered = df.loc[mask]
        preview = filtered.head(row_limit).copy()
        preview.insert(0, "#", preview.index.to_series().astype(int) + 1)
        st.dataframe(preview, use_container_width=True, height=390, hide_index=True)
        if search.strip():
            st.caption(f"Found {len(filtered):,} matching rows.")

    with quality_tab:
        st.markdown('<div class="section-title">Data Quality</div><div class="section-copy">Review common data issues. Nothing is removed or changed automatically.</div>', unsafe_allow_html=True)
        qc = st.columns(5)
        qc[0].metric("Missing values", f"{missing_cells:,}")
        qc[1].metric("Blank cells", f"{blank_cells:,}")
        qc[2].metric("Text NULL markers", f"{text_null_markers:,}")
        qc[3].metric("Duplicate rows", f"{duplicate_count:,}")
        qc[4].metric("Columns", f"{len(df.columns):,}")
        st.markdown("#### Column-level quality report")
        st.dataframe(quality_df, use_container_width=True, hide_index=True)
        duplicates = df[df.duplicated(keep=False)]
        if not duplicates.empty:
            with st.expander(f"Review duplicate records · {len(duplicates):,} rows involved"):
                st.dataframe(duplicates.head(200), use_container_width=True, hide_index=True)
        if missing_cells == 0 and blank_cells == 0 and text_null_markers == 0 and duplicate_count == 0:
            st.success("No issues were detected by these initial checks.")
        else:
            st.info("Review flagged values before deciding how they should be cleaned. Duplicate checks compare entire rows.")

    with quick_tab:
        st.markdown('<div class="section-title">Quick Analysis</div><div class="section-copy">Understand the size and shape of the dataset at a glance.</div>', unsafe_allow_html=True)
        q1, q2, q3, q4 = st.columns(4)
        q1.metric("Records", f"{len(df):,}")
        q2.metric("Numeric columns", f"{len(numeric_cols):,}")
        q3.metric("Text/category columns", f"{len(category_cols):,}")
        q4.metric("Rows with duplicates", f"{duplicate_count:,}")
        if numeric_cols:
            st.markdown("#### Numeric summary statistics")
            st.dataframe(df[numeric_cols].describe().T, use_container_width=True)
        else:
            st.info("No numeric columns were detected in this dataset.")
        with st.expander("Explore distinct values in a column"):
            selected_col = st.selectbox("Choose a column", list(df.columns), key="quick_distinct_col")
            counts = df[selected_col].astype("string").fillna("<MISSING>").value_counts(dropna=False).head(50).rename_axis("Value").reset_index(name="Count")
            st.dataframe(counts, use_container_width=True, hide_index=True)

    with visuals_tab:
        st.markdown('<div class="section-title">Visualizations</div><div class="section-copy">Choose columns and generate charts from the uploaded dataset.</div>', unsafe_allow_html=True)
        chart_left, chart_right = st.columns(2, gap="large")
        with chart_left:
            if numeric_cols:
                selected_numeric = st.selectbox("Numeric column · distribution", numeric_cols, key="numeric_chart_col")
                fig = px.histogram(df, x=selected_numeric, nbins=25, title=f"Distribution of {selected_numeric}", color_discrete_sequence=["#a855f7"])
                fig.update_traces(marker_line_color="#ffffff", marker_line_width=1)
                st.plotly_chart(chart_layout(fig), use_container_width=True)
            else:
                st.info("No numeric columns were detected for a distribution chart.")
        with chart_right:
            if category_cols:
                selected_category = st.selectbox("Category column · most common values", category_cols, key="category_chart_col")
                counts = df[selected_category].astype("string").fillna("<MISSING>").value_counts(dropna=False).head(15).rename_axis("Value").reset_index(name="Count")
                fig = px.bar(counts.sort_values("Count"), x="Count", y="Value", orientation="h", title=f"Top values in {selected_category}", color_discrete_sequence=["#ec4899"])
                st.plotly_chart(chart_layout(fig, height=max(330, 24 * len(counts) + 100)), use_container_width=True)
            else:
                st.info("No suitable text/category columns were detected for a category chart.")

    with corr_tab:
        st.markdown('<div class="section-title">Correlation</div><div class="section-copy">Explore relationships between numeric columns. Correlation does not prove causation.</div>', unsafe_allow_html=True)
        if len(numeric_cols) >= 2:
            corr = df[numeric_cols].corr(numeric_only=True)
            fig = px.imshow(corr, text_auto=True, aspect="auto", title="Correlation between numeric columns", color_continuous_scale=["#fce7f3", "#e9d5ff", "#60a5fa"], zmin=-1, zmax=1)
            st.plotly_chart(chart_layout(fig, height=max(390, 50 * len(numeric_cols))), use_container_width=True)
        else:
            st.info("Correlation needs at least two numeric columns. This dataset does not have enough numeric columns yet.")

    with download_tab:
        st.markdown('<div class="section-title">Download</div><div class="section-copy">Download the original imported data or supporting analysis files.</div>', unsafe_allow_html=True)
        down1, down2, down3 = st.columns(3)
        with down1:
            st.download_button("Download dataset as CSV", data=df.to_csv(index=False).encode("utf-8-sig"), file_name=f"{Path(source_label).stem}_data.csv", mime="text/csv", key="download_dataset_csv", use_container_width=True)
        with down2:
            xlsx_buffer = BytesIO()
            with pd.ExcelWriter(xlsx_buffer, engine="openpyxl") as writer:
                df.to_excel(writer, index=False, sheet_name="Dataset")
                quality_df.to_excel(writer, index=False, sheet_name="Data Quality")
                if numeric_cols:
                    df[numeric_cols].describe().T.to_excel(writer, sheet_name="Numeric Summary")
            st.download_button("Download Excel workbook", data=xlsx_buffer.getvalue(), file_name=f"{Path(source_label).stem}_analysis.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="download_dataset_xlsx", use_container_width=True)
        with down3:
            st.download_button("Download quality report", data=quality_df.to_csv(index=False).encode("utf-8-sig"), file_name=f"{Path(source_label).stem}_quality_report.csv", mime="text/csv", key="download_quality_csv", use_container_width=True)


# ---------- Main page: upload left, dataset information right ----------
left_col, right_col = st.columns([1, 1], gap="large", vertical_alignment="top")
with left_col:
    with st.container(border=True):
        st.markdown(
            '<div class="card-heading"><div class="heading-icon pink">⬆</div><div><div class="card-title">Upload your <span class="accent">dataset</span></div><div class="card-copy">Upload a CSV or Excel file to preview, check data quality and get instant insights.</div></div></div>',
            unsafe_allow_html=True,
        )
        uploaded_file = st.file_uploader(
            "Choose a dataset",
            type=SUPPORTED_FILE_TYPES,
            help="Supported formats: CSV, XLSX and XLS. Your source file stays unchanged.",
            key="dataset_file_upload",
        )
        st.markdown('<div class="upload-note">Supported formats: CSV, XLSX, XLS &nbsp; | &nbsp; Maximum size: 200MB</div>', unsafe_allow_html=True)

file_df = None
load_error = None
if uploaded_file is not None:
    try:
        file_df = read_uploaded_file(uploaded_file)
    except Exception as exc:
        load_error = str(exc)

with right_col:
    with st.container(border=True):
        if file_df is not None:
            render_snapshot(file_df, uploaded_file.name, uploaded_file.size, guess_encoding(uploaded_file))
        else:
            st.markdown(
                '<div class="empty-overview"><div class="empty-head"><div class="heading-icon blue">▥</div></div><h3>Dataset Overview</h3><p>Quick summary of your uploaded dataset will appear here.</p><p>After you upload a file, this panel shows rows, columns, missing values, blank cells, duplicate rows and basic file information.</p><div class="empty-tip">✦ One upload · Instant overview · Clear next steps</div></div>',
                unsafe_allow_html=True,
            )

if load_error:
    st.error(f"We couldn't read this file: {load_error}")
    st.info("Check that the file is not corrupted or password-protected, then try again.")
elif file_df is not None:
    render_dataset(file_df, uploaded_file.name)
else:
    st.markdown(
        '<div style="height:.6rem"></div>',
        unsafe_allow_html=True,
    )
    st.markdown('<div class="footer">DataScope · Upload once · Preview · Data Quality · Visual Analysis</div>', unsafe_allow_html=True)

if file_df is not None or load_error:
    st.markdown('<div class="footer">DataScope · Upload once · Preview · Data Quality · Visual Analysis</div>', unsafe_allow_html=True)
