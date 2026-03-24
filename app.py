import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import numpy as np
import html
import glob
import os
import re
from datetime import datetime

st.set_page_config(
    page_title="Industry Buy Pressure Dashboard (JP)",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("🔥 Industry Buy Pressure Dashboard (日本株)")
st.markdown("---")


def get_color_from_buy_pressure(buy_pressure):
    if pd.isna(buy_pressure):
        return "#808080"
    normalized = max(0.0, min(1.0, buy_pressure))
    if normalized >= 0.5:
        ratio = (normalized - 0.5) * 2
        r = int(255 * (1 - ratio))
        g = 255
        b = 0
    else:
        ratio = normalized * 2
        r = 255
        g = int(255 * ratio)
        b = 0
    return f"#{r:02x}{g:02x}{b:02x}"


def truncate_name(name, max_len=10):
    if pd.isna(name):
        return ""
    s = str(name).strip()
    if len(s) > max_len:
        return s[:max_len] + "…"
    return s


def get_buy_pressure_status(buy_pressure):
    if buy_pressure > 0.667:
        return "3 🔥 EXTREME"
    elif buy_pressure > 0.60:
        return "2 🚀 STRONG"
    elif buy_pressure > 0.55:
        return "1 📈 BUY"
    elif buy_pressure < 0.333:
        return "0a 💀 WEAK"
    elif buy_pressure < 0.45:
        return "0b ⚠️ CAUTION"
    else:
        return "0c ➖ NEUTRAL"


def get_buy_pressure_status_display(buy_pressure):
    if buy_pressure > 0.667:
        return "🔥 EXTREME"
    elif buy_pressure > 0.60:
        return "🚀 STRONG"
    elif buy_pressure > 0.55:
        return "📈 BUY"
    elif buy_pressure < 0.333:
        return "💀 WEAK"
    elif buy_pressure < 0.45:
        return "⚠️ CAUTION"
    else:
        return "➖ NEUTRAL"


STATUS_DEFINITIONS = [
    {"key": "extreme", "label": "🔥 EXTREME", "default": True},
    {"key": "strong",  "label": "🚀 STRONG",  "default": True},
    {"key": "buy",     "label": "📈 BUY",     "default": True},
]

FILTERABLE_KEYS = [sd["key"] for sd in STATUS_DEFINITIONS]


def get_status_key(buy_pressure):
    if buy_pressure > 0.667:
        return "extreme"
    elif buy_pressure > 0.60:
        return "strong"
    elif buy_pressure > 0.55:
        return "buy"
    elif buy_pressure < 0.333:
        return "weak"
    elif buy_pressure < 0.45:
        return "caution"
    else:
        return "neutral"


CUSTOM_RS_COLORSCALE = [
    [0.0, "#ff0000"],
    [0.4, "#ff8c00"],
    [0.79, "#d4c860"],
    [0.80, "#9acd32"],
    [1.0, "#006400"],
]


def find_latest_file(directory, prefix):
    pattern = os.path.join(directory, f"{prefix}*.xlsx")
    matched_files = glob.glob(pattern)
    if not matched_files:
        raise FileNotFoundError(
            f"'{directory}/' 内に '{prefix}*.xlsx' に一致するファイルが見つかりません。"
        )
    date_pattern = re.compile(r'(\d{8}_\d{6})\.xlsx$')
    files_with_dates = []
    for filepath in matched_files:
        filename = os.path.basename(filepath)
        match = date_pattern.search(filename)
        if match:
            files_with_dates.append((filepath, match.group(1)))
    if not files_with_dates:
        raise FileNotFoundError(
            f"'{directory}/' 内に日付パターン(YYYYMMDD_HHMMSS)を含むファイルが見つかりません。"
        )
    files_with_dates.sort(key=lambda x: x[1], reverse=True)
    return files_with_dates[0][0]


def get_data_date_from_filename(filename):
    match = re.search(r'(\d{8})_\d{6}', filename)
    if match:
        file_date = datetime.strptime(match.group(1), '%Y%m%d')
        return file_date.strftime('%Y-%m-%d')
    return "不明"


@st.cache_data
def load_data():
    DATA_DIR = "data"
    file_path = find_latest_file(DATA_DIR, "jp_full_screening_")
    file_name = os.path.basename(file_path)
    data_date = get_data_date_from_filename(file_name)
    xl = pd.ExcelFile(file_path)
    sheet_names = xl.sheet_names

    if 'Screening_Results' not in sheet_names:
        raise ValueError(f"'Screening_Results' シートが見つかりません。シート: {sheet_names}")
    df_screening_raw = pd.read_excel(file_path, sheet_name='Screening_Results')
    col_rename_screening = {
        'Code': 'Symbol', 'CoName': 'Company Name', 'YF_Industry': 'Industry',
        'YF_Sector': 'Sector', 'BP_Stock': 'Buy_Pressure',
    }
    df_screening_raw = df_screening_raw.rename(columns=col_rename_screening)
    df_screening_raw['Symbol'] = (
        df_screening_raw['Symbol'].astype(str)
        .str.replace(r'\.0$', '', regex=True)
        .str.replace(r'0$', '', regex=True)
    )
    if 'Fundamental_Score' not in df_screening_raw.columns:
        if 'Screening_Score' in df_screening_raw.columns and 'Technical_Score' in df_screening_raw.columns:
            df_screening_raw['Fundamental_Score'] = (
                df_screening_raw['Screening_Score'] - df_screening_raw['Technical_Score']
            )
    df_screening_filtered = df_screening_raw[df_screening_raw['Technical_Score'] >= 10].copy()
    required_cols = ['Symbol', 'Industry', 'Technical_Score', 'Screening_Score', 'Buy_Pressure', 'Company Name']
    for col in required_cols:
        if col not in df_screening_filtered.columns:
            raise ValueError(f"Screening_Results に '{col}' 列が見つかりません。")
    df_screening_filtered = df_screening_filtered[required_cols].copy()

    if 'Industry_Ranking' not in sheet_names:
        raise ValueError(f"'Industry_Ranking' シートが見つかりません。")
    df_ind_ranking = pd.read_excel(file_path, sheet_name='Industry_Ranking')
    df_industry = df_ind_ranking[['Industry', 'RS_Pct_CW', 'BP_CW']].copy()
    df_industry = df_industry.rename(columns={'RS_Pct_CW': 'RS_Rating', 'BP_CW': 'Buy_Pressure'})
    df_industry['RS_Rating'] = pd.to_numeric(df_industry['RS_Rating'], errors='coerce')
    df_industry['Buy_Pressure'] = pd.to_numeric(df_industry['Buy_Pressure'], errors='coerce')
    df_industry = df_industry.dropna(subset=['RS_Rating', 'Buy_Pressure'])
    df_all_industry = df_industry.copy()

    # --- 業種→セクター マッピング ---
    industry_sector_map = {}
    if 'Sector' in df_screening_raw.columns and 'Industry' in df_screening_raw.columns:
        sector_df = df_screening_raw[['Industry', 'Sector']].dropna().drop_duplicates()
        for industry in sector_df['Industry'].unique():
            sectors = sector_df[sector_df['Industry'] == industry]['Sector']
            industry_sector_map[industry] = sectors.mode().iloc[0] if len(sectors) > 0 else 'Unknown'

    # --- セクターランキング読み込み ---
    df_sector_ranking = None
    rs80_sectors = set()
    if 'Sector_Ranking' in sheet_names:
        df_sector_ranking = pd.read_excel(file_path, sheet_name='Sector_Ranking')
        if 'RS' in df_sector_ranking.columns and 'Sector' in df_sector_ranking.columns:
            rs80_sectors = set(
                df_sector_ranking[df_sector_ranking['RS'] >= 80]['Sector'].tolist()
            )

    # --- 条件①: 業種 RS ≥ 80 かつ BP > 0.55 ---
    industries_by_industry_condition = set(
        df_industry[
            (df_industry['RS_Rating'] >= 80) & (df_industry['Buy_Pressure'] > 0.55)
        ]['Industry'].tolist()
    )

    # --- 条件②: セクター RS ≥ 80 に属する業種 ---
    industries_by_sector_condition = set()
    for industry, sector in industry_sector_map.items():
        if sector in rs80_sectors:
            industries_by_sector_condition.add(industry)

    # --- 和集合 ---
    all_passed_industries = industries_by_industry_condition | industries_by_sector_condition

    df_industry_passed = df_industry[
        df_industry['Industry'].isin(all_passed_industries)
    ].copy()

    # --- 解説データ読み込み ---
    df_description = None
    desc_path = os.path.join(DATA_DIR, "Industry_Description_JP.xlsx")
    if not os.path.exists(desc_path):
        desc_path = os.path.join(DATA_DIR, "Industry_Description.xlsx")
    if os.path.exists(desc_path):
        try:
            df_description = pd.read_excel(desc_path)
        except Exception:
            df_description = None

    return (
        df_industry_passed, df_all_industry, df_screening_filtered, industry_sector_map,
        data_date, df_description, rs80_sectors,
    )


try:
    (
        df_industry, df_all_industry, df_screening, industry_sector_map,
        data_date, df_description, rs80_sectors,
    ) = load_data()

    # 条件内訳を表示
    ind_cond_count = len(df_all_industry[
        (df_all_industry['RS_Rating'] >= 80) & (df_all_industry['Buy_Pressure'] > 0.55)
    ])
    sector_cond_industries = set()
    for industry, sector in industry_sector_map.items():
        if sector in rs80_sectors:
            if industry in df_all_industry['Industry'].values:
                sector_cond_industries.add(industry)
    sector_only_count = len(sector_cond_industries - set(df_all_industry[
        (df_all_industry['RS_Rating'] >= 80) & (df_all_industry['Buy_Pressure'] > 0.55)
    ]['Industry'].tolist()))

    st.success(
        f"✅ データ読み込み成功: {len(df_industry)} 業種 (条件通過), "
        f"{len(df_all_industry)} 業種 (全体), {len(df_screening)} 銘柄"
    )
    st.caption(f"📅 データ日付: **{data_date}**")
    st.caption(
        f"📋 条件通過内訳: 業種RS≥80 & BP>0.55 → **{ind_cond_count}**件 ＋ "
        f"セクターRS≥80 → 追加 **{sector_only_count}**件 "
        f"（RS≥80セクター: {', '.join(sorted(rs80_sectors)) if rs80_sectors else 'なし'}）"
    )
except Exception as e:
    st.error(f"❌ データ読み込みエラー: {str(e)}")
    import traceback
    st.code(traceback.format_exc())
    st.stop()


with st.sidebar:
    st.header("📊 フィルター設定")
    min_tech_score = st.slider(
        "テクニカルスコア最小値", min_value=10,
        max_value=int(df_screening['Technical_Score'].max()), value=10, step=1
    )
    max_stocks_per_industry = st.slider(
        "業種ごとの最大表示銘柄数", min_value=5, max_value=30, value=15, step=5
    )
    all_industries_in_data = sorted(
        set(df_industry['Industry'].tolist()) | set(df_all_industry['Industry'].tolist())
    )
    selected_industries = st.multiselect(
        "業種選択（空白=条件通過業種のみ）", options=all_industries_in_data, default=None
    )
    st.markdown("---")
    st.markdown("### 🎨 カラーコード")
    st.markdown("- 🟢 **緑**: Buy Pressure 高い")
    st.markdown("- 🟡 **黄**: Buy Pressure 中程度")
    st.markdown("- 🔴 **赤**: Buy Pressure 低い")


df_screening_display = df_screening[df_screening['Technical_Score'] >= min_tech_score].copy()
df_screening_display['Fundamental_Score'] = (
    df_screening_display['Screening_Score'] - df_screening_display['Technical_Score']
)
if selected_industries:
    df_screening_display = df_screening_display[
        df_screening_display['Industry'].isin(selected_industries)
    ]
    df_industry_display = df_all_industry[df_all_industry['Industry'].isin(selected_industries)].copy()
else:
    df_industry_display = df_industry.copy()


def create_summary_data(df_screening_disp, df_industry_disp):
    industry_summary = []
    for _, industry_row in df_industry_disp.iterrows():
        industry = industry_row['Industry']
        stocks = df_screening_disp[df_screening_disp['Industry'] == industry]
        status = get_buy_pressure_status(industry_row['Buy_Pressure'])
        industry_summary.append({
            '業種': industry,
            'RS Rating': industry_row['RS_Rating'],
            'Buy Pressure': industry_row['Buy_Pressure'],
            'ステータス': status,
            '銘柄数': len(stocks),
            '平均テクニカルスコア': stocks['Technical_Score'].mean() if len(stocks) > 0 else 0,
            '平均スクリーニングスコア': stocks['Screening_Score'].mean() if len(stocks) > 0 else 0,
        })
    df_summary = pd.DataFrame(industry_summary)
    df_summary = df_summary.sort_values('RS Rating', ascending=False)
    return df_summary


df_summary = create_summary_data(df_screening_display, df_industry_display)

# --- 6タブ構成 ---
tab0, tab0b, tab1, tab2, tab3, tab4 = st.tabs([
    "✅ チェック", "✅ チェック②",
    "📈 テクニカルスコア別マトリックス", "🎯 スクリーニングスコア別マトリックス",
    "📊 業種サマリー", "📖 解説"
])


def style_symbol(row):
    styles = [''] * len(row)
    try:
        bp = float(row['Buy Pressure'])
        color = get_color_from_buy_pressure(bp)
        symbol_idx = row.index.get_loc('Symbol')
        styles[symbol_idx] = f'background-color: #000000; color: {color}; font-weight: bold; font-size: 16px;'
        bp_idx = row.index.get_loc('Buy Pressure')
        styles[bp_idx] = f'background-color: #000000; color: {color}; font-weight: bold;'
    except (ValueError, TypeError, KeyError):
        pass
    return styles


def style_symbol_black_bg(row):
    base = 'background-color: #000000; color: #fafafa;'
    styles = [base] * len(row)
    try:
        bp = float(row['Buy Pressure'])
        color = get_color_from_buy_pressure(bp)
        symbol_idx = row.index.get_loc('Symbol')
        styles[symbol_idx] = f'background-color: #000000; color: {color}; font-weight: bold; font-size: 16px;'
        bp_idx = row.index.get_loc('Buy Pressure')
        styles[bp_idx] = f'background-color: #000000; color: {color}; font-weight: bold;'
    except (ValueError, TypeError, KeyError):
        pass
    return styles


def create_industry_table(df_screening_disp, df_industry_disp, sort_by='Technical_Score'):
    df_industry_sorted = df_industry_disp.sort_values('RS_Rating', ascending=False)
    for _, industry_row in df_industry_sorted.iterrows():
        industry_name = industry_row['Industry']
        rs_rating = industry_row['RS_Rating']
        buy_pressure = industry_row['Buy_Pressure']
        stocks_in_industry = df_screening_disp[
            df_screening_disp['Industry'] == industry_name
        ].sort_values(sort_by, ascending=False).head(max_stocks_per_industry)
        if len(stocks_in_industry) == 0:
            continue
        st.markdown(f"### {industry_name}")
        col1, col2, col3, col4 = st.columns([3, 1, 1, 2])
        with col1:
            st.metric("業種", industry_name)
        with col2:
            st.metric("RS Rating", f"{rs_rating:.1f}")
        with col3:
            st.metric("Buy Pressure", f"{buy_pressure:.3f}")
        with col4:
            status = get_buy_pressure_status_display(buy_pressure)
            st.markdown(f"**{status}**")
        display_df = stocks_in_industry[
            ['Symbol', 'Company Name', 'Technical_Score', 'Screening_Score', 'Buy_Pressure']
        ].copy()
        display_df = display_df.reset_index(drop=True)
        display_df.index = display_df.index + 1
        display_df.index.name = 'No'
        display_df.columns = ['Symbol', 'Company Name', 'Technical Score', 'Screening Score', 'Buy Pressure']
        display_df['Company Name'] = display_df['Company Name'].apply(
            lambda x: str(x)[:40] if pd.notna(x) else ''
        )
        styled_df = display_df.style.apply(style_symbol_black_bg, axis=1)
        st.dataframe(styled_df, use_container_width=True, height=min(len(display_df) * 40 + 50, 650))
        st.markdown("---")


def get_colored_symbols_html(industry, score, df_screening_disp):
    stocks = df_screening_disp[
        (df_screening_disp['Industry'] == industry) &
        (df_screening_disp['Technical_Score'] == score)
    ].sort_values('Buy_Pressure', ascending=False)
    if len(stocks) == 0:
        return '', ''
    colored_spans = []
    plain_symbols = []
    for _, stock in stocks.iterrows():
        symbol = html.escape(str(stock['Symbol']))
        name = truncate_name(stock['Company Name'], 10)
        name_esc = html.escape(name)
        bp = stock['Buy_Pressure']
        color = get_color_from_buy_pressure(bp)
        colored_spans.append(
            f'<div class="stock-chip" title="{symbol} {html.escape(str(stock["Company Name"]))}">'
            f'<span style="color:{color}; font-weight:bold;">'
            f'<span class="sym-code">{symbol}</span> {name_esc}</span></div>'
        )
        plain_symbols.append(symbol)
    return ''.join(colored_spans), ', '.join(plain_symbols)


def get_colored_symbols_html_with_fs(industry, ts, fs, df_screening_disp):
    stocks = df_screening_disp[
        (df_screening_disp['Industry'] == industry) &
        (df_screening_disp['Technical_Score'] == ts) &
        (df_screening_disp['Fundamental_Score'] == fs)
    ].sort_values('Buy_Pressure', ascending=False)
    if len(stocks) == 0:
        return '', ''
    colored_spans = []
    plain_symbols = []
    for _, stock in stocks.iterrows():
        symbol = html.escape(str(stock['Symbol']))
        name = truncate_name(stock['Company Name'], 10)
        name_esc = html.escape(name)
        bp = stock['Buy_Pressure']
        color = get_color_from_buy_pressure(bp)
        colored_spans.append(
            f'<div class="stock-chip" title="{symbol} {html.escape(str(stock["Company Name"]))}">'
            f'<span data-symbol="{symbol}" style="color:{color}; font-weight:bold;">'
            f'<span class="sym-code">{symbol}</span> {name_esc}</span></div>'
        )
        plain_symbols.append(symbol)
    return ''.join(colored_spans), ', '.join(plain_symbols)


def get_qualified_stocks(df_screening_disp, df_industry_disp):
    rs80_industries = df_industry_disp[
        df_industry_disp['RS_Rating'] >= 80
    ]['Industry'].tolist()
    qualified = df_screening_disp[
        (df_screening_disp['Industry'].isin(rs80_industries)) &
        (df_screening_disp['Buy_Pressure'] > 0.55)
    ].sort_values(['Industry', 'Buy_Pressure'], ascending=[True, False])
    return qualified


def build_txt_by_industry(qualified_stocks, df_industry_disp, data_date):
    lines = []
    lines.append(f"=== Buy Pressure Qualified Stocks - JP ({data_date}) ===")
    lines.append(f"条件: 業種 RS Rating >= 80 & 個別銘柄 Buy Pressure > 0.55")
    lines.append(f"合計: {len(qualified_stocks)} 銘柄")
    lines.append("")
    industries = qualified_stocks['Industry'].unique()
    industry_rs = df_industry_disp.set_index('Industry')['RS_Rating']
    sorted_industries = sorted(industries, key=lambda x: industry_rs.get(x, 0), reverse=True)
    for industry in sorted_industries:
        stocks = qualified_stocks[qualified_stocks['Industry'] == industry]
        rs = industry_rs.get(industry, 0)
        symbols = stocks['Symbol'].tolist()
        lines.append(f"--- {industry} (RS: {rs:.1f}) ---")
        lines.append(', '.join(symbols))
        lines.append("")
    return '\n'.join(lines)


# ============================================================
# フィルター用 HTML / CSS / JS（ステータス＋セクター）
# ============================================================
def build_filter_html(uid, sector_list):
    status_cbs = ""
    for sd in STATUS_DEFINITIONS:
        checked = "checked" if sd["default"] else ""
        status_cbs += (
            f'<label class="sf-label">'
            f'<input type="checkbox" class="sf-cb-status" data-status="{sd["key"]}" {checked} '
            f'onchange="applyFilter_{uid}()" /> {sd["label"]}</label>'
        )

    sector_cbs = ""
    for sector in sorted(sector_list):
        sector_esc = html.escape(sector)
        sector_cbs += (
            f'<label class="sf-label sf-sector-label">'
            f'<input type="checkbox" class="sf-cb-sector" data-sector="{sector_esc}" checked '
            f'onchange="applyFilter_{uid}()" /> {sector_esc}</label>'
        )

    return f"""
    <div class="filter-bar" id="sf-bar-{uid}">
        <div class="filter-row">
            <span class="sf-title">ステータス:</span>
            {status_cbs}
        </div>
        <div class="filter-row">
            <span class="sf-title">セクター:</span>
            {sector_cbs}
            <button class="sf-btn sf-all" onclick="sfSectorAll_{uid}(true)">全ON</button>
            <button class="sf-btn sf-none" onclick="sfSectorAll_{uid}(false)">全OFF</button>
        </div>
    </div>
    """


def build_filter_css():
    return """
    .filter-bar {
        background-color: #181b22;
        border: 1px solid #444;
        border-radius: 8px;
        padding: 8px 14px;
        margin-bottom: 10px;
    }
    .filter-row {
        display: flex;
        align-items: center;
        flex-wrap: wrap;
        gap: 4px 12px;
        padding: 3px 0;
    }
    .filter-row + .filter-row {
        border-top: 1px solid #333;
        margin-top: 4px;
        padding-top: 6px;
    }
    .sf-title {
        color: #aaa;
        font-size: 12px;
        font-weight: bold;
        margin-right: 4px;
        white-space: nowrap;
    }
    .sf-label {
        color: #ddd;
        font-size: 12px;
        cursor: pointer;
        user-select: none;
        display: inline-flex;
        align-items: center;
        gap: 3px;
        white-space: nowrap;
    }
    .sf-sector-label {
        font-size: 11px;
        color: #bbb;
    }
    .sf-cb-status, .sf-cb-sector {
        accent-color: #00c853;
        width: 15px;
        height: 15px;
        cursor: pointer;
    }
    .sf-btn {
        border: 1px solid #555;
        border-radius: 4px;
        padding: 2px 8px;
        font-size: 10px;
        cursor: pointer;
        color: #ccc;
        background-color: #262730;
        margin-left: 4px;
    }
    .sf-btn:hover { background-color: #3a3d48; }
    .sf-all { border-color: #00c853; }
    .sf-none { border-color: #ff5252; }
    """


def build_filter_js(uid, table_id):
    filterable_js = str(FILTERABLE_KEYS)
    return f"""
    var FILTERABLE_KEYS_{uid} = {filterable_js};

    function getActiveStatuses_{uid}() {{
        var cbs = document.querySelectorAll('#sf-bar-{uid} .sf-cb-status');
        var active = [];
        cbs.forEach(function(cb) {{
            if (cb.checked) active.push(cb.getAttribute('data-status'));
        }});
        return active;
    }}
    function getActiveSectors_{uid}() {{
        var cbs = document.querySelectorAll('#sf-bar-{uid} .sf-cb-sector');
        var active = [];
        cbs.forEach(function(cb) {{
            if (cb.checked) active.push(cb.getAttribute('data-sector'));
        }});
        return active;
    }}
    function applyFilter_{uid}() {{
        var activeStatus = getActiveStatuses_{uid}();
        var activeSector = getActiveSectors_{uid}();
        var rows = document.querySelectorAll('#{table_id} tbody tr');
        rows.forEach(function(row) {{
            var st = row.getAttribute('data-status');
            var sec = row.getAttribute('data-sector') || '';

            var statusOk = false;
            if (st && FILTERABLE_KEYS_{uid}.indexOf(st) >= 0) {{
                statusOk = activeStatus.indexOf(st) >= 0;
            }}

            var sectorOk = activeSector.indexOf(sec) >= 0;

            row.style.display = (statusOk && sectorOk) ? '' : 'none';
        }});
    }}
    function sfSectorAll_{uid}(state) {{
        var cbs = document.querySelectorAll('#sf-bar-{uid} .sf-cb-sector');
        cbs.forEach(function(cb) {{ cb.checked = state; }});
        applyFilter_{uid}();
    }}
    document.addEventListener('DOMContentLoaded', function() {{ applyFilter_{uid}(); }});
    setTimeout(function() {{ applyFilter_{uid}(); }}, 100);
    """


def get_sectors_in_table(df_check, sector_map):
    sectors = set()
    for _, row in df_check.iterrows():
        industry = row['業種']
        sector = sector_map.get(industry, 'Unknown')
        sectors.add(sector)
    return sorted(sectors)


# ============================================================
# チェックタブ用
# ============================================================
def render_check_tab(df_check, df_screening_disp, df_industry_disp, table_id_suffix=""):
    st.header("Buy Pressure")
    max_symbols_per_row = []
    for _, row in df_check.iterrows():
        row_max = 0
        ts_values = sorted(df_screening_disp['Technical_Score'].unique(), reverse=True)
        for score in ts_values:
            count = len(df_screening_disp[
                (df_screening_disp['Industry'] == row['業種']) &
                (df_screening_disp['Technical_Score'] == score)
            ])
            row_max = max(row_max, count)
        max_symbols_per_row.append(row_max)

    ts_values = sorted(df_screening_disp['Technical_Score'].unique(), reverse=True)

    tid = f"check-table{table_id_suffix}"
    toast_id = f"copy-toast{table_id_suffix}"
    func_name = f"copySymbols{table_id_suffix.replace('-', '_')}"
    uid = f"ct{table_id_suffix.replace('-', '')}"

    sector_list = get_sectors_in_table(df_check, industry_sector_map)

    ts_headers = ''.join([f'<th>TS {ts}</th>' for ts in ts_values])

    filter_html = build_filter_html(uid, sector_list)

    table_html = f"""
    <style>
    {build_filter_css()}
    #{tid} {{ width: 100%; border-collapse: collapse; font-size: 11px; }}
    #{tid} th {{ background-color: #262730; color: #fafafa; padding: 6px 8px; text-align: left; border: 1px solid #444; font-size: 12px; }}
    #{tid} td {{ padding: 4px 6px; border: 1px solid #444; background-color: #0e1117; color: #fafafa; vertical-align: top; }}
    #{tid} tr:hover td {{ background-color: #1a1d24; }}
    #{tid} .stock-chip {{ padding: 1px 0; line-height: 1.4; white-space: nowrap; }}
    #{tid} .stock-chip .sym-code {{ display: inline-block; min-width: 36px; color: #888; font-size: 10px; }}
    .copyable{table_id_suffix} {{ cursor: pointer; position: relative; }}
    .copyable{table_id_suffix}:hover {{ background-color: #2a2d34 !important; }}
    #{toast_id} {{ position: fixed; top: 20px; right: 20px; background-color: #00c853; color: white;
                   padding: 10px 20px; border-radius: 8px; font-size: 14px; font-weight: bold;
                   z-index: 9999; opacity: 0; transition: opacity 0.3s; pointer-events: none; }}
    #{toast_id}.show {{ opacity: 1; }}
    </style>
    <div id="{toast_id}" class="copy-toast">📋 Copied!</div>
    {filter_html}
    <div style="overflow-x: auto;">
    <table id="{tid}">
    <thead><tr>
        <th>業種</th><th>RS</th><th>BP</th><th>ステータス</th>
        {ts_headers}
    </tr></thead><tbody>
    """
    for idx, row in df_check.iterrows():
        bp = row['Buy Pressure']
        bp_color = get_color_from_buy_pressure(bp)
        industry_name = str(row['業種'])
        industry_esc = html.escape(industry_name)
        rs = f"{row['RS Rating']:.1f}"
        bp_val = f"{bp:.3f}"
        status_raw = str(row['ステータス'])
        status_display = re.sub(r'^\d+[a-z]?\s+', '', status_raw)
        status = html.escape(status_display)
        status_key = get_status_key(bp)
        sector = html.escape(industry_sector_map.get(industry_name, 'Unknown'))
        table_html += (
            f'<tr data-status="{status_key}" data-sector="{sector}">'
            f'<td style="white-space:nowrap;">{industry_esc}</td><td>{rs}</td>'
        )
        table_html += f'<td style="color: {bp_color}; font-weight: bold;">{bp_val}</td>'
        table_html += f'<td style="white-space:nowrap;">{status}</td>'
        for score in ts_values:
            display_html, copy_text = get_colored_symbols_html(row['業種'], score, df_screening_disp)
            if display_html:
                escaped_copy = html.escape(copy_text).replace("'", "\\'")
                table_html += (
                    f'<td class="copyable{table_id_suffix}" '
                    f'onclick="{func_name}(this, \'{escaped_copy}\')" '
                    f'title="コピー: {escaped_copy}">{display_html}</td>'
                )
            else:
                table_html += '<td></td>'
        table_html += "</tr>"

    table_html += f"""
    </tbody></table></div>
    <script>
    function {func_name}(el, text) {{
        navigator.clipboard.writeText(text).then(function() {{
            var toast = document.getElementById('{toast_id}');
            toast.classList.add('show');
            el.style.backgroundColor = '#1b5e20';
            setTimeout(function() {{ toast.classList.remove('show'); el.style.backgroundColor = ''; }}, 1500);
        }});
    }}
    {build_filter_js(uid, tid)}
    </script>
    """
    total_height = 160
    for sym_count in max_symbols_per_row:
        row_h = max(50, sym_count * 22 + 16)
        total_height += row_h
    total_height = min(total_height, 5000)
    st.components.v1.html(table_html, height=total_height, scrolling=True)

    qualified_stocks = get_qualified_stocks(df_screening_disp, df_industry_disp)
    qualified_symbols = qualified_stocks['Symbol'].tolist()

    with st.expander(f"🎯 個別BP > 0.55 & 業種RS≧80 の銘柄（{len(qualified_symbols)}件）", expanded=False):
        if len(qualified_symbols) == 0:
            st.warning("⚠️ 条件を満たす銘柄はありません。")
        else:
            industry_rs = df_industry_disp.set_index('Industry')['RS_Rating']
            industries_in_qualified = qualified_stocks['Industry'].unique()
            sorted_industries = sorted(
                industries_in_qualified, key=lambda x: industry_rs.get(x, 0), reverse=True
            )
            all_symbols_text = ', '.join(qualified_symbols)
            copy_html = """
            <style>
            .q-toast { position: fixed; top: 20px; right: 20px; background-color: #00c853; color: white;
                       padding: 10px 20px; border-radius: 8px; font-size: 14px; font-weight: bold;
                       z-index: 9999; opacity: 0; transition: opacity 0.3s; pointer-events: none; }
            .q-toast.show { opacity: 1; }
            .q-section { margin-bottom: 16px; }
            .q-industry-title { font-size: 14px; font-weight: bold; color: #fafafa; margin-bottom: 4px; }
            .q-symbols { padding: 8px 12px; background-color: #0e1117; border: 1px solid #333;
                         border-radius: 6px; cursor: pointer; line-height: 2.0; transition: background-color 0.2s; }
            .q-symbols:hover { background-color: #1a2e1a; box-shadow: 0 0 8px rgba(0,200,83,0.2); }
            .q-all-box { margin-bottom: 20px; padding: 12px 16px; background-color: #0e1117;
                         border: 2px solid #00c853; border-radius: 8px; cursor: pointer;
                         transition: background-color 0.2s; line-height: 2.0; }
            .q-all-box:hover { background-color: #1a2e1a; box-shadow: 0 0 12px rgba(0,200,83,0.3); }
            .q-all-title { font-size: 13px; font-weight: bold; color: #00c853; margin-bottom: 6px; }
            .q-hint { font-size: 11px; color: #888; margin-top: 6px; }
            .q-chip { display: inline-block; margin: 2px 6px 2px 0; font-size: 12px; }
            .q-chip .q-code { color: #888; font-size: 10px; margin-right: 2px; }
            </style>
            <div id="q-toast" class="q-toast">📋 Copied!</div>
            """
            escaped_all = html.escape(all_symbols_text).replace("'", "\\'")
            all_colored = []
            for _, stk in qualified_stocks.iterrows():
                sym = html.escape(str(stk['Symbol']))
                name = truncate_name(stk['Company Name'], 10)
                name_esc = html.escape(name)
                c = get_color_from_buy_pressure(stk['Buy_Pressure'])
                all_colored.append(
                    f'<span class="q-chip" title="{sym} {html.escape(str(stk["Company Name"]))}">'
                    f'<span class="q-code">{sym}</span>'
                    f'<span style="color:{c};font-weight:bold;">{name_esc}</span></span>'
                )
            copy_html += f"""
            <div class="q-all-box" onclick="qCopy(this, '{escaped_all}')" title="コピー: {escaped_all}">
                <div class="q-all-title">📋 全銘柄まとめてコピー（{len(qualified_symbols)}件）— クリックでティッカーコードをコピー</div>
                <div>{' '.join(all_colored)}</div>
                <div class="q-hint">クリックでティッカーコードをコピー</div>
            </div>
            """
            for ind in sorted_industries:
                ind_stocks = qualified_stocks[qualified_stocks['Industry'] == ind]
                rs_val = industry_rs.get(ind, 0)
                syms = ind_stocks['Symbol'].tolist()
                plain = ', '.join(syms)
                escaped = html.escape(plain).replace("'", "\\'")
                colored_spans = []
                for _, stk in ind_stocks.iterrows():
                    sym = html.escape(str(stk['Symbol']))
                    name = truncate_name(stk['Company Name'], 10)
                    name_esc = html.escape(name)
                    c = get_color_from_buy_pressure(stk['Buy_Pressure'])
                    colored_spans.append(
                        f'<span class="q-chip" title="{sym} {html.escape(str(stk["Company Name"]))}">'
                        f'<span class="q-code">{sym}</span>'
                        f'<span style="color:{c};font-weight:bold;">{name_esc}</span></span>'
                    )
                ind_esc = html.escape(ind)
                copy_html += f"""
                <div class="q-section">
                    <div class="q-industry-title">{ind_esc} (RS: {rs_val:.1f}) — {len(syms)}件</div>
                    <div class="q-symbols" onclick="qCopy(this, '{escaped}')" title="コピー: {escaped}">
                        {' '.join(colored_spans)}
                        <div class="q-hint">📋 クリックでティッカーコードをコピー</div>
                    </div>
                </div>
                """
            copy_html += """
            <script>
            function qCopy(el, text) {
                navigator.clipboard.writeText(text).then(function() {
                    var toast = document.getElementById('q-toast');
                    toast.classList.add('show');
                    el.style.backgroundColor = '#1b5e20';
                    setTimeout(function() { toast.classList.remove('show'); el.style.backgroundColor = ''; }, 1500);
                });
            }
            </script>
            """
            section_count = len(sorted_industries)
            copy_height = 120 + section_count * 80 + len(qualified_symbols) * 2
            copy_height = min(max(copy_height, 300), 1500)
            st.components.v1.html(copy_html, height=copy_height, scrolling=True)

    if len(qualified_symbols) > 0:
        txt_content = build_txt_by_industry(qualified_stocks, df_industry_disp, data_date)
        st.download_button(
            label="📥 業種別銘柄リストをダウンロード (.txt)",
            data=txt_content.encode('utf-8'),
            file_name=f"qualified_stocks_jp_{data_date}.txt",
            mime="text/plain",
        )


# ============================================================
# チェック②タブ用（TS × FS 細分化）
# ============================================================
def render_check_tab_with_fs(df_check, df_screening_disp):
    st.header("Buy Pressure（TS × FS 細分化）")

    ts_values = sorted(df_screening_disp['Technical_Score'].unique(), reverse=True)
    global_max_fs = int(df_screening_disp['Fundamental_Score'].max())
    global_min_fs = int(df_screening_disp['Fundamental_Score'].min())
    fixed_fs_values = list(range(global_max_fs, global_min_fs - 1, -1))
    ts_fs_map = {}
    for ts in ts_values:
        ts_fs_map[ts] = fixed_fs_values
    all_sub_cols = []
    for ts in ts_values:
        for fs in ts_fs_map[ts]:
            all_sub_cols.append((ts, fs))
    num_rows = len(df_check)

    col_widths = [200, 85, 110, 130]
    left_positions = []
    cumulative = 0
    for w in col_widths:
        left_positions.append(cumulative)
        cumulative += w
    frozen_total_width = cumulative

    tid = "check-table-fs"
    toast_id = "copy-toast-fs"
    func_name = "copySymbolsFS"
    uid = "ctfs"

    sector_list = get_sectors_in_table(df_check, industry_sector_map)

    ts_colors_palette = ["#1b3a1b", "#2a4a1b", "#3a3a1b", "#4a3a1b", "#3a2a1b",
                         "#2b2b3a", "#3a1b2a", "#1b2a3a", "#3a3a2b", "#2a3a3a"]
    ts_header_colors = {}
    for i, ts in enumerate(ts_values):
        ts_header_colors[ts] = ts_colors_palette[i % len(ts_colors_palette)]
    header_row_h = 38

    filter_html = build_filter_html(uid, sector_list)

    style_css = f"""
    <style>
    {build_filter_css()}
    html, body {{ margin: 0; padding: 0; height: 100%; overflow: hidden; }}
    .search-bar {{
        position: sticky; top: 0; z-index: 10; background-color: #0e1117;
        padding: 10px 12px; display: flex; align-items: center; gap: 10px; border-bottom: 2px solid #444;
    }}
    .search-bar input {{
        background-color: #1a1d24; color: #fafafa; border: 1px solid #555; border-radius: 6px;
        padding: 8px 14px; font-size: 14px; width: 260px; outline: none;
    }}
    .search-bar input:focus {{ border-color: #00c853; box-shadow: 0 0 6px rgba(0,200,83,0.4); }}
    .search-bar input::placeholder {{ color: #888; }}
    .search-bar button {{
        background-color: #00c853; color: #fff; border: none; border-radius: 6px;
        padding: 8px 18px; font-size: 14px; font-weight: bold; cursor: pointer;
    }}
    .search-bar button:hover {{ background-color: #00e676; }}
    .search-bar .clear-btn {{ background-color: #555; }}
    .search-bar .clear-btn:hover {{ background-color: #777; }}
    .search-bar .result-text {{ color: #aaa; font-size: 13px; margin-left: 8px; }}
    .fs-scroll-wrapper {{ overflow: auto; height: calc(100vh - 160px); border: 1px solid #444; }}
    #{tid} {{ border-collapse: separate; border-spacing: 0; font-size: 11px; width: max-content; }}
    #{tid} th, #{tid} td {{
        padding: 4px 6px; border: 1px solid #444; background-color: #0e1117; color: #fafafa; line-height: 1.4;
    }}
    #{tid} td {{ vertical-align: top; }}
    #{tid} thead th {{ position: sticky; z-index: 3; background-color: #262730; white-space: nowrap; font-size: 12px; }}
    #{tid} thead tr:first-child th {{ top: 0; }}
    #{tid} thead tr:nth-child(2) th {{ top: {header_row_h}px; }}
    #{tid} .sticky-col {{ position: sticky; z-index: 2; background-color: #0e1117; }}
    #{tid} thead .sticky-col {{ z-index: 5; background-color: #262730; }}
    #{tid} .sticky-col-0 {{ left: {left_positions[0]}px; min-width: {col_widths[0]}px; max-width: {col_widths[0]}px; white-space: nowrap; }}
    #{tid} .sticky-col-1 {{ left: {left_positions[1]}px; min-width: {col_widths[1]}px; max-width: {col_widths[1]}px; text-align: right; }}
    #{tid} .sticky-col-2 {{ left: {left_positions[2]}px; min-width: {col_widths[2]}px; max-width: {col_widths[2]}px; text-align: right; }}
    #{tid} .sticky-col-3 {{ left: {left_positions[3]}px; min-width: {col_widths[3]}px; max-width: {col_widths[3]}px;
                            border-right: 3px solid #888; white-space: nowrap; }}
    #{tid} td.data-cell {{ min-width: 120px; }}
    #{tid} .stock-chip {{ padding: 1px 0; white-space: nowrap; }}
    #{tid} .stock-chip .sym-code {{ display: inline-block; min-width: 36px; color: #888; font-size: 10px; }}
    #{tid} tbody tr:hover td {{ background-color: #1a1d24; }}
    #{tid} tbody tr:hover .sticky-col {{ background-color: #1a1d24; }}
    .copyable-fs {{ cursor: pointer; }}
    .copyable-fs:hover {{ background-color: #2a2d34 !important; }}
    #{tid} td.search-hit {{
        background-color: rgba(0, 200, 83, 0.18) !important; box-shadow: inset 0 0 0 2px #00c853;
    }}
    #{tid} td.search-hit .search-match {{
        background-color: #00c853; color: #000; border-radius: 3px; padding: 1px 4px;
        animation: pulse-glow 1.2s ease-in-out 3;
    }}
    @keyframes pulse-glow {{
        0%, 100% {{ box-shadow: 0 0 4px #00c853; }}
        50% {{ box-shadow: 0 0 16px #00e676, 0 0 30px rgba(0,230,118,0.4); }}
    }}
    #{toast_id} {{
        position: fixed; top: 20px; right: 20px; background-color: #00c853; color: white;
        padding: 10px 20px; border-radius: 8px; font-size: 14px; font-weight: bold;
        z-index: 9999; opacity: 0; transition: opacity 0.3s; pointer-events: none;
    }}
    #{toast_id}.show {{ opacity: 1; }}
    </style>
    """

    table_html = style_css
    table_html += f'<div id="{toast_id}" class="copy-toast">📋 Copied!</div>'
    table_html += filter_html

    table_html += """
    <div class="search-bar" id="search-bar-area">
        <input type="text" id="symbol-search" placeholder="🔍 銘柄コードを入力 (例: 6946)"
               onkeydown="if(event.key==='Enter') searchSymbol();" />
        <button onclick="searchSymbol()">検索</button>
        <button class="clear-btn" onclick="clearSearchAndInput()">クリア</button>
        <span id="search-result" class="result-text"></span>
    </div>
    """

    table_html += '<div class="fs-scroll-wrapper" id="fs-scroll-wrapper">'
    table_html += f'<table id="{tid}">'
    table_html += "<thead><tr>"
    for i, label in enumerate(["業種", "RS", "BP", "ステータス"]):
        table_html += f'<th rowspan="2" class="sticky-col sticky-col-{i}">{label}</th>'
    for ts in ts_values:
        colspan = len(ts_fs_map[ts])
        bg = ts_header_colors.get(ts, "#262730")
        table_html += f'<th colspan="{colspan}" style="background-color:{bg}; text-align:center; font-size:13px;">TS {ts}</th>'
    table_html += "</tr><tr>"
    for ts in ts_values:
        bg = ts_header_colors.get(ts, "#262730")
        for fs in ts_fs_map[ts]:
            table_html += f'<th style="background-color:{bg}; font-size:11px; text-align:center;">FS {fs}</th>'
    table_html += "</tr></thead><tbody>"

    for _, row in df_check.iterrows():
        bp = row['Buy Pressure']
        bp_color = get_color_from_buy_pressure(bp)
        industry_name = str(row['業種'])
        industry_esc = html.escape(industry_name)
        rs = f"{row['RS Rating']:.1f}"
        bp_val = f"{bp:.3f}"
        status_raw = str(row['ステータス'])
        status_display = re.sub(r'^\d+[a-z]?\s+', '', status_raw)
        status = html.escape(status_display)
        status_key = get_status_key(bp)
        sector = html.escape(industry_sector_map.get(industry_name, 'Unknown'))

        table_html += f'<tr data-status="{status_key}" data-sector="{sector}">'
        table_html += f'<td class="sticky-col sticky-col-0">{industry_esc}</td>'
        table_html += f'<td class="sticky-col sticky-col-1">{rs}</td>'
        table_html += f'<td class="sticky-col sticky-col-2" style="color:{bp_color}; font-weight:bold;">{bp_val}</td>'
        table_html += f'<td class="sticky-col sticky-col-3">{status}</td>'

        for ts, fs in all_sub_cols:
            display_html, copy_text = get_colored_symbols_html_with_fs(
                industry_name, ts, fs, df_screening_disp
            )
            if display_html:
                escaped_copy = html.escape(copy_text).replace("'", "\\'")
                table_html += (
                    f'<td class="data-cell copyable-fs" '
                    f'onclick="{func_name}(this, \'{escaped_copy}\')" '
                    f'title="コピー: {escaped_copy}">{display_html}</td>'
                )
            else:
                table_html += '<td class="data-cell"></td>'
        table_html += "</tr>"

    table_html += "</tbody></table></div>"

    table_html += f"""
    <script>
    var FROZEN_WIDTH = {frozen_total_width};

    function {func_name}(el, text) {{
        navigator.clipboard.writeText(text).then(function() {{
            var toast = document.getElementById('{toast_id}');
            toast.classList.add('show');
            el.style.backgroundColor = '#1b5e20';
            setTimeout(function() {{ toast.classList.remove('show'); el.style.backgroundColor = ''; }}, 1500);
        }});
    }}

    {build_filter_js(uid, tid)}

    function clearHighlights() {{
        var table = document.getElementById('{tid}');
        if (!table) return;
        var hitCells = table.querySelectorAll('td.search-hit');
        hitCells.forEach(function(td) {{
            td.classList.remove('search-hit');
            var matchSpans = td.querySelectorAll('.search-match');
            matchSpans.forEach(function(sp) {{ sp.classList.remove('search-match'); }});
        }});
        document.getElementById('search-result').textContent = '';
    }}
    function clearSearchAndInput() {{
        document.getElementById('symbol-search').value = '';
        clearHighlights();
    }}
    function scrollToCell(cell) {{
        var wrapper = document.getElementById('fs-scroll-wrapper');
        if (!wrapper || !cell) return;
        var wrapperRect = wrapper.getBoundingClientRect();
        var cellOffsetLeft = cell.offsetLeft;
        var targetScrollLeft = cellOffsetLeft - FROZEN_WIDTH - 20;
        if (targetScrollLeft < 0) targetScrollLeft = 0;
        var cellRect = cell.getBoundingClientRect();
        var headerHeight = 80;
        var targetScrollTop = wrapper.scrollTop + (cellRect.top - wrapperRect.top) - headerHeight;
        if (targetScrollTop < 0) targetScrollTop = 0;
        wrapper.scrollTo({{ top: targetScrollTop, left: targetScrollLeft, behavior: 'smooth' }});
    }}
    function searchSymbol() {{
        var query = document.getElementById('symbol-search').value.trim().toUpperCase();
        var resultEl = document.getElementById('search-result');
        var table = document.getElementById('{tid}');
        clearHighlights();
        document.getElementById('symbol-search').value = query;
        if (!query) {{ resultEl.textContent = ''; return; }}
        var keywords = query.split(/[,\\s]+/).filter(function(k) {{ return k.length > 0; }});
        var hitCount = 0;
        var firstHit = null;
        var allSpans = table.querySelectorAll('td.data-cell span[data-symbol]');
        allSpans.forEach(function(span) {{
            var sym = span.getAttribute('data-symbol').toUpperCase();
            var matched = false;
            for (var i = 0; i < keywords.length; i++) {{
                if (sym === keywords[i]) {{ matched = true; break; }}
            }}
            if (matched) {{
                span.classList.add('search-match');
                var parentTd = span.closest('td');
                if (parentTd && !parentTd.classList.contains('search-hit')) {{
                    parentTd.classList.add('search-hit');
                    hitCount++;
                    if (!firstHit) firstHit = parentTd;
                }}
            }}
        }});
        if (hitCount > 0) {{
            resultEl.textContent = '✅ ' + hitCount + ' 件ヒット';
            resultEl.style.color = '#00c853';
            scrollToCell(firstHit);
        }} else {{
            resultEl.textContent = '❌ 該当なし';
            resultEl.style.color = '#ff5252';
        }}
    }}
    document.addEventListener('click', function(e) {{
        var table = document.getElementById('{tid}');
        var searchBar = document.getElementById('search-bar-area');
        var toast = document.getElementById('{toast_id}');
        if (table && table.contains(e.target)) return;
        if (searchBar && searchBar.contains(e.target)) return;
        if (toast && toast.contains(e.target)) return;
        clearSearchAndInput();
    }});
    </script>
    """

    row_height = 42
    header_height = 90
    search_bar_height = 60
    filter_bar_height = 80
    padding = 20
    calculated = filter_bar_height + search_bar_height + header_height + num_rows * row_height + padding
    iframe_height = min(calculated, 2000)
    st.components.v1.html(table_html, height=iframe_height, scrolling=True)


# ============================================================
# 解説タブ用
# ============================================================
def render_kaisetsu_tab(df_industry_disp, df_screening_disp, df_description):
    st.header("📖 注目業種の解説")
    st.markdown(
        "RS Rating **80以上** かつ Buy Pressure **BUY基準以上（> 0.55）** を満たす業種を抽出し、"
        "それぞれの業種について解説します。"
    )
    st.markdown("---")
    df_qualified = df_industry_disp[
        (df_industry_disp['RS_Rating'] >= 80) & (df_industry_disp['Buy_Pressure'] > 0.55)
    ].sort_values('RS_Rating', ascending=False).copy()
    if len(df_qualified) == 0:
        st.warning("⚠️ 現在、RS Rating ≥ 80 かつ Buy Pressure > 0.55 を満たす業種はありません。")
        return
    st.success(f"✅ 条件を満たす業種: **{len(df_qualified)}** 件")
    summary_rows = []
    for _, row in df_qualified.iterrows():
        status = get_buy_pressure_status_display(row['Buy_Pressure'])
        stock_count = len(df_screening_disp[df_screening_disp['Industry'] == row['Industry']])
        summary_rows.append({
            '業種': row['Industry'], 'RS Rating': row['RS_Rating'],
            'Buy Pressure': row['Buy_Pressure'], 'ステータス': status, '対象銘柄数': stock_count,
        })
    df_qualified_summary = pd.DataFrame(summary_rows)
    st.dataframe(df_qualified_summary, use_container_width=True, hide_index=True)
    st.markdown("---")

    desc_map = {}
    desc_available = False
    if df_description is not None and len(df_description) > 0:
        desc_available = True
        industry_col = None
        for col in df_description.columns:
            if col.lower() in ['industry', '業種', 'industry_name']:
                industry_col = col
                break
        if industry_col is None:
            industry_col = df_description.columns[0]
        desc_cols = [c for c in df_description.columns if c != industry_col]
        for _, desc_row in df_description.iterrows():
            ind_name = str(desc_row[industry_col]).strip()
            desc_data = {}
            for col in desc_cols:
                val = desc_row[col]
                if pd.notna(val) and str(val).strip():
                    desc_data[col] = str(val).strip()
            desc_map[ind_name] = desc_data

    for _, row in df_qualified.iterrows():
        industry_name = row['Industry']
        rs_rating = row['RS_Rating']
        buy_pressure = row['Buy_Pressure']
        status = get_buy_pressure_status_display(buy_pressure)
        st.markdown(f"## {industry_name}")
        col1, col2, col3 = st.columns([1, 1, 1])
        with col1:
            st.metric("RS Rating", f"{rs_rating:.1f}")
        with col2:
            st.metric("Buy Pressure", f"{buy_pressure:.3f}")
        with col3:
            st.markdown(f"**ステータス: {status}**")
        if desc_available and industry_name in desc_map:
            desc_data = desc_map[industry_name]
            if desc_data:
                for col_name, col_value in desc_data.items():
                    st.markdown(f"**{col_name}**")
                    st.markdown(col_value)
            else:
                st.info(f"ℹ️ '{industry_name}' の解説データは空です。")
        elif desc_available:
            matched_key = None
            for key in desc_map:
                if key.lower() == industry_name.lower():
                    matched_key = key
                    break
            if matched_key is None:
                for key in desc_map:
                    if key.lower() in industry_name.lower() or industry_name.lower() in key.lower():
                        matched_key = key
                        break
            if matched_key:
                desc_data = desc_map[matched_key]
                if desc_data:
                    for col_name, col_value in desc_data.items():
                        st.markdown(f"**{col_name}**")
                        st.markdown(col_value)
                else:
                    st.info(f"ℹ️ '{industry_name}' の解説データは空です。")
            else:
                st.info(f"ℹ️ '{industry_name}' の解説データが見つかりません。")
        else:
            st.warning("⚠️ Industry_Description.xlsx が読み込めなかったため、解説を表示できません。")
        stocks_in_industry = df_screening_disp[
            df_screening_disp['Industry'] == industry_name
        ].sort_values('Technical_Score', ascending=False).head(10)
        if len(stocks_in_industry) > 0:
            with st.expander(f"📋 {industry_name} の対象銘柄（上位{min(10, len(stocks_in_industry))}件）", expanded=False):
                display_df = stocks_in_industry[
                    ['Symbol', 'Company Name', 'Technical_Score', 'Screening_Score', 'Buy_Pressure']
                ].copy()
                display_df = display_df.reset_index(drop=True)
                display_df.index = display_df.index + 1
                display_df.index.name = 'No'
                display_df.columns = ['Symbol', 'Company Name', 'Technical Score', 'Screening Score', 'Buy Pressure']
                display_df['Company Name'] = display_df['Company Name'].apply(
                    lambda x: str(x)[:40] if pd.notna(x) else ''
                )
                styled_df = display_df.style.apply(style_symbol_black_bg, axis=1)
                st.dataframe(styled_df, use_container_width=True)
        st.markdown("---")


# ============================================================
# タブの描画
# ============================================================
with tab0:
    df_check = df_summary[['業種', 'RS Rating', 'Buy Pressure', 'ステータス']].copy()
    render_check_tab(df_check, df_screening_display, df_industry_display, table_id_suffix="")

with tab0b:
    df_check2 = df_summary[['業種', 'RS Rating', 'Buy Pressure', 'ステータス']].copy()
    render_check_tab_with_fs(df_check2, df_screening_display)

with tab1:
    st.header("テクニカルスコア別 業種×銘柄マトリックス")
    create_industry_table(df_screening_display, df_industry_display, sort_by='Technical_Score')

with tab2:
    st.header("スクリーニングスコア (テクニカル+ファンダメンタル) 別 業種×銘柄マトリックス")
    create_industry_table(df_screening_display, df_industry_display, sort_by='Screening_Score')

with tab3:
    st.header("業種別サマリー統計")
    st.dataframe(
        df_summary, use_container_width=True, height=600,
        column_config={
            'ステータス': st.column_config.TextColumn('ステータス', help='クリックでソート', width='medium'),
        },
    )
    st.subheader("RS Rating vs Buy Pressure")
    STATUS_COLOR_MAP = {
        '0a 💀 WEAK': '#636EFA', '0b ⚠️ CAUTION': '#EF553B', '0c ➖ NEUTRAL': '#00CC96',
        '1 📈 BUY': '#1a3ab5', '2 🚀 STRONG': '#6fa8dc', '3 🔥 EXTREME': '#d84315',
    }
    STATUS_ORDER = [
        '0a 💀 WEAK', '0b ⚠️ CAUTION', '0c ➖ NEUTRAL',
        '1 📈 BUY', '2 🚀 STRONG', '3 🔥 EXTREME',
    ]
    fig = px.scatter(
        df_summary, x='RS Rating', y='Buy Pressure', size='銘柄数', color='ステータス',
        hover_data=['業種', '平均テクニカルスコア'], text='業種', title='業種別 RS Rating vs Buy Pressure',
        color_discrete_map=STATUS_COLOR_MAP, category_orders={'ステータス': STATUS_ORDER},
    )
    fig.update_traces(textposition='top center')
    fig.update_layout(height=700, yaxis=dict(range=[0.3, 1]))
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("業種別BPランキング")
    df_bp_ranking = df_all_industry.copy()
    df_bp_ranking['Sector'] = df_bp_ranking['Industry'].map(industry_sector_map).fillna('Unknown')
    sector_avg_bp = df_bp_ranking.groupby('Sector')['Buy_Pressure'].mean().sort_values(ascending=False)
    sorted_sectors = sector_avg_bp.index.tolist()
    for sector in sorted_sectors:
        df_sector = df_bp_ranking[df_bp_ranking['Sector'] == sector].copy()
        df_sector = df_sector.sort_values('RS_Rating', ascending=True)
        if len(df_sector) == 0:
            continue
        sector_avg = df_sector['Buy_Pressure'].mean()
        rs80_count = len(df_sector[df_sector['RS_Rating'] >= 80])
        total_count = len(df_sector)
        st.markdown(f"#### 📂 {sector}（平均BP: {sector_avg:.3f}　RS≧80: {rs80_count}/{total_count}）")
        fig_sector = px.bar(
            df_sector, x='Buy_Pressure', y='Industry', orientation='h', color='RS_Rating',
            color_continuous_scale=CUSTOM_RS_COLORSCALE, range_color=[0, 100],
            labels={'Buy_Pressure': 'Buy Pressure', 'Industry': '業種', 'RS_Rating': 'RS Rating'},
        )
        fig_sector.add_vline(
            x=0.550, line_dash="dot", line_color="black", line_width=2,
            annotation_text="BUY (0.550)", annotation_position="top",
            annotation_font_size=11, annotation_font_color="black",
        )
        fig_sector.update_layout(
            height=max(len(df_sector) * 30 + 80, 150), yaxis=dict(dtick=1),
            coloraxis_colorbar=dict(title='RS Rating'), margin=dict(t=40, b=20), showlegend=False,
        )
        st.plotly_chart(fig_sector, use_container_width=True)

with tab4:
    render_kaisetsu_tab(df_industry_display, df_screening_display, df_description)

st.markdown("---")
st.markdown(
    f'<div style="text-align: center; color: gray; font-size: 12px;">'
    f'Industry Buy Pressure Dashboard (JP) | Data: {data_date}</div>',
    unsafe_allow_html=True
)
