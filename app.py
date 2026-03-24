import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime, timedelta
import os
import glob
import numpy as np
import gc
import psutil
import yfinance as yf

# ページ設定
st.set_page_config(
    page_title="米国株RSダッシュボード",
    page_icon="📈",
    layout="wide"
)

# レスポンシブ対応のカスタムCSS
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    .stDeployButton {display:none;}
    header {visibility: hidden;}
    footer {visibility: hidden;}
    [data-testid="stToolbar"] {display: none;}
    .styles_viewerBadge__1yB5_ {display: none;}
    .viewerBadge_container__1QSob {display: none;}
    .viewerBadge_link__1S137 {display: none;}
    .viewerBadge_text__1JaDK {display: none;}
    button[kind="header"] {display: none;}
    .stActionButton {display: none;}
    [data-testid="manage-app-button"] {display: none;}
    .stAppDeployButton {display: none;}
    #stDecoration {display: none;}
    div[data-testid="stStatusWidget"] {display: none;}
    .streamlit-footer {display: none;}
    ._profileContainer_gzau3_53 {display: none;}
    ._profilePreview_gzau3_63 {display: none;}
    @media (max-width: 768px) {
        [data-testid="stSidebar"] { width: 250px !important; }
        .main .block-container {
            padding-left: 1rem !important;
            padding-right: 1rem !important;
            padding-top: 1rem !important;
        }
        h1 { font-size: 1.5rem !important; }
        h2 { font-size: 1.3rem !important; }
        h3 { font-size: 1.1rem !important; }
        .dataframe { font-size: 0.8rem !important; }
        .stButton > button { width: 100% !important; }
        .stNumberInput { font-size: 0.9rem !important; }
        [data-testid="column"] { width: 100% !important; flex: 100% !important; }
        .stTabs [data-baseweb="tab-list"] button {
            font-size: 0.9rem !important;
            padding: 8px 12px !important;
        }
        .market-status { padding: 15px !important; margin-bottom: 15px !important; }
        .market-status h2 { font-size: 1.2rem !important; }
        .market-status p { font-size: 0.9rem !important; }
        .js-plotly-plot { min-height: 400px !important; }
    }
    @media (min-width: 769px) and (max-width: 1024px) {
        .main .block-container {
            padding-left: 2rem !important;
            padding-right: 2rem !important;
        }
        h1 { font-size: 1.8rem !important; }
    }
    .market-status { border-radius: 10px; transition: all 0.3s ease; }
    .dataframe-container { overflow-x: auto; -webkit-overflow-scrolling: touch; }
</style>
""", unsafe_allow_html=True)

st.title("📈 米国株RS分析ダッシュボード")
st.markdown("---")

DATA_FOLDER = "data"

if 'debug_mode' not in st.session_state:
    st.session_state.debug_mode = False


def get_memory_usage():
    process = psutil.Process()
    memory_info = process.memory_info()
    return memory_info.rss / 1024 / 1024


def get_display_date(screening_date):
    return screening_date - timedelta(days=1)


def get_year_month_from_date(date):
    return date.strftime('%Y年%m月')


def filter_data_by_month(all_data, selected_month):
    filtered = []
    for data in all_data:
        display_date = get_display_date(data['date'])
        if get_year_month_from_date(display_date) == selected_month:
            filtered.append(data)
    return filtered


def get_available_months_by_display_date(all_data):
    months = set()
    for data in all_data:
        display_date = get_display_date(data['date'])
        months.add(get_year_month_from_date(display_date))
    return sorted(list(months), reverse=True)


@st.cache_data(ttl=3600)
def get_days_since_high_from_yahoo(symbols, period="3mo"):
    days_since_high_dict = {}
    progress_bar = st.progress(0)
    status_text = st.empty()
    for idx, symbol in enumerate(symbols):
        try:
            ticker = yf.Ticker(symbol)
            hist = ticker.history(period=period)
            if not hist.empty:
                max_high_date = hist['High'].idxmax()
                max_high_idx = hist.index.get_loc(max_high_date)
                current_idx = len(hist) - 1
                days_since = current_idx - max_high_idx
                days_since_high_dict[symbol] = days_since
            else:
                days_since_high_dict[symbol] = None
        except Exception:
            days_since_high_dict[symbol] = None
        progress = (idx + 1) / len(symbols)
        progress_bar.progress(progress)
        status_text.text(f"Yahoo Finance データ取得中... {idx + 1}/{len(symbols)}")
    progress_bar.empty()
    status_text.empty()
    return days_since_high_dict


@st.cache_data(ttl=300)
def load_data_from_github_optimized(max_files=None, debug=False):
    all_data = []
    debug_info = {
        'total_files': 0, 'loaded_files': 0, 'failed_files': [],
        'memory_before': get_memory_usage(), 'file_details': []
    }
    try:
        if os.path.exists(DATA_FOLDER):
            excel_files = glob.glob(os.path.join(DATA_FOLDER, "*.xlsx")) + \
                         glob.glob(os.path.join(DATA_FOLDER, "*.xls"))
            excel_files = sorted(excel_files)
            debug_info['total_files'] = len(excel_files)
            if max_files:
                excel_files = excel_files[:max_files]
            for idx, file_path in enumerate(excel_files):
                try:
                    filename = os.path.basename(file_path)
                    file_size = os.path.getsize(file_path) / (1024 * 1024)
                    if debug:
                        debug_info['file_details'].append({
                            'index': idx + 1, 'filename': filename,
                            'size_mb': round(file_size, 2), 'memory_before': get_memory_usage()
                        })
                    date_str = filename.split('_')[2] if len(filename.split('_')) > 2 else None
                    if date_str:
                        try:
                            date = datetime.strptime(date_str, '%Y%m%d')
                        except:
                            date = datetime.now()
                    else:
                        date = datetime.now()
                    with pd.ExcelFile(file_path) as excel:
                        df_main = excel.parse('Screening_Results')
                        df_main['Date'] = date
                        essential_columns = [
                            'Symbol', 'Company Name', 'Sector', 'Industry',
                            'Screening_Score', 'Technical_Score', 'Fundamental_Score',
                            'RS_Score', 'Individual_RS', 'Individual_RS_Percentile',
                            'Sector_RS', 'Sector_RS_Percentile',
                            'Industry_RS', 'Industry_RS_Percentile',
                            'Current_Price', 'MA21', 'MA50', 'MA150',
                            'ATR_Pct_from_MA50', 'ADR', 'Date',
                            'Days_Since_High', 'High_52W', 'sales_accel_3_qtrs', 'eps_accel_3_qtrs'
                        ]
                        available_columns = [col for col in essential_columns if col in df_main.columns]
                        df_main = df_main[available_columns]
                        market_summary = None
                        try:
                            if 'Market_Summary' in excel.sheet_names:
                                market_summary_df = excel.parse('Market_Summary', nrows=5)
                                if len(market_summary_df) >= 4:
                                    market_summary = {
                                        'status': market_summary_df.iloc[3, 1],
                                        'score': market_summary_df.iloc[2, 1]
                                    }
                        except:
                            pass
                    all_data.append({
                        'date': date, 'display_date': get_display_date(date),
                        'main': df_main, 'market_summary': market_summary,
                        'filename': filename, 'filepath': file_path
                    })
                    debug_info['loaded_files'] += 1
                    if debug:
                        debug_info['file_details'][-1]['memory_after'] = get_memory_usage()
                        debug_info['file_details'][-1]['status'] = 'success'
                    if idx % 5 == 0:
                        gc.collect()
                except Exception as e:
                    debug_info['failed_files'].append({'filename': filename, 'error': str(e)})
                    if debug:
                        debug_info['file_details'][-1]['status'] = 'failed'
                        debug_info['file_details'][-1]['error'] = str(e)
    except Exception as e:
        st.error(f"データ読み込み中に重大なエラーが発生しました: {e}")
    finally:
        debug_info['memory_after'] = get_memory_usage()
        debug_info['memory_used'] = debug_info['memory_after'] - debug_info['memory_before']
        gc.collect()
    return all_data, debug_info


# サイドバー
st.sidebar.header("📁 データ管理")
st.sidebar.checkbox("🔧 デバッグモード", key="debug_mode")

if st.session_state.debug_mode:
    st.sidebar.markdown("### 💾 システム情報")
    st.sidebar.info(f"現在のメモリ使用量: {get_memory_usage():.2f} MB")
    max_files = st.sidebar.number_input(
        "最大読み込みファイル数（0=無制限）", min_value=0, max_value=100, value=0,
        help="メモリ不足の場合は制限してください"
    )
else:
    max_files = None

if st.sidebar.button("🔄 最新データを再読み込み", help="GitHubから最新のファイルを読み込みます"):
    st.cache_data.clear()
    st.rerun()

with st.spinner("データ読み込み中..."):
    all_data, debug_info = load_data_from_github_optimized(
        max_files=max_files if max_files and max_files > 0 else None,
        debug=st.session_state.debug_mode
    )

if st.session_state.debug_mode and debug_info:
    with st.sidebar.expander("📊 読み込み統計"):
        st.write(f"検出ファイル数: {debug_info['total_files']}")
        st.write(f"読み込み成功: {debug_info['loaded_files']}")
        st.write(f"読み込み失敗: {len(debug_info['failed_files'])}")
        st.write(f"メモリ使用量: {debug_info['memory_used']:.2f} MB")
        if debug_info['failed_files']:
            st.error("失敗したファイル:")
            for failed in debug_info['failed_files']:
                st.write(f"- {failed['filename']}: {failed['error']}")

if all_data:
    st.sidebar.success(f"✅ {len(all_data)}ファイル読み込み完了")
    all_data.sort(key=lambda x: x['date'])
    latest_data = all_data[-1]['main']
    latest_market = all_data[-1]['market_summary']
    available_months = get_available_months_by_display_date(all_data)

    if latest_market:
        st.markdown("### 📊 現在のマーケット状況")

        def get_status_colors(status):
            if 'Strong Positive' in status:
                return "#d4edda", "#155724", "🟢"
            elif 'Positive' in status:
                return "#e8f5e9", "#2e7d32", "🟢"
            elif 'Neutral' in status:
                return "#fff3cd", "#856404", "🟡"
            elif 'Negative' in status:
                return "#ffebee", "#c62828", "🔴"
            else:
                return "#f8d7da", "#721c24", "🔴"

        status = latest_market['status']
        bg_color, text_color, emoji = get_status_colors(status)
        st.markdown(
            f"""
            <div class="market-status" style="padding: 20px; border-radius: 10px; background-color: {bg_color}; margin-bottom: 20px;">
                <h2 style="margin: 0; color: {text_color};">{emoji} {status}</h2>
                <p style="margin: 5px 0 0 0; font-size: 18px; color: {text_color};">スコア: {latest_market['score']}</p>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("---")

    # ===== 6タブ定義 =====
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📊 セクターRS",
        "🏭 業種RS",
        "📈 マーケット詳細",
        "🚀 モメンタム銘柄",
        "💎 ベース銘柄タイプ①",
        "📋 運用ルール"
    ])

    # ===== タブ1: セクターRS分析 =====
    with tab1:
        st.header("セクターRSランキング推移")
        if len(available_months) > 0:
            col1, col2 = st.columns([2, 8])
            with col1:
                selected_month_sector = st.selectbox(
                    "表示する月を選択", available_months, index=0,
                    key="sector_month_selector"
                )
            month_data = filter_data_by_month(all_data, selected_month_sector)
            if len(month_data) > 1:
                sector_timeseries = []
                for data_point in month_data:
                    df = data_point['main']
                    display_date = data_point['display_date']
                    sector_rs = df.groupby('Sector').agg({
                        'Sector_RS': 'first', 'Sector_RS_Percentile': 'first'
                    }).reset_index()
                    sector_rs = sector_rs.sort_values('Sector_RS', ascending=False).reset_index(drop=True)
                    sector_rs['RS_Ranking'] = range(1, len(sector_rs) + 1)
                    sector_rs['Date'] = display_date
                    sector_timeseries.append(sector_rs)
                sector_ts_df = pd.concat(sector_timeseries, ignore_index=True)
                pivot_data = sector_ts_df.pivot(index='Sector', columns='Date', values='RS_Ranking')
                latest_date = pivot_data.columns[-1]
                pivot_data = pivot_data.sort_values(by=latest_date, ascending=False)
                max_rank = int(pivot_data.values.max())
                inverted_data = max_rank + 1 - pivot_data.values
                display_dates = [d.strftime('%m/%d') for d in pivot_data.columns]
                fig = go.Figure(data=go.Heatmap(
                    z=inverted_data, x=display_dates, y=pivot_data.index,
                    colorscale='RdYlGn', text=pivot_data.values,
                    texttemplate='%{text}', textfont={"size": 9},
                    colorbar=dict(
                        title="ランク", tickmode='array',
                        tickvals=[max_rank, max_rank * 0.75, max_rank * 0.5, max_rank * 0.25, 1],
                        ticktext=['1', str(int(max_rank * 0.25)), str(int(max_rank * 0.5)),
                                  str(int(max_rank * 0.75)), str(max_rank)]
                    ),
                    hoverongaps=False, zmin=1, zmax=max_rank
                ))
                fig.update_layout(
                    title=f"セクターRSランキング推移 - {selected_month_sector}",
                    xaxis_title="日付", yaxis_title="セクター", height=600,
                    xaxis={'side': 'bottom'}, font=dict(size=10),
                    margin=dict(l=150, r=50, t=80, b=80)
                )
                st.plotly_chart(fig, use_container_width=True)
                st.info(f"📊 {selected_month_sector}のデータ数: {len(month_data)}日分")
                del sector_timeseries, sector_ts_df, pivot_data
                gc.collect()
            else:
                st.info(f"{selected_month_sector}のデータが不足しています（最低2日分必要）")

        st.subheader("最新セクターランキング")
        sector_ranking = latest_data.groupby('Sector').agg({
            'Sector_RS': 'first', 'Sector_RS_Percentile': 'first', 'Symbol': 'count'
        }).reset_index()
        sector_ranking.columns = ['Sector', 'Sector_RS', 'Percentile', 'Stock_Count']
        sector_ranking = sector_ranking.sort_values('Sector_RS', ascending=False).reset_index(drop=True)
        sector_ranking.insert(0, 'Rank', range(1, len(sector_ranking) + 1))
        st.dataframe(sector_ranking, use_container_width=True, height=400)

    # ===== タブ2: 業種RS分析 =====
    with tab2:
        st.header("業種RSランキング推移（上位20%）")
        if len(available_months) > 0:
            col1, col2 = st.columns([2, 8])
            with col1:
                selected_month_industry = st.selectbox(
                    "表示する月を選択", available_months, index=0,
                    key="industry_month_selector"
                )
            month_data_industry = filter_data_by_month(all_data, selected_month_industry)
            if len(month_data_industry) > 1:
                industry_timeseries = []
                for data_point in month_data_industry:
                    df = data_point['main']
                    display_date = data_point['display_date']
                    industry_rs = df.groupby('Industry').agg({
                        'Industry_RS': 'first', 'Industry_RS_Percentile': 'first'
                    }).reset_index()
                    industry_rs = industry_rs.sort_values('Industry_RS', ascending=False).reset_index(drop=True)
                    industry_rs['RS_Ranking'] = range(1, len(industry_rs) + 1)
                    industry_rs['Date'] = display_date
                    industry_timeseries.append(industry_rs)
                industry_ts_df = pd.concat(industry_timeseries, ignore_index=True)
                latest_industries = industry_ts_df[industry_ts_df['Date'] == month_data_industry[-1]['display_date']]
                top_20_pct_count = max(1, int(len(latest_industries) * 0.2))
                top_industries = latest_industries.nsmallest(top_20_pct_count, 'RS_Ranking')['Industry'].tolist()
                industry_ts_filtered = industry_ts_df[industry_ts_df['Industry'].isin(top_industries)]
                pivot_data = industry_ts_filtered.pivot(index='Industry', columns='Date', values='RS_Ranking')
                latest_date = pivot_data.columns[-1]
                pivot_data = pivot_data.sort_values(by=latest_date, ascending=False)
                max_rank = int(pivot_data.values.max())
                inverted_data = max_rank + 1 - pivot_data.values
                display_dates = [d.strftime('%m/%d') for d in pivot_data.columns]
                fig = go.Figure(data=go.Heatmap(
                    z=inverted_data, x=display_dates, y=pivot_data.index,
                    colorscale='RdYlGn', text=pivot_data.values,
                    texttemplate='%{text}', textfont={"size": 8},
                    colorbar=dict(
                        title="ランク", tickmode='array',
                        tickvals=[max_rank, max_rank * 0.8, max_rank * 0.6, max_rank * 0.4, max_rank * 0.2, 1],
                        ticktext=['1', str(int(max_rank * 0.2 + 1)), str(int(max_rank * 0.4 + 1)),
                                  str(int(max_rank * 0.6 + 1)), str(int(max_rank * 0.8 + 1)), str(max_rank)]
                    ),
                    hoverongaps=False, zmin=1, zmax=max_rank
                ))
                fig.update_layout(
                    title=f"業種RSランキング推移（上位{top_20_pct_count}業種） - {selected_month_industry}",
                    xaxis_title="日付", yaxis_title="業種", height=800,
                    xaxis={'side': 'bottom'}, font=dict(size=9),
                    margin=dict(l=180, r=50, t=80, b=80)
                )
                st.plotly_chart(fig, use_container_width=True)
                st.info(f"📊 {selected_month_industry}のデータ数: {len(month_data_industry)}日分")
                del industry_timeseries, industry_ts_df, pivot_data
                gc.collect()
            else:
                st.info(f"{selected_month_industry}のデータが不足しています（最低2日分必要）")

        st.subheader("最新業種ランキング（上位30）")
        industry_ranking = latest_data.groupby('Industry').agg({
            'Industry_RS': 'first', 'Industry_RS_Percentile': 'first',
            'Sector': 'first', 'Symbol': 'count'
        }).reset_index()
        industry_ranking.columns = ['Industry', 'Industry_RS', 'Percentile', 'Sector', 'Stock_Count']
        industry_ranking = industry_ranking.sort_values('Industry_RS', ascending=False).reset_index(drop=True)
        industry_ranking.insert(0, 'Rank', range(1, len(industry_ranking) + 1))
        industry_ranking = industry_ranking.head(30)
        st.dataframe(industry_ranking, use_container_width=True, height=500)

        st.markdown("---")
        st.subheader("セクターローテーション")
        sectors_rotation = [
            {'name': 'Technology', 'color': '#A9A9A9', 'value': 1},
            {'name': 'Materials', 'color': '#808080', 'value': 1},
            {'name': 'Consumer Discretionary', 'color': '#696969', 'value': 1},
            {'name': 'Energy', 'color': '#FFD700', 'value': 1},
            {'name': 'Utilities', 'color': '#4169E1', 'value': 1},
            {'name': 'Consumer Staples', 'color': '#6495ED', 'value': 1},
            {'name': 'Healthcare', 'color': '#FF8C00', 'value': 1},
            {'name': 'Financials', 'color': '#FFA500', 'value': 1},
        ]
        colors = [s['color'] for s in sectors_rotation]
        values = [s['value'] for s in sectors_rotation]
        labels = [''] * len(sectors_rotation)
        fig_rotation = go.Figure(data=[go.Pie(
            labels=labels, values=values,
            marker=dict(colors=colors, line=dict(color='white', width=2)),
            textinfo='none', hoverinfo='skip', hole=0.0, direction='clockwise', sort=False
        )])
        fig_rotation.add_shape(type="line", x0=0, y0=-1.1, x1=0, y1=1.1,
                               line=dict(color="black", width=4), xref='x', yref='y')
        fig_rotation.add_shape(type="line", x0=-1.1, y0=0, x1=1.1, y1=0,
                               line=dict(color="black", width=4), xref='x', yref='y')
        fig_rotation.add_annotation(x=0, y=1.2, text="景気が強い", showarrow=False,
                                    font=dict(size=14, color="black", weight="bold"), xref='x', yref='y')
        fig_rotation.add_annotation(x=0, y=-1.2, text="景気が弱い", showarrow=False,
                                    font=dict(size=14, color="black", weight="bold"), xref='x', yref='y')
        fig_rotation.add_annotation(x=1.2, y=0, text="金利が高い", showarrow=False,
                                    font=dict(size=14, color="black", weight="bold"), textangle=-90, xref='x', yref='y')
        fig_rotation.add_annotation(x=-1.2, y=0, text="金利が低い", showarrow=False,
                                    font=dict(size=14, color="black", weight="bold"), textangle=90, xref='x', yref='y')
        fig_rotation.update_layout(
            showlegend=False, height=700, width=700,
            margin=dict(l=80, r=80, t=80, b=80),
            paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
            xaxis=dict(showgrid=False, showticklabels=False, zeroline=False, range=[-1.4, 1.4]),
            yaxis=dict(showgrid=False, showticklabels=False, zeroline=False, range=[-1.4, 1.4],
                       scaleanchor='x', scaleratio=1)
        )
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.plotly_chart(fig_rotation, use_container_width=False)
        st.caption("セクターローテーションの概念図")

    # ===== タブ3: マーケット詳細 =====
    with tab3:
        st.header("📈 マーケット状況の詳細判定")
        if latest_market:
            status = latest_market['status']
            bg_color, text_color, emoji = get_status_colors(status)
            st.markdown(
                f"""
                <div style="padding: 15px; border-radius: 10px; background-color: {bg_color}; margin-bottom: 20px; border-left: 5px solid {text_color};">
                    <h3 style="margin: 0; color: {text_color};">{emoji} 現在のマーケット状況: {status}</h3>
                    <p style="margin: 5px 0 0 0; color: {text_color};">スコア: {latest_market['score']}</p>
                </div>
                """,
                unsafe_allow_html=True
            )
        try:
            latest_data_info = all_data[-1]
            if latest_data_info.get('filepath'):
                latest_file_path = latest_data_info['filepath']
                with pd.ExcelFile(latest_file_path) as excel:
                    if 'Market_Condition' in excel.sheet_names:
                        market_condition_df = excel.parse('Market_Condition')
                display_date = latest_data_info['display_date'].strftime('%Y年%m月%d日')
                st.caption(f"📅 データ日付: {display_date}")
                columns_to_keep = [col for col in market_condition_df.columns
                                   if 'スコア' not in col and 'Score' not in col.lower()]
                market_condition_df = market_condition_df[columns_to_keep]

                def color_judgment(row):
                    judgment = row['判定']
                    if judgment == 'positive':
                        return ['background-color: #d4edda'] * len(row)
                    elif judgment == 'neutral':
                        return ['background-color: #fff3cd'] * len(row)
                    elif judgment == 'negative':
                        return ['background-color: #ffebee'] * len(row)
                    return [''] * len(row)

                styled_condition = market_condition_df.style.apply(color_judgment, axis=1)
                st.dataframe(styled_condition, use_container_width=True, hide_index=True, height=600)
                st.markdown("""
                **判定の色分け:**
                - 🟢 **Positive (緑)**: 強気シグナル
                - 🟡 **Neutral (黄)**: 中立シグナル
                - 🔴 **Negative (赤)**: 弱気シグナル
                """)
                st.markdown("---")
                st.subheader("📊 判定の統計")
                col1, col2, col3 = st.columns(3)
                positive_count = (market_condition_df['判定'] == 'positive').sum()
                neutral_count = (market_condition_df['判定'] == 'neutral').sum()
                negative_count = (market_condition_df['判定'] == 'negative').sum()
                total_count = len(market_condition_df)
                with col1:
                    st.metric("Positive", f"{positive_count}/{total_count}",
                              f"{positive_count / total_count * 100:.1f}%")
                with col2:
                    st.metric("Neutral", f"{neutral_count}/{total_count}",
                              f"{neutral_count / total_count * 100:.1f}%")
                with col3:
                    st.metric("Negative", f"{negative_count}/{total_count}",
                              f"{negative_count / total_count * 100:.1f}%")
                st.markdown("---")
                st.subheader("📊 判定の内訳")
                fig = go.Figure(data=[go.Pie(
                    labels=['Positive', 'Neutral', 'Negative'],
                    values=[positive_count, neutral_count, negative_count],
                    marker_colors=['#2e7d32', '#ffc107', '#dc3545']
                )])
                fig.update_layout(height=400, showlegend=True)
                st.plotly_chart(fig, use_container_width=True)
                st.markdown("---")
                st.subheader("📈 マーケットスコアの推移")
                col1, col2 = st.columns([2, 8])
                with col1:
                    selected_month_market = st.selectbox(
                        "表示する月を選択", available_months, index=0,
                        key="market_month_selector"
                    )
                month_data_market = filter_data_by_month(all_data, selected_month_market)
                score_data = []
                for data in month_data_market:
                    if data.get('market_summary') and data['market_summary'].get('score'):
                        score_str = data['market_summary']['score']
                        if isinstance(score_str, str):
                            score_value = float(score_str.replace('%', ''))
                        else:
                            score_value = float(score_str)
                        score_data.append({
                            'Date': data['display_date'], 'Score': score_value,
                            'Status': data['market_summary']['status']
                        })
                if score_data:
                    score_df = pd.DataFrame(score_data)
                    fig = go.Figure()
                    fig.add_trace(go.Scatter(
                        x=score_df['Date'], y=score_df['Score'],
                        mode='lines+markers', name='マーケットスコア',
                        line=dict(color='#2196F3', width=3), marker=dict(size=8),
                        hovertemplate='<b>日付</b>: %{x|%Y/%m/%d}<br><b>スコア</b>: %{y:.1f}%<br><extra></extra>'
                    ))
                    fig.add_hrect(y0=80, y1=100, fillcolor="green", opacity=0.1,
                                  annotation_text="Strong Positive", annotation_position="right")
                    fig.add_hrect(y0=60, y1=80, fillcolor="lightgreen", opacity=0.1,
                                  annotation_text="Positive", annotation_position="right")
                    fig.add_hrect(y0=40, y1=60, fillcolor="yellow", opacity=0.1,
                                  annotation_text="Neutral", annotation_position="right")
                    fig.add_hrect(y0=20, y1=40, fillcolor="orange", opacity=0.1,
                                  annotation_text="Negative", annotation_position="right")
                    fig.add_hrect(y0=0, y1=20, fillcolor="red", opacity=0.1,
                                  annotation_text="Strong Negative", annotation_position="right")
                    fig.update_layout(
                        title=f"マーケットスコアの日次推移 - {selected_month_market}",
                        xaxis_title="日付", yaxis_title="スコア (%)",
                        yaxis=dict(range=[0, 100]), height=500,
                        hovermode='x unified', showlegend=True
                    )
                    st.plotly_chart(fig, use_container_width=True)
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("現在のスコア", f"{score_df.iloc[-1]['Score']:.1f}%")
                    with col2:
                        st.metric("月平均スコア", f"{score_df['Score'].mean():.1f}%")
                    with col3:
                        st.metric("月内最高スコア", f"{score_df['Score'].max():.1f}%")
                    with col4:
                        st.metric("月内最低スコア", f"{score_df['Score'].min():.1f}%")
                    st.info(f"📊 {selected_month_market}のデータ数: {len(score_data)}日分")
                else:
                    st.info(f"{selected_month_market}のスコアデータがありません。")
            else:
                st.info("アップロードされたファイルではMarket_Conditionシートを表示できません。")
        except Exception as e:
            st.warning(f"マーケット状況の詳細を読み込めませんでした: {e}")
            if st.session_state.debug_mode:
                with st.expander("デバッグ情報"):
                    st.write(f"エラー詳細: {str(e)}")

    # ===== タブ4: モメンタム銘柄 =====
    with tab4:
        st.header("条件フィルタリング")
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("📊 テクニカル条件")
            enable_technical = st.checkbox("テクニカル条件を有効にする", value=True, key="enable_tech")
            if enable_technical:
                st.markdown("**ATR条件**")
                atr_min = st.number_input("ATR from MA50 最小値", value=2.0, step=0.1, key="atr_min")
                atr_max = st.number_input("ATR from MA50 最大値", value=5.0, step=0.1, key="atr_max")
                adr_min = st.number_input("ADR 最小値", value=4.0, step=0.5, key="adr_min")
                st.markdown("---")
                st.markdown("**移動平均線条件**")
                ma21_condition = st.checkbox("株価 > MA21（21日移動平均線）", value=True, key="ma21_check")
                ma50_condition = st.checkbox("株価 > MA50（50日移動平均線）", value=True, key="ma50_check")
                ma150_condition = st.checkbox("株価 > MA150（150日移動平均線）", value=True, key="ma150_check")
                ma_order_condition = st.checkbox("MA21 > MA50 > MA150（上昇トレンド）", value=True, key="ma_order_check")
            else:
                atr_min, atr_max, adr_min = 0.0, 100.0, 0.0
                ma21_condition = ma50_condition = ma150_condition = ma_order_condition = False
                st.info("テクニカル条件は無効になっています")
            st.markdown("---")
            st.markdown("**価格条件**")
            price_min = st.number_input("株価 最小値 ($", value=10.0, step=1.0, key="price_min")
            st.markdown("---")
            st.markdown("**ファンダメンタル条件**")
            enable_fundamental = st.checkbox("ファンダメンタル条件を有効にする", value=False, key="enable_fund")
            if enable_fundamental:
                fundamental_score_min = st.number_input(
                    "ファンダメンタルスコア最小値", min_value=0, max_value=10, value=10,
                    step=1, key="fund_score"
                )
            else:
                fundamental_score_min = 0
                st.info("ファンダメンタル条件は無効になっています")
        with col2:
            st.subheader("📈 RS条件")
            enable_rs = st.checkbox("RS条件を有効にする", value=True, key="enable_rs")
            if enable_rs:
                individual_rs_min = st.number_input("Individual RS Percentile 最小値", value=80, step=1, key="ind_rs_min")
                sector_rs_min = st.number_input("Sector RS Percentile 最小値", value=79, step=5, key="sec_rs_min")
                industry_rs_min = st.number_input("Industry RS Percentile 最小値", value=80, step=5, key="ind_rs_min2")
            else:
                individual_rs_min = sector_rs_min = industry_rs_min = 0
                st.info("RS条件は無効になっています")
            st.markdown("---")
            status_text = f"""
            **📋 現在の設定:**\n
            **テクニカル条件:** {'✅ 有効' if enable_technical else '❌ 無効'}
            """
            if enable_technical:
                status_text += f"""
            - ATR: {atr_min}% ~ {atr_max}%
            - ADR: {adr_min}% 以上
            - MA21条件: {'✅' if ma21_condition else '❌'}
            - MA50条件: {'✅' if ma50_condition else '❌'}
            - MA150条件: {'✅' if ma150_condition else '❌'}
            - MA順列: {'✅' if ma_order_condition else '❌'}
            """
            status_text += f"""
            \n**RS条件:** {'✅ 有効' if enable_rs else '❌ 無効'}
            """
            if enable_rs:
                status_text += f"""
            - 個別RS: {individual_rs_min}% 以上
            - セクターRS: {sector_rs_min}% 以上
            - 業種RS: {industry_rs_min}% 以上
            """
            status_text += f"""
            \n**価格条件:**
            - 最低価格: ${price_min}
            \n**ファンダメンタル条件:** {'✅ 有効' if enable_fundamental else '❌ 無効'}
            """
            if enable_fundamental:
                status_text += f"""
            - 最小スコア: {fundamental_score_min}点以上
            """
            st.info(status_text)

        st.markdown("---")
        filtered_df = latest_data.copy()
        if enable_technical:
            if 'ATR_Pct_from_MA50' in filtered_df.columns:
                filtered_df = filtered_df[(filtered_df['ATR_Pct_from_MA50'] >= atr_min) & (filtered_df['ATR_Pct_from_MA50'] <= atr_max)]
            if 'ADR' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['ADR'] >= adr_min]
            if ma21_condition and 'MA21' in filtered_df.columns and 'Current_Price' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['Current_Price'] > filtered_df['MA21']]
            if ma50_condition and 'MA50' in filtered_df.columns and 'Current_Price' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['Current_Price'] > filtered_df['MA50']]
            if ma150_condition and 'MA150' in filtered_df.columns and 'Current_Price' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['Current_Price'] > filtered_df['MA150']]
            if ma_order_condition and all(col in filtered_df.columns for col in ['MA21', 'MA50', 'MA150']):
                filtered_df = filtered_df[(filtered_df['MA21'] > filtered_df['MA50']) & (filtered_df['MA50'] > filtered_df['MA150'])]
        if enable_rs:
            if 'Individual_RS_Percentile' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['Individual_RS_Percentile'] >= individual_rs_min]
            if 'Sector_RS_Percentile' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['Sector_RS_Percentile'] >= sector_rs_min]
            if 'Industry_RS_Percentile' in filtered_df.columns:
                filtered_df = filtered_df[filtered_df['Industry_RS_Percentile'] >= industry_rs_min]
        if 'Current_Price' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Current_Price'] >= price_min]
        if enable_fundamental and 'Fundamental_Score' in filtered_df.columns:
            filtered_df = filtered_df[filtered_df['Fundamental_Score'] >= fundamental_score_min]

        st.subheader(f"🚀 フィルタリング結果: {len(filtered_df)}銘柄")
        if len(filtered_df) > 0:
            display_columns = [
                'Symbol', 'Company Name', 'Sector', 'Industry',
                'Screening_Score', 'Technical_Score', 'Fundamental_Score',
                'RS_Score', 'Individual_RS_Percentile',
                'Sector_RS_Percentile', 'Industry_RS_Percentile',
                'Current_Price', 'MA21', 'MA50', 'MA150', 'ATR_Pct_from_MA50', 'ADR'
            ]
            available_columns = [col for col in display_columns if col in filtered_df.columns]
            st.dataframe(
                filtered_df[available_columns].sort_values('Screening_Score', ascending=False),
                use_container_width=True, height=600
            )
            with st.expander("📊 フィルタリング結果の統計"):
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("銘柄数", len(filtered_df))
                with col2:
                    if 'Screening_Score' in filtered_df.columns:
                        st.metric("平均スコア", f"{filtered_df['Screening_Score'].mean():.1f}")
                with col3:
                    if 'Individual_RS_Percentile' in filtered_df.columns:
                        st.metric("平均個別RS", f"{filtered_df['Individual_RS_Percentile'].mean():.1f}%")
                with col4:
                    if 'ADR' in filtered_df.columns:
                        st.metric("平均ADR", f"{filtered_df['ADR'].mean():.1f}%")
                if 'Sector' in filtered_df.columns:
                    st.markdown("**セクター分布:**")
                    st.bar_chart(filtered_df['Sector'].value_counts())
            col1, col2 = st.columns(2)
            with col1:
                csv = filtered_df[available_columns].to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 CSVダウンロード（全データ）", data=csv,
                    file_name=f'filtered_stocks_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv', mime='text/csv'
                )
            with col2:
                if 'Symbol' in filtered_df.columns:
                    sorted_df = filtered_df.sort_values('Screening_Score', ascending=False)
                    symbols_list = sorted_df['Symbol'].dropna().astype(str).tolist()
                    if symbols_list:
                        st.download_button(
                            label="📝 Symbolリストダウンロード（TXT）",
                            data=','.join(symbols_list),
                            file_name=f'symbols_{datetime.now().strftime("%Y%m%d_%H%M%S")}.txt',
                            mime='text/plain'
                        )
            if 'Symbol' in filtered_df.columns:
                with st.expander("📌 Symbolリスト表示（TradingView用）"):
                    sorted_df = filtered_df.sort_values('Screening_Score', ascending=False)
                    symbols_list = sorted_df['Symbol'].dropna().astype(str).tolist()
                    if symbols_list:
                        st.markdown("**カンマ区切り（コピー用）:**")
                        st.code(','.join(symbols_list), language=None)
                        st.markdown("---")
                        st.success(f"✅ 合計 {len(symbols_list)} 銘柄")
                        if len(symbols_list) > 10:
                            st.info(f"📊 上位10銘柄: {', '.join(symbols_list[:10])}")
        else:
            st.warning("⚠️ 条件に合致する銘柄がありません。条件を緩和してください。")

    # ===== タブ5: ベース銘柄タイプ① =====
    with tab5:
        st.header("💎 ベース銘柄タイプ①スクリーニング")
        st.info("""
        **ベース銘柄とは？**\n
        ベースとは、ざっくり言うと、力を蓄えている期間に現れるもの。\n
        • **ベース3**の上へブレイクが成功する確率は約**67％**
        • **ベース4**の上昇ブレイクの成功確率は約**20％** (IBDより)\n
        💡 なるべく**ベース1**か**ベース2**の時に仕込みたい
        """)
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("📊 基本条件")
            st.markdown("**価格条件**")
            base_price_min = st.number_input("株価 最小値 ($", value=5.0, step=1.0, key="base_price_min")
            enable_price_max = st.checkbox("株価上限を設定", value=False, key="enable_price_max")
            base_price_max = st.number_input("株価 最大値 ($", value=500.0, step=10.0, key="base_price_max") if enable_price_max else float('inf')
            st.markdown("---")
            st.markdown("**ファンダメンタル条件**")
            base_fundamental_min = st.number_input("ファンダメンタルスコア最小値", min_value=0, max_value=10, value=1, step=1, key="base_fund_score")
            st.markdown("---")
            st.markdown("**テクニカル条件**")
            base_technical_min = st.number_input("テクニカルスコア最小値", min_value=0, max_value=14, value=8, step=1, key="base_tech_score")
            st.markdown("---")
            st.markdown("**ATR条件**")
            base_atr_min = st.number_input("ATR from MA50 最小値 (%)", value=2.0, step=0.1, key="base_atr_min")
            base_atr_max = st.number_input("ATR from MA50 最大値 (%)", value=5.0, step=0.1, key="base_atr_max")
            st.markdown("---")
            st.markdown("**ボラティリティ条件**")
            base_adr_max = st.number_input("ADR 最大値 (%)", value=4.0, step=0.5, key="base_adr_max")
            st.markdown("---")
            st.markdown("**高値更新条件**")
            enable_high_condition = st.checkbox("高値更新フィルターを有効にする", value=True, key="base_high_check")
            if enable_high_condition:
                col_a, col_b = st.columns(2)
                with col_a:
                    high_days = st.number_input("高値更新なし期間（日）", min_value=5, max_value=60, value=15, step=1, key="base_high_days")
                with col_b:
                    use_yahoo_finance = st.checkbox("Yahoo Financeを使用", value=True, key="use_yahoo")
                if use_yahoo_finance:
                    period_options = {"1ヶ月": "1mo", "2ヶ月": "2mo", "3ヶ月": "3mo", "6ヶ月": "6mo"}
                    selected_period = st.selectbox("分析期間", options=list(period_options.keys()), index=2, key="yahoo_period")
                    yahoo_period = period_options[selected_period]
                st.info("📊 Yahoo Financeから最新の高値更新データを取得して押し目買いポイントを判定します")
            else:
                high_days = 0
                use_yahoo_finance = False
        with col2:
            st.subheader("📈 RS条件")
            base_individual_rs_min = st.number_input("Individual RS Percentile 最小値", value=60, step=5, key="base_ind_rs")
            base_sector_rs_min = st.number_input("Sector RS Percentile 最小値", value=60, step=5, key="base_sec_rs")
            base_industry_rs_min = st.number_input("Industry RS Percentile 最小値", value=60, step=5, key="base_ind_rs2")
            st.markdown("---")
            st.markdown("**移動平均線条件**")
            base_ma21_condition = st.checkbox("株価 > MA21（21日移動平均線）", value=False, key="base_ma21_check")
            base_ma50_condition = st.checkbox("株価 > MA50（50日移動平均線）", value=False, key="base_ma50_check")
            base_ma150_condition = st.checkbox("株価 > MA150（150日移動平均線）", value=False, key="base_ma150_check")
            st.markdown("---")
            st.markdown("**セクターフィルター**")
            available_sectors = sorted(latest_data['Sector'].unique().tolist())
            selected_sectors = st.multiselect("対象セクターを選択（空欄=全セクター）", available_sectors, key="base_sectors")
            st.markdown("---")
            st.markdown("**売上加速条件**")
            enable_sales_accel = st.checkbox("売上加速フィルターを有効にする", value=False, key="enable_sales_accel")
            if enable_sales_accel:
                st.info("✅ 直近3四半期で売上が加速している銘柄のみを表示")
            st.markdown("---")
            st.markdown("**EPS加速条件**")
            enable_eps_accel = st.checkbox("EPS加速フィルターを有効にする", value=False, key="enable_eps_accel")
            if enable_eps_accel:
                st.info("✅ 直近3四半期でEPSが加速している銘柄のみを表示")

        st.markdown("---")
        base_filtered_df = latest_data.copy()
        if 'Current_Price' in base_filtered_df.columns:
            base_filtered_df = base_filtered_df[(base_filtered_df['Current_Price'] >= base_price_min) & (base_filtered_df['Current_Price'] <= base_price_max)]
        if 'Fundamental_Score' in base_filtered_df.columns:
            base_filtered_df = base_filtered_df[base_filtered_df['Fundamental_Score'] >= base_fundamental_min]
        if enable_sales_accel:
            if 'sales_accel_3_qtrs' in base_filtered_df.columns:
                base_filtered_df = base_filtered_df[base_filtered_df['sales_accel_3_qtrs'].astype(str).str.upper() == 'YES']
                st.success(f"✅ 売上加速フィルター適用: {len(base_filtered_df)}銘柄")
            else:
                st.warning("⚠️ 'sales_accel_3_qtrs'列が見つかりません。この条件はスキップされます。")
        if enable_eps_accel:
            if 'eps_accel_3_qtrs' in base_filtered_df.columns:
                base_filtered_df = base_filtered_df[base_filtered_df['eps_accel_3_qtrs'].astype(str).str.upper() == 'YES']
                st.success(f"✅ EPS加速フィルター適用: {len(base_filtered_df)}銘柄")
            else:
                st.warning("⚠️ 'eps_accel_3_qtrs'列が見つかりません。この条件はスキップされます。")
        if 'Technical_Score' in base_filtered_df.columns:
            base_filtered_df = base_filtered_df[base_filtered_df['Technical_Score'] >= base_technical_min]
        if 'ATR_Pct_from_MA50' in base_filtered_df.columns:
            base_filtered_df = base_filtered_df[(base_filtered_df['ATR_Pct_from_MA50'] >= base_atr_min) & (base_filtered_df['ATR_Pct_from_MA50'] <= base_atr_max)]
        if 'ADR' in base_filtered_df.columns:
            base_filtered_df = base_filtered_df[base_filtered_df['ADR'] <= base_adr_max]
        if 'Individual_RS_Percentile' in base_filtered_df.columns:
            base_filtered_df = base_filtered_df[base_filtered_df['Individual_RS_Percentile'] >= base_individual_rs_min]
        if 'Sector_RS_Percentile' in base_filtered_df.columns:
            base_filtered_df = base_filtered_df[base_filtered_df['Sector_RS_Percentile'] >= base_sector_rs_min]
        if 'Industry_RS_Percentile' in base_filtered_df.columns:
            base_filtered_df = base_filtered_df[base_filtered_df['Industry_RS_Percentile'] >= base_industry_rs_min]
        if base_ma21_condition and 'MA21' in base_filtered_df.columns and 'Current_Price' in base_filtered_df.columns:
            base_filtered_df = base_filtered_df[base_filtered_df['Current_Price'] > base_filtered_df['MA21']]
        if base_ma50_condition and 'MA50' in base_filtered_df.columns and 'Current_Price' in base_filtered_df.columns:
            base_filtered_df = base_filtered_df[base_filtered_df['Current_Price'] > base_filtered_df['MA50']]
        if base_ma150_condition and 'MA150' in base_filtered_df.columns and 'Current_Price' in base_filtered_df.columns:
            base_filtered_df = base_filtered_df[base_filtered_df['Current_Price'] > base_filtered_df['MA150']]
        if selected_sectors:
            base_filtered_df = base_filtered_df[base_filtered_df['Sector'].isin(selected_sectors)]
        if enable_high_condition and use_yahoo_finance and len(base_filtered_df) > 0:
            with st.spinner(f"Yahoo Financeから{len(base_filtered_df)}銘柄の高値更新データを取得中..."):
                symbols = base_filtered_df['Symbol'].dropna().astype(str).tolist()
                days_since_high_dict = get_days_since_high_from_yahoo(symbols, period=yahoo_period)
                base_filtered_df['Days_Since_High_Yahoo'] = base_filtered_df['Symbol'].map(days_since_high_dict)
                base_filtered_df = base_filtered_df[base_filtered_df['Days_Since_High_Yahoo'].notna()]
                base_filtered_df = base_filtered_df[base_filtered_df['Days_Since_High_Yahoo'] >= high_days]
                st.success(f"✅ Yahoo Finance高値更新フィルター適用: {high_days}日以上高値更新なし")
        elif enable_high_condition and not use_yahoo_finance:
            if 'Days_Since_High' in base_filtered_df.columns:
                base_filtered_df = base_filtered_df[base_filtered_df['Days_Since_High'] >= high_days]
                st.success(f"✅ 高値更新フィルター適用: {high_days}日以上高値更新なし")
            elif 'High_52W' in base_filtered_df.columns and 'Current_Price' in base_filtered_df.columns:
                base_filtered_df['Price_from_52W_High'] = (base_filtered_df['Current_Price'] / base_filtered_df['High_52W'] - 1) * 100
                base_filtered_df = base_filtered_df[base_filtered_df['Price_from_52W_High'] < -10]
                st.warning("⚠️ Days_Since_High列が見つからないため、52週高値から10%以上下落した銘柄を選定しています")
            else:
                st.warning("⚠️ 高値更新データが利用できません。この条件はスキップされます。")

        st.subheader(f"💎 ベース銘柄タイプ①候補: {len(base_filtered_df)}銘柄")
        if len(base_filtered_df) > 0:
            display_columns = [
                'Symbol', 'Company Name', 'Sector', 'Industry',
                'Fundamental_Score', 'Technical_Score', 'Screening_Score',
                'Individual_RS_Percentile', 'Sector_RS_Percentile', 'Industry_RS_Percentile',
                'Current_Price', 'MA21', 'MA50', 'MA150', 'ATR_Pct_from_MA50', 'ADR'
            ]
            if 'sales_accel_3_qtrs' in base_filtered_df.columns:
                display_columns.insert(7, 'sales_accel_3_qtrs')
            if 'eps_accel_3_qtrs' in base_filtered_df.columns:
                display_columns.insert(8, 'eps_accel_3_qtrs')
            if 'Days_Since_High_Yahoo' in base_filtered_df.columns:
                display_columns.append('Days_Since_High_Yahoo')
            if 'Days_Since_High' in base_filtered_df.columns:
                display_columns.append('Days_Since_High')
            if 'Price_from_52W_High' in base_filtered_df.columns:
                display_columns.append('Price_from_52W_High')
            available_columns = [col for col in display_columns if col in base_filtered_df.columns]
            st.dataframe(
                base_filtered_df[available_columns].sort_values('Fundamental_Score', ascending=False),
                use_container_width=True, height=600
            )
            with st.expander("📊 ベース銘柄タイプ①の統計"):
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("銘柄数", len(base_filtered_df))
                with col2:
                    if 'Fundamental_Score' in base_filtered_df.columns:
                        st.metric("平均ファンダメンタル", f"{base_filtered_df['Fundamental_Score'].mean():.1f}")
                with col3:
                    if 'ADR' in base_filtered_df.columns:
                        st.metric("平均ADR", f"{base_filtered_df['ADR'].mean():.1f}%")
                with col4:
                    if 'ATR_Pct_from_MA50' in base_filtered_df.columns:
                        st.metric("平均ATR", f"{base_filtered_df['ATR_Pct_from_MA50'].mean():.1f}%")
                if 'sales_accel_3_qtrs' in base_filtered_df.columns or 'eps_accel_3_qtrs' in base_filtered_df.columns:
                    st.markdown("---")
                    st.markdown("**📈 成長指標の統計**")
                    col1, col2 = st.columns(2)
                    with col1:
                        if 'sales_accel_3_qtrs' in base_filtered_df.columns:
                            sales_accel_count = (base_filtered_df['sales_accel_3_qtrs'].astype(str).str.upper() == 'YES').sum()
                            st.metric("売上加速銘柄数", f"{sales_accel_count}銘柄")
                            st.metric("売上加速比率", f"{sales_accel_count / len(base_filtered_df) * 100:.1f}%")
                    with col2:
                        if 'eps_accel_3_qtrs' in base_filtered_df.columns:
                            eps_accel_count = (base_filtered_df['eps_accel_3_qtrs'].astype(str).str.upper() == 'YES').sum()
                            st.metric("EPS加速銘柄数", f"{eps_accel_count}銘柄")
                            st.metric("EPS加速比率", f"{eps_accel_count / len(base_filtered_df) * 100:.1f}%")
                if 'Days_Since_High_Yahoo' in base_filtered_df.columns:
                    st.markdown("---")
                    st.markdown("**📈 高値更新の統計（Yahoo Finance）**")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("平均高値更新なし日数", f"{base_filtered_df['Days_Since_High_Yahoo'].mean():.0f}日")
                    with col2:
                        st.metric("最大高値更新なし日数", f"{base_filtered_df['Days_Since_High_Yahoo'].max():.0f}日")
                    with col3:
                        st.metric("最小高値更新なし日数", f"{base_filtered_df['Days_Since_High_Yahoo'].min():.0f}日")
                elif 'Days_Since_High' in base_filtered_df.columns:
                    st.markdown("---")
                    st.markdown("**📈 高値更新の統計**")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("平均高値更新なし日数", f"{base_filtered_df['Days_Since_High'].mean():.0f}日")
                    with col2:
                        st.metric("最大高値更新なし日数", f"{base_filtered_df['Days_Since_High'].max():.0f}日")
                    with col3:
                        st.metric("最小高値更新なし日数", f"{base_filtered_df['Days_Since_High'].min():.0f}日")
                if 'Sector' in base_filtered_df.columns:
                    st.markdown("**セクター分布:**")
                    st.bar_chart(base_filtered_df['Sector'].value_counts())
                st.markdown("---")
                st.markdown("**📊 上位10銘柄の特徴**")
                top_10 = base_filtered_df.head(10)
                col1, col2 = st.columns(2)
                with col1:
                    if 'Fundamental_Score' in top_10.columns:
                        st.metric("平均ファンダメンタルスコア", f"{top_10['Fundamental_Score'].mean():.1f}")
                    if 'Technical_Score' in top_10.columns:
                        st.metric("平均テクニカルスコア", f"{top_10['Technical_Score'].mean():.1f}")
                with col2:
                    if 'Individual_RS_Percentile' in top_10.columns:
                        st.metric("平均Individual RS", f"{top_10['Individual_RS_Percentile'].mean():.0f}%")
                    if 'Days_Since_High_Yahoo' in top_10.columns:
                        st.metric("平均高値更新なし日数(Yahoo)", f"{top_10['Days_Since_High_Yahoo'].mean():.0f}日")
                    elif 'Days_Since_High' in top_10.columns:
                        st.metric("平均高値更新なし日数", f"{top_10['Days_Since_High'].mean():.0f}日")
                    elif 'ADR' in top_10.columns:
                        st.metric("平均ADR", f"{top_10['ADR'].mean():.1f}%")
            col1, col2 = st.columns(2)
            with col1:
                csv = base_filtered_df[available_columns].to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 ベース銘柄タイプ①CSVダウンロード", data=csv,
                    file_name=f'base_stocks_type1_{datetime.now().strftime("%Y%m%d_%H%M%S")}.csv', mime='text/csv'
                )
            with col2:
                if 'Symbol' in base_filtered_df.columns:
                    sorted_df = base_filtered_df.sort_values('Fundamental_Score', ascending=False)
                    symbols_list = sorted_df['Symbol'].dropna().astype(str).tolist()
                    if symbols_list:
                        st.download_button(
                            label="📝 Symbolリストダウンロード",
                            data=','.join(symbols_list),
                            file_name=f'base_symbols_type1_{datetime.now().strftime("%Y%m%d_%H%M%S")}.txt',
                            mime='text/plain'
                        )
            if 'Symbol' in base_filtered_df.columns:
                with st.expander("📌 Symbolリスト表示（TradingView用）"):
                    sorted_df = base_filtered_df.sort_values('Fundamental_Score', ascending=False)
                    symbols_list = sorted_df['Symbol'].dropna().astype(str).tolist()
                    if symbols_list:
                        st.markdown("**カンマ区切り（コピー用）:**")
                        st.code(','.join(symbols_list), language=None)
                        st.markdown("---")
                        st.success(f"✅ 合計 {len(symbols_list)} 銘柄")
                        if len(symbols_list) > 10:
                            st.info(f"📊 上位10銘柄: {', '.join(symbols_list[:10])}")
        else:
            st.warning("⚠️ 条件に合致するベース銘柄タイプ①がありません。条件を緩和してください。")

        st.markdown("---")
        with st.expander("💡 ベース銘柄タイプ①の投資戦略"):
            st.markdown("""
            ### ベースのカウント方法

            1. **1回目のベース**は形成前に**30%以上**株価が上昇していること
            2. ベースの正規エントリーポイントから**20%以上**株価が上昇した後に形成されたベースは次のベースとなる
            3. ベースの正規エントリーポイントから**20%未満**の株価上昇で形成されたベースは、同じ回のベースとする
            4. 前回のベース安値を次回のベース安値が下回った場合、カウントはリセットされ、1回目となる。ただし、前回ベースの安値から30%以上株価が上昇していない場合は、ルール1が適用されないため、ベースと見なされない
            5. 市況が弱気相場（主要指数が**20%以上下落**した調整相場）になった場合、カウントはリセットされる

            ---

            ### ベースの基本事項

            • ベース形成前には安値から**30%以上**の価格上昇が必要

            • 前ベースがある場合は前ベースのエントリーポイントから**20%以上**の価格上昇が必要

            • エントリーポイントを価格が上抜ける際は**出来高の急増**を伴わないと、ブレイクアウト成功確率は下がる
              （目安は**50日平均出来高の1.5倍以上**）
              ※書籍には1.4倍以上とあるが、反例があったため1.5倍以上を推奨

            • 実際のエントリーはエントリーポイント上抜けから**5％以内**を推奨

            • 長期下降トレンド（第4ステージ）中はブレイクアウト成功率は下がる

            ---

            ### 📊 成功確率まとめ

            | ベース回数 | 成功確率 | 推奨度 |
            |-----------|---------|--------|
            | ベース1 | 高い | ⭐⭐⭐ |
            | ベース2 | 高い | ⭐⭐⭐ |
            | ベース3 | 約67% | ⭐⭐ |
            | ベース4 | 約20% | ⭐ |
            """)

    # ===== タブ6: 運用ルール =====
    with tab6:
        st.header("📋 株式運用ルール")
        st.markdown("### 🎯 エントリーポリシー")
        col1, col2 = st.columns([1, 2])
        with col1:
            st.markdown("""
            **基本方針**
            - ✅ Strong Positive時
            - ✅ Positive時
            - ⚠️ Neutral時は様子見
            - ❌ Negative時は購入しない
            - ❌ Strong Negative時は購入しない
            """)
        with col2:
            st.info("""
            **銘柄選定基準**

            1. **最優先**: テクニカルスコアが高い銘柄（最高14点）
            2. **次点**: ファンダメンタルスコアも高い銘柄（最高10点）
            3. **理想**: テクニカル+ファンダメンタル両方が高得点

            💡 「モメンタム銘柄」タブで条件フィルタリングを活用してください
            """)
        st.markdown("---")
        st.markdown("### 💼 ポジション比率")
        position_data = pd.DataFrame({
            'Market Condition': ['Strong Positive', 'Positive', 'Neutral', 'Negative', 'Strong Negative'],
            'ポジション比率': ['80–100%', '60–80%', '40–60%', '20–40%', '0–20%'],
            '推奨アクション': ['積極的に投資', '通常投資', '慎重に判断', '防御的姿勢', 'ポジション最小化']
        })

        def highlight_current_status(row):
            if latest_market and row['Market Condition'] == latest_market['status']:
                return ['background-color: #ffffcc'] * len(row)
            return [''] * len(row)

        styled_df = position_data.style.apply(highlight_current_status, axis=1)
        st.dataframe(styled_df, use_container_width=True, hide_index=True)
        st.markdown("---")
        st.markdown("### 🛑 エントリー時の損切り（Stop Loss）ルール")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("#### 📊 損切りライン一覧")
            stoploss_data = pd.DataFrame({
                'Market Condition': ['Positive以上', 'Neutral', 'Negative以下'],
                '損切りライン': ['-8%', '-6%', '-4%'],
                '考え方': ['強気相場で余裕を持つ', '中立相場で慎重に', '弱気相場で早めの撤退']
            })

            def highlight_stoploss(row):
                if latest_market:
                    status = latest_market['status']
                    if ('Positive' in status and row['Market Condition'] == 'Positive以上') or \
                       ('Neutral' in status and row['Market Condition'] == 'Neutral') or \
                       ('Negative' in status and row['Market Condition'] == 'Negative以下'):
                        return ['background-color: #ffffcc'] * len(row)
                return [''] * len(row)

            styled_stoploss = stoploss_data.style.apply(highlight_stoploss, axis=1)
            st.dataframe(styled_stoploss, use_container_width=True, hide_index=True)
        with col2:
            st.warning("""
            **⚠️ 重要な注意事項**

            - 損切りラインは**必ず守る**こと
            - 感情的な判断を避ける
            - マーケット状況が変わったら即座に対応
            - 損切り後は冷静に次の機会を待つ
            """)

else:
    st.warning("⚠️ dataフォルダにExcelファイルが見つかりません")
    st.info("""
    ### 📖 セットアップ方法

    1. **GitHubリポジトリに `data` フォルダを作成**
    2. **Excelファイルを `data` フォルダにコミット**
    3. **GitHubにプッシュ**
    4. **Streamlit Cloudで再デプロイ**
    """)

gc.collect()
