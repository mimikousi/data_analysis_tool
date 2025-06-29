import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import plotly.graph_objects as go
import io
import base64
from typing import List, Dict, Optional

# カスタムモジュールのインポート
from src.data_processor import DataProcessor
from src.outlier_removal import OutlierRemover
from src.visualization import DataVisualizer
from src.statistics import StatisticsCalculator
from src.report_generator import ReportGenerator

# ページ設定
st.set_page_config(
    page_title="プロセスデータ解析アプリ",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# セッション状態の初期化
def initialize_session_state():
    """セッション状態を初期化"""
    if 'data_processor' not in st.session_state:
        st.session_state.data_processor = DataProcessor()
    if 'outlier_remover' not in st.session_state:
        st.session_state.outlier_remover = OutlierRemover()
    if 'visualizer' not in st.session_state:
        st.session_state.visualizer = DataVisualizer()
    if 'stats_calculator' not in st.session_state:
        st.session_state.stats_calculator = StatisticsCalculator()
    if 'report_generator' not in st.session_state:
        st.session_state.report_generator = ReportGenerator()
    if 'data_loaded' not in st.session_state:
        st.session_state.data_loaded = False
    if 'current_data' not in st.session_state:
        st.session_state.current_data = None
    if 'outlier_removal_active' not in st.session_state:
        st.session_state.outlier_removal_active = False
    if 'selected_trend_variables' not in st.session_state:
        st.session_state.selected_trend_variables = []

def main():
    """メイン関数"""
    initialize_session_state()
    
    # サイドバー
    st.sidebar.title("📊 プロセスデータ解析")
    st.sidebar.markdown("---")
    
    # ファイルアップロード
    uploaded_file = st.sidebar.file_uploader(
        "データファイルをアップロード",
        type=['csv', 'xlsx', 'xlsm'],
        help="CSV、Excel（.xlsx, .xlsm）ファイルをサポート"
    )
    
    if uploaded_file is not None:
        # データ読み込み
        if not st.session_state.data_loaded or st.sidebar.button("🔄 データを再読み込み"):
            with st.spinner("データを読み込み中..."):
                data = st.session_state.data_processor.load_file(uploaded_file)
                if data is not None:
                    st.session_state.current_data = data
                    st.session_state.data_loaded = True
                    st.session_state.outlier_remover.initialize_data(data)
                    st.success("✅ データの読み込みが完了しました")
                else:
                    st.error("❌ データの読み込みに失敗しました")
                    return
    
    # データが読み込まれていない場合
    if not st.session_state.data_loaded:
        st.title("プロセスデータ解析アプリ")
        st.markdown("""
        ## 📋 このアプリについて
        
        化学業界のプロセスデータ・品質データを対象とした時系列データ解析アプリケーションです。
        
        ### 🔧 主な機能
        - **データ読み込み**: CSV、Excel形式のデータファイル対応
        - **外れ値除去**: 範囲指定や統計的手法による除去、履歴管理
        - **トレンド分析**: 多変数時系列グラフ、統計線表示
        - **相関分析**: 散布図マトリックス、相関ヒートマップ
        - **分布分析**: ヒストグラム、正規性検定
        - **レポート出力**: PDF形式の自動レポート生成
        
        ### 🚀 使い方
        1. 左サイドバーからデータファイルをアップロード
        2. データの確認と外れ値除去（必要に応じて）
        3. 各種分析機能を使用
        4. 結果をエクスポート
        """)
        return
    
    # データが読み込まれている場合のメイン処理
    data = st.session_state.current_data
    
    # サイドバーメニュー
    st.sidebar.markdown("---")
    menu_options = [
        "📊 データ概要",
        "🔧 外れ値除去",
        "📈 トレンド分析",
        "🔍 相関分析", 
        "📊 分布分析",
        "📄 レポート生成"
    ]
    
    selected_menu = st.sidebar.selectbox("分析メニュー", menu_options)
    
    # メイン画面の表示
    if selected_menu == "📊 データ概要":
        show_data_overview(data)
    elif selected_menu == "🔧 外れ値除去":
        show_outlier_removal(data)
    elif selected_menu == "📈 トレンド分析":
        show_trend_analysis(data)
    elif selected_menu == "🔍 相関分析":
        show_correlation_analysis(data)
    elif selected_menu == "📊 分布分析":
        show_distribution_analysis(data)
    elif selected_menu == "📄 レポート生成":
        show_report_generation(data)

def show_data_overview(data: pd.DataFrame):
    """データ概要を表示"""
    st.title("📊 データ概要")
    
    # 基本情報
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("データ行数", f"{len(data):,}")
    with col2:
        st.metric("データ列数", len(data.columns))
    with col3:
        if isinstance(data.index, pd.DatetimeIndex):
            period_days = (data.index.max() - data.index.min()).days
            st.metric("期間（日）", period_days)
        else:
            st.metric("期間", "N/A")
    with col4:
        memory_mb = data.memory_usage(deep=True).sum() / 1024 / 1024
        st.metric("メモリ使用量", f"{memory_mb:.1f} MB")
    
    # データプレビュー
    st.subheader("📋 データプレビュー")
    
    # 表示行数の選択
    preview_rows = st.selectbox("表示行数", [10, 20, 50, 100], index=0)
    st.dataframe(data.head(preview_rows), use_container_width=True)
    
    # 統計情報
    st.subheader("📈 基本統計情報")
    
    numeric_cols = data.select_dtypes(include=[np.number]).columns.tolist()
    if numeric_cols:
        stats_df = st.session_state.stats_calculator.calculate_basic_statistics(data, numeric_cols)
        st.dataframe(stats_df, use_container_width=True)
    else:
        st.warning("数値型のカラムが見つかりません")
    
    # データ品質チェック
    st.subheader("🔍 データ品質")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**欠損値情報**")
        null_info = data.isnull().sum()
        null_info_with_missing = null_info[null_info > 0]
        if len(null_info_with_missing) > 0:
            st.write(null_info_with_missing)
            
            # 欠損値除去ボタン
            st.write("**欠損値除去**")
            removal_method = st.radio(
                "除去方法を選択",
                ["行を削除", "列を削除", "前の値で補完", "平均値で補完"],
                horizontal=True,
                key="missing_removal_method"
            )
            
            if st.button("🗑️ 欠損値を除去"):
                if removal_method == "行を削除":
                    cleaned_data = data.dropna()
                    removed_count = len(data) - len(cleaned_data)
                    st.session_state.current_data = cleaned_data
                    st.success(f"✅ {removed_count}行の欠損値を含む行を削除しました")
                    
                elif removal_method == "列を削除":
                    cleaned_data = data.dropna(axis=1)
                    removed_count = len(data.columns) - len(cleaned_data.columns)
                    st.session_state.current_data = cleaned_data
                    st.success(f"✅ {removed_count}列の欠損値を含む列を削除しました")
                    
                elif removal_method == "前の値で補完":
                    cleaned_data = data.ffill()
                    filled_count = data.isnull().sum().sum()
                    st.session_state.current_data = cleaned_data
                    st.success(f"✅ {filled_count}個の欠損値を前の値で補完しました")
                    
                elif removal_method == "平均値で補完":
                    numeric_cols_for_fill = data.select_dtypes(include=[np.number]).columns
                    cleaned_data = data.copy()
                    for col in numeric_cols_for_fill:
                        cleaned_data[col].fillna(cleaned_data[col].mean(), inplace=True)
                    filled_count = data[numeric_cols_for_fill].isnull().sum().sum()
                    st.session_state.current_data = cleaned_data
                    st.success(f"✅ {filled_count}個の数値欠損値を平均値で補完しました")
                
                st.rerun()
        else:
            st.success("欠損値はありません")
    
    with col2:
        st.write("**データ型情報**")
        dtype_info = data.dtypes.value_counts()
        st.write(dtype_info)

def show_outlier_removal(data: pd.DataFrame):
    """外れ値除去機能を表示"""
    st.title("🔧 外れ値除去")
    
    numeric_cols = data.select_dtypes(include=[np.number]).columns.tolist()
    
    if not numeric_cols:
        st.warning("数値型のカラムが見つかりません")
        return
    
    # カラム選択
    selected_column = st.selectbox("対象カラムを選択", numeric_cols)
    
    if selected_column:
        # 現在のデータを取得
        current_data = st.session_state.current_data
        
        # トレンドグラフ表示
        st.subheader(f"📈 {selected_column} のトレンド")
        
        fig = go.Figure()
        fig.add_trace(go.Scatter(
            x=current_data.index,
            y=current_data[selected_column],
            mode='lines',
            name=selected_column,
            line=dict(color='blue')
        ))
        
        fig.update_layout(
            title=f"{selected_column} の時系列データ",
            xaxis_title="時間",
            yaxis_title="値",
            height=400
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # 外れ値除去方法の選択
        st.subheader("🎯 外れ値除去方法")
        
        removal_method = st.radio(
            "除去方法を選択",
            ["範囲指定", "統計的手法"],
            horizontal=True
        )
        
        if removal_method == "範囲指定":
            show_range_based_removal(current_data, selected_column)
        else:
            show_statistical_removal(current_data, selected_column)
        
        # 履歴管理
        show_outlier_history()

def show_range_based_removal(data: pd.DataFrame, column: str):
    """範囲指定による外れ値除去"""
    st.write("**範囲指定による外れ値除去**")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**Y軸範囲（値の範囲）**")
        col_data = data[column].dropna()
        min_val, max_val = float(col_data.min()), float(col_data.max())
        
        y_min = st.number_input("下限値", value=min_val, step=(max_val-min_val)/100)
        y_max = st.number_input("上限値", value=max_val, step=(max_val-min_val)/100)
        
    with col2:
        st.write("**X軸範囲（時間範囲）**")
        if isinstance(data.index, pd.DatetimeIndex):
            date_min, date_max = data.index.min(), data.index.max()
            
            x_min = st.date_input("開始日", value=date_min.date())
            x_max = st.date_input("終了日", value=date_max.date())
            
            # 日付をdatetimeに変換
            x_min = pd.to_datetime(x_min)
            x_max = pd.to_datetime(x_max)
        else:
            st.info("時間範囲は日時インデックスの場合のみ利用可能です")
            x_min, x_max = None, None
    
    # 除去範囲の可視化
    if st.button("🎯 除去範囲をプレビュー"):
        fig = go.Figure()
        
        # 元データ
        fig.add_trace(go.Scatter(
            x=data.index,
            y=data[column],
            mode='lines',
            name='元データ',
            line=dict(color='blue')
        ))
        
        # 除去範囲のハイライト
        if x_min and x_max:
            mask = (data.index >= x_min) & (data.index <= x_max)
            value_mask = (data[column] >= y_min) & (data[column] <= y_max)
            remove_mask = mask & value_mask
        else:
            remove_mask = (data[column] >= y_min) & (data[column] <= y_max)
        
        if remove_mask.any():
            fig.add_trace(go.Scatter(
                x=data.index[remove_mask],
                y=data[column][remove_mask],
                mode='markers',
                name='除去対象',
                marker=dict(color='red', size=8)
            ))
        
        # 範囲線の表示
        fig.add_hline(y=y_min, line_dash="dash", line_color="red", annotation_text="下限")
        fig.add_hline(y=y_max, line_dash="dash", line_color="red", annotation_text="上限")
        
        fig.update_layout(
            title=f"{column} - 除去範囲プレビュー",
            xaxis_title="時間",
            yaxis_title="値",
            height=400
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # 除去件数の表示
        remove_count = remove_mask.sum()
        st.info(f"除去対象: {remove_count} 件 ({remove_count/len(data)*100:.2f}%)")
    
    # 除去実行
    if st.button("❌ 外れ値を除去"):
        x_range = (x_min, x_max) if x_min and x_max else None
        y_range = (y_min, y_max)
        
        filtered_data = st.session_state.outlier_remover.remove_outliers_by_range(
            st.session_state.current_data, column, x_range, y_range
        )
        
        st.session_state.current_data = filtered_data
        st.success(f"✅ 外れ値除去が完了しました")
        st.rerun()

def show_statistical_removal(data: pd.DataFrame, column: str):
    """統計的手法による外れ値除去"""
    st.write("**統計的手法による外れ値除去**")
    
    col1, col2 = st.columns(2)
    
    with col1:
        method = st.selectbox("統計手法", ["IQR法", "Z-score法"])
        method_key = "iqr" if method == "IQR法" else "zscore"
    
    with col2:
        if method == "IQR法":
            multiplier = st.slider("IQR係数", 1.0, 3.0, 1.5, 0.1)
        else:
            multiplier = st.slider("標準偏差係数", 1.0, 5.0, 3.0, 0.1)
    
    # 外れ値候補の表示
    if st.button("🔍 外れ値候補を確認"):
        outlier_mask = st.session_state.outlier_remover.get_outlier_candidates(
            data, column, method_key, multiplier
        )
        
        if outlier_mask.any():
            outlier_count = outlier_mask.sum()
            st.info(f"外れ値候補: {outlier_count} 件 ({outlier_count/len(data)*100:.2f}%)")
            
            # 外れ値候補の可視化
            fig = go.Figure()
            
            # 正常データ
            fig.add_trace(go.Scatter(
                x=data.index[~outlier_mask],
                y=data[column][~outlier_mask],
                mode='lines+markers',
                name='正常データ',
                line=dict(color='blue'),
                marker=dict(size=4)
            ))
            
            # 外れ値候補
            fig.add_trace(go.Scatter(
                x=data.index[outlier_mask],
                y=data[column][outlier_mask],
                mode='markers',
                name='外れ値候補',
                marker=dict(color='red', size=8)
            ))
            
            fig.update_layout(
                title=f"{column} - 外れ値候補 ({method})",
                xaxis_title="時間",
                yaxis_title="値",
                height=400
            )
            
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.success("外れ値候補は見つかりませんでした")
    
    # 除去実行
    if st.button("❌ 外れ値を除去"):
        filtered_data = st.session_state.outlier_remover.remove_outliers_by_statistical_method(
            st.session_state.current_data, column, method_key, multiplier
        )
        
        st.session_state.current_data = filtered_data
        st.success(f"✅ 外れ値除去が完了しました")
        st.rerun()

def show_outlier_history():
    """外れ値除去履歴を表示"""
    st.subheader("📚 外れ値除去履歴")
    
    history_df = st.session_state.outlier_remover.get_history_summary()
    
    if not history_df.empty:
        st.dataframe(history_df, use_container_width=True)
        
        # 復元機能
        col1, col2 = st.columns(2)
        
        with col1:
            operation_id = st.selectbox("復元したい操作を選択", 
                                      range(len(history_df)), 
                                      format_func=lambda x: f"操作 {x}: {history_df.iloc[x]['対象カラム']}")
        
        with col2:
            if st.button("🔙 この操作時点に復元"):
                restored_data = st.session_state.outlier_remover.restore_to_operation(operation_id)
                if not restored_data.empty:
                    st.session_state.current_data = restored_data
                    st.success("✅ データが復元されました")
                    st.rerun()
        
        # 履歴クリア
        if st.button("🗑️ 履歴をクリア"):
            st.session_state.outlier_remover.clear_all_history()
            st.success("✅ 履歴がクリアされました")
            st.rerun()
    else:
        st.info("外れ値除去履歴はありません")

def show_trend_analysis(data: pd.DataFrame):
    """トレンド分析を表示"""
    st.title("📈 トレンド分析")
    
    numeric_cols = data.select_dtypes(include=[np.number]).columns.tolist()
    
    if not numeric_cols:
        st.warning("数値型のカラムが見つかりません")
        return
    
    # カラム選択
    st.subheader("📊 表示設定")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**Y1軸変数**")
        default_y1 = numeric_cols[:1] if numeric_cols else []
        y1_columns = st.multiselect("Y1軸に表示する変数", numeric_cols, default=default_y1)
        
        if y1_columns:
            y1_range_enabled = st.checkbox("Y1軸範囲を手動設定")
            if y1_range_enabled:
                y1_min = st.number_input("Y1軸 最小値", value=0.0)
                y1_max = st.number_input("Y1軸 最大値", value=100.0)
                y1_range = (y1_min, y1_max)
            else:
                y1_range = None
    
    with col2:
        st.write("**Y2軸変数（オプション）**")
        y2_columns = st.multiselect("Y2軸に表示する変数", numeric_cols)
        
        if y2_columns:
            y2_range_enabled = st.checkbox("Y2軸範囲を手動設定")
            if y2_range_enabled:
                y2_min = st.number_input("Y2軸 最小値", value=0.0)
                y2_max = st.number_input("Y2軸 最大値", value=100.0)
                y2_range = (y2_min, y2_max)
            else:
                y2_range = None
        else:
            y2_range = None
    
    # 統計線設定
    st.subheader("📊 統計線設定")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        show_statistics = st.checkbox("統計線を表示", value=True)
    
    with col2:
        if show_statistics:
            sigma_multiplier = st.selectbox("σ係数", [1, 2, 3, 4, 5], index=0)
        else:
            sigma_multiplier = 1
    
    with col3:
        if show_statistics:
            period_stats = st.checkbox("期間指定統計")
        else:
            period_stats = False
    
    # 期間指定
    statistics_period = None
    if show_statistics and period_stats and isinstance(data.index, pd.DatetimeIndex):
        st.write("**統計計算期間**")
        col1, col2 = st.columns(2)
        
        with col1:
            start_date = st.date_input("開始日", value=data.index.min().date())
        with col2:
            end_date = st.date_input("終了日", value=data.index.max().date())
        
        statistics_period = (pd.to_datetime(start_date), pd.to_datetime(end_date))
    
    # グラフ生成
    if y1_columns:
        all_columns = y1_columns + y2_columns
        
        # 選択された変数をセッション状態に保存
        st.session_state.selected_trend_variables = all_columns
        
        fig = st.session_state.visualizer.create_trend_line_chart(
            data, all_columns, y1_columns, y2_columns,
            y1_range, y2_range, show_statistics, statistics_period, sigma_multiplier
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        # エクスポート機能
        st.subheader("📥 エクスポート")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("PNG ダウンロード"):
                img_bytes = st.session_state.visualizer.export_figure(fig, 'png')
                st.download_button(
                    label="💾 PNG ファイルをダウンロード",
                    data=img_bytes,
                    file_name=f"trend_chart_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png",
                    mime="image/png"
                )
        
        with col2:
            if st.button("SVG ダウンロード"):
                img_bytes = st.session_state.visualizer.export_figure(fig, 'svg')
                st.download_button(
                    label="💾 SVG ファイルをダウンロード",
                    data=img_bytes,
                    file_name=f"trend_chart_{datetime.now().strftime('%Y%m%d_%H%M%S')}.svg",
                    mime="image/svg+xml"
                )
        
        with col3:
            if st.button("PDF ダウンロード"):
                img_bytes = st.session_state.visualizer.export_figure(fig, 'pdf')
                st.download_button(
                    label="💾 PDF ファイルをダウンロード",
                    data=img_bytes,
                    file_name=f"trend_chart_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                    mime="application/pdf"
                )

def show_correlation_analysis(data: pd.DataFrame):
    """相関分析を表示"""
    st.title("🔍 相関分析")
    
    numeric_cols = data.select_dtypes(include=[np.number]).columns.tolist()
    
    if len(numeric_cols) < 2:
        st.warning("相関分析には2つ以上の数値型カラムが必要です")
        return
    
    # カラム選択（トレンド分析で選択された変数をデフォルトに）
    if st.session_state.selected_trend_variables:
        default_corr_cols = [col for col in st.session_state.selected_trend_variables if col in numeric_cols]
    else:
        default_corr_cols = numeric_cols[:5]
        
    selected_columns = st.multiselect(
        "分析対象の変数を選択", 
        numeric_cols, 
        default=default_corr_cols
    )
    
    if len(selected_columns) < 2:
        st.warning("2つ以上の変数を選択してください")
        return
    
    # タブ分割
    tab1, tab2, tab3 = st.tabs(["散布図マトリックス", "相関ヒートマップ", "相関統計"])
    
    with tab1:
        st.subheader("📊 散布図マトリックス")
        
        if len(selected_columns) <= 6:  # 6変数以下の場合のみ表示
            fig_scatter = st.session_state.visualizer.create_scatter_matrix(data, selected_columns)
            st.plotly_chart(fig_scatter, use_container_width=True)
        else:
            st.warning("散布図マトリックスは6変数以下で表示可能です")
    
    with tab2:
        st.subheader("🔥 相関ヒートマップ")
        
        fig_heatmap = st.session_state.visualizer.create_correlation_heatmap(data, selected_columns)
        st.plotly_chart(fig_heatmap, use_container_width=True)
    
    with tab3:
        st.subheader("📊 相関統計")
        
        # 相関行列
        corr_matrix = st.session_state.stats_calculator.calculate_correlation_matrix(data, selected_columns)
        st.write("**相関行列**")
        st.dataframe(corr_matrix.round(3), use_container_width=True)
        
        # 相関係数と有意性
        corr_with_significance = st.session_state.stats_calculator.calculate_correlation_with_significance(
            data, selected_columns
        )
        
        if corr_with_significance:
            st.write("**相関係数と有意性検定**")
            
            corr_results = []
            for pair, result in corr_with_significance.items():
                var1, var2 = pair.split('_vs_')
                corr_results.append({
                    '変数1': var1,
                    '変数2': var2,
                    '相関係数': round(result['correlation'], 3),
                    'p値': f"{result['p_value']:.3e}",
                    'サンプル数': result['sample_size'],
                    '有意性(p<0.05)': '✓' if result['significant_005'] else '✗',
                    '有意性(p<0.01)': '✓' if result['significant_001'] else '✗'
                })
            
            corr_df = pd.DataFrame(corr_results)
            st.dataframe(corr_df, use_container_width=True)

def show_distribution_analysis(data: pd.DataFrame):
    """分布分析を表示"""
    st.title("📊 分布分析")
    
    numeric_cols = data.select_dtypes(include=[np.number]).columns.tolist()
    
    if not numeric_cols:
        st.warning("数値型のカラムが見つかりません")
        return
    
    # カラム選択（トレンド分析で選択された変数をデフォルトに）
    if st.session_state.selected_trend_variables:
        default_dist_cols = [col for col in st.session_state.selected_trend_variables if col in numeric_cols]
    else:
        default_dist_cols = numeric_cols[:4]
        
    selected_columns = st.multiselect(
        "分析対象の変数を選択", 
        numeric_cols, 
        default=default_dist_cols
    )
    
    if not selected_columns:
        st.warning("分析対象の変数を選択してください")
        return
    
    # タブ分割
    tab1, tab2 = st.tabs(["ヒストグラム", "正規性検定"])
    
    with tab1:
        st.subheader("📊 ヒストグラム")
        
        # ビン数設定
        bins = st.slider("ビン数", 10, 100, 30)
        
        # ヒストグラム表示
        fig_hist = st.session_state.visualizer.create_histogram_grid(data, selected_columns, bins)
        st.plotly_chart(fig_hist, use_container_width=True)
    
    with tab2:
        st.subheader("🔬 正規性検定")
        
        # 正規性検定の実行
        normality_results = st.session_state.stats_calculator.perform_normality_tests(data, selected_columns)
        
        if not normality_results.empty:
            st.dataframe(normality_results, use_container_width=True)
            
            # 結果の解釈
            st.write("**結果の解釈**")
            st.write("""
            - **Shapiro-Wilk検定**: サンプルサイズが小さい場合に適用
            - **Kolmogorov-Smirnov検定**: 連続分布の適合度検定
            - **Anderson-Darling検定**: 正規分布に対してより敏感な検定
            
            p値 > 0.05 の場合、正規分布に従うと判断できます。
            """)
        else:
            st.warning("正規性検定を実行できませんでした")

def show_report_generation(data: pd.DataFrame):
    """レポート生成を表示"""
    st.title("📄 レポート生成")
    
    st.subheader("📋 レポート設定")
    
    # レポートに含める分析項目
    include_basic_stats = st.checkbox("基本統計量", value=True)
    include_correlation = st.checkbox("相関分析", value=True)
    include_normality = st.checkbox("正規性検定", value=True)
    include_outlier_history = st.checkbox("外れ値除去履歴", value=True)
    include_figures = st.checkbox("グラフ", value=True)
    
    if st.button("📊 レポートを生成"):
        with st.spinner("レポートを生成中..."):
            try:
                # 分析結果の準備
                analysis_results = {}
                
                numeric_cols = data.select_dtypes(include=[np.number]).columns.tolist()
                
                if include_basic_stats and numeric_cols:
                    analysis_results['basic_stats'] = (
                        st.session_state.stats_calculator.calculate_basic_statistics(data, numeric_cols)
                    )
                
                if include_correlation and len(numeric_cols) >= 2:
                    analysis_results['correlation_matrix'] = (
                        st.session_state.stats_calculator.calculate_correlation_matrix(data, numeric_cols)
                    )
                
                if include_normality and numeric_cols:
                    analysis_results['normality_tests'] = (
                        st.session_state.stats_calculator.perform_normality_tests(data, numeric_cols)
                    )
                
                # 外れ値除去履歴
                outlier_history = None
                if include_outlier_history:
                    history_export = st.session_state.outlier_remover.export_history()
                    outlier_history = history_export.get('history', [])
                
                # グラフの生成
                figures = {}
                if include_figures and numeric_cols:
                    # トレンドグラフ
                    figures['トレンド分析'] = st.session_state.visualizer.create_trend_line_chart(
                        data, numeric_cols[:3], numeric_cols[:3]
                    )
                    
                    # 相関ヒートマップ
                    if len(numeric_cols) >= 2:
                        figures['相関ヒートマップ'] = st.session_state.visualizer.create_correlation_heatmap(
                            data, numeric_cols[:5]
                        )
                    
                    # ヒストグラム
                    figures['分布分析'] = st.session_state.visualizer.create_histogram_grid(
                        data, numeric_cols[:4]
                    )
                
                # PDFレポート生成（条件付き）
                pdf_bytes = st.session_state.report_generator.generate_analysis_report(
                    data, analysis_results, figures, outlier_history
                )
                
                if pdf_bytes:
                    # ダウンロードボタン
                    st.success("✅ レポートが生成されました")
                    
                    st.download_button(
                        label="📥 PDFレポートをダウンロード",
                        data=pdf_bytes,
                        file_name=f"process_data_analysis_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
                        mime="application/pdf"
                    )
                else:
                    # テキストサマリーを表示
                    st.info("📄 レポートサマリー")
                    
                    summary_text = st.session_state.report_generator.export_data_summary(data)
                    st.text_area("データサマリー", summary_text, height=300)
                    
                    # テキストサマリーのダウンロード
                    st.download_button(
                        label="📥 テキストサマリーをダウンロード",
                        data=summary_text,
                        file_name=f"data_summary_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt",
                        mime="text/plain"
                    )
                
            except Exception as e:
                st.error(f"レポート生成エラー: {str(e)}")
    
    # データエクスポート
    st.subheader("💾 データエクスポート")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("CSV エクスポート"):
            csv_data = st.session_state.data_processor.export_to_csv()
            st.download_button(
                label="📥 CSV ファイルをダウンロード",
                data=csv_data,
                file_name=f"processed_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv"
            )
    
    with col2:
        if st.button("Excel エクスポート"):
            excel_data = st.session_state.data_processor.export_to_excel()
            st.download_button(
                label="📥 Excel ファイルをダウンロード",
                data=excel_data,
                file_name=f"processed_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()