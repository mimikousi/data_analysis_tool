import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import streamlit as st
from typing import List, Dict, Tuple, Optional, Union
from datetime import datetime
import io
import base64

class DataVisualizer:
    def __init__(self):
        self.color_palette = px.colors.qualitative.Set3
        
    def create_trend_line_chart(self,
                               data: pd.DataFrame,
                               columns: List[str],
                               y1_columns: List[str] = None,
                               y2_columns: List[str] = None,
                               y1_range: Tuple[float, float] = None,
                               y2_range: Tuple[float, float] = None,
                               show_statistics: bool = True,
                               statistics_period: Tuple[datetime, datetime] = None,
                               sigma_multiplier: int = 1) -> go.Figure:
        """トレンド折れ線グラフを作成"""
        
        if data.empty or not columns:
            return go.Figure()
            
        # Y1軸とY2軸の設定
        if y1_columns is None:
            y1_columns = columns
        if y2_columns is None:
            y2_columns = []
            
        # サブプロットの作成（Y2軸がある場合は secondary_y=True）
        has_y2 = len(y2_columns) > 0
        fig = make_subplots(specs=[[{"secondary_y": has_y2}]])
        
        # Y1軸のプロット
        for i, col in enumerate(y1_columns):
            if col in data.columns:
                fig.add_trace(
                    go.Scatter(
                        x=data.index,
                        y=data[col],
                        mode='lines',
                        name=col,
                        line=dict(color=self.color_palette[i % len(self.color_palette)]),
                        yaxis='y'
                    )
                )
                
        # Y2軸のプロット
        for i, col in enumerate(y2_columns):
            if col in data.columns:
                fig.add_trace(
                    go.Scatter(
                        x=data.index,
                        y=data[col],
                        mode='lines',
                        name=col,
                        line=dict(color=self.color_palette[(i + len(y1_columns)) % len(self.color_palette)]),
                        yaxis='y2'
                    ),
                    secondary_y=True
                )
        
        # 統計線の追加
        if show_statistics:
            self._add_statistics_lines(fig, data, y1_columns, y2_columns, 
                                     statistics_period, sigma_multiplier, has_y2)
        
        # レイアウトの設定
        fig.update_layout(
            title="時系列トレンドグラフ",
            xaxis_title="時間",
            hovermode='x unified',
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            )
        )
        
        # Y軸の範囲設定
        if y1_range:
            fig.update_yaxes(range=y1_range, secondary_y=False)
        if y2_range and has_y2:
            fig.update_yaxes(range=y2_range, secondary_y=True)
            
        # Y軸タイトル
        fig.update_yaxes(title_text="Y1軸", secondary_y=False)
        if has_y2:
            fig.update_yaxes(title_text="Y2軸", secondary_y=True)
            
        return fig
    
    def _add_statistics_lines(self, fig, data, y1_columns, y2_columns, 
                             statistics_period, sigma_multiplier, has_y2):
        """統計線を追加"""
        
        # 統計計算用のデータを選択
        if statistics_period:
            mask = (data.index >= statistics_period[0]) & (data.index <= statistics_period[1])
            stats_data = data.loc[mask]
            # 期間のハイライト表示
            fig.add_vrect(
                x0=statistics_period[0], x1=statistics_period[1],
                fillcolor="lightgray", opacity=0.2,
                line_width=0
            )
        else:
            stats_data = data
            
        # Y1軸の統計線
        for col in y1_columns:
            if col in stats_data.columns:
                self._add_column_statistics(fig, data, stats_data, col, sigma_multiplier, False)
                
        # Y2軸の統計線
        if has_y2:
            for col in y2_columns:
                if col in stats_data.columns:
                    self._add_column_statistics(fig, data, stats_data, col, sigma_multiplier, True)
    
    def _add_column_statistics(self, fig, full_data, stats_data, column, sigma_multiplier, secondary_y):
        """指定カラムの統計線を追加"""
        
        col_data = stats_data[column].dropna()
        if col_data.empty:
            return
            
        mean_val = col_data.mean()
        std_val = col_data.std()
        
        # 線の色を軸によって変更
        line_color = "red" if not secondary_y else "green"
        sigma_color = "orange" if not secondary_y else "lightgreen"
        
        # 平均線
        fig.add_hline(
            y=mean_val,
            line_dash="dash",
            line_color=line_color,
            annotation_text=f"{column} 平均",
            yref="y2" if secondary_y else "y"
        )
        
        # 平均±σ線
        for sign, name in [(1, "上"), (-1, "下")]:
            sigma_val = mean_val + sign * sigma_multiplier * std_val
            fig.add_hline(
                y=sigma_val,
                line_dash="dot",
                line_color=sigma_color,
                annotation_text=f"{column} 平均{name}±{sigma_multiplier}σ",
                yref="y2" if secondary_y else "y"
            )
    
    def create_scatter_matrix(self, data: pd.DataFrame, columns: List[str]) -> go.Figure:
        """散布図マトリックスを作成"""
        
        if data.empty or len(columns) < 2:
            return go.Figure()
            
        n_cols = len(columns)
        
        # サブプロットの作成
        fig = make_subplots(
            rows=n_cols, cols=n_cols,
            subplot_titles=[f"{columns[i]} vs {columns[j]}" 
                          for i in range(n_cols) for j in range(n_cols)]
        )
        
        # 各散布図を作成
        for i, col1 in enumerate(columns):
            for j, col2 in enumerate(columns):
                if i == j:
                    # 対角線上はヒストグラム
                    fig.add_trace(
                        go.Histogram(x=data[col1], name=f"{col1} 分布"),
                        row=i+1, col=j+1
                    )
                else:
                    # 散布図
                    clean_data = data[[col1, col2]].dropna()
                    
                    fig.add_trace(
                        go.Scatter(
                            x=clean_data[col2],
                            y=clean_data[col1],
                            mode='markers',
                            name=f"{col1} vs {col2}",
                            marker=dict(size=4, opacity=0.6)
                        ),
                        row=i+1, col=j+1
                    )
                    
                    # 線形回帰線を追加
                    if len(clean_data) > 1:
                        self._add_regression_line(fig, clean_data, col1, col2, i+1, j+1)
        
        # レイアウト設定
        fig.update_layout(
            title="散布図マトリックス",
            showlegend=False,
            height=300 * n_cols,  # 高さを増加
            width=200 * n_cols   # 幅も調整
        )
        
        return fig
    
    def _add_regression_line(self, fig, data, col1, col2, row, col):
        """回帰直線を追加"""
        
        try:
            # 線形回帰計算
            x = data[col2].values
            y = data[col1].values
            
            # 回帰係数の計算
            coef = np.polyfit(x, y, 1)
            regression_line = np.poly1d(coef)
            
            # 相関係数の計算
            correlation = np.corrcoef(x, y)[0, 1]
            r_squared = correlation ** 2
            
            # 回帰直線をプロット
            x_range = np.linspace(x.min(), x.max(), 100)
            fig.add_trace(
                go.Scatter(
                    x=x_range,
                    y=regression_line(x_range),
                    mode='lines',
                    name=f"回帰直線 (r={correlation:.3f})",
                    line=dict(color='red', dash='dash')
                ),
                row=row, col=col
            )
            
            # 統計情報をアノテーションで追加
            fig.add_annotation(
                x=x.max(), y=y.min(),
                text=f"y = {coef[0]:.3f}x + {coef[1]:.3f}<br>r = {correlation:.3f}<br>R² = {r_squared:.3f}",
                showarrow=False,
                row=row, col=col,
                xanchor="right", yanchor="bottom"
            )
            
        except Exception as e:
            st.warning(f"回帰分析エラー ({col1} vs {col2}): {str(e)}")
    
    def create_histogram_grid(self, data: pd.DataFrame, columns: List[str], 
                             bins: int = 30) -> go.Figure:
        """ヒストグラムグリッドを作成"""
        
        if data.empty or not columns:
            return go.Figure()
            
        n_cols = min(3, len(columns))  # 最大3列
        n_rows = (len(columns) + n_cols - 1) // n_cols
        
        fig = make_subplots(
            rows=n_rows, cols=n_cols,
            subplot_titles=columns,
            vertical_spacing=0.08
        )
        
        for i, col in enumerate(columns):
            row = i // n_cols + 1
            col_pos = i % n_cols + 1
            
            if col in data.columns:
                col_data = data[col].dropna()
                
                if not col_data.empty:
                    # ヒストグラム
                    fig.add_trace(
                        go.Histogram(
                            x=col_data,
                            nbinsx=bins,
                            name=f"{col} 分布",
                            opacity=0.7
                        ),
                        row=row, col=col_pos
                    )
                    
                    # 正規分布曲線を追加
                    self._add_normal_distribution_curve(fig, col_data, row, col_pos)
        
        fig.update_layout(
            title="データ分布ヒストグラム",
            showlegend=False,
            height=300 * n_rows
        )
        
        return fig
    
    def _add_normal_distribution_curve(self, fig, data, row, col):
        """正規分布曲線を追加"""
        
        try:
            from scipy import stats
            
            mean = data.mean()
            std = data.std()
            
            # 正規分布の範囲
            x_range = np.linspace(data.min(), data.max(), 100)
            normal_curve = stats.norm.pdf(x_range, mean, std)
            
            # スケーリング（ヒストグラムの面積に合わせる）
            bin_width = (data.max() - data.min()) / 30  # デフォルトのbin数
            scale_factor = len(data) * bin_width
            normal_curve_scaled = normal_curve * scale_factor
            
            fig.add_trace(
                go.Scatter(
                    x=x_range,
                    y=normal_curve_scaled,
                    mode='lines',
                    name='正規分布',
                    line=dict(color='red', width=2)
                ),
                row=row, col=col
            )
            
        except ImportError:
            pass  # scipyが利用できない場合はスキップ
        except Exception as e:
            st.warning(f"正規分布曲線の描画エラー: {str(e)}")
    
    def create_correlation_heatmap(self, data: pd.DataFrame, columns: List[str]) -> go.Figure:
        """相関ヒートマップを作成"""
        
        if data.empty or len(columns) < 2:
            return go.Figure()
            
        # 相関行列の計算
        corr_matrix = data[columns].corr()
        
        fig = go.Figure(data=go.Heatmap(
            z=corr_matrix.values,
            x=corr_matrix.columns,
            y=corr_matrix.columns,
            colorscale='RdBu',
            zmid=0,
            text=np.round(corr_matrix.values, 3),
            texttemplate="%{text}",
            textfont={"size": 10},
            hoverongaps=False
        ))
        
        fig.update_layout(
            title="相関ヒートマップ",
            xaxis_title="変数",
            yaxis_title="変数"
        )
        
        return fig
    
    def export_figure(self, fig: go.Figure, format: str = 'png', 
                     width: int = 1200, height: int = 800) -> bytes:
        """グラフを画像としてエクスポート"""
        
        try:
            if format.lower() == 'png':
                img_bytes = fig.to_image(format="png", width=width, height=height, engine="kaleido")
            elif format.lower() == 'svg':
                img_bytes = fig.to_image(format="svg", width=width, height=height, engine="kaleido")
            elif format.lower() == 'pdf':
                img_bytes = fig.to_image(format="pdf", width=width, height=height, engine="kaleido")
            else:
                raise ValueError(f"サポートされていない形式: {format}")
                
            return img_bytes
            
        except Exception as e:
            st.warning(f"画像エクスポート機能はローカル環境でのみ利用可能です: {str(e)}")
            return b''
    
    def create_download_link(self, fig: go.Figure, filename: str, format: str = 'png') -> str:
        """ダウンロードリンクを作成"""
        
        img_bytes = self.export_figure(fig, format)
        if img_bytes:
            b64 = base64.b64encode(img_bytes).decode()
            href = f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}.{format}">📥 {filename}.{format} をダウンロード</a>'
            return href
        return ""
    
    def get_figure_stats(self, fig: go.Figure) -> Dict:
        """グラフの統計情報を取得"""
        
        stats = {
            'trace_count': len(fig.data),
            'trace_types': [trace.type for trace in fig.data],
            'has_annotations': len(fig.layout.annotations) > 0,
            'title': fig.layout.title.text if fig.layout.title else None
        }
        
        return stats