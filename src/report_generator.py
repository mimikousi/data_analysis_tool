import pandas as pd
import numpy as np
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas
from datetime import datetime
import io
import base64
import plotly.graph_objects as go
from typing import Dict, List, Optional, Union
import streamlit as st
import tempfile
import os

class ReportGenerator:
    def __init__(self):
        self.styles = getSampleStyleSheet()
        self.custom_styles = self._create_custom_styles()
        
    def _create_custom_styles(self):
        """カスタムスタイルを作成"""
        custom_styles = {}
        
        # タイトルスタイル
        custom_styles['Title'] = ParagraphStyle(
            'CustomTitle',
            parent=self.styles['Title'],
            fontSize=18,
            spaceAfter=30,
            alignment=TA_CENTER
        )
        
        # 見出しスタイル
        custom_styles['Heading1'] = ParagraphStyle(
            'CustomHeading1',
            parent=self.styles['Heading1'],
            fontSize=14,
            spaceAfter=12,
            textColor=colors.darkblue
        )
        
        custom_styles['Heading2'] = ParagraphStyle(
            'CustomHeading2',
            parent=self.styles['Heading2'],
            fontSize=12,
            spaceAfter=10,
            textColor=colors.darkgreen
        )
        
        # 本文スタイル
        custom_styles['Normal'] = ParagraphStyle(
            'CustomNormal',
            parent=self.styles['Normal'],
            fontSize=10,
            spaceAfter=6
        )
        
        return custom_styles
    
    def generate_analysis_report(self, 
                               data: pd.DataFrame,
                               analysis_results: Dict,
                               figures: Dict[str, go.Figure] = None,
                               outlier_history: List[Dict] = None) -> bytes:
        """総合解析レポートを生成"""
        
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=0.5*inch, bottomMargin=0.5*inch)
        story = []
        
        # タイトルページ
        story.extend(self._create_title_page())
        
        # データ概要
        story.extend(self._create_data_overview(data))
        
        # 外れ値除去履歴
        if outlier_history:
            story.extend(self._create_outlier_history_section(outlier_history))
        
        # 統計分析結果
        story.extend(self._create_statistics_section(analysis_results))
        
        # グラフ
        if figures:
            story.extend(self._create_figures_section(figures))
        
        # 結論・まとめ
        story.extend(self._create_conclusion_section(data, analysis_results))
        
        # PDFを構築
        doc.build(story)
        buffer.seek(0)
        return buffer.getvalue()
    
    def _create_title_page(self) -> List:
        """タイトルページを作成"""
        elements = []
        
        # メインタイトル
        title = Paragraph("プロセスデータ解析レポート", self.custom_styles['Title'])
        elements.append(title)
        elements.append(Spacer(1, 0.5*inch))
        
        # 作成日時
        now = datetime.now()
        date_str = now.strftime("%Y年%m月%d日 %H時%M分")
        date_para = Paragraph(f"作成日時: {date_str}", self.custom_styles['Normal'])
        elements.append(date_para)
        elements.append(Spacer(1, 0.3*inch))
        
        # 自動生成の注記
        note = Paragraph("このレポートは自動生成されました。", self.custom_styles['Normal'])
        elements.append(note)
        elements.append(PageBreak())
        
        return elements
    
    def _create_data_overview(self, data: pd.DataFrame) -> List:
        """データ概要セクションを作成"""
        elements = []
        
        elements.append(Paragraph("1. データ概要", self.custom_styles['Heading1']))
        
        # 基本情報
        info_data = [
            ["項目", "値"],
            ["データ行数", f"{len(data):,} 行"],
            ["データ列数", f"{len(data.columns)} 列"],
            ["データ期間", f"{data.index.min()} ～ {data.index.max()}"],
            ["期間日数", f"{(data.index.max() - data.index.min()).days} 日"],
            ["メモリ使用量", f"{data.memory_usage(deep=True).sum() / 1024 / 1024:.1f} MB"]
        ]
        
        table = Table(info_data, colWidths=[2*inch, 3*inch])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        elements.append(table)
        elements.append(Spacer(1, 0.3*inch))
        
        # 変数一覧
        elements.append(Paragraph("1.1 変数一覧", self.custom_styles['Heading2']))
        
        numeric_cols = data.select_dtypes(include=[np.number]).columns.tolist()
        variable_data = [["変数名", "データ型", "欠損値数", "欠損率"]]
        
        for col in data.columns:
            null_count = data[col].isnull().sum()
            null_rate = (null_count / len(data)) * 100
            variable_data.append([
                col,
                str(data[col].dtype),
                str(null_count),
                f"{null_rate:.1f}%"
            ])
        
        var_table = Table(variable_data, colWidths=[2*inch, 1.5*inch, 1*inch, 1*inch])
        var_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        elements.append(var_table)
        elements.append(Spacer(1, 0.3*inch))
        
        return elements
    
    def _create_outlier_history_section(self, outlier_history: List[Dict]) -> List:
        """外れ値除去履歴セクションを作成"""
        elements = []
        
        elements.append(Paragraph("2. 外れ値除去履歴", self.custom_styles['Heading1']))
        
        if not outlier_history:
            elements.append(Paragraph("外れ値除去は実行されませんでした。", self.custom_styles['Normal']))
            return elements
        
        # 履歴テーブル
        history_data = [["操作日時", "対象変数", "除去方法", "除去件数"]]
        
        for hist in outlier_history:
            timestamp = hist.get('timestamp', datetime.now()).strftime("%Y-%m-%d %H:%M")
            column = hist.get('column', 'N/A')
            method = self._get_method_description(hist)
            removed_count = hist.get('removed_count', 0)
            
            history_data.append([timestamp, column, method, str(removed_count)])
        
        hist_table = Table(history_data, colWidths=[1.5*inch, 1.5*inch, 2.5*inch, 1*inch])
        hist_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        elements.append(hist_table)
        elements.append(Spacer(1, 0.3*inch))
        
        return elements
    
    def _create_statistics_section(self, analysis_results: Dict) -> List:
        """統計分析セクションを作成"""
        elements = []
        
        elements.append(Paragraph("3. 統計分析結果", self.custom_styles['Heading1']))
        
        # 基本統計量
        if 'basic_stats' in analysis_results:
            elements.append(Paragraph("3.1 基本統計量", self.custom_styles['Heading2']))
            
            stats_df = analysis_results['basic_stats']
            if not stats_df.empty:
                # 統計量テーブル（上位5変数のみ表示）
                stats_data = [["変数名", "平均値", "標準偏差", "最小値", "最大値"]]
                
                for _, row in stats_df.head(5).iterrows():
                    stats_data.append([
                        row['変数名'],
                        f"{row['平均値']:.3f}",
                        f"{row['標準偏差']:.3f}",
                        f"{row['最小値']:.3f}",
                        f"{row['最大値']:.3f}"
                    ])
                
                stats_table = Table(stats_data, colWidths=[1.5*inch, 1*inch, 1*inch, 1*inch, 1*inch])
                stats_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black)
                ]))
                
                elements.append(stats_table)
                elements.append(Spacer(1, 0.2*inch))
        
        # 相関分析
        if 'correlation_matrix' in analysis_results:
            elements.append(Paragraph("3.2 相関分析", self.custom_styles['Heading2']))
            
            corr_matrix = analysis_results['correlation_matrix']
            if not corr_matrix.empty and len(corr_matrix) > 1:
                # 高い相関を持つペアを抽出
                high_corr_pairs = []
                for i in range(len(corr_matrix.columns)):
                    for j in range(i+1, len(corr_matrix.columns)):
                        corr_val = corr_matrix.iloc[i, j]
                        if abs(corr_val) > 0.7:  # 相関係数の絶対値が0.7以上
                            high_corr_pairs.append((
                                corr_matrix.columns[i],
                                corr_matrix.columns[j],
                                corr_val
                            ))
                
                if high_corr_pairs:
                    corr_data = [["変数1", "変数2", "相関係数"]]
                    for var1, var2, corr in high_corr_pairs[:10]:  # 上位10ペア
                        corr_data.append([var1, var2, f"{corr:.3f}"])
                    
                    corr_table = Table(corr_data, colWidths=[2*inch, 2*inch, 1*inch])
                    corr_table.setStyle(TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, -1), 8),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                        ('GRID', (0, 0), (-1, -1), 1, colors.black)
                    ]))
                    
                    elements.append(corr_table)
                else:
                    elements.append(Paragraph("高い相関を持つ変数ペアは見つかりませんでした。", 
                                            self.custom_styles['Normal']))
                elements.append(Spacer(1, 0.2*inch))
        
        return elements
    
    def _create_figures_section(self, figures: Dict[str, go.Figure]) -> List:
        """グラフセクションを作成"""
        elements = []
        
        elements.append(Paragraph("4. グラフ", self.custom_styles['Heading1']))
        
        for fig_name, fig in figures.items():
            elements.append(Paragraph(f"4.{list(figures.keys()).index(fig_name)+1} {fig_name}", 
                                    self.custom_styles['Heading2']))
            
            # グラフを画像として保存
            try:
                img_bytes = fig.to_image(format="png", width=600, height=400)
                with tempfile.NamedTemporaryFile(delete=False, suffix='.png') as tmp_file:
                    tmp_file.write(img_bytes)
                    tmp_file.flush()
                    
                    img = Image(tmp_file.name, width=5*inch, height=3.3*inch)
                    elements.append(img)
                    elements.append(Spacer(1, 0.2*inch))
                    
                # 一時ファイルを削除
                os.unlink(tmp_file.name)
                
            except Exception as e:
                elements.append(Paragraph(f"グラフの生成に失敗しました: {str(e)}", 
                                        self.custom_styles['Normal']))
        
        return elements
    
    def _create_conclusion_section(self, data: pd.DataFrame, analysis_results: Dict) -> List:
        """結論セクションを作成"""
        elements = []
        
        elements.append(Paragraph("5. 結論・まとめ", self.custom_styles['Heading1']))
        
        # データの概要
        summary_text = f"""
        本解析では、{len(data)}行 × {len(data.columns)}列のプロセスデータを対象として、
        時系列分析、統計分析、相関分析を実施しました。
        
        期間: {data.index.min()} ～ {data.index.max()}
        解析期間: {(data.index.max() - data.index.min()).days} 日間
        """
        
        elements.append(Paragraph(summary_text, self.custom_styles['Normal']))
        elements.append(Spacer(1, 0.2*inch))
        
        # 主要な発見
        if 'basic_stats' in analysis_results:
            stats_df = analysis_results['basic_stats']
            if not stats_df.empty:
                # 変動係数の大きい変数を特定
                high_cv_vars = []
                for _, row in stats_df.iterrows():
                    if row['標準偏差'] > 0 and row['平均値'] != 0:
                        cv = abs(row['標準偏差'] / row['平均値'])
                        if cv > 0.5:  # 変動係数が50%以上
                            high_cv_vars.append((row['変数名'], cv))
                
                if high_cv_vars:
                    high_cv_vars.sort(key=lambda x: x[1], reverse=True)
                    var_names = [var[0] for var in high_cv_vars[:3]]
                    
                    findings_text = f"""
                    主要な発見:
                    - 変動の大きい変数: {', '.join(var_names)}
                    - これらの変数は特に注意深い監視が必要です
                    """
                    elements.append(Paragraph(findings_text, self.custom_styles['Normal']))
        
        # 推奨事項
        recommendations = """
        推奨事項:
        1. 定期的なデータ品質チェックの実施
        2. 異常値の早期発見のための監視システムの構築
        3. 相関の高い変数グループの統合的な分析
        4. トレンド変化の継続的な監視
        """
        
        elements.append(Paragraph(recommendations, self.custom_styles['Normal']))
        
        return elements
    
    def _get_method_description(self, operation: Dict) -> str:
        """操作の説明テキストを生成"""
        
        if 'method' in operation:
            method_names = {
                'iqr': 'IQR法',
                'zscore': 'Z-score法'
            }
            method_name = method_names.get(operation['method'], operation['method'])
            return f"{method_name}"
        elif 'x_range' in operation or 'y_range' in operation:
            return "範囲指定"
        else:
            return "不明"
    
    def export_data_summary(self, data: pd.DataFrame) -> str:
        """データサマリーをテキスト形式でエクスポート"""
        
        summary_lines = []
        summary_lines.append("=== データサマリー ===")
        summary_lines.append(f"作成日時: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        summary_lines.append("")
        
        summary_lines.append("【基本情報】")
        summary_lines.append(f"データ行数: {len(data):,}")
        summary_lines.append(f"データ列数: {len(data.columns)}")
        
        if isinstance(data.index, pd.DatetimeIndex):
            summary_lines.append(f"データ期間: {data.index.min()} ～ {data.index.max()}")
            summary_lines.append(f"期間日数: {(data.index.max() - data.index.min()).days}")
        
        summary_lines.append("")
        
        # 数値変数の統計
        numeric_cols = data.select_dtypes(include=[np.number]).columns.tolist()
        if numeric_cols:
            summary_lines.append("【数値変数統計】")
            for col in numeric_cols:
                col_data = data[col].dropna()
                if not col_data.empty:
                    summary_lines.append(f"{col}:")
                    summary_lines.append(f"  平均: {col_data.mean():.3f}")
                    summary_lines.append(f"  標準偏差: {col_data.std():.3f}")
                    summary_lines.append(f"  範囲: {col_data.min():.3f} ～ {col_data.max():.3f}")
                    summary_lines.append(f"  欠損率: {(data[col].isnull().sum() / len(data) * 100):.1f}%")
                    summary_lines.append("")
        
        return "\n".join(summary_lines)
    
    def create_quick_report(self, data: pd.DataFrame, title: str = "データ解析レポート") -> bytes:
        """簡易レポートを生成"""
        
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        story = []
        
        # タイトル
        story.append(Paragraph(title, self.custom_styles['Title']))
        story.append(Spacer(1, 0.3*inch))
        
        # データ概要
        story.extend(self._create_data_overview(data))
        
        doc.build(story)
        buffer.seek(0)
        return buffer.getvalue()