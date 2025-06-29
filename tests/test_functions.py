import pytest
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import sys
import os

# プロジェクトルートをパスに追加
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from src.data_processor import DataProcessor
from src.outlier_removal import OutlierRemover
from src.statistics import StatisticsCalculator
from src.visualization import DataVisualizer

class TestDataProcessor:
    """DataProcessorクラスのテスト"""
    
    def setup_method(self):
        """テスト前の準備"""
        self.processor = DataProcessor()
        
        # テスト用データの作成
        dates = pd.date_range('2023-01-01', periods=100, freq='H')
        data = {
            'temperature': np.random.normal(25, 5, 100),
            'pressure': np.random.normal(100, 10, 100),
            'flow_rate': np.random.normal(50, 8, 100)
        }
        self.test_data = pd.DataFrame(data, index=dates)
        self.processor.data = self.test_data
        self.processor.original_data = self.test_data.copy()
    
    def test_get_numeric_columns(self):
        """数値列の取得テスト"""
        numeric_cols = self.processor.get_numeric_columns()
        expected_cols = ['temperature', 'pressure', 'flow_rate']
        assert set(numeric_cols) == set(expected_cols)
    
    def test_get_data_info(self):
        """データ情報の取得テスト"""
        info = self.processor.get_data_info()
        
        assert info['shape'] == (100, 3)
        assert len(info['columns']) == 3
        assert 'date_range' in info
        assert info['date_range'] is not None
    
    def test_get_column_statistics(self):
        """カラム統計の取得テスト"""
        stats = self.processor.get_column_statistics('temperature')
        
        assert 'count' in stats
        assert 'mean' in stats
        assert 'std' in stats
        assert stats['count'] == 100
    
    def test_filter_data_by_date_range(self):
        """日付範囲フィルタリングテスト"""
        start_date = pd.to_datetime('2023-01-01 10:00:00')
        end_date = pd.to_datetime('2023-01-01 20:00:00')
        
        filtered = self.processor.filter_data_by_date_range(start_date, end_date)
        
        assert len(filtered) <= len(self.test_data)
        assert filtered.index.min() >= start_date
        assert filtered.index.max() <= end_date

class TestOutlierRemover:
    """OutlierRemoverクラスのテスト"""
    
    def setup_method(self):
        """テスト前の準備"""
        self.remover = OutlierRemover()
        
        # テスト用データの作成（外れ値を含む）
        dates = pd.date_range('2023-01-01', periods=100, freq='H')
        data = np.random.normal(25, 5, 100)
        data[50] = 100  # 外れ値を追加
        data[75] = -50  # 外れ値を追加
        
        self.test_data = pd.DataFrame({
            'temperature': data,
            'pressure': np.random.normal(100, 10, 100)
        }, index=dates)
        
        self.remover.initialize_data(self.test_data)
    
    def test_get_outlier_candidates_iqr(self):
        """IQR法による外れ値候補取得テスト"""
        outlier_mask = self.remover.get_outlier_candidates(
            self.test_data, 'temperature', 'iqr', 1.5
        )
        
        assert isinstance(outlier_mask, pd.Series)
        assert len(outlier_mask) == len(self.test_data)
        assert outlier_mask.dtype == bool
        assert outlier_mask.any()  # 外れ値が検出されるはず
    
    def test_get_outlier_candidates_zscore(self):
        """Z-score法による外れ値候補取得テスト"""
        outlier_mask = self.remover.get_outlier_candidates(
            self.test_data, 'temperature', 'zscore', 3.0
        )
        
        assert isinstance(outlier_mask, pd.Series)
        assert len(outlier_mask) == len(self.test_data)
        assert outlier_mask.dtype == bool
    
    def test_remove_outliers_by_statistical_method(self):
        """統計的手法による外れ値除去テスト"""
        original_len = len(self.test_data)
        
        filtered_data = self.remover.remove_outliers_by_statistical_method(
            self.test_data, 'temperature', 'iqr', 1.5
        )
        
        assert len(filtered_data) <= original_len
        assert isinstance(filtered_data, pd.DataFrame)
        
        # 履歴が記録されているかチェック
        history = self.remover.get_history_summary()
        assert len(history) > 0
    
    def test_remove_outliers_by_range(self):
        """範囲指定による外れ値除去テスト"""
        original_len = len(self.test_data)
        
        # 値の範囲指定
        y_range = (0, 50)  # temperatureの範囲を0-50に制限
        
        filtered_data = self.remover.remove_outliers_by_range(
            self.test_data, 'temperature', None, y_range
        )
        
        assert len(filtered_data) <= original_len
        assert isinstance(filtered_data, pd.DataFrame)
        
        # 指定範囲外の値が除去されているかチェック
        remaining_values = filtered_data['temperature']
        assert remaining_values.min() >= y_range[0]
        assert remaining_values.max() <= y_range[1]
    
    def test_history_management(self):
        """履歴管理テスト"""
        # 複数回の外れ値除去を実行
        self.remover.remove_outliers_by_statistical_method(
            self.test_data, 'temperature', 'iqr', 1.5
        )
        self.remover.remove_outliers_by_statistical_method(
            self.test_data, 'pressure', 'zscore', 3.0
        )
        
        history = self.remover.get_history_summary()
        assert len(history) == 2
        
        # 履歴の内容をチェック
        assert '操作ID' in history.columns
        assert '対象カラム' in history.columns
        assert '除去件数' in history.columns

class TestStatisticsCalculator:
    """StatisticsCalculatorクラスのテスト"""
    
    def setup_method(self):
        """テスト前の準備"""
        self.calculator = StatisticsCalculator()
        
        # テスト用データの作成
        dates = pd.date_range('2023-01-01', periods=100, freq='H')
        self.test_data = pd.DataFrame({
            'temperature': np.random.normal(25, 5, 100),
            'pressure': np.random.normal(100, 10, 100),
            'flow_rate': np.random.normal(50, 8, 100)
        }, index=dates)
    
    def test_calculate_basic_statistics(self):
        """基本統計量計算テスト"""
        stats_df = self.calculator.calculate_basic_statistics(self.test_data)
        
        assert isinstance(stats_df, pd.DataFrame)
        assert len(stats_df) == 3  # 3つの数値列
        assert '変数名' in stats_df.columns
        assert '平均値' in stats_df.columns
        assert '標準偏差' in stats_df.columns
        
        # 統計値の妥当性チェック
        temp_stats = stats_df[stats_df['変数名'] == 'temperature'].iloc[0]
        assert 20 < temp_stats['平均値'] < 30  # 大体25周辺
        assert temp_stats['標準偏差'] > 0
    
    def test_calculate_correlation_matrix(self):
        """相関行列計算テスト"""
        corr_matrix = self.calculator.calculate_correlation_matrix(self.test_data)
        
        assert isinstance(corr_matrix, pd.DataFrame)
        assert corr_matrix.shape == (3, 3)
        
        # 対角要素は1になるはず
        for i in range(3):
            assert abs(corr_matrix.iloc[i, i] - 1.0) < 0.001
    
    def test_perform_normality_tests(self):
        """正規性検定テスト"""
        normality_df = self.calculator.perform_normality_tests(self.test_data)
        
        assert isinstance(normality_df, pd.DataFrame)
        assert len(normality_df) == 3  # 3つの数値列
        assert 'Shapiro統計量' in normality_df.columns
        assert 'Shapiro p値' in normality_df.columns
        assert 'KS統計量' in normality_df.columns
    
    def test_calculate_outlier_statistics(self):
        """外れ値統計計算テスト"""
        stats = self.calculator.calculate_outlier_statistics(
            self.test_data, 'temperature', 'iqr'
        )
        
        assert isinstance(stats, dict)
        assert 'method' in stats
        assert 'total_points' in stats
        assert 'outlier_count' in stats
        assert 'lower_bound' in stats
        assert 'upper_bound' in stats
        assert stats['method'] == 'iqr'
        assert stats['total_points'] == 100

class TestDataVisualizer:
    """DataVisualizerクラスのテスト"""
    
    def setup_method(self):
        """テスト前の準備"""
        self.visualizer = DataVisualizer()
        
        # テスト用データの作成
        dates = pd.date_range('2023-01-01', periods=50, freq='H')
        self.test_data = pd.DataFrame({
            'temperature': np.random.normal(25, 5, 50),
            'pressure': np.random.normal(100, 10, 50),
            'flow_rate': np.random.normal(50, 8, 50)
        }, index=dates)
    
    def test_create_trend_line_chart(self):
        """トレンド折れ線グラフ作成テスト"""
        columns = ['temperature', 'pressure']
        fig = self.visualizer.create_trend_line_chart(
            self.test_data, columns, columns
        )
        
        assert fig is not None
        assert len(fig.data) >= len(columns)  # トレース数がカラム数以上
    
    def test_create_scatter_matrix(self):
        """散布図マトリックス作成テスト"""
        columns = ['temperature', 'pressure']
        fig = self.visualizer.create_scatter_matrix(self.test_data, columns)
        
        assert fig is not None
        assert len(fig.data) > 0
    
    def test_create_histogram_grid(self):
        """ヒストグラムグリッド作成テスト"""
        columns = ['temperature', 'pressure']
        fig = self.visualizer.create_histogram_grid(self.test_data, columns)
        
        assert fig is not None
        assert len(fig.data) >= len(columns)
    
    def test_create_correlation_heatmap(self):
        """相関ヒートマップ作成テスト"""
        columns = ['temperature', 'pressure', 'flow_rate']
        fig = self.visualizer.create_correlation_heatmap(self.test_data, columns)
        
        assert fig is not None
        assert len(fig.data) > 0
    
    def test_export_figure(self):
        """グラフエクスポートテスト"""
        columns = ['temperature']
        fig = self.visualizer.create_trend_line_chart(
            self.test_data, columns, columns
        )
        
        # PNG形式でエクスポート
        img_bytes = self.visualizer.export_figure(fig, 'png', 800, 600)
        assert isinstance(img_bytes, bytes)
        assert len(img_bytes) > 0

# テスト実行用の関数
def run_tests():
    """すべてのテストを実行"""
    pytest.main([__file__, '-v'])

if __name__ == "__main__":
    run_tests()