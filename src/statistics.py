import pandas as pd
import numpy as np
from scipy import stats
from typing import Dict, List, Tuple, Optional, Union
import streamlit as st

class StatisticsCalculator:
    def __init__(self):
        pass
    
    def calculate_basic_statistics(self, data: pd.DataFrame, columns: List[str] = None) -> pd.DataFrame:
        """基本統計量を計算"""
        
        if data.empty:
            return pd.DataFrame()
            
        if columns is None:
            columns = data.select_dtypes(include=[np.number]).columns.tolist()
        
        stats_data = []
        
        for col in columns:
            if col in data.columns:
                col_data = data[col].dropna()
                
                if not col_data.empty:
                    stats_dict = {
                        '変数名': col,
                        'データ数': len(col_data),
                        '平均値': col_data.mean(),
                        '標準偏差': col_data.std(),
                        '最小値': col_data.min(),
                        '25%': col_data.quantile(0.25),
                        '中央値': col_data.median(),
                        '75%': col_data.quantile(0.75),
                        '最大値': col_data.max(),
                        '歪度': col_data.skew(),
                        '尖度': col_data.kurtosis(),
                        '欠損値数': data[col].isnull().sum(),
                        '欠損率(%)': (data[col].isnull().sum() / len(data)) * 100
                    }
                    stats_data.append(stats_dict)
        
        return pd.DataFrame(stats_data)
    
    def calculate_correlation_matrix(self, data: pd.DataFrame, columns: List[str] = None) -> pd.DataFrame:
        """相関行列を計算"""
        
        if data.empty:
            return pd.DataFrame()
            
        if columns is None:
            columns = data.select_dtypes(include=[np.number]).columns.tolist()
        
        return data[columns].corr()
    
    def calculate_correlation_with_significance(self, data: pd.DataFrame, 
                                               columns: List[str] = None) -> Dict:
        """相関係数と有意性検定を計算"""
        
        if data.empty or len(columns) < 2:
            return {}
            
        if columns is None:
            columns = data.select_dtypes(include=[np.number]).columns.tolist()
        
        results = {}
        
        for i, col1 in enumerate(columns):
            for j, col2 in enumerate(columns):
                if i < j:  # 重複を避ける
                    clean_data = data[[col1, col2]].dropna()
                    
                    if len(clean_data) > 2:
                        corr, p_value = stats.pearsonr(clean_data[col1], clean_data[col2])
                        
                        results[f"{col1}_vs_{col2}"] = {
                            'correlation': corr,
                            'p_value': p_value,
                            'sample_size': len(clean_data),
                            'significant_005': p_value < 0.05,
                            'significant_001': p_value < 0.01
                        }
        
        return results
    
    def perform_normality_tests(self, data: pd.DataFrame, columns: List[str] = None) -> pd.DataFrame:
        """正規性検定を実行"""
        
        if data.empty:
            return pd.DataFrame()
            
        if columns is None:
            columns = data.select_dtypes(include=[np.number]).columns.tolist()
        
        test_results = []
        
        for col in columns:
            if col in data.columns:
                col_data = data[col].dropna()
                
                if len(col_data) >= 3:
                    try:
                        # Shapiro-Wilk検定
                        shapiro_stat, shapiro_p = stats.shapiro(col_data[:5000])  # 最大5000サンプル
                        
                        # Kolmogorov-Smirnov検定
                        ks_stat, ks_p = stats.kstest(col_data, 'norm', 
                                                    args=(col_data.mean(), col_data.std()))
                        
                        # Anderson-Darling検定
                        ad_result = stats.anderson(col_data, dist='norm')
                        
                        result_dict = {
                            '変数名': col,
                            'サンプル数': len(col_data),
                            'Shapiro統計量': shapiro_stat,
                            'Shapiro p値': shapiro_p,
                            'Shapiro正規性': 'Yes' if shapiro_p > 0.05 else 'No',
                            'KS統計量': ks_stat,
                            'KS p値': ks_p,
                            'KS正規性': 'Yes' if ks_p > 0.05 else 'No',
                            'AD統計量': ad_result.statistic,
                            'AD臨界値(5%)': ad_result.critical_values[2],
                            'AD正規性': 'Yes' if ad_result.statistic < ad_result.critical_values[2] else 'No'
                        }
                        
                        test_results.append(result_dict)
                        
                    except Exception as e:
                        st.warning(f"正規性検定エラー ({col}): {str(e)}")
        
        return pd.DataFrame(test_results)
    
    def calculate_outlier_statistics(self, data: pd.DataFrame, column: str, 
                                   method: str = 'iqr') -> Dict:
        """外れ値の統計情報を計算"""
        
        if data.empty or column not in data.columns:
            return {}
            
        col_data = data[column].dropna()
        
        if col_data.empty:
            return {}
        
        if method == 'iqr':
            Q1 = col_data.quantile(0.25)
            Q3 = col_data.quantile(0.75)
            IQR = Q3 - Q1
            lower_bound = Q1 - 1.5 * IQR
            upper_bound = Q3 + 1.5 * IQR
            
        elif method == 'zscore':
            mean = col_data.mean()
            std = col_data.std()
            lower_bound = mean - 3 * std
            upper_bound = mean + 3 * std
            
        else:
            return {}
        
        outliers = col_data[(col_data < lower_bound) | (col_data > upper_bound)]
        
        return {
            'method': method,
            'total_points': len(col_data),
            'outlier_count': len(outliers),
            'outlier_percentage': (len(outliers) / len(col_data)) * 100,
            'lower_bound': lower_bound,
            'upper_bound': upper_bound,
            'outlier_values': outliers.tolist()[:100]  # 最大100個
        }
    
    def calculate_time_series_statistics(self, data: pd.DataFrame, column: str) -> Dict:
        """時系列統計を計算"""
        
        if data.empty or column not in data.columns:
            return {}
            
        if not isinstance(data.index, pd.DatetimeIndex):
            return {}
            
        col_data = data[column].dropna()
        
        if col_data.empty:
            return {}
        
        # 基本統計
        stats_dict = {
            'period_start': col_data.index.min(),
            'period_end': col_data.index.max(),
            'total_duration': col_data.index.max() - col_data.index.min(),
            'data_points': len(col_data),
            'mean': col_data.mean(),
            'std': col_data.std(),
            'trend': None,
            'seasonality': None
        }
        
        # トレンド分析（線形回帰）
        try:
            x = np.arange(len(col_data))
            slope, intercept, r_value, p_value, std_err = stats.linregress(x, col_data.values)
            
            stats_dict['trend'] = {
                'slope': slope,
                'r_squared': r_value**2,
                'p_value': p_value,
                'trend_direction': 'increasing' if slope > 0 else 'decreasing' if slope < 0 else 'stable'
            }
        except:
            pass
        
        # 自己相関（ラグ1）
        try:
            if len(col_data) > 1:
                autocorr_lag1 = col_data.autocorr(lag=1)
                stats_dict['autocorr_lag1'] = autocorr_lag1
        except:
            pass
        
        return stats_dict
    
    def calculate_period_comparison(self, data: pd.DataFrame, column: str,
                                  period1: Tuple, period2: Tuple) -> Dict:
        """期間比較統計を計算"""
        
        if data.empty or column not in data.columns:
            return {}
            
        if not isinstance(data.index, pd.DatetimeIndex):
            return {}
        
        # 期間1のデータ
        mask1 = (data.index >= period1[0]) & (data.index <= period1[1])
        data1 = data.loc[mask1, column].dropna()
        
        # 期間2のデータ
        mask2 = (data.index >= period2[0]) & (data.index <= period2[1])
        data2 = data.loc[mask2, column].dropna()
        
        if data1.empty or data2.empty:
            return {}
        
        # 基本統計の比較
        comparison = {
            'period1': {
                'start': period1[0],
                'end': period1[1],
                'count': len(data1),
                'mean': data1.mean(),
                'std': data1.std(),
                'min': data1.min(),
                'max': data1.max()
            },
            'period2': {
                'start': period2[0],
                'end': period2[1],
                'count': len(data2),
                'mean': data2.mean(),
                'std': data2.std(),
                'min': data2.min(),
                'max': data2.max()
            }
        }
        
        # 統計的検定
        try:
            # t検定（平均の差の検定）
            t_stat, t_p = stats.ttest_ind(data1, data2)
            comparison['t_test'] = {
                'statistic': t_stat,
                'p_value': t_p,
                'significant': t_p < 0.05
            }
            
            # F検定（分散の差の検定）
            f_stat = data1.var() / data2.var() if data2.var() != 0 else np.inf
            f_p = 2 * min(stats.f.cdf(f_stat, len(data1)-1, len(data2)-1),
                         1 - stats.f.cdf(f_stat, len(data1)-1, len(data2)-1))
            comparison['f_test'] = {
                'statistic': f_stat,
                'p_value': f_p,
                'significant': f_p < 0.05
            }
            
            # Mann-Whitney U検定（ノンパラメトリック）
            u_stat, u_p = stats.mannwhitneyu(data1, data2, alternative='two-sided')
            comparison['mann_whitney'] = {
                'statistic': u_stat,
                'p_value': u_p,
                'significant': u_p < 0.05
            }
            
        except Exception as e:
            st.warning(f"統計検定エラー: {str(e)}")
        
        return comparison
    
    def generate_statistics_summary(self, data: pd.DataFrame, columns: List[str] = None) -> str:
        """統計サマリーをテキストで生成"""
        
        if data.empty:
            return "データが空です。"
            
        if columns is None:
            columns = data.select_dtypes(include=[np.number]).columns.tolist()
        
        summary_lines = []
        summary_lines.append("=== データ統計サマリー ===")
        summary_lines.append(f"データ期間: {data.index.min()} ～ {data.index.max()}")
        summary_lines.append(f"総データ数: {len(data):,} 行")
        summary_lines.append(f"変数数: {len(columns)} 個")
        summary_lines.append("")
        
        # 各変数の統計
        for col in columns:
            if col in data.columns:
                col_data = data[col].dropna()
                if not col_data.empty:
                    summary_lines.append(f"【{col}】")
                    summary_lines.append(f"  平均: {col_data.mean():.3f}")
                    summary_lines.append(f"  標準偏差: {col_data.std():.3f}")
                    summary_lines.append(f"  範囲: {col_data.min():.3f} ～ {col_data.max():.3f}")
                    summary_lines.append(f"  欠損率: {(data[col].isnull().sum() / len(data) * 100):.1f}%")
                    summary_lines.append("")
        
        return "\n".join(summary_lines)