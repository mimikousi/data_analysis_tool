import pandas as pd
import numpy as np
import streamlit as st
from datetime import datetime
from typing import Dict, List, Tuple, Optional
import copy

class OutlierRemovalHistory:
    def __init__(self):
        self.history = []
        self.current_data = None
        
    def add_operation(self, operation_data: Dict):
        """外れ値除去操作を履歴に追加"""
        operation_data['timestamp'] = datetime.now()
        operation_data['operation_id'] = len(self.history)
        self.history.append(operation_data)
    
    def get_history(self) -> List[Dict]:
        """履歴リストを取得"""
        return self.history
    
    def clear_history(self):
        """履歴をクリア"""
        self.history = []
    
    def remove_last_operation(self):
        """最後の操作を削除"""
        if self.history:
            self.history.pop()

class OutlierRemover:
    def __init__(self):
        self.history_manager = OutlierRemovalHistory()
        self.data_snapshots = {}  # operation_id -> DataFrame
        
    def initialize_data(self, data: pd.DataFrame):
        """初期データを設定"""
        self.history_manager.current_data = data.copy()
        self.data_snapshots[-1] = data.copy()  # 初期状態を保存
        self.history_manager.clear_history()
    
    def remove_outliers_by_range(self, 
                                data: pd.DataFrame,
                                column: str,
                                x_range: Tuple[datetime, datetime] = None,
                                y_range: Tuple[float, float] = None) -> pd.DataFrame:
        """範囲指定による外れ値除去"""
        
        if data is None or data.empty:
            return data
            
        # フィルタリング用のマスクを作成
        mask = pd.Series([True] * len(data), index=data.index)
        removed_indices = []
        
        # X軸（時間）範囲のフィルタリング
        if x_range is not None and isinstance(data.index, pd.DatetimeIndex):
            time_mask = (data.index >= x_range[0]) & (data.index <= x_range[1])
            mask = mask & time_mask
            
        # Y軸（値）範囲のフィルタリング
        if y_range is not None and column in data.columns:
            value_mask = (data[column] >= y_range[0]) & (data[column] <= y_range[1])
            mask = mask & value_mask
            
        # 除去対象のインデックスを記録
        removed_indices = data.index[mask].tolist()
        
        # データから除去
        filtered_data = data.drop(index=removed_indices)
        
        # 履歴に記録
        operation_data = {
            'column': column,
            'x_range': x_range,
            'y_range': y_range,
            'removed_count': len(removed_indices),
            'removed_indices': removed_indices,
            'data_shape_before': data.shape,
            'data_shape_after': filtered_data.shape
        }
        
        # 操作前のデータを保存
        operation_id = len(self.history_manager.history)
        self.data_snapshots[operation_id] = data.copy()
        
        # 履歴に追加
        self.history_manager.add_operation(operation_data)
        
        return filtered_data
    
    def remove_outliers_by_statistical_method(self,
                                             data: pd.DataFrame,
                                             column: str,
                                             method: str = 'iqr',
                                             multiplier: float = 1.5) -> pd.DataFrame:
        """統計的手法による外れ値除去"""
        
        if data is None or data.empty or column not in data.columns:
            return data
            
        col_data = data[column].dropna()
        
        if method == 'iqr':
            # IQR法
            Q1 = col_data.quantile(0.25)
            Q3 = col_data.quantile(0.75)
            IQR = Q3 - Q1
            lower_bound = Q1 - multiplier * IQR
            upper_bound = Q3 + multiplier * IQR
            
        elif method == 'zscore':
            # Z-score法
            mean = col_data.mean()
            std = col_data.std()
            lower_bound = mean - multiplier * std
            upper_bound = mean + multiplier * std
            
        else:
            st.error(f"サポートされていない統計手法: {method}")
            return data
            
        # 外れ値のマスクを作成
        outlier_mask = (data[column] < lower_bound) | (data[column] > upper_bound)
        removed_indices = data.index[outlier_mask].tolist()
        
        # データから除去
        filtered_data = data[~outlier_mask]
        
        # 履歴に記録
        operation_data = {
            'column': column,
            'method': method,
            'multiplier': multiplier,
            'lower_bound': lower_bound,
            'upper_bound': upper_bound,
            'removed_count': len(removed_indices),
            'removed_indices': removed_indices,
            'data_shape_before': data.shape,
            'data_shape_after': filtered_data.shape
        }
        
        # 操作前のデータを保存
        operation_id = len(self.history_manager.history)
        self.data_snapshots[operation_id] = data.copy()
        
        # 履歴に追加
        self.history_manager.add_operation(operation_data)
        
        return filtered_data
    
    def get_outlier_candidates(self, 
                              data: pd.DataFrame, 
                              column: str,
                              method: str = 'iqr',
                              multiplier: float = 1.5) -> pd.Series:
        """外れ値候補を取得（除去前の確認用）"""
        
        if data is None or data.empty or column not in data.columns:
            return pd.Series(dtype=bool)
            
        col_data = data[column].dropna()
        
        if method == 'iqr':
            Q1 = col_data.quantile(0.25)
            Q3 = col_data.quantile(0.75)
            IQR = Q3 - Q1
            lower_bound = Q1 - multiplier * IQR
            upper_bound = Q3 + multiplier * IQR
            
        elif method == 'zscore':
            mean = col_data.mean()
            std = col_data.std()
            lower_bound = mean - multiplier * std
            upper_bound = mean + multiplier * std
            
        else:
            return pd.Series(dtype=bool)
            
        return (data[column] < lower_bound) | (data[column] > upper_bound)
    
    def restore_to_operation(self, operation_id: int) -> pd.DataFrame:
        """指定した操作時点のデータに復元"""
        
        if operation_id in self.data_snapshots:
            # 指定した操作以降の履歴を削除
            self.history_manager.history = self.history_manager.history[:operation_id]
            
            # 対応するデータスナップショットも削除
            keys_to_remove = [k for k in self.data_snapshots.keys() if k > operation_id]
            for key in keys_to_remove:
                del self.data_snapshots[key]
                
            return self.data_snapshots[operation_id].copy()
        else:
            st.error(f"操作ID {operation_id} のデータが見つかりません")
            return pd.DataFrame()
    
    def get_history_summary(self) -> pd.DataFrame:
        """履歴の概要をDataFrameで取得"""
        
        if not self.history_manager.history:
            return pd.DataFrame()
            
        summary_data = []
        for i, op in enumerate(self.history_manager.history):
            summary_data.append({
                '操作ID': i,
                '操作日時': op['timestamp'].strftime('%Y-%m-%d %H:%M:%S'),
                '対象カラム': op['column'],
                '除去方法': self._get_method_description(op),
                '除去件数': op['removed_count'],
                '処理前行数': op['data_shape_before'][0],
                '処理後行数': op['data_shape_after'][0]
            })
            
        return pd.DataFrame(summary_data)
    
    def _get_method_description(self, operation: Dict) -> str:
        """操作の説明テキストを生成"""
        
        if 'method' in operation:
            # 統計的手法
            method_names = {
                'iqr': 'IQR法',
                'zscore': 'Z-score法'
            }
            method_name = method_names.get(operation['method'], operation['method'])
            return f"{method_name} (係数: {operation.get('multiplier', 'N/A')})"
            
        elif 'x_range' in operation or 'y_range' in operation:
            # 範囲指定
            desc_parts = []
            if operation.get('x_range'):
                start = operation['x_range'][0].strftime('%Y-%m-%d %H:%M')
                end = operation['x_range'][1].strftime('%Y-%m-%d %H:%M')
                desc_parts.append(f"時間範囲: {start}〜{end}")
            if operation.get('y_range'):
                desc_parts.append(f"値範囲: {operation['y_range'][0]:.2f}〜{operation['y_range'][1]:.2f}")
            return "範囲指定 (" + ", ".join(desc_parts) + ")"
            
        else:
            return "不明な方法"
    
    def export_history(self) -> Dict:
        """履歴をエクスポート"""
        return {
            'history': self.history_manager.history,
            'total_operations': len(self.history_manager.history)
        }
    
    def clear_all_history(self):
        """全履歴をクリア"""
        self.history_manager.clear_history()
        self.data_snapshots = {}
    
    def get_removed_data_summary(self, operation_id: int) -> pd.DataFrame:
        """指定した操作で除去されたデータの概要を取得"""
        
        if operation_id >= len(self.history_manager.history):
            return pd.DataFrame()
            
        operation = self.history_manager.history[operation_id]
        removed_indices = operation.get('removed_indices', [])
        
        if not removed_indices or operation_id not in self.data_snapshots:
            return pd.DataFrame()
            
        original_data = self.data_snapshots[operation_id]
        removed_data = original_data.loc[removed_indices]
        
        return removed_data