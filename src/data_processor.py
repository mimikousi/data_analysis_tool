import pandas as pd
import numpy as np
import streamlit as st
from typing import Tuple, List, Optional, Union
import io

class DataProcessor:
    def __init__(self):
        self.data = None
        self.original_data = None
        
    def load_file(self, uploaded_file) -> pd.DataFrame:
        """ファイルを読み込みDataFrameを返す"""
        try:
            file_extension = uploaded_file.name.split('.')[-1].lower()
            
            if file_extension == 'csv':
                # CSV読み込み（複数のエンコーディングを試行）
                encodings = ['utf-8', 'shift_jis', 'cp932', 'iso-2022-jp']
                for encoding in encodings:
                    try:
                        uploaded_file.seek(0)
                        data = pd.read_csv(uploaded_file, encoding=encoding, index_col=0, parse_dates=True)
                        break
                    except (UnicodeDecodeError, UnicodeError):
                        continue
                else:
                    st.error("CSVファイルの文字エンコーディングを識別できませんでした。")
                    return None
                    
            elif file_extension in ['xlsx', 'xlsm']:
                # Excel読み込み
                data = pd.read_excel(uploaded_file, index_col=0, parse_dates=True)
            else:
                st.error(f"サポートされていないファイル形式です: {file_extension}")
                return None
                
            # データ型の確認と変換
            data = self._process_datetime_index(data)
            data = self._process_numeric_columns(data)
            
            self.original_data = data.copy()
            self.data = data.copy()
            
            return data
            
        except Exception as e:
            st.error(f"ファイル読み込みエラー: {str(e)}")
            return None
    
    def _process_datetime_index(self, data: pd.DataFrame) -> pd.DataFrame:
        """インデックスを日時型に変換"""
        try:
            if not isinstance(data.index, pd.DatetimeIndex):
                data.index = pd.to_datetime(data.index)
        except:
            st.warning("インデックスを日時型に変換できませんでした。元の形式を保持します。")
        return data
    
    def _process_numeric_columns(self, data: pd.DataFrame) -> pd.DataFrame:
        """数値列の処理"""
        for col in data.columns:
            if data[col].dtype == 'object':
                try:
                    # 数値に変換を試行
                    data[col] = pd.to_numeric(data[col], errors='coerce')
                except:
                    pass
        return data
    
    def get_data_info(self) -> dict:
        """データの基本情報を取得"""
        if self.data is None:
            return {}
            
        info = {
            'shape': self.data.shape,
            'columns': list(self.data.columns),
            'dtypes': self.data.dtypes.to_dict(),
            'null_counts': self.data.isnull().sum().to_dict(),
            'memory_usage': self.data.memory_usage(deep=True).sum(),
            'date_range': None
        }
        
        # 日時範囲の取得
        if isinstance(self.data.index, pd.DatetimeIndex):
            info['date_range'] = {
                'start': self.data.index.min(),
                'end': self.data.index.max()
            }
            
        return info
    
    def get_numeric_columns(self) -> List[str]:
        """数値型のカラムリストを取得"""
        if self.data is None:
            return []
        return self.data.select_dtypes(include=[np.number]).columns.tolist()
    
    def get_column_statistics(self, column: str) -> dict:
        """指定カラムの統計情報を取得"""
        if self.data is None or column not in self.data.columns:
            return {}
            
        col_data = self.data[column].dropna()
        
        if col_data.empty:
            return {}
            
        stats = {
            'count': len(col_data),
            'mean': col_data.mean(),
            'std': col_data.std(),
            'min': col_data.min(),
            'max': col_data.max(),
            'median': col_data.median(),
            'q25': col_data.quantile(0.25),
            'q75': col_data.quantile(0.75),
            'skewness': col_data.skew(),
            'kurtosis': col_data.kurtosis()
        }
        
        return stats
    
    def filter_data_by_date_range(self, start_date, end_date) -> pd.DataFrame:
        """日付範囲でデータをフィルタリング"""
        if self.data is None:
            return pd.DataFrame()
            
        if isinstance(self.data.index, pd.DatetimeIndex):
            mask = (self.data.index >= start_date) & (self.data.index <= end_date)
            return self.data.loc[mask]
        else:
            return self.data
    
    def export_to_csv(self) -> bytes:
        """CSVとしてエクスポート"""
        if self.data is None:
            return b''
            
        output = io.StringIO()
        self.data.to_csv(output, encoding='utf-8')
        return output.getvalue().encode('utf-8')
    
    def export_to_excel(self) -> bytes:
        """Excelとしてエクスポート"""
        if self.data is None:
            return b''
            
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            self.data.to_excel(writer, sheet_name='Data')
        return output.getvalue()
    
    def update_data(self, new_data: pd.DataFrame):
        """データを更新"""
        self.data = new_data.copy()
    
    def reset_data(self):
        """オリジナルデータに戻す"""
        if self.original_data is not None:
            self.data = self.original_data.copy()
    
    def validate_data(self) -> Tuple[bool, List[str]]:
        """データの妥当性をチェック"""
        errors = []
        
        if self.data is None:
            errors.append("データが読み込まれていません")
            return False, errors
            
        if self.data.empty:
            errors.append("データが空です")
            
        if len(self.get_numeric_columns()) == 0:
            errors.append("数値列が見つかりません")
            
        # メモリ使用量チェック（100MBを超える場合は警告）
        memory_mb = self.data.memory_usage(deep=True).sum() / 1024 / 1024
        if memory_mb > 100:
            errors.append(f"データサイズが大きすぎます: {memory_mb:.1f}MB")
            
        return len(errors) == 0, errors