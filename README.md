# プロセスデータ解析アプリ

化学業界のプロセスデータ・品質データを対象とした時系列データ解析Webアプリケーション

## 🚀 機能

- **データ読み込み**: CSV、Excel形式のファイル対応
- **外れ値除去**: 範囲指定・統計的手法による除去、履歴管理
- **トレンド分析**: 多変数時系列グラフ、統計線表示
- **相関分析**: 散布図マトリックス、相関ヒートマップ
- **分布分析**: ヒストグラム、正規性検定
- **レポート出力**: PDF形式の自動レポート生成

## 📋 使用方法

### ローカル実行
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

### Streamlit Share
1. GitHubリポジトリをStreamlit Shareに連携
2. `streamlit_app.py`を指定してデプロイ

## 📊 サンプルデータ形式

CSVまたはExcelファイル（日時インデックス + 数値データ）
```
datetime,temperature,pressure,flow_rate
2023-01-01 00:00:00,25.1,100.2,50.3
2023-01-01 01:00:00,24.8,99.8,51.1
...
```

## 🔧 技術仕様

- **フレームワーク**: Streamlit
- **言語**: Python 3.9+
- **主要ライブラリ**: pandas, plotly, scipy, reportlab

## 📄 ライセンス

MIT License