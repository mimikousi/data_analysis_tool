# Process Data Analytics App - 仕様書

## 概要
化学業界のプロセスデータ・品質データを対象とした時系列データ解析アプリケーション。
Streamlitを使用してWebアプリとして構築し、Streamlit Shareで公開する。

## 目的
- 時系列データの確認・比較
- 統計解析
- 異常検知
- データクリーニング（外れ値除去）

## 技術仕様

### 開発環境
- **フレームワーク**: Streamlit
- **言語**: Python 3.9+
- **主要ライブラリ**: 
  - pandas（データ処理）
  - numpy（数値計算）
  - plotly（インタラクティブグラフ）
  - scipy（統計計算）
  - openpyxl（Excelファイル読み込み）
  - reportlab（PDFレポート生成）
  - streamlit-aggrid（高性能データテーブル）

### データ仕様
- **形式**: CSV (.csv), Excel (.xlsx, .xlsm)
- **サイズ**: 1-2万行 × 50列
- **構造**: 
  - インデックス: datetime型（日時分データ）
  - カラム: 変数名（数値データ）
- **更新頻度**: 1時間または5分間隔（将来対応）

## 機能仕様

### 1. データ読み込み機能
- ファイルアップロード（drag & drop対応）
- 対応形式: .csv, .xlsx, .xlsm
- データプレビュー表示
- カラム情報の表示（データ型、欠損値数など）

### 2. 外れ値除去機能
**目的**: データクリーニングによる解析精度向上

**機能詳細**:
- カラム選択（単一選択）
- トレンドグラフ表示（plotlyインタラクティブグラフ）
- 外れ値除去方法:
  - 数値入力による範囲指定（メイン方式）
    - X範囲（期間）: 開始日時 〜 終了日時（カレンダー形式）
    - Y範囲（値）: 下限値 〜 上限値
  - ドラッグ範囲選択（実装可能な場合の補助機能）
- 除去対象データのハイライト表示
- 除去実行後の新データフレーム生成
- **履歴管理機能**:
  - 除去操作1回ごとの履歴保存（上限なし）
  - 履歴リスト表示（操作日時、対象カラム、除去件数）
  - クリック操作による簡単復元機能
  - セッション終了時の履歴保持
  - **履歴リセット条件**:
    - Webアプリケーション再起動時
    - 新しいデータファイル読み込み時
  - 履歴データの自動キャッシュ
- 完了ボタンによるグラフ終了
- 複数カラムでの逐次実行対応

### 3. トレンド折れ線グラフ機能
**目的**: 多変数時系列データの可視化・統計解析

**機能詳細**:
- 複数カラム選択（チェックボックス）
- 軸設定:
  - Y1軸・Y2軸の選択
  - 軸範囲の手動設定（上限・下限・間隔）
  - デフォルト：自動スケール
- 統計線の表示:
  - 全期間統計:
    - 平均値線
    - 平均値±σ線（σ=1〜5選択可能）
  - 期間指定統計:
    - 期間選択（カレンダー形式：開始〜終了日時）
    - 選択期間の網掛け表示
    - 期間内平均値線
    - 期間内平均値±σ線（σ=1〜5選択可能）
- グラフエクスポート機能（PNG, SVG）

### 4. 散布図マトリックス機能
**目的**: 変数間相関関係の可視化・定量評価

**機能詳細**:
- 選択カラムの全組み合わせ散布図作成
- マトリックスレイアウト（n×nグリッド）
- グラフサイズ・配置の調整機能
- 各散布図に以下を表示:
  - 線形回帰直線
  - 回帰式（y = ax + b）
  - 相関係数（r）
  - 決定係数（R²）
- インタラクティブ機能（ズーム・パン）

### 5. ヒストグラム機能
**目的**: データ分布の確認・正規性検定

**機能詳細**:
- 選択カラム全てのヒストグラム作成
- 各ヒストグラムに以下を表示:
  - 実データ分布（バー）
  - 正規分布理論曲線（同一平均・標準偏差）
  - 基本統計情報（平均・標準偏差・歪度・尖度）
- ビン数の調整機能
- 正規性検定結果の表示（Shapiro-Wilk検定）

### 6. データエクスポート機能
**目的**: 解析結果の保存・共有

**機能詳細**:
- **データファイルエクスポート**:
  - クリーニング後データフレームのダウンロード
  - 対応形式: CSV, Excel (.xlsx)
  - ファイル名の自動生成（タイムスタンプ付き）
- **グラフ画像エクスポート**:
  - 対応形式: PNG, SVG, PDF
  - 高解像度オプション
  - カスタムサイズ指定
- **解析結果レポート自動生成**:
  - PDF形式の統合レポート
  - 含有内容：
    - データ概要（行数、列数、期間）
    - 外れ値除去履歴
    - 基本統計情報（平均、標準偏差、相関係数等）
    - 全グラフの自動埋め込み
    - 正規性検定結果
  - レポートテンプレートのカスタマイズ対応

### レイアウト構成
1. **サイドバー**: 
   - ファイルアップロード
   - 機能選択メニュー
   - パラメータ設定
2. **メインエリア**: 
   - グラフ表示
   - データテーブル
   - 結果出力
3. **フッター**: 
   - エクスポート機能
   - システム情報

### 操作フロー
1. データファイルアップロード
2. データプレビュー確認
3. 外れ値除去（任意）
   - 履歴管理による復元操作
   - 完了ボタンで次ステップへ
4. 解析機能の同時実行・表示
   - トレンド折れ線グラフ
   - 散布図マトリックス  
   - ヒストグラム
5. 結果確認・調整（スクロール表示）
6. エクスポート・レポート生成

## 技術的考慮事項

### パフォーマンス
- **データフレームキャッシュ機能（必須）**:
  - Streamlit session_stateによる状態管理
  - 外れ値除去履歴の自動保存（セッション内保持）
  - グラフ描画結果のキャッシュ
  - **履歴管理ポリシー**:
    - セッション継続中は全履歴保持
    - アプリ再起動時・新データ読み込み時にリセット
- **高速処理**:
  - 大容量データ（2万行×50列）の最適化
  - pandas操作の効率化
  - plotlyグラフの軽量化
- **インタラクティブ性重視**:
  - リアルタイムパラメータ更新
  - 即座のグラフ反映
  - スムーズなスクロール体験

### エラーハンドリング
- ファイル形式チェック
- データ型検証
- 欠損値処理

### セキュリティ
- ファイルサイズ制限
- アップロードファイル検証

## デプロイメント

### Streamlit Share設定
- requirements.txt作成
- GitHub連携設定
- 環境変数管理

### ファイル構成
```
project/
├── streamlit_app.py      # メインアプリケーション
├── requirements.txt      # 依存関係
├── README.md            # プロジェクト説明
├── CLAUDE.md            # 本仕様書
├── src/
│   ├── data_processor.py   # データ処理機能
│   ├── visualization.py   # グラフ作成機能
│   ├── statistics.py      # 統計計算機能
│   ├── outlier_removal.py # 外れ値除去・履歴管理
│   └── report_generator.py # レポート自動生成
└── tests/
    └── test_functions.py   # テストコード
```

## 今後の拡張予定
- リアルタイムデータ更新対応（API連携）
- バッチ処理機能（複数ファイル一括解析）
- カスタムレポートテンプレート作成
- データベース連携機能
- ユーザー設定の保存・復元

---
*作成日: 2025年6月29日*  
*バージョン: 1.1*  
*更新内容: 履歴管理機能、エクスポート機能、レポート自動生成機能の追加*