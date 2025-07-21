# Python版 Japan_Swing

日本の10エリア（北海道〜沖縄）の電力系統における連成スイング（同期安定性）をシミュレーションするプログラムのPython版です。

## 概要

MATLAB版を完全にPythonに移植し、同じ機能を提供します：
- 10エリア連成解析
- 可変発電機台数対応
- 擾乱シミュレーション
- リアルタイム可視化
- Excelパラメータ管理

## 環境要件

- Python 3.8以上
- 必要なライブラリ（requirements.txtを参照）

## インストール

```bash
cd python
pip install -r requirements.txt
```

## 使用方法

### 1. パラメータテンプレート生成（初回のみ）
```bash
python generate_area_template.py
```

### 2. シミュレーション実行
```bash
python simulate_area_network.py
```

### 3. 実行時の設定
- コンソールで可視化対象エリアを選択
- 擾乱を投入するエリアと発電機番号を指定
- 擾乱量（Δδ [rad]）を設定

## ファイル構成

```
python/
├── simulate_area_network.py       # メインシミュレーションスクリプト
├── generate_area_template.py      # Excelテンプレート生成スクリプト
├── requirements.txt               # Python依存関係
├── README_python.md              # このファイル
└── area_parameters_template.xlsx  # パラメータ設定ファイル（自動生成）
```

## MATLAB版との違い

- **GUI**: MATLABのダイアログの代わりにコンソール入力を使用
- **可視化**: matplotlibのアニメーション機能を使用
- **依存関係**: MATLABツールボックスの代わりにPythonライブラリを使用

## 機能詳細

### SwingSimulatorクラス
- エリア選択、擾乱設定、パラメータ管理を統合
- オブジェクト指向設計で保守性を向上

### 主要メソッド
- `run_simulation()`: メインシミュレーション実行
- `visualize_network()`: ネットワーク可視化
- `plot_coi_timeseries()`: COI時系列プロット
- `dynamics()`: 連成スイング方程式

### 可視化機能
- 上部: 地理マップ上のCOIベクトル表示
- 下部: 1D発電機角度プロット
- リアルタイムアニメーション

## 注意事項

- GUI環境でのmatplotlib表示が必要
- アニメーション速度は計算能力に依存
- 大量の発電機（総数>200）では処理が重くなる場合があります