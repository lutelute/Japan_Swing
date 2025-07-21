#!/bin/bash
# Python版 Japan_Swing 実行スクリプト

# 仮想環境をアクティベート
source venv/bin/activate

# シミュレーション実行
python simulate_area_network.py

# 仮想環境をデアクティベート
deactivate