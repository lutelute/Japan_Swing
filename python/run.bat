@echo off
REM Python版 Japan_Swing 実行スクリプト (Windows用)

REM 仮想環境をアクティベート
call venv\Scripts\activate.bat

REM シミュレーション実行
python simulate_area_network.py

REM 仮想環境をデアクティベート
call deactivate

pause