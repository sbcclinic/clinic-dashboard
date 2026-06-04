@echo off
chcp 65001 > nul
cd /d "C:\Users\宮城杏奈\Downloads\clinic_dashboard"

:: キャッシュフォルダを削除してExcelを必ず再読み込み
if exist ".streamlit\cache" rmdir /s /q ".streamlit\cache"
if exist "__pycache__" rmdir /s /q "__pycache__"

echo クリニックダッシュボードを起動します...
py -m streamlit run dashboard.py --server.port 8501 --browser.gatherUsageStats false
pause
