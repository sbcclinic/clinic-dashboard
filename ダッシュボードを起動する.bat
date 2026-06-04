@echo off
cd /d "%~dp0"
echo クリニックダッシュボードを起動します...
py -m streamlit run dashboard.py --server.port 8501 --browser.gatherUsageStats false
pause
