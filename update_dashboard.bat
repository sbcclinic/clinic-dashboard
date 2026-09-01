@echo off
chcp 65001 > nul
echo ========================================
echo  ダッシュボード月次更新
echo ========================================
echo.

cd /d "C:\Users\宮城杏奈\Box\総合企画部_特殊案件\その他\clinic_dashboard"

echo [1/3] HTMLを生成中...
"C:\Users\宮城杏奈\AppData\Local\Python\bin\python.exe" generate_html.py
if errorlevel 1 (
    echo エラー: HTML生成に失敗しました。
    pause
    exit /b 1
)

echo.
echo [2/3] GitHubへアップロード中...
git add index.html
git commit -m "月次更新: %date%"
git push origin main
if errorlevel 1 (
    echo エラー: GitHubへのアップロードに失敗しました。
    pause
    exit /b 1
)

echo.
echo ========================================
echo  完了しました！
echo  ダッシュボードURL:
echo  https://sbcclinic.github.io/clinic-dashboard/
echo ========================================
pause
