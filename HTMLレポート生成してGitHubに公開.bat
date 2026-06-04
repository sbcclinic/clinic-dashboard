@echo off
chcp 65001 > nul
cd /d "C:\Users\宮城杏奈\Downloads\clinic_dashboard"

echo HTMLレポートを生成中...
py generate_html.py

if %errorlevel% neq 0 (
    echo エラーが発生しました。
    pause
    exit /b 1
)

echo.
echo GitHubにアップロード中...
git add index.html
git commit -m "Update clinic dashboard report"
git push origin main

echo.
echo 完了しました！
echo 以下のURLで確認できます：
echo https://sbcmiyagi.github.io/clinic-dashboard/
echo.
pause
