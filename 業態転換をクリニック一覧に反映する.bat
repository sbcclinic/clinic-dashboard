@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo 「業態転換・名称変更履歴」の内容をクリニック一覧に反映します...
echo.
py sync_conversions.py

if %errorlevel% neq 0 (
    echo エラーが発生しました。
    pause
    exit /b 1
)

echo.
echo 内容を確認したら、いつも通りダッシュボードを更新してください。
pause
