@echo off
chcp 65001 > nul

set TASK_NAME=クリニックDB_月次自動更新
set BAT_PATH=C:\Users\宮城杏奈\Downloads\clinic_dashboard\start.bat

echo タスクスケジューラーに登録します...
echo.

schtasks /create ^
  /tn "%TASK_NAME%" ^
  /tr "%BAT_PATH%" ^
  /sc monthly ^
  /d 25 ^
  /st 09:00 ^
  /f

echo.
if %errorlevel% == 0 (
    echo 設定完了しました！
    echo 毎月25日 午前9時に自動でダッシュボードが起動します。
) else (
    echo エラーが発生しました。
    echo 管理者権限で実行し直してください。
)
echo.
pause
