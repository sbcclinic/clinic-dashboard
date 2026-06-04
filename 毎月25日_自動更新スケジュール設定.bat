@echo off
echo 毎月25日の自動更新スケジュールを設定します...

set TASK_NAME=クリニックダッシュボード月次更新
set BAT_PATH=%~dp0ダッシュボードを起動する.bat

schtasks /create /tn "%TASK_NAME%" /tr "%BAT_PATH%" /sc monthly /d 25 /st 09:00 /f

echo.
echo スケジュール設定が完了しました。
echo 毎月25日 午前9時に自動的にダッシュボードが起動します。
echo.
echo 設定を確認するには「タスクスケジューラ」アプリを開いてください。
pause
