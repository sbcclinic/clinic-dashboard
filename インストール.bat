@echo off
chcp 65001 > nul
echo.
echo === クリニックダッシュボード インストール ===
echo.

where py >nul 2>&1
if %errorlevel% == 0 (
    set PYCMD=py
    goto :install
)

where python >nul 2>&1
if %errorlevel% == 0 (
    set PYCMD=python
    goto :install
)

echo Pythonが見つかりません。
echo パソコンを再起動してから再度試してください。
pause
exit /b 1

:install
echo Pythonコマンド: %PYCMD%
echo.
echo インストール中... (数分かかります)
%PYCMD% -m pip install streamlit pandas openpyxl plotly
echo.
if %errorlevel% == 0 (
    echo ============================
    echo インストール完了しました！
    echo ============================
    echo.
    echo 次に「ダッシュボードを起動する.bat」を
    echo ダブルクリックしてください。
) else (
    echo エラーが発生しました。
    echo 上の画面をスクリーンショットで送ってください。
)
echo.
pause
