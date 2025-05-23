@echo off
setlocal enabledelayedexpansion

echo fam8 自動CSV取得・転記処理
echo ==============================

:: ▼ 昨日の日付を取得
for /f %%a in ('powershell -command "(Get-Date).AddDays(-1).ToString('yyyyMMdd')"') do set "target_date=%%a"

echo 使用日付: %target_date%
echo 処理を開始します...

rem --- 月替わりチェック ---
python month_checker.py %target_date%
if %ERRORLEVEL% neq 0 (
    echo 月替わりチェック処理でエラーが発生しました。
    set "HAS_ERROR=1"
    goto end
)

rem --- ブラウザ操作・CSV取得 ---
python browser_control.py %target_date%
if %ERRORLEVEL% neq 0 (
    echo CSV取得処理でエラーが発生しました。
    set "HAS_ERROR=1"
    goto end
)

rem --- Excel転記処理 ---
python excel_writer.py %target_date%
if %ERRORLEVEL% neq 0 (
    echo Excel転記処理でエラーが発生しました。
    set "HAS_ERROR=1"
    goto end
)

echo.
echo 全ての処理が完了しました！
set "HAS_ERROR=0"
goto end

:end
echo.
if "%HAS_ERROR%"=="1" (
    echo エラーにより処理を中断しました。
    echo ログを確認してください。
    echo.
    echo Enterキーを押すと終了します...
    pause >nul
)
exit /b 0
