@echo off
setlocal enabledelayedexpansion

echo fam8 自動CSV取得・転記処理
echo ==============================
echo 取得する日付を入力してください
echo [例1] 単日: 20250512
echo [例2] 範囲: 20250512-20250515
echo [注意] 未入力または1分経過で昨日の日付が自動適用されます
echo.

set /p target_date="日付を入力: " <nul
set /a counter=0
:loop
set /p target_date=
if defined target_date (
    goto run
)
timeout /t 1 >nul
set /a counter+=1
if !counter! GEQ 60 (
    set target_date=default
    goto run
)
goto loop

:run
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
