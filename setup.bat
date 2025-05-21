@echo off
setlocal
title モジュールインストール
color 0A

echo ===========================
echo Python必須モジュールインストール
echo ===========================
echo.

rem --- Pythonパスの確認 ---
if exist "%~dp0Python実行環境\python.exe" (
    set PYTHON_PATH=%~dp0Python実行環境\python.exe
    echo エンベッダブルPython(同階層)を使用します
) else if exist "%~dp0..\Python実行環境\python.exe" (
    set PYTHON_PATH=%~dp0..\Python実行環境\python.exe
    echo エンベッダブルPython(親階層)を使用します
) else (
    where python >nul 2>nul
    if %ERRORLEVEL% equ 0 (
        set PYTHON_PATH=python
        echo システムPythonを使用します
    ) else (
        echo エラー: Pythonが見つかりません。
        echo Python実行環境を配置するか、Pythonをインストールしてください。
        pause
        exit /b 1
    )
)

echo.
echo Pythonパス: %PYTHON_PATH%
echo.

echo loguruモジュールをインストールしています...
%PYTHON_PATH% -m pip install loguru
if %ERRORLEVEL% neq 0 (
    echo loguruのインストールに失敗しました。
    pause
    exit /b 1
)
echo loguruのインストールが完了しました。

echo.
echo xlwingsモジュールをインストールしています...
%PYTHON_PATH% -m pip install xlwings
if %ERRORLEVEL% neq 0 (
    echo xlwingsのインストールに失敗しました。
    pause
    exit /b 1
)
echo xlwingsのインストールが完了しました。

echo.
echo すべてのモジュールのインストールが完了しました！
echo.
pause
exit /b 0