@echo off
setlocal
title ���W���[���C���X�g�[��
color 0A

echo ===========================
echo Python�K�{���W���[���C���X�g�[��
echo ===========================
echo.

rem --- Python�p�X�̊m�F ---
if exist "%~dp0Python���s��\python.exe" (
    set PYTHON_PATH=%~dp0Python���s��\python.exe
    echo �G���x�b�_�u��Python(���K�w)���g�p���܂�
) else if exist "%~dp0..\Python���s��\python.exe" (
    set PYTHON_PATH=%~dp0..\Python���s��\python.exe
    echo �G���x�b�_�u��Python(�e�K�w)���g�p���܂�
) else (
    where python >nul 2>nul
    if %ERRORLEVEL% equ 0 (
        set PYTHON_PATH=python
        echo �V�X�e��Python���g�p���܂�
    ) else (
        echo �G���[: Python��������܂���B
        echo Python���s����z�u���邩�APython���C���X�g�[�����Ă��������B
        pause
        exit /b 1
    )
)

echo.
echo Python�p�X: %PYTHON_PATH%
echo.

echo loguru���W���[�����C���X�g�[�����Ă��܂�...
%PYTHON_PATH% -m pip install loguru
if %ERRORLEVEL% neq 0 (
    echo loguru�̃C���X�g�[���Ɏ��s���܂����B
    pause
    exit /b 1
)
echo loguru�̃C���X�g�[�����������܂����B

echo.
echo xlwings���W���[�����C���X�g�[�����Ă��܂�...
%PYTHON_PATH% -m pip install xlwings
if %ERRORLEVEL% neq 0 (
    echo xlwings�̃C���X�g�[���Ɏ��s���܂����B
    pause
    exit /b 1
)
echo xlwings�̃C���X�g�[�����������܂����B

echo.
echo ���ׂẴ��W���[���̃C���X�g�[�����������܂����I
echo.
pause
exit /b 0