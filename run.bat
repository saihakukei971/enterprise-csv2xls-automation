@echo off
setlocal enabledelayedexpansion

echo fam8 ����CSV�擾�E�]�L����
echo ==============================
echo �擾������t����͂��Ă�������
echo [��1] �P��: 20250512
echo [��2] �͈�: 20250512-20250515
echo [����] �����͂܂���1���o�߂ō���̓��t�������K�p����܂�
echo.

set /p target_date="���t�����: " <nul
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
echo �������J�n���܂�...

rem --- ���ւ��`�F�b�N ---
python month_checker.py %target_date%
if %ERRORLEVEL% neq 0 (
    echo ���ւ��`�F�b�N�����ŃG���[���������܂����B
    set "HAS_ERROR=1"
    goto end
)

rem --- �u���E�U����ECSV�擾 ---
python browser_control.py %target_date%
if %ERRORLEVEL% neq 0 (
    echo CSV�擾�����ŃG���[���������܂����B
    set "HAS_ERROR=1"
    goto end
)

rem --- Excel�]�L���� ---
python excel_writer.py %target_date%
if %ERRORLEVEL% neq 0 (
    echo Excel�]�L�����ŃG���[���������܂����B
    set "HAS_ERROR=1"
    goto end
)

echo.
echo �S�Ă̏������������܂����I
set "HAS_ERROR=0"
goto end

:end
echo.
if "%HAS_ERROR%"=="1" (
    echo �G���[�ɂ�菈���𒆒f���܂����B
    echo ���O���m�F���Ă��������B
    echo.
    echo Enter�L�[�������ƏI�����܂�...
    pause >nul
)
exit /b 0
