@echo off
setlocal enabledelayedexpansion

:: ===============================
:: �y1�z���t���擾�i�O���j
:: ===============================
for /f %%I in ('powershell -command "((Get-Date).AddDays(-1)).ToString('yyyyMMdd')"') do set YYYYMMDD=%%I

:: ===============================
:: �y2�z���O�t�H���_�̊m�F�E�쐬
:: ===============================
set LOG_DIR=C:\Users\rep03\Desktop\Python�t�@�C��\�}�C��_��Ĕ}��_�i��Report\log
if not exist %LOG_DIR% mkdir %LOG_DIR%

:: ===============================
:: �y3�z���O�t�@�C������ݒ�i1��1�A�㏑�������j
:: ===============================
set LOG_FILE=%LOG_DIR%\�}�C��_��Ĕ}�̃��|�[�g���s_%YYYYMMDD%��.log

:: ===============================
:: �y4�z�g�p���� Python �̃p�X���w��
:: ===============================
set PYTHON_EXEC=C:\Users\rep03\Desktop\Python�t�@�C��\Python���s��\python.exe
set SCRIPT_DIR=C:\Users\rep03\Desktop\Python�t�@�C��\�}�C��_��Ĕ}��_�i��Report\�T�C�g�擾
set BATCH_DIR=C:\Users\rep03\Desktop\Python�t�@�C��\�}�C��_��Ĕ}��_�i��Report

:: ===============================
:: �y5�zCSV�t�@�C���̃p�X��ݒ�i�t�H���_�����I�ɕύX�j
:: ===============================
set CSV_DIR=C:\Users\rep03\Desktop\Python�t�@�C��\�}�C��_��Ĕ}��_�i��Report\�}�C������CSV\%YYYYMMDD%
set MIME_CSV=%CSV_DIR%\�}�C�����|�[�g_%YYYYMMDD%.csv
set FAM8_CSV=%CSV_DIR%\fam8���|�[�g_%YYYYMMDD%.csv

:: ===============================
:: �y6�z�f�o�b�O�����L�^
:: ===============================
echo ==================================================== >> %LOG_FILE%
echo [%date% %time%] [INFO] �}�C��_��Ĕ}��_�������� �J�n >> %LOG_FILE%
echo ---------------------------------------------------- >> %LOG_FILE%
echo [%date% %time%] [INFO] ���s�f�B���N�g��: %CD% >> %LOG_FILE%
echo [%date% %time%] [INFO] �g�p����Python: %PYTHON_EXEC% >> %LOG_FILE%
echo [%date% %time%] [INFO] CSV�ۑ��t�H���_: %CSV_DIR% >> %LOG_FILE%
echo ==================================================== >> %LOG_FILE%

:: ===============================
:: �y7�zPython�̃p�X���m�F
:: ===============================
%PYTHON_EXEC% --version >> %LOG_FILE% 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [ERROR] Python��������܂���B���ϐ����m�F���Ă��������B >> %LOG_FILE%
    exit /b 1
)

:: ===============================
:: �y8�z�}�C���̃f�[�^�擾
:: ===============================
echo [%date% %time%] [INFO] �}�C������CSV�擾.py ���s�J�n >> %LOG_FILE%
%PYTHON_EXEC% "%SCRIPT_DIR%\�}�C������CSV�擾.py" >> %LOG_FILE% 2>&1
if %ERRORLEVEL% neq 0 (
    echo [%date% %time%] [ERROR] �}�C�� CSV �擾���s�i���g���C1��ځj >> %LOG_FILE%
    timeout /t 5 >nul
    %PYTHON_EXEC% "%SCRIPT_DIR%\�}�C������CSV�擾.py" >> %LOG_FILE% 2>&1
    if %ERRORLEVEL% neq 0 (
        echo [%date% %time%] [ERROR] �}�C�� CSV �擾���s�i���g���C2��ځj >> %LOG_FILE%
        timeout /t 5 >nul
        %PYTHON_EXEC% "%SCRIPT_DIR%\�}�C������CSV�擾.py" >> %LOG_FILE% 2>&1
        if %ERRORLEVEL% neq 0 (
            echo [%date% %time%] [CRITICAL] �}�C�� CSV �擾�Ɏ��s�A�������~ >> %LOG_FILE%
            exit /b 1
        )
    )
)
echo [%date% %time%] [SUCCESS] �}�C�� CSV �擾���� >> %LOG_FILE%

:: ===============================
:: �y9�zfam8�̃f�[�^�擾
:: ===============================
echo [%date% %time%] [INFO] fam8����CSV�擾.py ���s�J�n >> %LOG_FILE%
%PYTHON_EXEC% "%SCRIPT_DIR%\fam8����CSV�擾.py" >> %LOG_FILE% 2>&1
if %ERRORLEVEL% neq 0 (
    echo [%date% %time%] [ERROR] fam8 CSV �擾���s�i���g���C1��ځj >> %LOG_FILE%
    timeout /t 5 >nul
    %PYTHON_EXEC% "%SCRIPT_DIR%\fam8����CSV�擾.py" >> %LOG_FILE% 2>&1
    if %ERRORLEVEL% neq 0 (
        echo [%date% %time%] [ERROR] fam8 CSV �擾���s�i���g���C2��ځj >> %LOG_FILE%
        timeout /t 5 >nul
        %PYTHON_EXEC% "%SCRIPT_DIR%\fam8����CSV�擾.py" >> %LOG_FILE% 2>&1
        if %ERRORLEVEL% neq 0 (
            echo [%date% %time%] [CRITICAL] fam8 CSV �擾�Ɏ��s�A�������~ >> %LOG_FILE%
            exit /b 1
        )
    )
)
echo [%date% %time%] [SUCCESS] fam8 CSV �擾���� >> %LOG_FILE%

:: ===============================
:: �y10�zCSV�̊m�F
:: ===============================
echo [%date% %time%] [INFO] �}�C����fam8(�O����)��CSV�t�@�C���m�F�J�n >> %LOG_FILE%
if not exist "%MIME_CSV%" (
    echo [%date% %time%] [ERROR] %MIME_CSV% ��������܂���A�������~ >> %LOG_FILE%
    exit /b 1
)
if not exist "%FAM8_CSV%" (
    echo [%date% %time%] [ERROR] %FAM8_CSV% ��������܂���A�������~ >> %LOG_FILE%
    exit /b 1
)
echo [%date% %time%] [SUCCESS] �}�C����fam8(�O����)��CSV�m�F���� >> %LOG_FILE%

:CSV_UPLOAD
:: ===============================
:: �y11�zCSV�A�b�v���[�h����
:: ===============================
echo ==================================================== >> %LOG_FILE%
echo [%date% %time%] [INFO] ���o���|�[�g��CSV�A�b�v���[�h�����J�n >> %LOG_FILE%
echo ---------------------------------------------------- >> %LOG_FILE%
call %PYTHON_EXEC% "%BATCH_DIR%\CSV�𒊏o���|�[�g�փA�b�v���[�h(�Ɗ֐��}��).py" >> %LOG_FILE% 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [ERROR] ���o���|�[�g��CSV�A�b�v���[�h�����ŃG���[���� >> %LOG_FILE%
    exit /b 1
)
echo [%date% %time%] [SUCCESS] ���o���|�[�g��CSV�A�b�v���[�h���� >> %LOG_FILE%

:: ===============================
:: �y12�z�X�v���b�h�V�[�g�֐��K�p & �l�ϊ�
:: ===============================
echo ==================================================== >> %LOG_FILE%
echo [%date% %time%] [INFO] �X�v���b�h�V�[�g�s�Ɋ֐��}���ƒl�̂ݕϊ�.py ���s�J�n >> %LOG_FILE%
echo ---------------------------------------------------- >> %LOG_FILE%
call %PYTHON_EXEC% "%BATCH_DIR%\�X�v���b�h�V�[�g�s�Ɋ֐��}���ƒl�̂ݕϊ�.py" >> %LOG_FILE% 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [ERROR] �������|�[�g�̃X�v���b�h�V�[�g�����ŃG���[���� >> %LOG_FILE%
    exit /b 1
)
echo [%date% %time%] [SUCCESS] �������|�[�g�̃X�v���b�h�V�[�g�������� >> %LOG_FILE%
echo ==================================================== >> %LOG_FILE%

exit /b 0
