@echo off
setlocal enabledelayedexpansion

:: ===============================
:: 【1】日付を取得（前日）
:: ===============================
for /f %%I in ('powershell -command "((Get-Date).AddDays(-1)).ToString('yyyyMMdd')"') do set YYYYMMDD=%%I

:: ===============================
:: 【2】ログフォルダの確認・作成
:: ===============================
set LOG_DIR=C:\Users\rep03\Desktop\Pythonファイル\マイム_提案媒体_進捗Report\log
if not exist %LOG_DIR% mkdir %LOG_DIR%

:: ===============================
:: 【3】ログファイル名を設定（1日1つ、上書き方式）
:: ===============================
set LOG_FILE=%LOG_DIR%\マイム_提案媒体レポート実行_%YYYYMMDD%分.log

:: ===============================
:: 【4】使用する Python のパスを指定
:: ===============================
set PYTHON_EXEC=C:\Users\rep03\Desktop\Pythonファイル\Python実行環境\python.exe
set SCRIPT_DIR=C:\Users\rep03\Desktop\Pythonファイル\マイム_提案媒体_進捗Report\サイト取得
set BATCH_DIR=C:\Users\rep03\Desktop\Pythonファイル\マイム_提案媒体_進捗Report

:: ===============================
:: 【5】CSVファイルのパスを設定（フォルダも動的に変更）
:: ===============================
set CSV_DIR=C:\Users\rep03\Desktop\Pythonファイル\マイム_提案媒体_進捗Report\マイム日次CSV\%YYYYMMDD%
set MIME_CSV=%CSV_DIR%\マイムレポート_%YYYYMMDD%.csv
set FAM8_CSV=%CSV_DIR%\fam8レポート_%YYYYMMDD%.csv

:: ===============================
:: 【6】デバッグ情報を記録
:: ===============================
echo ==================================================== >> %LOG_FILE%
echo [%date% %time%] [INFO] マイム_提案媒体_日次処理 開始 >> %LOG_FILE%
echo ---------------------------------------------------- >> %LOG_FILE%
echo [%date% %time%] [INFO] 実行ディレクトリ: %CD% >> %LOG_FILE%
echo [%date% %time%] [INFO] 使用するPython: %PYTHON_EXEC% >> %LOG_FILE%
echo [%date% %time%] [INFO] CSV保存フォルダ: %CSV_DIR% >> %LOG_FILE%
echo ==================================================== >> %LOG_FILE%

:: ===============================
:: 【7】Pythonのパスを確認
:: ===============================
%PYTHON_EXEC% --version >> %LOG_FILE% 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [ERROR] Pythonが見つかりません。環境変数を確認してください。 >> %LOG_FILE%
    exit /b 1
)

:: ===============================
:: 【8】マイムのデータ取得
:: ===============================
echo [%date% %time%] [INFO] マイム日次CSV取得.py 実行開始 >> %LOG_FILE%
%PYTHON_EXEC% "%SCRIPT_DIR%\マイム日次CSV取得.py" >> %LOG_FILE% 2>&1
if %ERRORLEVEL% neq 0 (
    echo [%date% %time%] [ERROR] マイム CSV 取得失敗（リトライ1回目） >> %LOG_FILE%
    timeout /t 5 >nul
    %PYTHON_EXEC% "%SCRIPT_DIR%\マイム日次CSV取得.py" >> %LOG_FILE% 2>&1
    if %ERRORLEVEL% neq 0 (
        echo [%date% %time%] [ERROR] マイム CSV 取得失敗（リトライ2回目） >> %LOG_FILE%
        timeout /t 5 >nul
        %PYTHON_EXEC% "%SCRIPT_DIR%\マイム日次CSV取得.py" >> %LOG_FILE% 2>&1
        if %ERRORLEVEL% neq 0 (
            echo [%date% %time%] [CRITICAL] マイム CSV 取得に失敗、処理中止 >> %LOG_FILE%
            exit /b 1
        )
    )
)
echo [%date% %time%] [SUCCESS] マイム CSV 取得成功 >> %LOG_FILE%

:: ===============================
:: 【9】fam8のデータ取得
:: ===============================
echo [%date% %time%] [INFO] fam8日次CSV取得.py 実行開始 >> %LOG_FILE%
%PYTHON_EXEC% "%SCRIPT_DIR%\fam8日次CSV取得.py" >> %LOG_FILE% 2>&1
if %ERRORLEVEL% neq 0 (
    echo [%date% %time%] [ERROR] fam8 CSV 取得失敗（リトライ1回目） >> %LOG_FILE%
    timeout /t 5 >nul
    %PYTHON_EXEC% "%SCRIPT_DIR%\fam8日次CSV取得.py" >> %LOG_FILE% 2>&1
    if %ERRORLEVEL% neq 0 (
        echo [%date% %time%] [ERROR] fam8 CSV 取得失敗（リトライ2回目） >> %LOG_FILE%
        timeout /t 5 >nul
        %PYTHON_EXEC% "%SCRIPT_DIR%\fam8日次CSV取得.py" >> %LOG_FILE% 2>&1
        if %ERRORLEVEL% neq 0 (
            echo [%date% %time%] [CRITICAL] fam8 CSV 取得に失敗、処理中止 >> %LOG_FILE%
            exit /b 1
        )
    )
)
echo [%date% %time%] [SUCCESS] fam8 CSV 取得成功 >> %LOG_FILE%

:: ===============================
:: 【10】CSVの確認
:: ===============================
echo [%date% %time%] [INFO] マイムとfam8(前日分)のCSVファイル確認開始 >> %LOG_FILE%
if not exist "%MIME_CSV%" (
    echo [%date% %time%] [ERROR] %MIME_CSV% が見つかりません、処理中止 >> %LOG_FILE%
    exit /b 1
)
if not exist "%FAM8_CSV%" (
    echo [%date% %time%] [ERROR] %FAM8_CSV% が見つかりません、処理中止 >> %LOG_FILE%
    exit /b 1
)
echo [%date% %time%] [SUCCESS] マイムとfam8(前日分)のCSV確認完了 >> %LOG_FILE%

:CSV_UPLOAD
:: ===============================
:: 【11】CSVアップロード処理
:: ===============================
echo ==================================================== >> %LOG_FILE%
echo [%date% %time%] [INFO] 抽出レポートへCSVアップロード処理開始 >> %LOG_FILE%
echo ---------------------------------------------------- >> %LOG_FILE%
call %PYTHON_EXEC% "%BATCH_DIR%\CSVを抽出レポートへアップロード(と関数挿入).py" >> %LOG_FILE% 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [ERROR] 抽出レポートへCSVアップロード処理でエラー発生 >> %LOG_FILE%
    exit /b 1
)
echo [%date% %time%] [SUCCESS] 抽出レポートへCSVアップロード完了 >> %LOG_FILE%

:: ===============================
:: 【12】スプレッドシート関数適用 & 値変換
:: ===============================
echo ==================================================== >> %LOG_FILE%
echo [%date% %time%] [INFO] スプレッドシート行に関数挿入と値のみ変換.py 実行開始 >> %LOG_FILE%
echo ---------------------------------------------------- >> %LOG_FILE%
call %PYTHON_EXEC% "%BATCH_DIR%\スプレッドシート行に関数挿入と値のみ変換.py" >> %LOG_FILE% 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] [ERROR] 日時レポートのスプレッドシート処理でエラー発生 >> %LOG_FILE%
    exit /b 1
)
echo [%date% %time%] [SUCCESS] 日時レポートのスプレッドシート処理完了 >> %LOG_FILE%
echo ==================================================== >> %LOG_FILE%

exit /b 0
