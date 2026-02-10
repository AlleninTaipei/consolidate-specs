@echo off
chcp 65001 >nul 2>&1
echo ============================================
echo   consolidate_specs 規格彙整工具
echo ============================================
echo.

REM 取得腳本所在目錄
set "SCRIPT_DIR=%~dp0"
set "EXE_PATH=%SCRIPT_DIR%consolidate_specs.exe"

REM 確認 .exe 存在
if not exist "%EXE_PATH%" (
    echo [錯誤] 找不到 consolidate_specs.exe
    echo        請確認 consolidate_specs.exe 與 run.bat 放在同一個資料夾。
    echo.
    pause
    exit /b 1
)

REM 判斷是否有拖放的資料夾
if "%~1"=="" (
    echo 請輸入 spec_files 資料夾路徑（或將資料夾拖放到本檔案上）：
    echo.
    set /p "FOLDER_PATH=路徑: "
) else (
    set "FOLDER_PATH=%~1"
)

REM 移除路徑尾端的引號和反斜線
set "FOLDER_PATH=%FOLDER_PATH:"=%"

REM 確認資料夾存在
if not exist "%FOLDER_PATH%\" (
    echo.
    echo [錯誤] 找不到資料夾: %FOLDER_PATH%
    echo        請確認路徑是否正確。
    echo.
    pause
    exit /b 1
)

echo 處理資料夾: %FOLDER_PATH%
echo.

"%EXE_PATH%" "%FOLDER_PATH%"

echo.
echo ============================================
echo   執行完畢
echo ============================================
pause
