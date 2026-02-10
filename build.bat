@echo off
chcp 65001 >nul 2>&1
echo ============================================
echo   consolidate_specs 打包工具
echo ============================================
echo.

REM 取得腳本所在目錄
set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

REM Step 1: 建立虛擬環境
echo [1/4] 建立 Python 虛擬環境...
if exist venv (
    echo       虛擬環境已存在，跳過建立。
) else (
    python -m venv venv
    if errorlevel 1 (
        echo [錯誤] 建立虛擬環境失敗，請確認已安裝 Python 3.8 以上版本。
        pause
        exit /b 1
    )
)

REM Step 2: 啟用虛擬環境並安裝套件
echo [2/4] 安裝相依套件...
call venv\Scripts\activate.bat
pip install -r requirements.txt
if errorlevel 1 (
    echo [錯誤] 安裝套件失敗。
    pause
    exit /b 1
)

REM Step 3: 執行 PyInstaller 打包
echo [3/4] 執行 PyInstaller 打包...
pyinstaller --onefile --console --name consolidate_specs ^
    --hidden-import openpyxl ^
    --hidden-import xlrd ^
    consolidate_specs.py
if errorlevel 1 (
    echo [錯誤] PyInstaller 打包失敗。
    pause
    exit /b 1
)

REM Step 4: 複製 .exe 到專案根目錄
echo [4/4] 複製執行檔到專案根目錄...
copy /y "dist\consolidate_specs.exe" "consolidate_specs.exe"

echo.
echo ============================================
echo   打包完成！
echo   執行檔位置: %SCRIPT_DIR%consolidate_specs.exe
echo ============================================
echo.
pause
