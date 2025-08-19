@echo off
echo ========================================
echo    IMAGE CRAWLER - CAO ANH TU DONG
echo ========================================
echo.
echo Dang khoi dong ung dung...
echo.

REM Kiem tra Python
python --version >nul 2>&1
if errorlevel 1 (
    echo Loi: Python khong duoc cai dat hoac khong co trong PATH
    echo Vui long cai dat Python 3.7+ va thu lai
    pause
    exit /b 1
)

REM Kiem tra requirements
if not exist requirements.txt (
    echo Loi: Khong tim thay file requirements.txt
    pause
    exit /b 1
)

echo Cai dat dependencies...
pip install -r requirements.txt

if errorlevel 1 (
    echo Loi: Khong the cai dat dependencies
    pause
    exit /b 1
)

echo.
echo Khoi dong ung dung...
python main.py

pause
