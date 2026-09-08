@echo off
cd /d C:\antigravity\navercafe_auto

echo =========================================
echo Naver Cafe Auto Posting - Building EXE
echo =========================================echo Checking PyInstaller...
pyinstaller --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo PyInstaller is not installed. Installing...
    pip install pyinstaller
)

echo Building executable from spec file...
pyinstaller --noconfirm --clean NaverCafeAuto.spec

echo.
if exist "dist\NaverCafeAuto.exe" (
    echo =========================================
    echo Build Successful!
    echo Executable is located at: dist\NaverCafeAuto.exe
    echo =========================================
) else (
    echo =========================================
    echo Build Failed. Check the logs above.
    echo =========================================
)
pause
