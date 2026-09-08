@echo off
cd /d %~dp0

echo =========================================
echo Naver Cafe Auto Posting - Building navercafeauto2.exe
echo =========================================
echo Checking PyInstaller...
pyinstaller --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo PyInstaller is not installed. Installing...
    pip install pyinstaller
)

echo Building executable from spec file...
pyinstaller --noconfirm --clean navercafeauto2.spec

echo.
if exist "dist\navercafeauto2.exe" (
    echo =========================================
    echo Build Successful!
    echo Executable is located at: dist\navercafeauto2.exe
    echo =========================================
) else (
    echo =========================================
    echo Build Failed. Check the logs above.
    echo =========================================
)
pause
