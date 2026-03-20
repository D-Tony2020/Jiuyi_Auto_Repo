@echo off
echo ============================================
echo    Jiuyi Report Updater - Build Script
echo ============================================
echo.

echo [1/3] Installing dependencies...
pip install openpyxl pypdf pyinstaller
if errorlevel 1 (
    echo.
    echo Failed! Please make sure Python 3.8+ is installed.
    pause
    exit /b 1
)

echo.
echo [2/3] Building .exe ...
pyinstaller --onefile --windowed ^
    --name "ReportUpdater" ^
    --distpath "D:\jiuyi_build\dist" ^
    --workpath "D:\jiuyi_build\work" ^
    --specpath "D:\jiuyi_build" ^
    --clean ^
    main.py
if errorlevel 1 (
    echo.
    echo Build failed!
    pause
    exit /b 1
)

echo.
echo [3/3] Done!
echo .exe is at: D:\jiuyi_build\dist\ReportUpdater.exe
echo.
echo Please copy it to wherever you need.
echo.
pause
