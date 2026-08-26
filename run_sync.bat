@echo off
setlocal
cd /d "%~dp0"

set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

echo ========================================
echo ERP monthly actual and return-rate sync
echo Config: %~dp0config.json
echo ========================================

where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python was not found. Install Python 3.10 or newer.
    pause
    exit /b 1
)

python "%~dp0erp_excel_sync.py" --config "%~dp0config.json"
set "SYNC_EXIT_CODE=%ERRORLEVEL%"

if "%SYNC_EXIT_CODE%"=="0" (
    echo.
    echo [OK] Monthly sync completed. The previous workbook is kept as .bak.
) else (
    echo.
    echo [ERROR] Sync failed. Exit code: %SYNC_EXIT_CODE%
)

pause
exit /b %SYNC_EXIT_CODE%
