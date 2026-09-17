@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"

echo ========================================
echo ERP 月度实发数量与退货率同步
echo 配置文件：%~dp0config.json
echo ========================================

if not exist "%~dp0erp_excel_sync.py" (
    echo.
    echo [失败] 找不到 erp_excel_sync.py
    pause
    exit /b 1
)

if not exist "%~dp0config.json" (
    echo.
    echo [失败] 找不到 config.json
    pause
    exit /b 1
)

set "PYTHON_EXE="
set "PYTHON_ARGS="

if exist "%~dp0runtime\python.exe" (
    "%~dp0runtime\python.exe" -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul
    if not errorlevel 1 set "PYTHON_EXE=%~dp0runtime\python.exe"
)

if not defined PYTHON_EXE (
    python -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul
    if not errorlevel 1 set "PYTHON_EXE=python"
)

if not defined PYTHON_EXE (
    py -3 -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul
    if not errorlevel 1 (
        set "PYTHON_EXE=py"
        set "PYTHON_ARGS=-3"
    )
)

if not defined PYTHON_EXE (
    echo.
    echo [失败] 未找到 Python 3.10 或更高版本，请先安装 Python。
    pause
    exit /b 1
)

echo 使用 Python：
"%PYTHON_EXE%" %PYTHON_ARGS% --version
echo.
"%PYTHON_EXE%" %PYTHON_ARGS% "%~dp0erp_excel_sync.py" --config "%~dp0config.json"
set "SYNC_EXIT_CODE=%ERRORLEVEL%"

if "%SYNC_EXIT_CODE%"=="0" (
    echo.
    echo [成功] 月度同步已完成，原工作簿已保留为 .bak 备份。
) else if "%SYNC_EXIT_CODE%"=="2" (
    echo.
    echo [未替换] 关键货号校验未通过，请查看 reports 目录。
) else (
    echo.
    echo [失败] 同步程序退出，错误码：%SYNC_EXIT_CODE%
    echo 如果浏览器显示未登录，请在 ERP Chrome 窗口完成登录后重试。
)

pause
exit /b %SYNC_EXIT_CODE%
