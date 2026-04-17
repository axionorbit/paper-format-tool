@echo off
setlocal
chcp 65001 >nul

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [错误] 未找到虚拟环境解释器: .venv\Scripts\python.exe
    echo.
    echo 请先执行 setup_venv.bat 初始化环境，然后重试。
    pause
    exit /b 1
)

echo [启动] 使用虚拟环境运行论文格式助手...
".venv\Scripts\python.exe" "main.py"
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo.
    echo [异常] 程序退出码: %EXIT_CODE%
    pause
)

exit /b %EXIT_CODE%

