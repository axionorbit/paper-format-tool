@echo off
setlocal
chcp 65001 >nul

cd /d "%~dp0"

if not exist ".venv" (
    set "PY_BOOTSTRAP="
    where py >nul 2>&1
    if not errorlevel 1 set "PY_BOOTSTRAP=py -3.12"
    if "%PY_BOOTSTRAP%"=="" (
        where python >nul 2>&1
        if not errorlevel 1 set "PY_BOOTSTRAP=python"
    )
    if "%PY_BOOTSTRAP%"=="" (
        echo [错误] 未找到 Python 启动器（py 或 python）
        pause
        exit /b 1
    )

    echo [步骤] 创建虚拟环境 .venv ...
    %PY_BOOTSTRAP% -m venv .venv
)

if not exist ".venv\Scripts\python.exe" (
    echo [错误] 无法找到 .venv\Scripts\python.exe
    pause
    exit /b 1
)

echo [步骤] 升级 pip ...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :fail

echo [步骤] 安装依赖 requirements.txt ...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :fail

echo.
echo [完成] 环境已就绪。可双击 start.bat 启动程序。
pause
exit /b 0

:fail
echo.
echo [失败] 依赖安装失败，请检查网络或 pip 配置。
pause
exit /b 1
