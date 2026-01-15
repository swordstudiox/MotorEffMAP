@echo off
setlocal
cd /d "%~dp0"

echo 正在启动构建脚本...

:: 尝试激活虚拟环境 (如果存在)
if exist "venv\Scripts\activate.bat" (
    echo [INFO] 正在激活虚拟环境 venv...
    call venv\Scripts\activate.bat
)

:: 检查 Python 是否可用
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] 找不到 Python 环境。请确保已安装 Python 或激活了环境。
    pause
    exit /b
)

:: 运行 Python 构建脚本
python build_script.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] 构建过程中发生错误。
)

pause
