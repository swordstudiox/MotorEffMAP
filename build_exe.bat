@echo off
chcp 65001
echo ===================================================
echo      正在构建 MotorEffMAP 可执行文件...
echo ===================================================

:: 1. 检查并安装依赖 (PyInstaller 和 Pillow)
echo [1/5] 检查依赖环境...
pip show pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo PyInstaller 未安装，正在安装...
    pip install pyinstaller
) else (
    echo PyInstaller 已安装。
)

pip show pillow >nul 2>&1
if %errorlevel% neq 0 (
    echo Pillow (用于图标转换) 未安装，正在安装...
    pip install pillow
) else (
    echo Pillow 已安装。
)

:: 2. 处理图标 (PNG -> ICO)
echo [2/5] 正在准备图标...
if exist "图标.png" goto FoundIcon
echo 未找到 图标.png，将使用默认图标。
set "ICON_PARAM="
goto IconDone

:FoundIcon
echo 发现 图标.png，正在转换为 .ico 格式...
:: 注意：这里使用 python -c 时避免与批处理括号冲突
python -c "from PIL import Image; img = Image.open('图标.png'); img.save('MotorEffMAP.ico', format='ICO', sizes=[(256, 256)])"
set "ICON_PARAM=--icon=MotorEffMAP.ico"

:IconDone

:: 3. 清理旧的构建文件
echo [3/5] 清理旧的构建文件...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del *.spec

:: 4. 运行 PyInstaller
echo [4/5] 正在编译 exe (可能需要几分钟)...
:: --onedir:以此文件夹形式输出
:: --windowed: 不显示黑色控制台窗口
:: --name: 可执行文件名称
:: --clean: 清理缓存
:: %ICON_PARAM%: 图标参数
pyinstaller --noconfirm --onedir --windowed --name "MotorEffMAP" --clean %ICON_PARAM% run.py

:: 5. 复制配置文件和其他资源
echo [5/5] 复制配置文件...
copy /Y "MotorEffMAP.ini" "dist\MotorEffMAP\" >nul
copy /Y "目标.txt" "dist\MotorEffMAP\" >nul

if exist "MotorEffMAP.ico" (
    del "MotorEffMAP.ico"
)

echo ===================================================
echo                 构建完成!
echo ===================================================
echo 可执行文件夹位于: dist\MotorEffMAP
echo 请进入该文件夹运行 MotorEffMAP.exe
echo ===================================================
pause
