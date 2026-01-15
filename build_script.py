import os
import subprocess
import sys
import shutil

def install_package(package):
    print(f"[INFO] Installing {package}...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

def main():
    print("===================================================")
    print("      正在构建 MotorEffMAP 可执行文件 (Python脚本)")
    print("===================================================")

    # 1. Check and Install Dependencies
    print("[1/5] 检查依赖环境...")
    try:
        import PyInstaller
    except ImportError:
        install_package("pyinstaller")

    try:
        from PIL import Image
    except ImportError:
        install_package("pillow")
        from PIL import Image

    # 2. Icon Processing
    print("[2/5] 正在准备图标...")
    icon_param = []
    temp_icon = "MotorEffMAP.ico"
    
    if os.path.exists("图标.png"):
        print("发现 图标.png，正在转换为 .ico 格式...")
        try:
            img = Image.open("图标.png")
            # Clear previous if exists
            if os.path.exists(temp_icon):
                os.remove(temp_icon)
            
            # Save with multiple sizes for better Windows compatibility
            img.save(temp_icon, format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
            icon_param = [f"--icon={temp_icon}"]
            print("图标转换成功。")
        except Exception as e:
            print(f"[WARN] 图标转换失败: {e}")
            print("将使用默认图标。")
    else:
        print("未找到 图标.png，将使用默认图标。")

    # 3. Clean previous build
    print("[3/5] 清理旧的构建文件...")
    for folder in ["build", "dist"]:
        if os.path.exists(folder):
            try:
                shutil.rmtree(folder)
            except Exception as e:
                print(f"[WARN] 无法清理 {folder}: {e}")
                
    if os.path.exists("MotorEffMAP.spec"):
        os.remove("MotorEffMAP.spec")

    # 4. Run PyInstaller
    print("[4/5] 正在编译 exe (可能需要几分钟)...")
    # Use python -m PyInstaller to ensure we use the installed module in current env
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--onedir",
        "--windowed",
        "--name", "MotorEffMAP",
        "--clean"
    ] + icon_param + ["run.py"]
    
    try:
        subprocess.check_call(cmd)
    except subprocess.CalledProcessError as e:
        print(f"\n[ERROR] 编译失败，错误代码: {e.returncode}")
        sys.exit(1)

    # 5. Copy Resources
    print("[5/5] 复制配置文件...")
    dist_dir = os.path.join("dist", "MotorEffMAP")
    if not os.path.exists(dist_dir):
        os.makedirs(dist_dir)
    
    files_to_copy = ["MotorEffMAP.ini"]
    for file_name in files_to_copy:
        if os.path.exists(file_name):
            try:
                shutil.copy(file_name, dist_dir)
                print(f"已复制: {file_name}")
            except Exception as e:
                print(f"[WARN] 无法复制 {file_name}: {e}")
    
    # Copy icon for runtime window use
    if os.path.exists(temp_icon):
        try:
            shutil.copy(temp_icon, dist_dir)
            print(f"已复制图标文件: {temp_icon}")
        except:
             pass

    # Clean temp icon
    if os.path.exists(temp_icon):
        try:
            os.remove(temp_icon)
        except:
            pass

    print("\n===================================================")
    print("                 构建完成!")
    print("===================================================")
    print(f"可执行文件夹位于: {os.path.abspath(dist_dir)}")
    print("请进入该文件夹运行 MotorEffMAP.exe")
    print("===================================================")

if __name__ == "__main__":
    main()
