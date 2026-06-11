import os
import subprocess
import sys
import shutil
import configparser
import re
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
VENV_PYTHON = PROJECT_ROOT / "venv" / "Scripts" / "python.exe"


EXCLUDED_MODULES = [
    # 其他 GUI 框架 / 交互环境
    "tkinter",
    "PyQt5",
    "PyQt6",
    "PySide2",
    "IPython",
    "jupyter",
    "notebook",
    "pytest",
    "sympy",
    "test",

    # 当前项目不使用的深度学习 / 推理库；全局 Python 环境中存在时容易被 PyInstaller hooks 误收集
    "torch",
    "torchvision",
    "torchaudio",
    "onnx",
    "onnxruntime",
    "tensorflow",
    "keras",
    "sklearn",

    # PySide6 中当前程序未使用的模块
    "PySide6.Qt3DAnimation",
    "PySide6.Qt3DCore",
    "PySide6.Qt3DExtras",
    "PySide6.Qt3DInput",
    "PySide6.Qt3DLogic",
    "PySide6.Qt3DRender",
    "PySide6.QtBluetooth",
    "PySide6.QtCharts",
    "PySide6.QtConcurrent",
    "PySide6.QtDataVisualization",
    "PySide6.QtDesigner",
    "PySide6.QtHelp",
    "PySide6.QtLocation",
    "PySide6.QtMultimedia",
    "PySide6.QtMultimediaWidgets",
    "PySide6.QtNetworkAuth",
    "PySide6.QtOpenGL",
    "PySide6.QtOpenGLWidgets",
    "PySide6.QtPdf",
    "PySide6.QtPdfWidgets",
    "PySide6.QtPositioning",
    "PySide6.QtPrintSupport",
    "PySide6.QtQml",
    "PySide6.QtQuick",
    "PySide6.QtQuick3D",
    "PySide6.QtQuickControls2",
    "PySide6.QtQuickWidgets",
    "PySide6.QtRemoteObjects",
    "PySide6.QtScxml",
    "PySide6.QtSensors",
    "PySide6.QtSerialPort",
    "PySide6.QtSql",
    "PySide6.QtStateMachine",
    "PySide6.QtSvg",
    "PySide6.QtSvgWidgets",
    "PySide6.QtTest",
    "PySide6.QtTextToSpeech",
    "PySide6.QtUiTools",
    "PySide6.QtWebChannel",
    "PySide6.QtWebEngineCore",
    "PySide6.QtWebEngineQuick",
    "PySide6.QtWebEngineWidgets",
    "PySide6.QtWebSockets",
    "PySide6.QtXml",

    # Pillow 可选格式和 Tk 绑定；matplotlib 保存 PNG 不需要这些
    "PIL._avif",
    "PIL._webp",
    "PIL._imagingtk",
    "PIL.ImageTk",

]


PRUNE_AFTER_BUILD = [
    "_internal/PySide6/Qt6Pdf.dll",
    "_internal/PySide6/Qt6Qml.dll",
    "_internal/PySide6/Qt6QmlMeta.dll",
    "_internal/PySide6/Qt6QmlModels.dll",
    "_internal/PySide6/Qt6QmlWorkerScript.dll",
    "_internal/PySide6/Qt6Quick.dll",
    "_internal/PySide6/Qt6VirtualKeyboard.dll",
    "_internal/PySide6/translations",
]


def get_exclude_args():
    args = []
    for module in EXCLUDED_MODULES:
        args.extend(["--exclude-module", module])
    return args


def get_dir_size(path):
    total = 0
    root = Path(path)
    if not root.exists():
        return 0
    for file_path in root.rglob("*"):
        if file_path.is_file():
            try:
                total += file_path.stat().st_size
            except OSError:
                pass
    return total


def format_size(size):
    return f"{size / 1024 / 1024:.2f} MB"


def read_version_label():
    version_path = PROJECT_ROOT / "version.ini"
    if not version_path.exists():
        return ""

    parser = configparser.ConfigParser()
    try:
        parser.read(version_path, encoding="utf-8")
    except Exception as e:
        print(f"[WARN] 读取 version.ini 失败: {e}")
        return ""

    version = parser["version"] if parser.has_section("version") else parser.defaults()
    build_date = version.get("build_date", "").strip()
    code = version.get("code", "").strip()
    return "-".join(part for part in [build_date, code] if part)


def sanitize_dist_name(name):
    text = re.sub(r"[^\w.-]+", "_", name, flags=re.UNICODE)
    return text.strip("._ ") or "MotorEffMAP"


def get_dist_dir_name():
    version_label = read_version_label()
    if not version_label:
        return "MotorEffMAP"
    return sanitize_dist_name(f"MotorEffMAP_{version_label}")


def print_build_size_report(dist_dir):
    dist_path = Path(dist_dir)
    if not dist_path.exists():
        return

    print("\n[INFO] 构建体积统计:")
    print(f"  总体积: {format_size(get_dir_size(dist_path))}")

    entries = []
    for child in dist_path.iterdir():
        size = get_dir_size(child) if child.is_dir() else child.stat().st_size
        entries.append((size, child.name))

    for size, name in sorted(entries, reverse=True)[:10]:
        print(f"  {format_size(size):>10}  {name}")

    largest_files = []
    for file_path in dist_path.rglob("*"):
        if file_path.is_file():
            try:
                largest_files.append((file_path.stat().st_size, file_path))
            except OSError:
                pass

    print("\n[INFO] 最大文件 Top 15:")
    for size, file_path in sorted(largest_files, reverse=True)[:15]:
        rel_path = file_path.relative_to(dist_path)
        print(f"  {format_size(size):>10}  {rel_path}")


def prune_build_output(dist_dir):
    dist_path = Path(dist_dir)
    removed_size = 0

    print("\n[INFO] 清理未使用的打包文件:")
    for rel_path in PRUNE_AFTER_BUILD:
        target = dist_path / Path(rel_path)
        if not target.exists():
            continue

        size = get_dir_size(target) if target.is_dir() else target.stat().st_size
        try:
            if target.is_dir():
                shutil.rmtree(target)
            else:
                target.unlink()
            removed_size += size
            print(f"  已移除 {format_size(size):>10}  {rel_path}")
        except OSError as e:
            print(f"  [WARN] 无法移除 {rel_path}: {e}")

    if removed_size == 0:
        print("  没有可清理项。")
    else:
        print(f"  合计减少: {format_size(removed_size)}")


def validate_build_output(dist_dir):
    dist_path = Path(dist_dir)
    required_files = [
        dist_path / "MotorEffMAP.exe",
        dist_path / "MotorEffMAP.ini",
        dist_path / "version.ini",
    ]
    missing_files = [path.name for path in required_files if not path.exists()]
    if missing_files:
        raise RuntimeError(f"构建产物缺少必要文件: {', '.join(missing_files)}")

    exe_path = dist_path / "MotorEffMAP.exe"
    if exe_path.stat().st_size <= 0:
        raise RuntimeError("构建产物 MotorEffMAP.exe 大小异常。")

    print("[INFO] 构建产物完整性检查通过。")


def install_package(package):
    print(f"[INFO] Installing {package}...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])


def ensure_project_venv():
    if not VENV_PYTHON.exists():
        print("[ERROR] 未找到项目虚拟环境: venv\\Scripts\\python.exe")
        print("请先在项目根目录创建并安装依赖:")
        print("  python -m venv venv")
        print("  venv\\Scripts\\python.exe -m pip install -r requirements.txt")
        print("  venv\\Scripts\\python.exe -m pip install pyinstaller")
        sys.exit(1)

    current_python = Path(sys.executable).resolve()
    target_python = VENV_PYTHON.resolve()
    if current_python != target_python:
        print(f"[INFO] 当前 Python: {current_python}")
        print(f"[INFO] 切换到项目虚拟环境: {target_python}")
        code = subprocess.call([str(target_python), str(Path(__file__).resolve()), *sys.argv[1:]])
        sys.exit(code)


def main():
    ensure_project_venv()

    print("===================================================")
    print("      正在构建 MotorEffMAP 可执行文件 (Python脚本)")
    print("===================================================")

    # 1. Check and Install Dependencies
    print("[1/5] 检查依赖环境...")
    try:
        import PyInstaller
    except ImportError:
        install_package("pyinstaller")

    # 2. Icon Processing
    print("[2/5] 正在准备图标...")
    icon_param = []
    icon_file = "MotorEffMAP.ico"

    if os.path.exists(icon_file):
        icon_param = [f"--icon={icon_file}"]
        print(f"使用现有图标: {icon_file}")
    else:
        print(f"[WARN] 未找到 {icon_file}，将使用默认图标。")

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
    ] + icon_param + get_exclude_args() + ["run.py"]

    try:
        subprocess.check_call(cmd)
    except subprocess.CalledProcessError as e:
        print(f"\n[ERROR] 编译失败，错误代码: {e.returncode}")
        sys.exit(1)

    # 5. Copy Resources
    print("[5/5] 复制配置文件...")
    default_dist_dir = Path("dist") / "MotorEffMAP"
    dist_dir = Path("dist") / get_dist_dir_name()
    if default_dist_dir != dist_dir:
        if dist_dir.exists():
            shutil.rmtree(dist_dir)
        if default_dist_dir.exists():
            default_dist_dir.rename(dist_dir)
            print(f"发布目录已命名为: {dist_dir}")

    if not os.path.exists(dist_dir):
        os.makedirs(dist_dir)

    files_to_copy = ["MotorEffMAP.ini", "version.ini"]
    for file_name in files_to_copy:
        if os.path.exists(file_name):
            try:
                shutil.copy(file_name, dist_dir)
                print(f"已复制: {file_name}")
            except Exception as e:
                print(f"[WARN] 无法复制 {file_name}: {e}")

    # Copy icon for runtime window use
    if os.path.exists(icon_file):
        try:
            shutil.copy(icon_file, dist_dir)
            print(f"已复制图标文件: {icon_file}")
        except:
             pass

    prune_build_output(dist_dir)
    try:
        validate_build_output(dist_dir)
    except RuntimeError as e:
        print(f"[ERROR] {e}")
        sys.exit(1)
    print_build_size_report(dist_dir)

    print("\n===================================================")
    print("                 构建完成!")
    print("===================================================")
    print(f"可执行文件夹位于: {os.path.abspath(dist_dir)}")
    print("请进入该文件夹运行 MotorEffMAP.exe")
    print("===================================================")

if __name__ == "__main__":
    main()
