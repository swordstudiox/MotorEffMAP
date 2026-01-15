import sys
import os
from PySide6.QtGui import QIcon
from MotorEffMAP_GUI import MainWindow, QApplication, QFont
import matplotlib

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # 设置应用程序图标
    # 在 frozen (exe) 模式下，路径可能是 executable 目录
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        
    icon_path = os.path.join(base_dir, "MotorEffMAP.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    # Matplotlib 和 QT 全局字体修复
    try:
        matplotlib.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'Arial']
        matplotlib.rcParams['axes.unicode_minus'] = False
    except:
        pass
    
    font = QFont("Microsoft YaHei")
    font.setPointSize(9)
    app.setFont(font)
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
