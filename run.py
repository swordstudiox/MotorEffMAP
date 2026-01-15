import sys
from MotorEffMAP_GUI import MainWindow, QApplication, QFont
import matplotlib

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
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
