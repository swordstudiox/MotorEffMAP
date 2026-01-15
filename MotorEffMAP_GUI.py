import sys
import os
import datetime
import logging
import configparser
import subprocess
import platform
import pandas as pd
import numpy as np

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                               QHBoxLayout, QPushButton, QLabel, QFileDialog, 
                               QTabWidget, QFormLayout, QLineEdit, QTextEdit,
                               QScrollArea, QMessageBox, QGroupBox, QSplitter, QListWidget,
                               QProgressBar)
from PySide6.QtCore import Qt, QTimer, QSize
from PySide6.QtGui import QAction, QIcon, QFont

import matplotlib
matplotlib.use('QtAgg')
from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg, NavigationToolbar2QT
from matplotlib.figure import Figure
import matplotlib.pyplot as plt

from MotorEffMAP_Logic import MotorEffLogic

# Setup Logging
class QTextEditLogger(logging.Handler):
    def __init__(self, widget):
        super().__init__()
        self.widget = widget

    def emit(self, record):
        msg = self.format(record)
        self.widget.append(msg)
        if record.levelno >= logging.ERROR:
            # Trigger alert in GUI thread? 
            # Ideally use signals, but direct append is safe enough for simple cases usually.
            # To be safe regarding threads, we might just append. 
            pass

class AspectRatioWidget(QWidget):
    """
    一个包含 FigureCanvas 并保持固定长宽比（居中）的小部件，
    不管容器大小如何变化。
    """
    def __init__(self, figure, aspect_ratio=1.25, parent=None):
        super().__init__(parent)
        self.aspect_ratio = aspect_ratio
        
        # 实例化画布
        self.canvas = FigureCanvasQTAgg(figure)
        self.canvas.setParent(self)
        
        # 重要：允许画布被父级逻辑自由调整大小
        from PySide6.QtWidgets import QSizePolicy
        self.canvas.setSizePolicy(QSizePolicy.Ignored, QSizePolicy.Ignored)

    def resizeEvent(self, event):
        w = event.size().width()
        h = event.size().height()
        
        if w <= 0 or h <= 0: return

        # 计算目标尺寸
        if w / h > self.aspect_ratio:
            # 容器太宽；以高度为准
            target_h = h
            target_w = int(h * self.aspect_ratio)
        else:
            # 容器太高；以宽度为准
            target_w = w
            target_h = int(w / self.aspect_ratio)
            
        # 居中画布
        x = (w - target_w) // 2
        y = (h - target_h) // 2
        
        self.canvas.setGeometry(x, y, target_w, target_h)
        # super().resizeEvent(event) # 无需调用 super 的基本 QWidget resize


class SignatureLabel(QLabel):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setText("Author: Unknown | Date: " + datetime.datetime.now().strftime("%Y-%m-%d"))
        self.setStyleSheet("color: gray; font-size: 10px;")
        self.setVisible(False)
        self.original_text = ""
        
    def set_signature(self, author, date):
        self.original_text = f"Author: {author} | Date: {date}"
        self.setText(self.original_text)
        
    def enterEvent(self, event):
        self.setVisible(True)
        super().enterEvent(event)

    def leaveEvent(self, event):
        self.setVisible(False)
        super().leaveEvent(event)

# 悬停区域包装器，用于检测该区域的鼠标悬停，而不仅仅是标签（因为标签仅在可见时才工作）
# 实际上用户要求："显示在右下角... 仅当鼠标经过时"。
# 通常隐藏的标签无法接收鼠标事件。
# 所以我们创建一个始终可见的容器（透明），并在悬停时显示其中的文本。
class SignatureWidget(QWidget):
    def __init__(self, author="RunDa", date="2026-01-15"):
        super().__init__()
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(0,0,5,0)
        self.layout.setAlignment(Qt.AlignRight | Qt.AlignBottom)
        
        self.label = QLabel(f"Author: {author} | Date: {date}")
        self.label.setStyleSheet("color: #555; font-size: 11px; background: rgba(255,255,255,0.8); border-radius: 3px; padding: 2px;")
        self.label.setVisible(False)
        self.layout.addWidget(self.label)
        
        self.setMouseTracking(True)
        
    def enterEvent(self, event):
        self.label.setVisible(True)
        super().enterEvent(event)
        
    def leaveEvent(self, event):
        self.label.setVisible(False)
        super().leaveEvent(event)

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("电驱效率MAP绘制工具 (Python 版)")
        self.resize(1200, 800)
        
        # Determine base path for config file (Robust for PyInstaller Exe)
        if getattr(sys, 'frozen', False):
            # If run as exe
            base_dir = os.path.dirname(sys.executable)
        else:
            # If run as script
            base_dir = os.path.dirname(os.path.abspath(__file__))
            
        self.ini_path = os.path.join(base_dir, "MotorEffMAP.ini")
        
        if not os.path.exists(self.ini_path):
            # 如果不存在创建默认？或者警告。
            pass
            
        self.config_dict = {}
        self.raw_config_obj = None # 用于跟踪 sections
        
        self.logic = None
        self.data_files = []
        
        self.init_ui()
        self.setup_logging()
        self.reload_config()

    def setup_logging(self):
        # 文件处理器
        file_handler = logging.FileHandler("MotorEffMAP.log", mode='a', encoding='utf-8')
        file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        
        # GUI 处理器 (控制台)
        self.gui_log_handler = QTextEditLogger(self.log_text)
        self.gui_log_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))
        
        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        root_logger.addHandler(file_handler)
        root_logger.addHandler(self.gui_log_handler)
        
        logging.info("应用程序已启动。")

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        
        main_layout = QVBoxLayout(central)
        
        # 选项卡
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # --- 选项卡 1: 处理 ---
        self.tab_run = QWidget()
        self.init_run_tab()
        self.tabs.addTab(self.tab_run, "处理与分析")
        
        # --- 选项卡 2: 设置 ---
        self.tab_config = QWidget()
        self.init_config_tab()
        self.tabs.addTab(self.tab_config, "配置")
        
        # 日志区域
        log_group = QGroupBox("日志")
        # 固定 GroupBox 高度以紧密贴合内容 (文本 80 + 条 10 + 边距/标题 ~40 = ~130)
        # 稍微增加以防止重叠
        log_group.setFixedHeight(130) 
        
        log_layout = QVBoxLayout(log_group)
        log_layout.setSpacing(5) # 增加间距以分隔文本和进度条
        log_layout.setContentsMargins(5, 15, 5, 5) 
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFixedHeight(80) 
        log_layout.addWidget(self.log_text)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFixedHeight(10)   # 10px 高度
        # 小条的样式/字体调整
        self.progress_bar.setStyleSheet("QProgressBar { border: 1px solid #aaa; border-radius: 0px; text-align: center; font-size: 8px; margin: 0px;} QProgressBar::chunk { background-color: #4CAF50; }")
        log_layout.addWidget(self.progress_bar)

        main_layout.addWidget(log_group)
        
        # 签名
        self.sig_widget = SignatureWidget(data=datetime.datetime.now().strftime("%Y-%m-%d"))
        main_layout.addWidget(self.sig_widget)

    def init_run_tab(self):
        layout = QHBoxLayout(self.tab_run)
        
        # 左侧面板: 控件
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_panel.setMaximumWidth(250)
        
        self.btn_load = QPushButton("选择数据文件")
        self.btn_load.clicked.connect(self.select_files)
        self.btn_load.setStyleSheet("padding: 8px; font-weight: bold;")
        left_layout.addWidget(self.btn_load)
        
        self.list_widget = QListWidget()
        self.list_widget.currentRowChanged.connect(self.on_list_selection)
        left_layout.addWidget(self.list_widget)

        self.btn_process = QPushButton("处理并保存所有")
        self.btn_process.clicked.connect(self.run_process_all)
        self.btn_process.setStyleSheet("background-color: #4CAF50; color: white; padding: 10px; font-weight: bold;")
        left_layout.addWidget(self.btn_process)

        # 可视化选项
        group_vis = QGroupBox("视图")
        vis_layout = QVBoxLayout(group_vis)
        self.btn_view_mcu = QPushButton("MCU效率")
        self.btn_view_motor = QPushButton("电机效率")
        self.btn_view_sys = QPushButton("系统效率")
        
        self.btn_view_mcu.clicked.connect(lambda: self.switch_plot('MCU'))
        self.btn_view_motor.clicked.connect(lambda: self.switch_plot('Motor'))
        self.btn_view_sys.clicked.connect(lambda: self.switch_plot('SYS'))
        
        vis_layout.addWidget(self.btn_view_mcu)
        vis_layout.addWidget(self.btn_view_motor)
        vis_layout.addWidget(self.btn_view_sys)

        # 占比按钮
        self.btn_view_ratio = QPushButton("效率占比")
        self.btn_view_ratio.clicked.connect(self.show_ratio_plot)
        vis_layout.addWidget(self.btn_view_ratio)

        left_layout.addWidget(group_vis)
        
        left_layout.addStretch()
        
        layout.addWidget(left_panel)
        
        # 右侧面板: 绘图
        self.plot_area = QWidget()
        self.plot_layout = QVBoxLayout(self.plot_area)
        
        self.figure = Figure()
        
        # 使用 AspectRatioWidget 包装器进行 UI 显示以保持 25x20cm (1.25 AR) 比例
        # 9.84 / 7.87 大约 1.2503
        self.ar_container = AspectRatioWidget(self.figure, aspect_ratio=9.84/7.87)
        self.canvas = self.ar_container.canvas 
        
        self.toolbar = NavigationToolbar2QT(self.canvas, self)
        
        self.plot_layout.addWidget(self.toolbar)
        # 添加容器而不是直接添加画布
        self.plot_layout.addWidget(self.ar_container)
        
        layout.addWidget(self.plot_area)
        
        # 存储绘图数据
        self.current_results = {}

    def init_config_tab(self):
        layout = QVBoxLayout(self.tab_config)
        
        # 表单区域
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        self.config_form_widget = QWidget()
        self.config_form_layout = QFormLayout(self.config_form_widget)
        # 增加一些间距
        self.config_form_layout.setSpacing(10)
        self.config_form_layout.setContentsMargins(10, 10, 10, 10)
        
        scroll.setWidget(self.config_form_widget)
        layout.addWidget(scroll)
        
        # 按钮区域
        btn_layout = QVBoxLayout()
        
        self.btn_save_config = QPushButton("保存并重载 (Save & Reload)")
        self.btn_save_config.setStyleSheet("height: 40px; font-weight: bold; font-size: 11pt;")
        self.btn_save_config.clicked.connect(self.save_config)
        btn_layout.addWidget(self.btn_save_config)
        
        self.btn_open_ini = QPushButton("打开配置文件 (Open INI File)")
        self.btn_open_ini.clicked.connect(self.open_ini_file)
        btn_layout.addWidget(self.btn_open_ini)
        
        layout.addLayout(btn_layout)
        
        self.config_fields = {} # Key -> QLineEdit
        self.current_encoding = 'utf-8' # 默认

    def open_ini_file(self):
        try:
            if sys.platform == 'win32':
                os.startfile(self.ini_path)
            else:
                opener = "open" if sys.platform == "darwin" else "xdg-open"
                subprocess.call([opener, self.ini_path])
        except Exception as e:
            QMessageBox.warning(self, "错误", f"无法打开文件: {e}")

    def parse_ini_file(self, path):
        # 即使没有 section 也能读取 INI 的辅助函数
        if not os.path.exists(path):
            return configparser.ConfigParser(), {}

        content = ""
        encoding_used = 'utf-8'
        try:
            with open(path, 'r', encoding='utf-8') as f:
                content = f.read()
        except UnicodeDecodeError:
            try:
                with open(path, 'r', encoding='gb18030') as f:
                    content = f.read()
                encoding_used = 'gb18030'
            except Exception as e:
                raise Exception(f"读取配置文件失败。未知编码。 ({e})")
        
        self.current_encoding = encoding_used

        # 检查初始 section
        has_section = False
        for line in content.splitlines():
            line = line.strip()
            if line.startswith('[') and line.endswith(']'):
                has_section = True
                break
            if line and not line.startswith('#') and not line.startswith(';'):
                break # 在 section 之前找到 key

        if not has_section:
            content = '[DEFAULT]\n' + content
            
        parser = configparser.ConfigParser()
        parser.optionxform = str # 大小写敏感
        parser.read_string(content)
        
        # 扁平化以便使用的 dict，但保留对象用于保存
        flat_dict = {}
        # 手动包含 DEFAULT section，因为 parser.sections() 会跳过它
        if 'DEFAULT' in parser:
             for key, val in parser['DEFAULT'].items():
                val_clean = val.split(';')[0].split('#')[0].strip()
                flat_dict[key] = val_clean

        for section in parser.sections():
            for key, val in parser.items(section):
                # 移除行内注释
                val_clean = val.split(';')[0].split('#')[0].strip()
                flat_dict[key] = val_clean
        
        return parser, flat_dict

    def reload_config(self):
        try:
            self.raw_config_obj, self.config_dict = self.parse_ini_file(self.ini_path)
            
            # 重建 UI
            # 清除布局
            while self.config_form_layout.count():
                child = self.config_form_layout.takeAt(0)
                if child.widget(): child.widget().deleteLater()
            
            self.config_fields = {}
            
            # 组合要迭代的 sections，确保包含 DEFAULT (如果有条目)
            sections_to_show = []
            if 'DEFAULT' in self.raw_config_obj and len(self.raw_config_obj['DEFAULT']) > 0:
                sections_to_show.append('DEFAULT')
            sections_to_show.extend(self.raw_config_obj.sections())

            for section in sections_to_show:
                # 仅当有多个 section 或它是命名 section 时添加 Section 标签
                if section != "DEFAULT" or (section == "DEFAULT" and len(sections_to_show) > 1):
                    sec_label = QLabel(f"[{section}]")
                    sec_label.setStyleSheet("font-weight: bold; font-size: 11pt; color: #333; margin-top: 10px;")
                    self.config_form_layout.addRow(sec_label)
                elif section == "DEFAULT" and len(sections_to_show) == 1:
                     # 单个 default section - 也许添加通用标题或保持干净？
                     # 让我们添加一个微妙的分隔符或什么都不加。
                     pass

                # 尽可能使用 parser[section] 以避免继承重复，
                # 尽管 items() 对于值通常更安全。
                # 对于 DEFAULT，items() 很好。对于其他，items() 包含默认值。
                # 为了避免编辑器中出现重复，我们应该明确迭代 section 的 keys。
                
                keys = self.raw_config_obj[section].keys()
                
                for key in keys:
                    val = self.raw_config_obj[section][key]
                    edit = QLineEdit(val)
                    # 按要求应用样式
                    edit.setStyleSheet("""
                        QLineEdit {
                            font-family: 'Microsoft YaHei', 'Segoe UI', sans-serif;
                            font-size: 10pt;
                            padding: 4px; 
                            border: 1px solid #aaa;
                            border-radius: 3px;
                            background-color: white;
                        }
                        QLineEdit:focus {
                            border: 1px solid #3399ff;
                        }
                    """)
                    self.config_fields[(section, key)] = edit
                    self.config_form_layout.addRow(key + ":", edit)
            # 初始化逻辑类
            self.logic = MotorEffLogic(self.config_dict)
            
        except Exception as e:
            logging.error(f"加载配置失败: {e}")
            QMessageBox.critical(self, "配置错误", str(e))

    def save_config(self):
        try:
            # 更新 parser 对象 (保持在内存中同步)
            for (section, key), edit in self.config_fields.items():
                self.raw_config_obj.set(section, key, edit.text())
            
            # 为 write_ini_file 准备扁平化数据
            flat_data = {}
            for (section, key), edit in self.config_fields.items():
                flat_data[key] = edit.text()
            
            # 使用自定义写入以保留注释
            self.write_ini_file(self.ini_path, flat_data)
                
            logging.info("配置已保存。")
            self.reload_config() # Reload to refresh internal state
            QMessageBox.information(self, "成功", "配置已保存并重新加载。")
            
        except Exception as e:
            logging.error(f"保存配置失败: {e}")
            QMessageBox.critical(self, "保存错误", str(e))

    def write_ini_file(self, file_path, data):
        """
        自定义写入函数以尽可能保持结构。
        读取行，替换 key=value，写回。
        """
        try:
            with open(file_path, 'r', encoding=self.current_encoding) as f:
                lines = f.readlines()
        except FileNotFoundError:
            lines = []
            
        new_lines = []
        
        for line in lines:
            stripped = line.strip()
            # 直接传递 注释/空行/section 头
            if not stripped or stripped.startswith(';') or stripped.startswith('#') or stripped.startswith('['):
                new_lines.append(line)
                continue
                
            if '=' in line:
                parts = line.split('=', 1)
                key = parts[0].strip()
                
                # 检查此 key 是否存在于我们的数据中
                if key in data:
                    new_val = data[key]
                    # 如果有，保留原始间距/注释？
                    # 简单起见，我们重建 "Key = Value" 但尝试保留行内注释 (如果存在)
                    # 然而，我们的 data[key] 只是值。
                    # 让我们看看原始行是否有注释
                    
                    comment_part = ""
                    # 非常简单的注释检查
                    if '#' in parts[1]:
                        comment_part = " #" + parts[1].split('#', 1)[1]
                    elif ';' in parts[1]:
                        comment_part = " ;" + parts[1].split(';', 1)[1]
                        
                    # 保持缩进？假设通常 keys 没有缩进的标准 INI
                    new_lines.append(f"{key} = {new_val}{comment_part}\n")
                else:
                    # Key 不在我们编辑的数据中？保持原样
                    new_lines.append(line)
            else:
                new_lines.append(line)
        
        # 写回
        with open(file_path, 'w', encoding=self.current_encoding) as f:
            f.writelines(new_lines)

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "选择数据文件", "", "Excel Files (*.xls *.xlsx)")
        if files:
            self.data_files = files
            self.list_widget.clear()
            # 迭代文件和 sheets 来填充列表
            self.all_results = [] # list of dicts {'file': path, 'sheet': name, 'disp': str}
            
            for fpath in files:
                success, msg = self.logic.load_data(fpath)
                if success:
                    fname = os.path.basename(fpath)
                    for sheet_name in self.logic.sheets_dict.keys():
                        disp_name = f"{fname} - {sheet_name}"
                        self.all_results.append({
                            'file': fpath,
                            'sheet': sheet_name,
                            'disp': disp_name
                        })
                        self.list_widget.addItem(disp_name)
                    logging.info(f"已加载 {fname}: {len(self.logic.sheets_dict)} 个 sheets。")
                else:
                     logging.error(f"加载 {fpath} 失败: {msg}")

    def on_list_selection(self, row):
        if row < 0 or row >= len(self.all_results):
            return
            
        item = self.all_results[row]
        fpath = item['file']
        sheet = item['sheet']
        
        logging.info(f"已选择: {item['disp']}")
        
        # 加载特定 sheet 到逻辑类
        if self.logic.current_file != fpath:
             self.logic.load_data(fpath)
        
        self.logic.set_current_sheet(sheet)
        self.process_current_data()

    def run_process_all(self):
        """批量处理所有加载的 sheet 并保存结果"""
        if not self.all_results:
             QMessageBox.warning(self, "无数据", "请先加载文件。")
             return
        
        # 如果有现有绘图清除它
        self.figure.clear()
        self.canvas.draw()
        
        total = len(self.all_results)
        self.progress_bar.setValue(0)
        
        for i, item in enumerate(self.all_results):
            self.list_widget.setCurrentRow(i)
            QApplication.processEvents() # 允许 UI 更新
            
            # 更新进度
            # 我们希望在进行中显示进度。
            # 逻辑由 on_list_selection -> process_current_data 处理。
            # process_current_data 将在单个文件结束时将进度设置为 100。
            # 但这里我们是批量处理。
            percent = int((i + 1) / total * 100)
            self.progress_bar.setValue(percent)

        # QMessageBox.information(self, "完成", f"已处理 {total} 个 sheets。")
        self.progress_bar.setValue(100)

    def process_current_data(self):
        # 如果不是由批量触发，重置单个运行的进度
        # (我们不容易区分源，但设置为 50%...100% 没问题)
        self.progress_bar.setValue(10)
        
        direction, state = self.logic.filter_data()
        if direction is None:
            logging.error("映射数据列失败。请检查配置。")
            self.progress_bar.setValue(0)
            return

        # 全局标准化术语
        # Logic 返回 "电动"/"发电". 我们要 "驱动"/"发电".
        if state == "电动": state = "驱动"

        self.logic.normalization()
        self.progress_bar.setValue(30)
        
        self.logic.get_external_characteristics()
        self.progress_bar.setValue(50)

        # 存储结果用于切换视图
        self.current_results = {
            'file': os.path.basename(self.logic.current_file),
            'direction': direction,
            'state': state
        }
        
        # 1. 绘制并保存效率 MAPs (根据配置 MCU, Motor, SYS)
        if self.config_dict.get('MCUMAP', '0') == '1':
             self.switch_plot('MCU', save_png=True)
             
        if self.config_dict.get('MotorMAP', '0') == '1':
             # 我们切换绘制并保存，但对于 GUI 显示我们可能希望停留在一个上
             self.switch_plot('Motor', save_png=True)

        if self.config_dict.get('SYSMAP', '0') == '1':
             self.switch_plot('SYS', save_png=True)
        
        self.progress_bar.setValue(80) 
        
        # 2. 计算并保存区域占比 (Excel 和 绘图)
        self.process_area_ratios()
        
        self.progress_bar.setValue(100)

        # 重置视图到 MCU 或默认
        # self.switch_plot('MCU')

    def _collect_ratio_data(self):
        """辅助函数：收集占比数据。"""
        calc_mcu = self.config_dict.get('MCUAreaRatioCalculation', '0') == '1'
        calc_motor = self.config_dict.get('MotorAreaRatioCalculation', '0') == '1'
        calc_sys = self.config_dict.get('SYSAreaRatioCalculation', '0') == '1'
        
        def get_ratios_data(eff_type):
            res = self.logic.process_map_data(eff_type)
            if res:
                # 预期解包 5 个值 (Logic 已更新)
                try:
                    _, _, _, Z_Eff, geo_mask = res
                    return self.logic.calculate_area_ratios(Z_Eff, geo_mask)
                except ValueError:
                    # 如果未更新的回退？不，我们更新了逻辑。
                    # 但如果用户缓存了奇怪的东西？
                    _, _, _, Z_Eff = res
                    return self.logic.calculate_area_ratios(Z_Eff)
            return []

        mcu = get_ratios_data('Eff_MCU') if calc_mcu else []
        motor = get_ratios_data('Eff_Motor') if calc_motor else []
        sys_r = get_ratios_data('Eff_SYS') if calc_sys else []
        
        return mcu, motor, sys_r

    def _plot_ratio_on_axes(self, ax, mcu_ratios, motor_ratios, sys_ratios, title):
        """辅助函数：在给定坐标轴上绘制占比图。"""
        if mcu_ratios:
            levels = [r['Level'] for r in mcu_ratios if r['Ratio'] > 0]
            ratios = [r['Ratio'] for r in mcu_ratios if r['Ratio'] > 0]
            if levels: ax.plot(levels, ratios, '-*b', linewidth=1.5, label='控制器 (MCU)')
        
        if motor_ratios:
            levels = [r['Level'] for r in motor_ratios if r['Ratio'] > 0]
            ratios = [r['Ratio'] for r in motor_ratios if r['Ratio'] > 0]
            if levels: ax.plot(levels, ratios, '-og', linewidth=1.5, label='电机 (Motor)')
                
        if sys_ratios:
            levels = [r['Level'] for r in sys_ratios if r['Ratio'] > 0]
            ratios = [r['Ratio'] for r in sys_ratios if r['Ratio'] > 0]
            if levels: ax.plot(levels, ratios, '-+m', linewidth=1.5, label='系统 (System)')
        
        ax.grid(True)
        ax.set_xlim(80, 100)
        ax.set_ylim(0, 100)
        ax.set_xticks(range(80, 101, 1))
        ax.set_yticks(range(0, 101, 10))
        
        ax.set_xlabel('效率 [%]')
        ax.set_ylabel('效率区域占比 [%]')
        ax.set_title(title, fontsize=15)
        ax.legend()

    def show_ratio_plot(self):
        """查看占比按钮的参数。"""
        if self.logic is None or self.logic.processed_df is None:
            QMessageBox.warning(self, "无数据", "请先处理数据。")
            return

        mcu, motor, sys_r = self._collect_ratio_data()
        
        self.figure.clear()
        ax = self.figure.add_subplot(111)
        
        # 标题构造
        veh_code = self.config_dict.get('VehicleCode', '')
        udc_val = "0"
        if self.logic.u_dc is not None:
             udc_val = str(int(round(self.logic.u_dc.mean())))
        direction = self.current_results.get('direction', '')
        state = self.current_results.get('state', '')
        
        if state == "电动": state = "驱动"
        
        title_str = f"{veh_code}-{udc_val}V-{direction}{state}-效率区域占比图"
        
        self._plot_ratio_on_axes(ax, mcu, motor, sys_r, title_str)
        self.canvas.draw()

    def process_area_ratios(self):
        """计算区域占比，保存到 Excel，并绘制占比图。"""
        # 检查是否有任何占比计算启用
        calc_mcu = self.config_dict.get('MCUAreaRatioCalculation', '0') == '1'
        calc_motor = self.config_dict.get('MotorAreaRatioCalculation', '0') == '1'
        calc_sys = self.config_dict.get('SYSAreaRatioCalculation', '0') == '1'
        
        if not (calc_mcu or calc_motor or calc_sys):
            return

        # 准备文件名
        veh_code = self.config_dict.get('VehicleCode', '')
        
        udc_val = "0"
        if self.logic.u_dc is not None:
             udc_val = str(int(round(self.logic.u_dc.mean())))
             
        direction = self.current_results.get('direction', '')
        state = self.current_results.get('state', '')
        
        # 文件术语标准化
        if state == "电动": state = "驱动"
        
        base_name = f"{veh_code}{udc_val}V{direction}{state}"
        excel_name = f"{base_name}效率占比.xlsx"
        plot_name = f"{veh_code}-{udc_val}V-{direction}{state}-效率区域占比图.png"
        title_str = f"{veh_code}-{udc_val}V-{direction}{state}-效率区域占比图"
        
        # 收集数据
        mcu_ratios, motor_ratios, sys_ratios = self._collect_ratio_data()
        
        # --- 1. Excel 写入 ---
        # 获取 levels 文本 (使用可用的任意一个)
        row_labels = []
        if mcu_ratios: row_labels = [f"≥{r['Level']}" for r in mcu_ratios]
        elif motor_ratios: row_labels = [f"≥{r['Level']}" for r in motor_ratios]
        elif sys_ratios: row_labels = [f"≥{r['Level']}" for r in sys_ratios]
        
        if not row_labels:
            return # 无数据
            
        # 构造主 DF
        df = pd.DataFrame({'效率区间': row_labels})
        if calc_mcu: df['控制器效率占比'] = [r['Ratio'] for r in mcu_ratios] if mcu_ratios else []
        if calc_motor: df['电机效率占比'] = [r['Ratio'] for r in motor_ratios] if motor_ratios else []
        if calc_sys: df['系统效率占比'] = [r['Ratio'] for r in sys_ratios] if sys_ratios else []
        
        try:
            with pd.ExcelWriter(excel_name, engine='openpyxl') as writer:
                header_df = pd.DataFrame(columns=['效率区间','控制器效率占比','电机效率占比','系统效率占比'])
                header_df.to_excel(writer, sheet_name='数据占比', index=False, startrow=0)
                df.to_excel(writer, sheet_name='数据占比', index=False, header=False, startrow=2)
            logging.info(f"保存区域占比 Excel: {excel_name}")
        except Exception as e:
            logging.error(f"保存 Excel {excel_name} 失败: {e}")

        # --- 2. 占比绘图 ---
        # 匹配 MATLAB 尺寸 (大约 25x20cm)
        fig = Figure(figsize=(9.84, 7.87), dpi=100) # figsize 以英寸为单位
        # 边距 [1.5 1.5 1 1] cm 大约
        fig.subplots_adjust(left=0.08, bottom=0.08, right=0.96, top=0.94)
        
        ax = fig.add_subplot(111)
        self._plot_ratio_on_axes(ax, mcu_ratios, motor_ratios, sys_ratios, title_str)
        
        try:
            fig.savefig(plot_name, dpi=200) # 固定 DPI, 不使用 bbox_inches='tight'
            logging.info(f"保存占比图: {plot_name}")
        except Exception as e:
            logging.error(f"保存图表 {plot_name} 失败: {e}")
            
        # 在 GUI 上绘制
        self.show_ratio_plot()


    def run_process(self):
        pass # 已弃用，支持列表选择

    def switch_plot(self, eff_type_short, save_png=False):
        if not hasattr(self.logic, 'processed_df') or self.logic.processed_df is None:
            return

        map_key = ""
        title_part = ""
        
        if eff_type_short == 'MCU':
            map_key = 'Eff_MCU'
            title_part = "MCU"
        elif eff_type_short == 'Motor':
            map_key = 'Eff_Motor'
            title_part = "Motor"
        elif eff_type_short == 'SYS':
            map_key = 'Eff_SYS'
            title_part = "System"
            
        # 检查是否在配置中启用
        # 逻辑有配置，我们可以检查。
        # 但目前按需计算。
        
        res = self.logic.process_map_data(map_key)
        if res is None:
            logging.warning(f"无法处理 {map_key} 的 map")
            return
            
        try:
             XI, YI, ZI_Power, ZI_Eff, _ = res
        except ValueError:
             XI, YI, ZI_Power, ZI_Eff = res
        
        # 绘图
        self.figure.clear()
        
        # 注意：对于屏幕显示，不要强制绝对英寸尺寸。
        # 这允许 AspectRatioWidget (UI) 处理大小调整。
        # 如果我们强制它，如果小部件很小，Matplotlib 可能会裁剪内容。
        # self.figure.set_size_inches(9.84, 7.87) <-- 为屏幕逻辑移除

        # 为屏幕显示使用更宽松的边距以确保标签可见
        # (MATLAB 边距对于较小的 GUI 窗口来说太紧了)
        self.figure.subplots_adjust(left=0.12, bottom=0.12, right=0.95, top=0.94)
        
        ax = self.figure.add_subplot(111)
        
        # 来自配置的等高线设置
        eff_steps_str = self.config_dict.get('EffMAPStep', '90 85 80 70')
        raw_levels = [float(x) for x in eff_steps_str.replace(';', ' ').replace(',', ' ').split()]
        eff_levels = sorted(list(set(raw_levels)))
        
        # 修复：确保包含 100%，以便最高效率区域 (>max_configured_level) 被着色
        if eff_levels and eff_levels[-1] < 100:
             eff_levels.append(100.0)

        # 基本 ContourF
        try:
            # 填充等高线
            cf = ax.contourf(XI, YI, ZI_Eff, levels=eff_levels, cmap='jet')
            # 移除 Colorbar 以匹配 MATLAB 参考和固定边距
            # self.figure.colorbar(cf, ax=ax, label='Efficiency (%)')
            
            # 效率线 (黑色)
            line_levels = [l for l in eff_levels if l < 100]
            ce = ax.contour(XI, YI, ZI_Eff, levels=line_levels, colors='k', linewidths=0.5)
            
            # 严格根据请求标记
            # 使用下面的碰撞检测来清理重叠而不是跳过级别
            ax.clabel(ce, levels=line_levels, inline=True, fontsize=8, fmt='%1.0f')

            # 添加图例 (效率和功率) 以匹配 MATLAB
            from matplotlib.lines import Line2D
            
            # 效率：黑色边框，深蓝色填充 (模拟深效率区)
            # MATLAB 通常使用填充椭圆。我们使用圆形标记来模拟。
            # 减小 markersize 以适配 UI
            line_eff = Line2D([0], [0], color='k', linewidth=0.5, linestyle='-',
                              marker='o', markersize=4, markerfacecolor='#000080', markeredgecolor='k',
                              label='效率')

            # 功率：绿线，空心中心
            # 减小 markersize 以适配 UI
            line_pwr = Line2D([0], [0], color='green', linewidth=0.8, linestyle='-',
                              marker='o', markersize=4, markerfacecolor='none', markeredgecolor='green',
                              label='功率')

            legend_elements = [line_eff, line_pwr]
            
            # 右上角图例，带框，黑色边框，非圆角
            # 改为半透明并允许拖动，以解决 UI 界面遮挡图形的问题
            # 减小 fontsize (9 -> 7) 以减小整体面积
            leg = ax.legend(handles=legend_elements, loc='upper right', fancybox=False, edgecolor='k', framealpha=0.7, fontsize=7)
            leg.get_frame().set_linewidth(0.5) # 更细的边框
            try:
                leg.set_draggable(True)
            except:
                pass
            
        except Exception as e:
            logging.error(f"等高线图错误: {e}")
        
        # 功率等高线 (绿色)
        # 如果未设置，计算级别
        power_levels_str = self.config_dict.get('PowerMAPStep', '')
        if power_levels_str:
             normalized_str = power_levels_str.replace(',', ' ').replace(';', ' ')
             power_levels = sorted([float(x) for x in normalized_str.split()])
        else:
             power_levels = 10 
             
        try:
            cp = ax.contour(XI, YI, ZI_Power, levels=power_levels, colors='green',  linewidths=0.8)
            ax.clabel(cp, inline=True, fontsize=8, fmt='%1.0f')
        except Exception as e:
            pass
        
        # 轴设置 (移到重叠移除之前以确保正确的坐标变换)
        try:
            xstep = int(self.config_dict.get('xstep', 200)) # Default 200
            ystep = int(self.config_dict.get('ystep', 10))  # Default 10
            
            if XI.size > 0 and YI.size > 0:
                 max_x = np.nanmax(XI)
                 max_y = np.nanmax(YI)
                 
                 # 收紧限制以匹配 MATLAB (移除额外的步长填充)
                 end_x = int(np.ceil(max_x / xstep)) * xstep
                 end_y = int(np.ceil(max_y / ystep)) * ystep
                 
                 # 设置包含上限的刻度 (arange 是不包含的)
                 ax.set_xticks(np.arange(0, end_x + xstep/2, xstep))
                 ax.set_yticks(np.arange(0, end_y + ystep/2, ystep))
                 
                 ax.set_xlim(0, end_x) # 显式设置限制
                 ax.set_ylim(0, end_y)
        except:
            pass            

        # 调整 Label/Tick 字体大小以避免 UI 重叠
        ax.tick_params(axis='both', which='major', labelsize=8) 
        ax.set_xlabel('转速 [rpm]', fontsize=9) 
        ax.set_ylabel('扭矩 [N.m]', fontsize=9)
        
        # 标题设置
        veh_code = self.config_dict.get('VehicleCode', '')
        
        udc_val = "0"
        if self.logic.u_dc is not None:
             udc_val = str(int(round(self.logic.u_dc.mean())))

        direction = self.current_results.get('direction', '') 
        state = self.current_results.get('state', '') 
        
        # 术语标准化
        # 电动 -> 驱动 (Drive)
        # 发电 -> 发电 (Generate) - 不是 '制动' (Brake)
        if state == "电动": state_str = "驱动"
        elif state == "发电": state_str = "发电"
        else: state_str = state

        map_name_cn = ""
        if eff_type_short == 'MCU': map_name_cn = "控制器效率MAP"
        elif eff_type_short == 'Motor': map_name_cn = "电机效率MAP"
        elif eff_type_short == 'SYS': map_name_cn = "系统效率MAP"
        
        final_title = f"{veh_code}-{udc_val}V-{direction}{state_str}-{map_name_cn}"
        ax.set_title(final_title, fontsize=12)
        ax.grid(True, linestyle=':', alpha=0.6)
        
        # 强制绘制以计算标签位置
        self.canvas.draw()
            
        # 后处理：移除重叠标签
        # 现在布局已定，检查重叠
        all_texts = []
        for text_obj in ax.texts:
             all_texts.append(text_obj)
             
        if len(all_texts) > 1:
            renderer = self.canvas.get_renderer()
            bboxes = [] 
            texts_to_remove = []
            
            # 辅助函数：检查文本是否为功率 (绿色) 或 效率 (黑色)
            # 默认 matplotlib 黑色是 (0,0,0,1) 或 'k'
            # 功率设置为 'green'
            def is_black(t):
                c = t.get_color()
                return c == 'k' or c == 'black' or c == (0,0,0,1) or c == (0,0,0)

            # 优先级：效率 (黑色) > 功率 (绿色)
            # 排序：黑色优先
            all_texts.sort(key=lambda x: 1 if is_black(x) else 0, reverse=True)
            
            for t in all_texts:
                try:
                    bbox = t.get_window_extent(renderer)
                    bbox = bbox.expanded(1.2, 1.2) # 边距
                    
                    overlap = False
                    for existing_bbox in bboxes:
                        if bbox.overlaps(existing_bbox):
                            overlap = True
                            break
                    
                    if overlap:
                        texts_to_remove.append(t)
                    else:
                        bboxes.append(bbox)
                except:
                    pass
            
            for t in texts_to_remove:
                t.set_visible(False)

        # 如果我们隐藏了东西，重新绘制？
        # 如果我们只是更改了可见性属性，通常不需要严格要求，但这是个好习惯。
        # self.canvas.draw() 
        
        # 保存区域占比到 Excel (需求)
        try:
             if len(res) == 5:
                  _, _, _, _, geo_mask = res
                  ratios = self.logic.calculate_area_ratios(ZI_Eff, geo_mask)
             else:
                  ratios = self.logic.calculate_area_ratios(ZI_Eff)
        except:
             ratios = self.logic.calculate_area_ratios(ZI_Eff)
             
        # Log removed as requested
        # logging.info(f"--- Area Ratios for {title_part} ---")
        # for r in ratios:
        #    logging.info(f"Efficiency >= {r['Level']}% : {r['Ratio']:.2f}%")
            
        # 自动保存图像
        # 文件名构造镜像 MATLAB
        # [VehicleCode]-[Voltage]V-[Direction][State]-[Type]EfficiencyMAP.png
        try:
            udc_val = "0"
            if self.logic.u_dc is not None:
                udc_val = str(int(round(self.logic.u_dc.mean())))
                
            # 使用已经标准化 (驱动/发电) 的 'state_str' 而不是原始 state
            fname = f"{veh_code}-{udc_val}V-{self.current_results.get('direction')}{state_str}-{title_part.replace(' ', '')}EfficiencyMAP.png"
            
            # 清理文件名
            fname = "".join([c for c in fname if c.isalnum() or c in (' ', '.', '-', '_')]).strip()
            
            # 使用固定 DPI 并移除 bbox_inches='tight' 以尊重上面设置的确切图形尺寸
            # MATLAB 默认 'print' 通常使用 150 DPI 或屏幕 DPI。我们将为了平衡使用 200 DPI。
            
            # 临时：切换尺寸和边距以在文件中实现 MATLAB 对等
            old_size = self.figure.get_size_inches()
            self.figure.set_size_inches(9.84, 7.87) # 25x20cm 严格
            
            # 对文件应用严格的 MATLAB 边距
            self.figure.subplots_adjust(left=0.08, bottom=0.08, right=0.96, top=0.94)

            self.figure.savefig(fname, dpi=200)
            
            # 恢复：切换回以前的尺寸和安全的 UI 边距
            self.figure.set_size_inches(old_size)
            # 恢复 UI 的安全边距，以便标签在屏幕上可见
            self.figure.subplots_adjust(left=0.12, bottom=0.12, right=0.95, top=0.94) 
            self.canvas.draw()
            
            logging.info(f"已保存图形到 {fname}")
        except Exception as e:
            logging.error(f"保存图像失败: {e}")

class SignatureWidget(QWidget):
    def __init__(self, author="sword", data="2026-01-15"):
        super().__init__()
        # 透明背景，右下锚点
        self.layout = QHBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 10, 0) # 右边距
        self.layout.setAlignment(Qt.AlignRight)
        
        self.lbl = QLabel(f"Author: {author} | Date: {data}")
        # 样式：默认隐藏
        self.lbl.setStyleSheet("background-color: #EEE; border: 1px solid #CCC; padding: 1px; font-size: 9px; color: #333;")
        self.lbl.hide()
        
        self.layout.addWidget(self.lbl)
        
        # 我们需要检测此小部件区域上的鼠标。
        self.setMouseTracking(True)
        self.setFixedHeight(20)
        
    def enterEvent(self, event):
        self.lbl.show()
        super().enterEvent(event)
        
    def leaveEvent(self, event):
        self.lbl.hide()
        super().leaveEvent(event)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    
    # Matplotlib 和 QT 全局字体修复
    # 尝试多种常用中文字体
    matplotlib.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'Arial']
    matplotlib.rcParams['axes.unicode_minus'] = False
    
    font = QFont("Microsoft YaHei")
    font.setPointSize(9)
    app.setFont(font)
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

