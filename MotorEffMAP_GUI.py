import sys
import os
import datetime
import logging
import configparser
import subprocess
import platform
import re
import pandas as pd
import numpy as np

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                               QHBoxLayout, QPushButton, QLabel, QFileDialog,
                               QTabWidget, QFormLayout, QLineEdit, QTextEdit,
                               QScrollArea, QMessageBox, QGroupBox, QListWidget,
                               QProgressBar, QComboBox)
from PySide6.QtCore import Qt
from PySide6.QtGui import QFont

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


class MainWindow(QMainWindow):
    CONFIG_LABELS = {
        'VehicleCode': '车型代号',
        'Speed': '数据里的转速名称',
        'Toqrue': '数据里的扭矩名称',
        'P_Motor': '数据里的功率名称',
        'Eff_MCU': '数据里的控制器效率名称',
        'Eff_Motor': '数据里的电机效率名称',
        'Eff_SYS': '数据里的系统效率名称',
        'U_dc': '数据里的母线电压名称',
        'customUdc': '固定母线电压（填写后会覆盖U_dc）',
        'MCUMAP': '控制器效率MAP绘制',
        'MCUAreaRatioCalculation': '控制器效率区域占比计算',
        'MotorMAP': '电机效率MAP绘制',
        'MotorAreaRatioCalculation': '电机效率区域占比计算',
        'SYSMAP': '系统效率MAP绘制',
        'SYSAreaRatioCalculation': '系统效率区域占比计算',
        'EffMAPStep': '效率等高线级别',
        'PowerMAPStep': '功率等高线级别',
        'xstep': '横轴刻度步长',
        'ystep': '纵轴刻度步长',
        'StartSpeed': '起始转速',
        'StartTorque': '起始扭矩',
        'SpeedGrid': '转速网格步长',
        'TorqueGrid': '扭矩网格步长',
        'MaxGridPoints': '最大网格点数',
        'customSpeedDirection': '固定转向名称',
        'customMotionState': '固定工况状态名称',
    }
    SWITCH_CONFIG_KEYS = {
        'MCUMAP',
        'MCUAreaRatioCalculation',
        'MotorMAP',
        'MotorAreaRatioCalculation',
        'SYSMAP',
        'SYSAreaRatioCalculation',
    }
    EFFICIENCY_MAP_OUTPUTS = (
        ('Eff_MCU', 'MCUMAP', 'MCU', '控制器'),
        ('Eff_Motor', 'MotorMAP', 'Motor', '电机'),
        ('Eff_SYS', 'SYSMAP', 'SYS', '系统'),
    )
    EFFICIENCY_RATIO_OUTPUTS = (
        ('Eff_MCU', 'MCUAreaRatioCalculation', '控制器'),
        ('Eff_Motor', 'MotorAreaRatioCalculation', '电机'),
        ('Eff_SYS', 'SYSAreaRatioCalculation', '系统'),
    )
    EXPORT_FIGURE_SIZE = (9.84, 7.87)
    EXPORT_ASPECT_RATIO = EXPORT_FIGURE_SIZE[0] / EXPORT_FIGURE_SIZE[1]
    FIGURE_LAYOUT = {
        'left': 0.08,
        'bottom': 0.08,
        'right': 0.96,
        'top': 0.94,
    }

    def __init__(self):
        super().__init__()
        self.resize(1200, 800)

        # Determine base path for config file (Robust for PyInstaller Exe)
        if getattr(sys, 'frozen', False):
            # If run as exe
            base_dir = os.path.dirname(sys.executable)
        else:
            # If run as script
            base_dir = os.path.dirname(os.path.abspath(__file__))

        self.base_dir = base_dir
        self.setWindowTitle(self.build_window_title())
        self.ini_path = os.path.join(base_dir, "MotorEffMAP.ini")

        if not os.path.exists(self.ini_path):
            # 如果不存在创建默认？或者警告。
            pass

        self.config_dict = {}
        self.raw_config_obj = None # 用于跟踪 sections

        self.logic = None
        self.data_files = []
        self.all_results = []

        self.init_ui()
        self.setup_logging()
        self.reload_config()

    def build_window_title(self):
        base_title = "电驱效率MAP绘制工具 (Python 版)"
        version_path = os.path.join(self.base_dir, "version.ini")
        if not os.path.exists(version_path):
            return base_title

        parser = configparser.ConfigParser()
        try:
            parser.read(version_path, encoding="utf-8")
        except Exception as e:
            logging.warning(f"读取版本文件失败: {e}")
            return base_title

        version = parser["version"] if parser.has_section("version") else parser.defaults()
        build_date = version.get("build_date", "").strip()
        code = version.get("code", "").strip()
        version_text = "-".join(part for part in [build_date, code] if part)
        if not version_text:
            return base_title
        return f"{base_title}  {version_text}"

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

        self.btn_view_mcu.clicked.connect(lambda: self.show_map_plot('MCU'))
        self.btn_view_motor.clicked.connect(lambda: self.show_map_plot('Motor'))
        self.btn_view_sys.clicked.connect(lambda: self.show_map_plot('SYS'))

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

        self.figure = Figure(figsize=self.EXPORT_FIGURE_SIZE, dpi=100)
        self.apply_figure_layout(self.figure)

        # 使用 AspectRatioWidget 包装器进行 UI 显示以保持 25x20cm (1.25 AR) 比例
        self.ar_container = AspectRatioWidget(self.figure, aspect_ratio=self.EXPORT_ASPECT_RATIO)
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
        encoding_used = 'utf-8-sig'
        for encoding in ('utf-8-sig', 'gb18030'):
            try:
                with open(path, 'r', encoding=encoding) as f:
                    content = f.read()
                encoding_used = encoding
                break
            except UnicodeDecodeError:
                continue
            except Exception as e:
                raise Exception(f"读取配置文件失败。({e})")
        else:
            raise Exception("读取配置文件失败。未知编码，请保存为 UTF-8 或 GB18030。")

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

    def get_config_display_label(self, key):
        description = self.CONFIG_LABELS.get(key)
        if not description:
            return f"{key}:"
        return f"{key}（{description}）:"

    def create_config_editor(self, key, value):
        if key in self.SWITCH_CONFIG_KEYS:
            combo = QComboBox()
            combo.addItem("开启", "1")
            combo.addItem("关闭", "0")
            combo.setCurrentIndex(0 if str(value).strip() == "1" else 1)
            combo.setStyleSheet("""
                QComboBox {
                    font-family: 'Microsoft YaHei', 'Segoe UI', sans-serif;
                    font-size: 10pt;
                    padding: 4px;
                    border: 1px solid #aaa;
                    border-radius: 3px;
                    background-color: white;
                }
                QComboBox:focus {
                    border: 1px solid #3399ff;
                }
            """)
            return combo

        edit = QLineEdit(value)
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
        return edit

    def get_config_editor_value(self, editor):
        if isinstance(editor, QComboBox):
            return editor.currentData()
        return editor.text()

    def apply_figure_layout(self, figure=None):
        target = figure or self.figure
        target.subplots_adjust(**self.FIGURE_LAYOUT)

    def is_config_switch_on(self, key):
        return self.config_dict.get(key, '0') == '1'

    def should_use_efficiency_output(self, eff_type, switch_key, display_name):
        if not self.is_config_switch_on(switch_key):
            return False
        if self.logic.has_efficiency_data(eff_type):
            return True
        logging.warning(f"{display_name}输出已开启，但 {eff_type} 未配置或没有有效数据，已跳过。")
        return False

    def get_active_ratio_outputs(self):
        active = {}
        for eff_type, switch_key, display_name in self.EFFICIENCY_RATIO_OUTPUTS:
            active[eff_type] = self.should_use_efficiency_output(
                eff_type,
                switch_key,
                f"{display_name}效率占比",
            )
        return active

    def get_non_negative_config_float(self, key, default=0):
        try:
            value = float(self.config_dict.get(key, default) or default)
        except (TypeError, ValueError):
            logging.warning(f"{key} 配置无效，已按 {default} 处理。")
            return float(default)
        if not np.isfinite(value) or value < 0:
            logging.warning(f"{key} 配置无效，已按 {default} 处理。")
            return float(default)
        return value

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
                    edit = self.create_config_editor(key, val)
                    self.config_fields[(section, key)] = edit
                    self.config_form_layout.addRow(self.get_config_display_label(key), edit)
            # 初始化逻辑类
            self.logic = MotorEffLogic(self.config_dict)
            self.clear_loaded_results()

        except Exception as e:
            logging.error(f"加载配置失败: {e}")
            QMessageBox.critical(self, "配置错误", str(e))

    def clear_loaded_results(self):
        """清空已加载数据状态，避免配置重载后列表指向旧逻辑对象。"""
        self.data_files = []
        self.all_results = []
        self.current_results = {}
        if hasattr(self, 'list_widget'):
            self.list_widget.clear()

    def reload_runtime_config(self):
        """重新读取运行时配置，并同步到现有逻辑对象。"""
        self.raw_config_obj, self.config_dict = self.parse_ini_file(self.ini_path)
        if self.logic is None:
            self.logic = MotorEffLogic(self.config_dict)
        else:
            self.logic.config = self.config_dict
        logging.info("已重新读取 INI 配置。")

    def save_config(self):
        try:
            # 更新 parser 对象 (保持在内存中同步)
            for (section, key), edit in self.config_fields.items():
                self.raw_config_obj.set(section, key, self.get_config_editor_value(edit))

            # 为 write_ini_file 准备扁平化数据
            flat_data = {}
            for (section, key), edit in self.config_fields.items():
                flat_data[key] = self.get_config_editor_value(edit)

            # 使用自定义写入以保留注释
            self.write_ini_file(self.ini_path, flat_data)

            logging.info("配置已保存。")
            self.reload_config() # Reload to refresh internal state
            QMessageBox.information(self, "成功", "配置已保存并重新加载。")

        except Exception as e:
            logging.error(f"保存配置失败: {e}")
            QMessageBox.critical(self, "保存错误", str(e))

    def handle_processing_error(self, message):
        logging.error(message)
        self.progress_bar.setValue(0)
        QMessageBox.warning(self, "处理错误", message)

    def sanitize_filename_component(self, value):
        text = str(value or "").strip()
        text = re.sub(r"[^\w\u4e00-\u9fff.-]+", "_", text)
        text = re.sub(r"_+", "_", text).strip("._ ")
        return text or "未命名"

    def build_output_stem(self, suffix=""):
        source_file = self.current_results.get('file', '') if hasattr(self, 'current_results') else ''
        source_stem = os.path.splitext(os.path.basename(source_file))[0]
        sheet = self.current_results.get('sheet', '') if hasattr(self, 'current_results') else ''
        veh_code = self.config_dict.get('VehicleCode', '')

        udc_val = "0"
        if self.logic is not None and self.logic.u_dc is not None:
             udc_val = str(int(round(self.logic.u_dc.mean())))

        direction = self.current_results.get('direction', '') if hasattr(self, 'current_results') else ''
        state = self.current_results.get('state', '') if hasattr(self, 'current_results') else ''
        if state == "电动":
            state = "驱动"

        parts = [
            source_stem,
            sheet,
            f"{veh_code}-{udc_val}V-{direction}{state}",
            suffix,
        ]
        return "_".join(self.sanitize_filename_component(p) for p in parts if p)

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
        written_keys = set()

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

                    # 空值不保留等号后的空格，避免写出行尾空白。
                    value_text = f" {new_val}" if str(new_val) else ""
                    new_lines.append(f"{key} ={value_text}{comment_part}\n")
                    written_keys.add(key)
                else:
                    # Key 不在我们编辑的数据中？保持原样
                    new_lines.append(line)
            else:
                new_lines.append(line)

        missing_keys = [key for key in data if key not in written_keys]
        if missing_keys:
            if new_lines and new_lines[-1].strip():
                new_lines.append("\n")
            new_lines.append("# 由程序补充的新增配置项\n")
            for key in missing_keys:
                new_val = data[key]
                value_text = f" {new_val}" if str(new_val) else ""
                new_lines.append(f"{key} ={value_text}\n")

        # 写回
        self.current_encoding = 'utf-8-sig'
        with open(file_path, 'w', encoding=self.current_encoding, newline='\n') as f:
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
        self.process_result_item(item)

    def process_result_item(self, item):
        fpath = item['file']
        sheet = item['sheet']

        logging.info(f"已选择: {item['disp']}")

        # 加载特定 sheet 到逻辑类
        if getattr(self.logic, 'current_file', None) != fpath:
             success, msg = self.logic.load_data(fpath)
             if not success:
                 self.handle_processing_error(f"加载 {fpath} 失败: {msg}")
                 return False

        if not self.logic.set_current_sheet(sheet):
            self.handle_processing_error(f"未找到工作表: {sheet}")
            return False

        return self.process_current_data()

    def run_process_all(self):
        """批量处理所有加载的 sheet 并保存结果"""
        if not self.all_results:
             QMessageBox.warning(self, "无数据", "请先加载文件。")
             return

        try:
            self.reload_runtime_config()
        except Exception as e:
            logging.error(f"重新读取配置失败: {e}")
            QMessageBox.critical(self, "配置错误", f"重新读取配置失败: {e}")
            return

        # 如果有现有绘图清除它
        self.figure.clear()
        self.canvas.draw()

        total = len(self.all_results)
        self.progress_bar.setValue(0)

        for i, item in enumerate(self.all_results):
            old_block = self.list_widget.blockSignals(True)
            self.list_widget.setCurrentRow(i)
            self.list_widget.blockSignals(old_block)
            self.process_result_item(item)
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
            detail = getattr(self.logic, 'last_error', '') or "请检查配置。"
            logging.error(f"映射数据列失败。{detail}")
            self.progress_bar.setValue(0)
            return False

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
            'sheet': getattr(self.logic, 'current_sheet', ''),
            'direction': direction,
            'state': state
        }

        # 1. 绘制并保存效率 MAPs：开关开启且对应效率列可用才输出
        for eff_type, switch_key, plot_key, display_name in self.EFFICIENCY_MAP_OUTPUTS:
            if self.should_use_efficiency_output(eff_type, switch_key, f"{display_name}效率MAP"):
                if not self.switch_plot(plot_key, save_png=True):
                    return False

        self.progress_bar.setValue(80)

        # 2. 计算并保存区域占比 (Excel 和 绘图)
        try:
            self.process_area_ratios()
        except ValueError as e:
            self.handle_processing_error(str(e))
            return False

        self.progress_bar.setValue(100)

        # 重置视图到 MCU 或默认
        # self.switch_plot('MCU')
        return True

    def _collect_ratio_data(self, active_efficiency_types=None):
        """辅助函数：收集占比数据。"""
        active_efficiency_types = active_efficiency_types or {}

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

        ratio_data = {}
        for eff_type, switch_key, display_name in self.EFFICIENCY_RATIO_OUTPUTS:
            if active_efficiency_types.get(eff_type):
                ratio_data[eff_type] = get_ratios_data(eff_type)
            else:
                ratio_data[eff_type] = []

        return ratio_data['Eff_MCU'], ratio_data['Eff_Motor'], ratio_data['Eff_SYS']

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
            return False

        try:
            self.reload_runtime_config()
        except Exception as e:
            logging.error(f"重新读取配置失败: {e}")
            QMessageBox.critical(self, "配置错误", f"重新读取配置失败: {e}")
            return False

        mcu, motor, sys_r = self._collect_ratio_data(self.get_active_ratio_outputs())
        if not (mcu or motor or sys_r):
            self.figure.clear()
            self.canvas.draw()
            logging.info("效率占比计算未启用或无有效数据，已按配置跳过显示。")
            return False

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

        self.apply_figure_layout()
        self._plot_ratio_on_axes(ax, mcu, motor, sys_r, title_str)
        self.canvas.draw()
        return True

    def process_area_ratios(self):
        """计算区域占比，保存到 Excel，并绘制占比图。"""
        # 检查是否有任何占比计算启用
        active_ratios = self.get_active_ratio_outputs()
        calc_mcu = active_ratios['Eff_MCU']
        calc_motor = active_ratios['Eff_Motor']
        calc_sys = active_ratios['Eff_SYS']

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

        base_name = self.build_output_stem("效率占比")
        excel_name = f"{base_name}.xlsx"
        plot_name = f"{base_name}.png"
        title_str = f"{veh_code}-{udc_val}V-{direction}{state}-效率区域占比图"

        # 收集数据
        mcu_ratios, motor_ratios, sys_ratios = self._collect_ratio_data(active_ratios)

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
        if calc_mcu and mcu_ratios:
            df['控制器效率占比'] = [r['Ratio'] for r in mcu_ratios]
        if calc_motor and motor_ratios:
            df['电机效率占比'] = [r['Ratio'] for r in motor_ratios]
        if calc_sys and sys_ratios:
            df['系统效率占比'] = [r['Ratio'] for r in sys_ratios]

        try:
            with pd.ExcelWriter(excel_name, engine='openpyxl') as writer:
                header_df = pd.DataFrame(columns=['效率区间','控制器效率占比','电机效率占比','系统效率占比'])
                header_df.to_excel(writer, sheet_name='数据占比', index=False, startrow=0)
                df.to_excel(writer, sheet_name='数据占比', index=False, header=False, startrow=2)
            logging.info(f"保存区域占比 Excel: {excel_name}")
        except Exception as e:
            logging.error(f"保存 Excel {excel_name} 失败: {e}")

        # --- 2. 占比绘图 ---
        # 使用固定导出尺寸，保证区域占比图在不同机器上的版式一致
        fig = Figure(figsize=self.EXPORT_FIGURE_SIZE, dpi=100) # figsize 以英寸为单位
        self.apply_figure_layout(fig)

        ax = fig.add_subplot(111)
        self._plot_ratio_on_axes(ax, mcu_ratios, motor_ratios, sys_ratios, title_str)

        try:
            fig.savefig(plot_name, dpi=200) # 固定 DPI, 不使用 bbox_inches='tight'
            logging.info(f"保存占比图: {plot_name}")
        except Exception as e:
            logging.error(f"保存图表 {plot_name} 失败: {e}")

        # 在 GUI 上绘制
        self.show_ratio_plot()

    def show_map_plot(self, eff_type_short):
        """按最新 INI 配置显示单个效率 MAP。"""
        try:
            self.reload_runtime_config()
        except Exception as e:
            logging.error(f"重新读取配置失败: {e}")
            QMessageBox.critical(self, "配置错误", f"重新读取配置失败: {e}")
            return False

        output_by_plot_key = {
            plot_key: (eff_type, switch_key, display_name)
            for eff_type, switch_key, plot_key, display_name in self.EFFICIENCY_MAP_OUTPUTS
        }
        output = output_by_plot_key.get(eff_type_short)
        if output is None:
            return False

        eff_type, switch_key, display_name = output
        if not self.should_use_efficiency_output(eff_type, switch_key, f"{display_name}效率MAP"):
            self.figure.clear()
            self.canvas.draw()
            logging.info(f"{display_name}效率MAP未启用，已按配置跳过显示。")
            return False

        return self.switch_plot(eff_type_short)

    def switch_plot(self, eff_type_short, save_png=False):
        if not hasattr(self.logic, 'processed_df') or self.logic.processed_df is None:
            return False

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

        try:
            res = self.logic.process_map_data(map_key)
        except ValueError as e:
            self.handle_processing_error(str(e))
            return False

        if res is None:
            logging.warning(f"无法处理 {map_key} 的 map")
            return False

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

        self.apply_figure_layout()

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
            # 当前版式不显示 Colorbar，避免挤占主图区域
            # self.figure.colorbar(cf, ax=ax, label='Efficiency (%)')

            # 效率线 (黑色)
            line_levels = [l for l in eff_levels if l < 100]
            ce = ax.contour(XI, YI, ZI_Eff, levels=line_levels, colors='k', linewidths=0.5)

            # 严格根据请求标记
            # 使用下面的碰撞检测来清理重叠而不是跳过级别
            ax.clabel(ce, levels=line_levels, inline=True, fontsize=8, fmt='%1.0f')

            # 添加效率和功率图例
            from matplotlib.lines import Line2D

            # 效率：黑色边框，深蓝色填充，用圆形标记表示高效率区域
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
            self.handle_processing_error(f"等高线图生成失败: {e}")
            return False

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
                start_x = self.get_non_negative_config_float('StartSpeed', 0)
                start_y = self.get_non_negative_config_float('StartTorque', 0)
                start_x = min(start_x, max_x)
                start_y = min(start_y, max_y)

                # 按数据上限收紧坐标范围，避免额外空白区域
                end_x = int(np.ceil(max_x / xstep)) * xstep
                end_y = int(np.ceil(max_y / ystep)) * ystep

                # 设置包含上限的刻度 (arange 是不包含的)
                ax.set_xticks(np.arange(start_x, end_x + xstep/2, xstep))
                ax.set_yticks(np.arange(start_y, end_y + ystep/2, ystep))

                ax.set_xlim(start_x, end_x) # 显式设置限制
                ax.set_ylim(start_y, end_y)
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

        if save_png:
            try:
                fname = f"{self.build_output_stem(title_part.replace(' ', '') + 'EfficiencyMAP')}.png"

                # 使用固定 DPI 并移除 bbox_inches='tight' 以尊重上面设置的确切图形尺寸
                # 使用 200 DPI 平衡导出清晰度和文件体积。

                # 导出时切换到固定尺寸，确保输出图片版式稳定
                old_size = self.figure.get_size_inches()
                self.figure.set_size_inches(*self.EXPORT_FIGURE_SIZE) # 25x20cm 严格

                # 对文件应用固定边距
                self.apply_figure_layout()

                self.figure.savefig(fname, dpi=200)

                # 恢复：切换回以前的尺寸和安全的 UI 边距
                self.figure.set_size_inches(old_size)
                self.apply_figure_layout()
                self.canvas.draw()

                logging.info(f"已保存图形到 {fname}")
            except Exception as e:
                logging.error(f"保存图像失败: {e}")
        return True

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
