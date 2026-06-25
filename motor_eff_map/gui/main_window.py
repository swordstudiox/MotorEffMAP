import sys
import os
import datetime
import logging
import configparser
import subprocess

from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                               QHBoxLayout, QPushButton, QLabel,
                               QTabWidget, QFormLayout, QTextEdit,
                               QScrollArea, QMessageBox, QGroupBox, QListWidget,
                               QProgressBar, QCheckBox)
from PySide6.QtGui import QFont

import matplotlib
matplotlib.use('QtAgg')
from matplotlib.backends.backend_qtagg import NavigationToolbar2QT
from matplotlib.figure import Figure

from motor_eff_map.logic import MotorEffLogic
from motor_eff_map.gui.config_schema import (
    CONFIG_LABELS,
    DEFAULT_CONFIG_VALUES,
    DUAL_Y_FIGURE_LAYOUT,
    EFFICIENCY_MAP_OUTPUTS,
    EFFICIENCY_RATIO_OUTPUTS,
    EXPORT_ASPECT_RATIO,
    EXPORT_FIGURE_SIZE,
    FIGURE_LAYOUT,
    SWITCH_CONFIG_KEYS,
)
from motor_eff_map.gui.config_editor import ConfigEditorMixin
from motor_eff_map.gui.output_naming import OutputNamingMixin
from motor_eff_map.gui.plot_helpers import PlotHelperMixin
from motor_eff_map.gui.plotters import (
    EfficiencyMapPlotterMixin,
    ExternalCharacteristicsPlotterMixin,
    LossMapPlotterMixin,
    RatioPlotterMixin,
    SpeedPowerPlotterMixin,
)
from motor_eff_map.gui.processing_controller import ProcessingControllerMixin
from motor_eff_map.gui.widgets import AspectRatioWidget, QTextEditLogger, SignatureWidget

class MainWindow(
    ConfigEditorMixin,
    PlotHelperMixin,
    OutputNamingMixin,
    ProcessingControllerMixin,
    RatioPlotterMixin,
    ExternalCharacteristicsPlotterMixin,
    SpeedPowerPlotterMixin,
    LossMapPlotterMixin,
    EfficiencyMapPlotterMixin,
    QMainWindow,
):
    CONFIG_LABELS = CONFIG_LABELS
    SWITCH_CONFIG_KEYS = SWITCH_CONFIG_KEYS
    EFFICIENCY_MAP_OUTPUTS = EFFICIENCY_MAP_OUTPUTS
    EFFICIENCY_RATIO_OUTPUTS = EFFICIENCY_RATIO_OUTPUTS
    DEFAULT_CONFIG_VALUES = DEFAULT_CONFIG_VALUES
    EXPORT_FIGURE_SIZE = EXPORT_FIGURE_SIZE
    EXPORT_ASPECT_RATIO = EXPORT_ASPECT_RATIO
    FIGURE_LAYOUT = FIGURE_LAYOUT
    DUAL_Y_FIGURE_LAYOUT = DUAL_Y_FIGURE_LAYOUT

    def __init__(self):
        super().__init__()
        self.resize(1200, 800)

        # Determine base path for config file (Robust for PyInstaller Exe)
        if getattr(sys, 'frozen', False):
            # If run as exe
            base_dir = os.path.dirname(sys.executable)
        else:
            # If run as script
            base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))

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
        self.plot_image_cache = {}
        self.is_batch_running = False
        self._batch_thread = None
        self._batch_worker = None
        self._batch_error = ""
        self._close_after_batch = False

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
        self.file_log_handler = logging.FileHandler("MotorEffMAP.log", mode='a', encoding='utf-8')
        self.file_log_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

        # GUI 处理器 (控制台)
        self.gui_log_handler = QTextEditLogger(self.log_text)
        self.gui_log_handler.setFormatter(logging.Formatter('%(asctime)s - %(message)s'))

        root_logger = logging.getLogger()
        root_logger.setLevel(logging.INFO)
        root_logger.addHandler(self.file_log_handler)
        root_logger.addHandler(self.gui_log_handler)

        logging.info("应用程序已启动。")

    def cleanup_logging(self):
        root_logger = logging.getLogger()
        for handler_name in ("gui_log_handler", "file_log_handler"):
            handler = getattr(self, handler_name, None)
            if handler is None:
                continue
            if handler in root_logger.handlers:
                root_logger.removeHandler(handler)
            handler.close()
            setattr(self, handler_name, None)

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
        self.btn_view_mcu = QPushButton("电控转速-扭矩-效率MAP")
        self.btn_view_motor = QPushButton("电机转速-扭矩-效率MAP")
        self.btn_view_sys = QPushButton("系统转速-扭矩-效率MAP")
        self.btn_view_external = QPushButton("外特性")
        self.chk_external_values = QCheckBox("显示数值")
        self.chk_external_values.setChecked(True)
        self.btn_view_mcu_speed_power = QPushButton("电控转速-功率-效率MAP")
        self.btn_view_motor_speed_power = QPushButton("电机转速-功率-效率MAP")
        self.btn_view_sys_speed_power = QPushButton("系统转速-功率-效率MAP")
        self.btn_view_mcu_loss = QPushButton("电控损耗MAP")
        self.btn_view_motor_loss = QPushButton("电机损耗MAP")
        self.btn_view_sys_loss = QPushButton("系统损耗MAP")

        self.btn_view_mcu.clicked.connect(lambda: self.show_map_plot('MCU'))
        self.btn_view_motor.clicked.connect(lambda: self.show_map_plot('Motor'))
        self.btn_view_sys.clicked.connect(lambda: self.show_map_plot('SYS'))
        self.btn_view_external.clicked.connect(self.show_external_characteristics_plot)
        self.btn_view_mcu_speed_power.clicked.connect(lambda: self.show_speed_power_efficiency_plot('MCU'))
        self.btn_view_motor_speed_power.clicked.connect(lambda: self.show_speed_power_efficiency_plot('Motor'))
        self.btn_view_sys_speed_power.clicked.connect(lambda: self.show_speed_power_efficiency_plot('SYS'))
        self.btn_view_mcu_loss.clicked.connect(lambda: self.show_loss_map_plot('MCU'))
        self.btn_view_motor_loss.clicked.connect(lambda: self.show_loss_map_plot('Motor'))
        self.btn_view_sys_loss.clicked.connect(lambda: self.show_loss_map_plot('SYS'))

        vis_layout.addWidget(self.btn_view_mcu)
        vis_layout.addWidget(self.btn_view_motor)
        vis_layout.addWidget(self.btn_view_sys)
        external_layout = QHBoxLayout()
        external_layout.addWidget(self.btn_view_external)
        external_layout.addWidget(self.chk_external_values)
        vis_layout.addLayout(external_layout)
        vis_layout.addWidget(self.btn_view_mcu_speed_power)
        vis_layout.addWidget(self.btn_view_motor_speed_power)
        vis_layout.addWidget(self.btn_view_sys_speed_power)
        vis_layout.addWidget(self.btn_view_mcu_loss)
        vis_layout.addWidget(self.btn_view_motor_loss)
        vis_layout.addWidget(self.btn_view_sys_loss)

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

        self.config_fields = {}
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

    def get_efficiency_output_by_short(self, eff_type_short):
        output_by_plot_key = {
            plot_key: (eff_type, switch_key, plot_key, display_name)
            for eff_type, switch_key, plot_key, display_name in self.EFFICIENCY_MAP_OUTPUTS
        }
        return output_by_plot_key.get(eff_type_short)

    def get_efficiency_map_cache_suffix(self, eff_type_short):
        output = self.get_efficiency_output_by_short(eff_type_short)
        if output is None:
            return None
        return f"{output[3]}EfficiencyMAP"

    def get_standard_plot_title(self, map_name_cn):
        veh_code = self.config_dict.get('VehicleCode', '')
        udc_val = "0"
        if self.logic is not None and self.logic.u_dc is not None:
             udc_val = str(int(round(self.logic.u_dc.mean())))

        direction = self.current_results.get('direction', '') if hasattr(self, 'current_results') else ''
        state = self.current_results.get('state', '') if hasattr(self, 'current_results') else ''
        if state == "电动":
            state = "驱动"

        return f"{veh_code}-{udc_val}V-{direction}{state}-{map_name_cn}"

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

                keys = self.get_config_section_keys(self.raw_config_obj, section)

                for key in keys:
                    if key == "__name__":
                        continue
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
        self.plot_image_cache = {}
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

    def closeEvent(self, event):
        worker = getattr(self, "_batch_worker", None)
        thread = getattr(self, "_batch_thread", None)
        if worker is not None and thread is not None and thread.isRunning():
            worker.request_stop()
            thread.quit()
            if not thread.wait(3000):
                self._close_after_batch = True
                event.ignore()
                logging.info("批处理仍在结束中，已延后关闭窗口。")
                return
        self.cleanup_logging()
        super().closeEvent(event)


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
