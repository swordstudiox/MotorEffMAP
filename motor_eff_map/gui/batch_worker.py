import logging
from pathlib import Path

import pandas as pd
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.figure import Figure
from PySide6.QtCore import QObject, Signal

from motor_eff_map.gui.config_schema import (
    DUAL_Y_FIGURE_LAYOUT,
    EFFICIENCY_MAP_OUTPUTS,
    EFFICIENCY_RATIO_OUTPUTS,
    EXPORT_ASPECT_RATIO,
    EXPORT_FIGURE_SIZE,
    FIGURE_LAYOUT,
)
from motor_eff_map.gui.output_naming import OutputNamingMixin
from motor_eff_map.gui.plot_helpers import PlotHelperMixin
from motor_eff_map.gui.plotters import (
    EfficiencyMapPlotterMixin,
    ExternalCharacteristicsPlotterMixin,
    LossMapPlotterMixin,
    RatioPlotterMixin,
    SpeedPowerPlotterMixin,
)
from motor_eff_map.logic import MotorEffLogic


class _BoolFlag:
    def __init__(self, value):
        self._value = bool(value)

    def isChecked(self):
        return self._value


class BatchExportContext(
    PlotHelperMixin,
    OutputNamingMixin,
    RatioPlotterMixin,
    ExternalCharacteristicsPlotterMixin,
    SpeedPowerPlotterMixin,
    LossMapPlotterMixin,
    EfficiencyMapPlotterMixin,
):
    EXPORT_FIGURE_SIZE = EXPORT_FIGURE_SIZE
    EXPORT_ASPECT_RATIO = EXPORT_ASPECT_RATIO
    FIGURE_LAYOUT = FIGURE_LAYOUT
    DUAL_Y_FIGURE_LAYOUT = DUAL_Y_FIGURE_LAYOUT
    EFFICIENCY_MAP_OUTPUTS = EFFICIENCY_MAP_OUTPUTS
    EFFICIENCY_RATIO_OUTPUTS = EFFICIENCY_RATIO_OUTPUTS

    def __init__(self, config_dict, show_external_values=True):
        self.config_dict = dict(config_dict)
        self.logic = MotorEffLogic(self.config_dict)
        self.current_results = {}
        self.plot_image_cache = {}
        self.figure = Figure(figsize=self.EXPORT_FIGURE_SIZE, dpi=100)
        self.canvas = FigureCanvasAgg(self.figure)
        self.chk_external_values = _BoolFlag(show_external_values)
        self.last_error = ""

    def reload_runtime_config(self):
        return None

    def is_config_switch_on(self, key):
        return self.config_dict.get(key, "0") == "1"

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
        veh_code = self.config_dict.get("VehicleCode", "")
        udc_val = "0"
        if self.logic is not None and self.logic.u_dc is not None:
            udc_val = str(int(round(self.logic.u_dc.mean())))

        direction = self.current_results.get("direction", "")
        state = self.current_results.get("state", "")
        if state == "电动":
            state = "驱动"

        return f"{veh_code}-{udc_val}V-{direction}{state}-{map_name_cn}"

    def handle_processing_error(self, message):
        self.last_error = str(message)
        logging.error(message)

    def _report_progress(self, progress_callback, percent):
        if progress_callback is not None:
            progress_callback(int(percent))

    def process_item(self, item, progress_callback=None):
        fpath = item["file"]
        sheet = item["sheet"]
        if getattr(self.logic, "current_file", None) != fpath:
            success, msg = self.logic.load_data(fpath)
            if not success:
                raise ValueError(f"加载 {fpath} 失败: {msg}")

        if not self.logic.set_current_sheet(sheet):
            raise ValueError(f"未找到工作表: {sheet}")

        self._report_progress(progress_callback, 5)

        direction, state = self.logic.filter_data()
        if direction is None:
            detail = getattr(self.logic, "last_error", "") or "请检查配置。"
            raise ValueError(f"映射数据列失败。{detail}")

        if state == "电动":
            state = "驱动"
        self._report_progress(progress_callback, 15)

        self.logic.normalization()
        self._report_progress(progress_callback, 25)

        self.logic.get_external_characteristics()
        self._report_progress(progress_callback, 35)

        self.current_results = {
            "source_file": self.logic.current_file,
            "file": Path(self.logic.current_file).name,
            "sheet": getattr(self.logic, "current_sheet", ""),
            "direction": direction,
            "state": state,
        }
        self._report_progress(progress_callback, 40)

        for eff_type, switch_key, plot_key, display_name in self.EFFICIENCY_MAP_OUTPUTS:
            if self.should_use_efficiency_output(eff_type, switch_key, f"{display_name}效率MAP"):
                if not self.switch_plot(plot_key, save_png=True):
                    raise ValueError(f"{display_name}效率MAP保存失败")
        self._report_progress(progress_callback, 55)

        if self.is_config_switch_on("ExternalCharacteristicPlot"):
            if not self.show_external_characteristics_plot(save_png=True):
                raise ValueError("外特性曲线图保存失败")
        self._report_progress(progress_callback, 65)

        if self.is_config_switch_on("SpeedPowerMAP"):
            for eff_type, _, plot_key, display_name in self.EFFICIENCY_MAP_OUTPUTS:
                if self.logic.has_efficiency_data(eff_type):
                    if not self.show_speed_power_efficiency_plot(plot_key, save_png=True):
                        raise ValueError(f"{display_name}转速-功率效率MAP保存失败")
                else:
                    logging.warning(f"{display_name}转速-功率效率MAP输出已开启，但 {eff_type} 未配置或没有有效数据，已跳过。")
        self._report_progress(progress_callback, 78)

        if self.is_config_switch_on("LossMAP"):
            for eff_type, _, plot_key, display_name in self.EFFICIENCY_MAP_OUTPUTS:
                if self.logic.has_efficiency_data(eff_type):
                    if not self.show_loss_map_plot(plot_key, save_png=True):
                        raise ValueError(f"{display_name}损耗MAP保存失败")
                else:
                    logging.warning(f"{display_name}损耗MAP输出已开启，但 {eff_type} 未配置或没有有效数据，已跳过。")
        self._report_progress(progress_callback, 90)

        self.process_area_ratios()
        self._report_progress(progress_callback, 95)
        return True

    def process_area_ratios(self):
        active_ratios = self.get_active_ratio_outputs()
        calc_mcu = active_ratios["Eff_MCU"]
        calc_motor = active_ratios["Eff_Motor"]
        calc_sys = active_ratios["Eff_SYS"]

        if not (calc_mcu or calc_motor or calc_sys):
            return None

        veh_code = self.config_dict.get("VehicleCode", "")
        udc_val = "0"
        if self.logic.u_dc is not None:
            udc_val = str(int(round(self.logic.u_dc.mean())))

        direction = self.current_results.get("direction", "")
        state = self.current_results.get("state", "")
        if state == "电动":
            state = "驱动"

        base_name = self.build_output_stem("效率占比")
        excel_name = f"{base_name}.xlsx"
        plot_name = f"{base_name}.png"
        title_str = f"{veh_code}-{udc_val}V-{direction}{state}-效率区域占比图"

        mcu_ratios, motor_ratios, sys_ratios = self._collect_ratio_data(active_ratios)

        row_labels = []
        if mcu_ratios:
            row_labels = [f"≥{r['Level']}" for r in mcu_ratios]
        elif motor_ratios:
            row_labels = [f"≥{r['Level']}" for r in motor_ratios]
        elif sys_ratios:
            row_labels = [f"≥{r['Level']}" for r in sys_ratios]

        if not row_labels:
            return None

        df = pd.DataFrame({"效率区间": row_labels})
        if calc_mcu and mcu_ratios:
            df["控制器效率占比"] = [r["Ratio"] for r in mcu_ratios]
        if calc_motor and motor_ratios:
            df["电机效率占比"] = [r["Ratio"] for r in motor_ratios]
        if calc_sys and sys_ratios:
            df["系统效率占比"] = [r["Ratio"] for r in sys_ratios]

        try:
            with pd.ExcelWriter(excel_name, engine="openpyxl") as writer:
                header_df = pd.DataFrame(columns=["效率区间", "控制器效率占比", "电机效率占比", "系统效率占比"])
                header_df.to_excel(writer, sheet_name="数据占比", index=False, startrow=0)
                df.to_excel(writer, sheet_name="数据占比", index=False, header=False, startrow=2)
            logging.info(f"保存区域占比 Excel: {excel_name}")
        except Exception as e:
            logging.error(f"保存 Excel {excel_name} 失败: {e}")
            raise

        fig = Figure(figsize=self.EXPORT_FIGURE_SIZE, dpi=100)
        self.apply_figure_layout(fig)
        ax = fig.add_subplot(111)
        self._plot_ratio_on_axes(ax, mcu_ratios, motor_ratios, sys_ratios, title_str)

        try:
            fig.savefig(plot_name, dpi=200)
            self.register_plot_cache("效率占比", plot_name)
            logging.info(f"保存占比图: {plot_name}")
        except Exception as e:
            logging.error(f"保存图表 {plot_name} 失败: {e}")
            raise

        return True


class BatchWorker(QObject):
    progress_changed = Signal(int)
    item_changed = Signal(str)
    plot_cached = Signal(object)
    log_message = Signal(str)
    finished = Signal(bool)
    failed = Signal(str)

    def __init__(self, items, config_dict, show_external_values=True):
        super().__init__()
        self.items = list(items)
        self.context = BatchExportContext(config_dict, show_external_values=show_external_values)
        self._stopped = False
        self._current_index = 0
        self._total = len(self.items)

    def request_stop(self):
        self._stopped = True

    def _emit_item_progress(self, item_percent):
        total = self._total or 1
        item_percent = max(0, min(99, int(item_percent)))
        percent = int((self._current_index + item_percent / 100) / total * 100)
        self.progress_changed.emit(percent)

    def _process_item_with_progress(self, item):
        try:
            return self.context.process_item(item, progress_callback=self._emit_item_progress)
        except TypeError as exc:
            if "progress_callback" not in str(exc):
                raise
            return self.context.process_item(item)

    def run(self):
        total = len(self.items)
        self._total = total
        try:
            for index, item in enumerate(self.items):
                if self._stopped:
                    self.finished.emit(False)
                    return
                self._current_index = index
                self.item_changed.emit(item.get("disp", ""))
                before_cache = dict(getattr(self.context, "plot_image_cache", {}))
                self._process_item_with_progress(item)
                after_cache = dict(getattr(self.context, "plot_image_cache", {}))
                new_cache = {
                    key: value
                    for key, value in after_cache.items()
                    if before_cache.get(key) != value
                }
                if new_cache:
                    self.plot_cached.emit(new_cache)
                percent = int((index + 1) / total * 100) if total else 100
                self.progress_changed.emit(percent)
            self.finished.emit(True)
        except Exception as exc:
            self.failed.emit(str(exc))
            self.finished.emit(False)
