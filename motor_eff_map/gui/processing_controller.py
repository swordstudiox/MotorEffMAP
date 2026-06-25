import logging
import os

from PySide6.QtCore import QThread
from PySide6.QtWidgets import QFileDialog, QMessageBox

from motor_eff_map.gui.batch_worker import BatchWorker


class ProcessingControllerMixin:
    def select_files(self):
        if getattr(self, "is_batch_running", False):
            return

        files, _ = QFileDialog.getOpenFileNames(self, "选择数据文件", "", "Excel Files (*.xls *.xlsx)")
        if files:
            self.data_files = files
            self.plot_image_cache = {}
            self.list_widget.clear()
            self.all_results = []

            for fpath in files:
                success, msg = self.logic.load_data(fpath)
                if success:
                    fname = os.path.basename(fpath)
                    for sheet_name in self.logic.sheets_dict.keys():
                        disp_name = f"{fname} - {sheet_name}"
                        self.all_results.append({
                            "file": fpath,
                            "sheet": sheet_name,
                            "disp": disp_name,
                        })
                        self.list_widget.addItem(disp_name)
                    logging.info(f"已加载 {fname}: {len(self.logic.sheets_dict)} 个 sheets。")
                else:
                    logging.error(f"加载 {fpath} 失败: {msg}")

    def on_list_selection(self, row):
        if getattr(self, "is_batch_running", False):
            return
        if row < 0 or row >= len(self.all_results):
            return

        item = self.all_results[row]
        self.process_result_item(item)

    def process_result_item(self, item, update_progress=True):
        fpath = item["file"]
        sheet = item["sheet"]

        logging.info(f"已选择: {item['disp']}")

        if getattr(self.logic, "current_file", None) != fpath:
            success, msg = self.logic.load_data(fpath)
            if not success:
                self.handle_processing_error(f"加载 {fpath} 失败: {msg}")
                return False

        if not self.logic.set_current_sheet(sheet):
            self.handle_processing_error(f"未找到工作表: {sheet}")
            return False

        return self.process_current_data(update_progress=update_progress)

    def get_selected_result_item(self):
        row = -1
        list_widget = getattr(self, "list_widget", None)
        if list_widget is not None:
            row = list_widget.currentRow()
        if row < 0 and len(getattr(self, "all_results", [])) == 1:
            row = 0
        if row < 0 or row >= len(getattr(self, "all_results", [])):
            return None
        return self.all_results[row]

    def prepare_selected_result_cache_context(self):
        item = self.get_selected_result_item()
        if item is None:
            return None

        current = dict(getattr(self, "current_results", {}) or {})
        source_file = item["file"]
        sheet = item["sheet"]
        current.update({
            "source_file": source_file,
            "file": os.path.basename(source_file),
            "sheet": sheet,
        })
        self.current_results = current
        return item

    def show_cached_plot_for_selected_result(self, suffix):
        self.prepare_selected_result_cache_context()
        return self.show_cached_plot(suffix)

    def get_cached_plot_suffixes(self):
        suffixes = []
        for _, _, plot_key, display_name in self.EFFICIENCY_MAP_OUTPUTS:
            cache_suffix = self.get_efficiency_map_cache_suffix(plot_key)
            if cache_suffix:
                suffixes.append(cache_suffix)
        suffixes.append("外特性曲线")
        for _, _, _, display_name in self.EFFICIENCY_MAP_OUTPUTS:
            suffixes.append(f"{display_name}SpeedPowerEfficiencyMAP")
        for _, _, _, display_name in self.EFFICIENCY_MAP_OUTPUTS:
            suffixes.append(f"{display_name}LossMAP")
        suffixes.append("效率占比")
        return suffixes

    def show_first_cached_plot(self):
        if not getattr(self, "plot_image_cache", {}):
            return False
        if not getattr(self, "all_results", []):
            return False

        list_widget = getattr(self, "list_widget", None)
        if list_widget is not None and list_widget.currentRow() < 0:
            can_block = hasattr(list_widget, "blockSignals")
            try:
                if can_block:
                    list_widget.blockSignals(True)
                list_widget.setCurrentRow(0)
            finally:
                if can_block:
                    list_widget.blockSignals(False)

        item = self.get_selected_result_item()
        if item is None:
            item = self.all_results[0]
            current = dict(getattr(self, "current_results", {}) or {})
            source_file = item["file"]
            current.update({
                "source_file": source_file,
                "file": os.path.basename(source_file),
                "sheet": item["sheet"],
            })
            self.current_results = current

        for suffix in self.get_cached_plot_suffixes():
            if self.show_cached_plot_for_selected_result(suffix):
                return True
        return False

    def is_processed_data_for_item(self, item):
        if self.logic is None or getattr(self.logic, "processed_df", None) is None:
            return False
        current_file = getattr(self.logic, "current_file", "")
        current_sheet = getattr(self.logic, "current_sheet", "")
        if not current_file:
            return False
        file_matches = os.path.normcase(os.path.abspath(current_file)) == os.path.normcase(os.path.abspath(item["file"]))
        sheet_matches = not current_sheet or current_sheet == item["sheet"]
        return file_matches and sheet_matches

    def ensure_selected_result_ready(self, update_progress=False):
        item = self.get_selected_result_item()
        if self.logic is not None and getattr(self.logic, "processed_df", None) is not None and item is None:
            return True

        if item is not None and self.is_processed_data_for_item(item):
            return True

        if item is None:
            return False

        return self.process_result_item(item, update_progress=update_progress)

    def run_process_all(self):
        """批量处理所有加载的 sheet 并保存结果。"""
        if getattr(self, "is_batch_running", False):
            return
        if not self.all_results:
            QMessageBox.warning(self, "无数据", "请先加载文件。")
            return

        try:
            self.reload_runtime_config()
        except Exception as e:
            logging.error(f"重新读取配置失败: {e}")
            QMessageBox.critical(self, "配置错误", f"重新读取配置失败: {e}")
            return

        self.plot_image_cache = {}
        self.figure.clear()
        self.canvas.draw()
        self.progress_bar.setValue(0)
        self.set_batch_running(True)
        self._batch_error = ""

        worker = self._create_batch_worker()
        self._start_batch_worker(worker)

    def _create_batch_worker(self):
        show_external_values = True
        if hasattr(self, "chk_external_values"):
            show_external_values = self.chk_external_values.isChecked()
        return BatchWorker(self.all_results, self.config_dict, show_external_values=show_external_values)

    def _start_batch_worker(self, worker):
        thread = QThread(self)
        self._batch_thread = thread
        self._batch_worker = worker
        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.progress_changed.connect(self._on_batch_progress)
        worker.item_changed.connect(self._on_batch_item_changed)
        worker.plot_cached.connect(self._on_batch_plot_cached)
        worker.failed.connect(self._on_batch_failed)
        worker.finished.connect(self._on_batch_finished)
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)
        thread.finished.connect(self._on_batch_thread_finished)
        thread.start()

    def _on_batch_progress(self, percent):
        self.progress_bar.setValue(int(percent))

    def _on_batch_item_changed(self, display_name):
        if display_name:
            logging.info(f"正在处理: {display_name}")

    def _on_batch_failed(self, message):
        self._batch_error = message
        logging.error(message)

    def _on_batch_plot_cached(self, entries):
        if not hasattr(self, "plot_image_cache"):
            self.plot_image_cache = {}
        self.plot_image_cache.update(dict(entries or {}))

    def _on_batch_finished(self, success):
        self.set_batch_running(False)
        if success:
            self.progress_bar.setValue(100)
            logging.info("批处理完成。")
            self.show_first_cached_plot()
        else:
            self.progress_bar.setValue(0)
            message = getattr(self, "_batch_error", "") or "批处理失败。"
            if not getattr(self, "_close_after_batch", False):
                QMessageBox.warning(self, "处理错误", message)

    def _on_batch_thread_finished(self):
        self._batch_thread = None
        self._batch_worker = None
        if getattr(self, "_close_after_batch", False):
            self._close_after_batch = False
            self.close()

    def set_batch_running(self, running):
        self.is_batch_running = bool(running)
        for name in ("btn_process", "btn_load", "btn_save_config", "btn_open_ini"):
            widget = getattr(self, name, None)
            if widget is not None:
                widget.setEnabled(not running)
        list_widget = getattr(self, "list_widget", None)
        if list_widget is not None:
            list_widget.setEnabled(not running)

    def _set_processing_progress(self, percent, update_progress=True):
        if update_progress:
            self.progress_bar.setValue(int(percent))

    def process_current_data(self, update_progress=True):
        self._set_processing_progress(10, update_progress)

        direction, state = self.logic.filter_data()
        if direction is None:
            detail = getattr(self.logic, "last_error", "") or "请检查配置。"
            logging.error(f"映射数据列失败。{detail}")
            self._set_processing_progress(0, update_progress)
            return False

        if state == "电动":
            state = "驱动"

        self.logic.normalization()
        self._set_processing_progress(30, update_progress)

        self.logic.get_external_characteristics()
        self._set_processing_progress(50, update_progress)

        self.current_results = {
            "source_file": self.logic.current_file,
            "file": os.path.basename(self.logic.current_file),
            "sheet": getattr(self.logic, "current_sheet", ""),
            "direction": direction,
            "state": state,
        }

        for eff_type, switch_key, plot_key, display_name in self.EFFICIENCY_MAP_OUTPUTS:
            if self.should_use_efficiency_output(eff_type, switch_key, f"{display_name}效率MAP"):
                if not self.switch_plot(plot_key, save_png=True):
                    return False

        if self.is_config_switch_on("ExternalCharacteristicPlot"):
            if not self.show_external_characteristics_plot(save_png=True):
                return False

        if self.is_config_switch_on("SpeedPowerMAP"):
            for eff_type, _, plot_key, display_name in self.EFFICIENCY_MAP_OUTPUTS:
                if self.logic.has_efficiency_data(eff_type):
                    if not self.show_speed_power_efficiency_plot(plot_key, save_png=True):
                        return False
                else:
                    logging.warning(f"{display_name}转速-功率效率MAP输出已开启，但 {eff_type} 未配置或没有有效数据，已跳过。")

        if self.is_config_switch_on("LossMAP"):
            for eff_type, _, plot_key, display_name in self.EFFICIENCY_MAP_OUTPUTS:
                if self.logic.has_efficiency_data(eff_type):
                    if not self.show_loss_map_plot(plot_key, save_png=True):
                        return False
                else:
                    logging.warning(f"{display_name}损耗MAP输出已开启，但 {eff_type} 未配置或没有有效数据，已跳过。")

        self._set_processing_progress(80, update_progress)

        try:
            self.process_area_ratios()
        except ValueError as e:
            self.handle_processing_error(str(e))
            return False

        self._set_processing_progress(100, update_progress)
        return True
