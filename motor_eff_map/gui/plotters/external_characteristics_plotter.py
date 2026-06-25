import logging

import numpy as np
from PySide6.QtWidgets import QMessageBox


class ExternalCharacteristicsPlotterMixin:
    def show_external_characteristics_plot(self, save_png=False):
        """绘制外特性曲线：转速-扭矩与转速-功率。"""
        if not save_png:
            show_cached = getattr(self, "show_cached_plot_for_selected_result", self.show_cached_plot)
            if show_cached("外特性曲线"):
                return True

            try:
                self.reload_runtime_config()
            except Exception as e:
                logging.error(f"重新读取配置失败: {e}")
                QMessageBox.critical(self, "配置错误", f"重新读取配置失败: {e}")
                return False

        elif self.logic is None or self.logic.processed_df is None:
            return False

        if not self.is_config_switch_on('ExternalCharacteristicPlot'):
            self.figure.clear()
            self.canvas.draw()
            logging.info("外特性曲线图未启用，已按配置跳过显示。")
            return False

        if not save_png:
            if show_cached("外特性曲线"):
                return True

            if hasattr(self, "ensure_selected_result_ready") and not self.ensure_selected_result_ready():
                QMessageBox.warning(self, "无数据", "请先选择要查看的 sheet。")
                return False
            if show_cached("外特性曲线"):
                return True

        curve = self.logic.get_external_characteristics_data()
        if curve.empty:
            logging.warning("外特性曲线无有效数据。")
            return False

        self.figure.clear()
        self.apply_figure_layout(layout=self.DUAL_Y_FIGURE_LAYOUT)
        ax_torque = self.figure.add_subplot(111)
        ax_power = ax_torque.twinx()

        speed = curve['Speed'].to_numpy(dtype=float)
        torque = curve['Torque'].to_numpy(dtype=float)
        power = curve['P_Motor'].to_numpy(dtype=float)

        line_torque, = ax_torque.plot(speed, torque, '-ob', linewidth=1.5, markersize=4, label='扭矩')
        line_power, = ax_power.plot(speed, power, '-sr', linewidth=1.5, markersize=4, label='功率')

        show_values = self.chk_external_values.isChecked()
        if show_values:
            for x, y in zip(speed, torque):
                ax_torque.annotate(f"{y:.1f}", (x, y), textcoords="offset points", xytext=(0, 6),
                                   ha='center', fontsize=7, color='blue')
            for x, y in zip(speed, power):
                ax_power.annotate(f"{y:.1f}", (x, y), textcoords="offset points", xytext=(0, -12),
                                  ha='center', fontsize=7, color='red')

        try:
            speed_step = self.get_positive_config_float('xstepSpeed', 500)
            if len(speed) > 0 and speed_step > 0:
                end_x = int(np.ceil(np.nanmax(speed) / speed_step)) * speed_step
                ax_torque.set_xticks(np.arange(0, end_x + speed_step/2, speed_step))
                ax_torque.set_xlim(0, end_x)
        except:
            pass

        ax_torque.set_xlabel('转速 [rpm]', fontsize=9)
        ax_torque.set_ylabel('扭矩 [N.m]', fontsize=9, color='blue')
        ax_power.set_ylabel('功率 [kW]', fontsize=9, color='red')
        ax_torque.tick_params(axis='both', which='major', labelsize=8)
        ax_power.tick_params(axis='y', which='major', labelsize=8, colors='red')
        ax_torque.tick_params(axis='y', colors='blue')
        ax_torque.grid(True, linestyle=':', alpha=0.6)
        ax_torque.set_title(self.get_standard_plot_title("外特性曲线图"), fontsize=12)
        ax_torque.legend([line_torque, line_power], ['扭矩', '功率'], loc='upper right', fontsize=8)

        self.canvas.draw()
        if save_png:
            return self.save_current_figure("外特性曲线", layout=self.DUAL_Y_FIGURE_LAYOUT)
        return True
