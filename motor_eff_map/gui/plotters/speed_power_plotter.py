import logging

import numpy as np
from PySide6.QtWidgets import QMessageBox


class SpeedPowerPlotterMixin:
    def show_speed_power_efficiency_plot(self, eff_type_short, save_png=False):
        """绘制转速-功率-效率 MAP。"""
        if not save_png:
            output = self.get_efficiency_output_by_short(eff_type_short)
            if output is None:
                return False

            _, _, _, display_name = output
            show_cached = getattr(self, "show_cached_plot_for_selected_result", self.show_cached_plot)
            if show_cached(f"{display_name}SpeedPowerEfficiencyMAP"):
                return True

            try:
                self.reload_runtime_config()
            except Exception as e:
                logging.error(f"重新读取配置失败: {e}")
                QMessageBox.critical(self, "配置错误", f"重新读取配置失败: {e}")
                return False

        elif self.logic is None or self.logic.processed_df is None:
            return False

        if not self.is_config_switch_on('SpeedPowerMAP'):
            self.figure.clear()
            self.canvas.draw()
            logging.info("转速-功率效率MAP未启用，已按配置跳过显示。")
            return False

        output = self.get_efficiency_output_by_short(eff_type_short)
        if output is None:
            return False

        eff_type, _, _, display_name = output
        if not save_png:
            if show_cached(f"{display_name}SpeedPowerEfficiencyMAP"):
                return True

            if hasattr(self, "ensure_selected_result_ready") and not self.ensure_selected_result_ready():
                QMessageBox.warning(self, "无数据", "请先选择要查看的 sheet。")
                return False
            if show_cached(f"{display_name}SpeedPowerEfficiencyMAP"):
                return True

        if not self.logic.has_efficiency_data(eff_type):
            logging.warning(f"{display_name}转速-功率效率MAP输出已开启，但 {eff_type} 未配置或没有有效数据，已跳过。")
            return False

        try:
            res = self.logic.process_speed_power_efficiency_map_data(eff_type)
        except ValueError as e:
            self.handle_processing_error(str(e))
            return False

        if res is None:
            logging.warning(f"无法处理 {eff_type} 的转速-功率效率MAP")
            return False

        XI, YI, ZI_Eff, _ = res
        self.figure.clear()
        self.apply_figure_layout()
        ax = self.figure.add_subplot(111)

        try:
            eff_levels = self.parse_contour_levels('EffMAPStep', [70, 80, 85, 90, 95, 100])
            if eff_levels and eff_levels[-1] < 100:
                eff_levels.append(100.0)
            x_grid, y_grid, z_grid = self.prepare_masked_contour_grid(XI, YI, ZI_Eff)
            if z_grid.count() < 3:
                raise ValueError("有效绘图点少于 3 个，无法生成转速-功率效率MAP。")

            cf = ax.contourf(x_grid, y_grid, z_grid, levels=eff_levels, cmap='jet')
            z_min = float(z_grid.min())
            z_max = float(z_grid.max())
            line_levels = [l for l in eff_levels if l < 100 and z_min <= l <= z_max]
            if line_levels:
                ce = ax.contour(x_grid, y_grid, z_grid, levels=line_levels, colors='k', linewidths=0.5)
                ax.clabel(ce, levels=line_levels, inline=True, fontsize=8, fmt='%1.0f')
            self.add_contour_legend(ax, [
                {
                    "label": "效率",
                    "color": "k",
                    "linewidth": 0.5,
                    "marker": "o",
                    "markerfacecolor": "#000080",
                    "markeredgecolor": "k",
                },
            ])
            start_x = self.get_non_negative_config_float('StartSpeed', 0)
            if hasattr(self.logic, 'get_start_power_cutoff'):
                start_y = self.logic.get_start_power_cutoff()
            else:
                start_y = self.get_non_negative_config_float('StartPower', 0)
            self.apply_axis_ticks(
                ax,
                float(np.nanmax(XI)),
                float(np.nanmax(YI)),
                'xstepSpeed',
                'ystepPower',
                500,
                20,
                start_x=start_x,
                start_y=start_y,
            )
        except Exception as e:
            logging.error(f"转速-功率效率MAP生成失败: {e}")
            self.handle_processing_error(f"转速-功率效率MAP生成失败: {e}")
            return False

        ax.set_xlabel('转速 [rpm]', fontsize=9)
        ax.set_ylabel('功率 [kW]', fontsize=9)
        ax.tick_params(axis='both', which='major', labelsize=8)
        ax.set_title(self.get_standard_plot_title(f"{display_name}转速-功率效率MAP"), fontsize=12)
        ax.grid(True, linestyle=':', alpha=0.6)
        self.canvas.draw()

        if save_png:
            return self.save_current_figure(f"{display_name}SpeedPowerEfficiencyMAP")
        return True
