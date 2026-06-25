import logging

import numpy as np
from PySide6.QtWidgets import QMessageBox


class LossMapPlotterMixin:
    def show_loss_map_plot(self, eff_type_short, save_png=False):
        """绘制转速-扭矩-损耗 MAP。"""
        if not save_png:
            output = self.get_efficiency_output_by_short(eff_type_short)
            if output is None:
                return False

            _, _, _, display_name = output
            show_cached = getattr(self, "show_cached_plot_for_selected_result", self.show_cached_plot)
            if show_cached(f"{display_name}LossMAP"):
                return True

            try:
                self.reload_runtime_config()
            except Exception as e:
                logging.error(f"重新读取配置失败: {e}")
                QMessageBox.critical(self, "配置错误", f"重新读取配置失败: {e}")
                return False

        elif self.logic is None or self.logic.processed_df is None:
            return False

        if not self.is_config_switch_on('LossMAP'):
            self.figure.clear()
            self.canvas.draw()
            logging.info("损耗MAP未启用，已按配置跳过显示。")
            return False

        output = self.get_efficiency_output_by_short(eff_type_short)
        if output is None:
            return False

        eff_type, _, _, display_name = output
        if not save_png:
            if show_cached(f"{display_name}LossMAP"):
                return True

            if hasattr(self, "ensure_selected_result_ready") and not self.ensure_selected_result_ready():
                QMessageBox.warning(self, "无数据", "请先选择要查看的 sheet。")
                return False
            if show_cached(f"{display_name}LossMAP"):
                return True

        if not self.logic.has_efficiency_data(eff_type):
            logging.warning(f"{display_name}损耗MAP输出已开启，但 {eff_type} 未配置或没有有效数据，已跳过。")
            return False

        try:
            res = self.logic.process_torque_loss_map_data(eff_type)
        except ValueError as e:
            self.handle_processing_error(str(e))
            return False

        if res is None:
            logging.warning(f"无法处理 {eff_type} 的损耗MAP")
            return False

        XI, YI, ZI_Loss, _ = res
        self.figure.clear()
        self.apply_figure_layout()
        ax = self.figure.add_subplot(111)

        try:
            loss_levels = self.parse_contour_levels('LossMAPStep', 12)
            x_grid, y_grid, z_grid = self.prepare_masked_contour_grid(XI, YI, ZI_Loss)
            if z_grid.count() < 3:
                raise ValueError("有效绘图点少于 3 个，无法生成损耗MAP。")

            cf = ax.contourf(x_grid, y_grid, z_grid, levels=loss_levels, cmap='jet')
            if isinstance(loss_levels, (list, tuple, np.ndarray)):
                z_min = float(z_grid.min())
                z_max = float(z_grid.max())
                line_levels = [
                    l for l in loss_levels
                    if z_min <= l <= z_max
                ]
            else:
                line_levels = loss_levels
            if line_levels:
                ce = ax.contour(x_grid, y_grid, z_grid, levels=line_levels, colors='k', linewidths=0.5)
                ax.clabel(ce, inline=True, fontsize=8, fmt='%1.0f')
            self.add_contour_legend(ax, [
                {
                    "label": "损耗",
                    "color": "k",
                    "linewidth": 0.5,
                    "marker": "o",
                    "markerfacecolor": "none",
                    "markeredgecolor": "k",
                },
            ])
            max_x = float(np.nanmax(XI))
            max_y = float(np.nanmax(YI))
            self.apply_axis_ticks(
                ax,
                max_x,
                max_y,
                'xstepSpeed',
                'ystepTorque',
                500,
                20,
                start_x=self.get_non_negative_config_float('StartSpeed', 0),
                start_y=self.get_non_negative_config_float('StartTorque', 0),
            )
        except Exception as e:
            logging.error(f"损耗MAP生成失败: {e}")
            self.handle_processing_error(f"损耗MAP生成失败: {e}")
            return False

        ax.set_xlabel('转速 [rpm]', fontsize=9)
        ax.set_ylabel('扭矩 [N.m]', fontsize=9)
        ax.tick_params(axis='both', which='major', labelsize=8)
        ax.set_title(self.get_standard_plot_title(f"{display_name}损耗MAP"), fontsize=12)
        ax.grid(True, linestyle=':', alpha=0.6)
        self.canvas.draw()

        if save_png:
            return self.save_current_figure(f"{display_name}LossMAP")
        return True
