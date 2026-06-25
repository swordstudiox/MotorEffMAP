import logging

import pandas as pd
from matplotlib.figure import Figure
from PySide6.QtWidgets import QMessageBox


class RatioPlotterMixin:
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

    def process_area_ratios(self):
        """计算区域占比，保存到 Excel，并绘制占比图。"""
        active_ratios = self.get_active_ratio_outputs()
        calc_mcu = active_ratios['Eff_MCU']
        calc_motor = active_ratios['Eff_Motor']
        calc_sys = active_ratios['Eff_SYS']

        if not (calc_mcu or calc_motor or calc_sys):
            return

        veh_code = self.config_dict.get('VehicleCode', '')

        udc_val = "0"
        if self.logic.u_dc is not None:
            udc_val = str(int(round(self.logic.u_dc.mean())))

        direction = self.current_results.get('direction', '')
        state = self.current_results.get('state', '')
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
            return

        df = pd.DataFrame({'效率区间': row_labels})
        if calc_mcu and mcu_ratios:
            df['控制器效率占比'] = [r['Ratio'] for r in mcu_ratios]
        if calc_motor and motor_ratios:
            df['电机效率占比'] = [r['Ratio'] for r in motor_ratios]
        if calc_sys and sys_ratios:
            df['系统效率占比'] = [r['Ratio'] for r in sys_ratios]

        try:
            with pd.ExcelWriter(excel_name, engine='openpyxl') as writer:
                header_df = pd.DataFrame(columns=['效率区间', '控制器效率占比', '电机效率占比', '系统效率占比'])
                header_df.to_excel(writer, sheet_name='数据占比', index=False, startrow=0)
                df.to_excel(writer, sheet_name='数据占比', index=False, header=False, startrow=2)
            logging.info(f"保存区域占比 Excel: {excel_name}")
        except Exception as e:
            logging.error(f"保存 Excel {excel_name} 失败: {e}")
            raise ValueError(f"保存 Excel {excel_name} 失败: {e}") from e

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
            raise ValueError(f"保存图表 {plot_name} 失败: {e}") from e

        self.show_ratio_plot()

    def show_ratio_plot(self):
        """查看占比按钮的参数。"""
        show_cached = getattr(self, "show_cached_plot_for_selected_result", self.show_cached_plot)
        if show_cached("效率占比"):
            return True

        try:
            self.reload_runtime_config()
        except Exception as e:
            logging.error(f"重新读取配置失败: {e}")
            QMessageBox.critical(self, "配置错误", f"重新读取配置失败: {e}")
            return False

        if show_cached("效率占比"):
            return True

        if hasattr(self, "ensure_selected_result_ready") and not self.ensure_selected_result_ready():
            QMessageBox.warning(self, "无数据", "请先选择要查看的 sheet。")
            return False
        if show_cached("效率占比"):
            return True

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
