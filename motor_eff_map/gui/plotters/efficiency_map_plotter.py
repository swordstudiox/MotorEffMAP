import logging

import numpy as np
from PySide6.QtWidgets import QMessageBox


class EfficiencyMapPlotterMixin:
    def show_map_plot(self, eff_type_short):
        """按最新 INI 配置显示单个效率 MAP。"""
        output_by_plot_key = {
            plot_key: (eff_type, switch_key, display_name)
            for eff_type, switch_key, plot_key, display_name in self.EFFICIENCY_MAP_OUTPUTS
        }
        output = output_by_plot_key.get(eff_type_short)
        if output is None:
            return False

        cache_suffix = self.get_efficiency_map_cache_suffix(eff_type_short)
        if cache_suffix:
            show_cached = getattr(self, "show_cached_plot_for_selected_result", self.show_cached_plot)
            if show_cached(cache_suffix):
                return True

        try:
            self.reload_runtime_config()
        except Exception as e:
            logging.error(f"重新读取配置失败: {e}")
            QMessageBox.critical(self, "配置错误", f"重新读取配置失败: {e}")
            return False

        eff_type, switch_key, display_name = output
        if not self.is_config_switch_on(switch_key):
            self.figure.clear()
            self.canvas.draw()
            logging.info(f"{display_name}效率MAP未启用，已按配置跳过显示。")
            return False

        if cache_suffix:
            show_cached = getattr(self, "show_cached_plot_for_selected_result", self.show_cached_plot)
            if show_cached(cache_suffix):
                return True

        if hasattr(self, "ensure_selected_result_ready") and not self.ensure_selected_result_ready():
            QMessageBox.warning(self, "无数据", "请先选择要查看的 sheet。")
            return False

        if cache_suffix:
            show_cached = getattr(self, "show_cached_plot_for_selected_result", self.show_cached_plot)
            if show_cached(cache_suffix):
                return True

        if not self.should_use_efficiency_output(eff_type, switch_key, f"{display_name}效率MAP"):
            self.figure.clear()
            self.canvas.draw()
            logging.info(f"{display_name}效率MAP未启用，已按配置跳过显示。")
            return False

        return self.switch_plot(eff_type_short)

    def switch_plot(self, eff_type_short, save_png=False):
        if not hasattr(self.logic, 'processed_df') or self.logic.processed_df is None:
            return False

        output = self.get_efficiency_output_by_short(eff_type_short)
        if output is None:
            return False
        map_key, _, _, title_part = output

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

            self.add_contour_legend(ax, [
                {
                    "label": "效率",
                    "color": "k",
                    "linewidth": 0.5,
                    "marker": "o",
                    "markerfacecolor": "#000080",
                    "markeredgecolor": "k",
                },
                {
                    "label": "功率",
                    "color": "green",
                    "linewidth": 0.8,
                    "marker": "o",
                    "markerfacecolor": "none",
                    "markeredgecolor": "green",
                },
            ])

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
            if XI.size > 0 and YI.size > 0:
                max_x = np.nanmax(XI)
                max_y = np.nanmax(YI)
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
                fname = f"{self.build_output_stem(title_part + 'EfficiencyMAP')}.png"

                # 使用固定 DPI 并移除 bbox_inches='tight' 以尊重上面设置的确切图形尺寸
                # 使用 200 DPI 平衡导出清晰度和文件体积。

                # 导出时切换到固定尺寸，确保输出图片版式稳定
                old_size = self.figure.get_size_inches()
                self.figure.set_size_inches(*self.EXPORT_FIGURE_SIZE) # 25x20cm 严格

                # 对文件应用固定边距
                self.apply_figure_layout()

                self.figure.savefig(fname, dpi=200)
                cache_suffix = self.get_efficiency_map_cache_suffix(eff_type_short)
                if cache_suffix:
                    self.register_plot_cache(cache_suffix, fname)

                # 恢复：切换回以前的尺寸和安全的 UI 边距
                self.figure.set_size_inches(old_size)
                self.apply_figure_layout()
                self.canvas.draw()

                logging.info(f"已保存图形到 {fname}")
            except Exception as e:
                logging.error(f"保存图像失败: {e}")
        return True
