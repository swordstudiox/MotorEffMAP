import logging
import os

import numpy as np
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D


class PlotHelperMixin:
    def apply_figure_layout(self, figure=None, layout=None):
        target = figure or self.figure
        target.subplots_adjust(**(layout or self.FIGURE_LAYOUT))

    def parse_contour_levels(self, key, default):
        step_str = str(self.config_dict.get(key, "") or "").strip()
        if not step_str:
            return default
        normalized_str = step_str.replace(",", " ").replace(";", " ")
        return sorted(list(set(float(x) for x in normalized_str.split())))

    def prepare_masked_contour_grid(self, xi, yi, zi):
        """将 ragged MAP 网格转为 contourf 可用的有限坐标 + masked 值。"""
        yi_filled = np.array(yi, dtype=float, copy=True)
        for col_idx in range(yi_filled.shape[1]):
            finite_col = np.isfinite(yi_filled[:, col_idx])
            fill_value = float(np.nanmax(yi_filled[:, col_idx])) if finite_col.any() else 0.0
            yi_filled[~finite_col, col_idx] = fill_value

        return xi, yi_filled, np.ma.masked_invalid(zi)

    def add_contour_legend(self, ax, entries):
        handles = [
            Line2D(
                [0],
                [0],
                color=entry.get("color", "k"),
                linewidth=entry.get("linewidth", 0.5),
                linestyle=entry.get("linestyle", "-"),
                marker=entry.get("marker", "o"),
                markersize=entry.get("markersize", 4),
                markerfacecolor=entry.get("markerfacecolor", "none"),
                markeredgecolor=entry.get("markeredgecolor", entry.get("color", "k")),
                label=entry["label"],
            )
            for entry in entries
        ]
        legend = ax.legend(handles=handles, loc="upper right", fancybox=False, edgecolor="k", framealpha=0.7, fontsize=7)
        legend.get_frame().set_linewidth(0.5)
        try:
            legend.set_draggable(True)
        except Exception:
            pass
        return legend

    def save_current_figure(self, suffix, layout=None):
        try:
            fname = f"{self.build_output_stem(suffix)}.png"
            old_size = self.figure.get_size_inches()
            self.figure.set_size_inches(*self.EXPORT_FIGURE_SIZE)
            self.apply_figure_layout(layout=layout)
            self.figure.savefig(fname, dpi=200)
            self.register_plot_cache(suffix, fname)
            self.figure.set_size_inches(old_size)
            self.apply_figure_layout(layout=layout)
            self.canvas.draw()
            logging.info(f"已保存图形到 {fname}")
            return True
        except Exception as e:
            logging.error(f"保存图像失败: {e}")
            return False

    def get_plot_cache_key(self, suffix):
        current_results = getattr(self, "current_results", {})
        current_file = current_results.get("source_file", "")
        if not current_file and getattr(self, "logic", None) is not None:
            current_file = getattr(self.logic, "current_file", "")
        sheet = current_results.get("sheet", "")
        file_key = os.path.normcase(os.path.abspath(current_file)) if current_file else ""
        config_key = tuple(sorted((str(k), str(v)) for k, v in getattr(self, "config_dict", {}).items()))
        view_state_key = ()
        if suffix == "外特性曲线" and hasattr(self, "chk_external_values"):
            view_state_key = (("show_values", self.chk_external_values.isChecked()),)
        return (file_key, sheet, suffix, config_key, view_state_key)

    def register_plot_cache(self, suffix, path):
        if not hasattr(self, "plot_image_cache"):
            self.plot_image_cache = {}
        self.plot_image_cache[self.get_plot_cache_key(suffix)] = os.path.abspath(path)

    def show_cached_plot(self, suffix):
        cache = getattr(self, "plot_image_cache", {})
        path = cache.get(self.get_plot_cache_key(suffix))
        if not path or not os.path.exists(path):
            return False

        try:
            image = plt.imread(path)
            self.figure.clear()
            self.apply_figure_layout()
            ax = self.figure.add_subplot(111)
            ax.imshow(image)
            ax.axis("off")
            self.canvas.draw()
            logging.info(f"已从缓存显示图像: {path}")
            return True
        except Exception as e:
            logging.warning(f"读取缓存图像失败，改为重新绘制: {e}")
            return False

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

    def get_positive_config_float(self, key, default):
        try:
            value = float(self.config_dict.get(key, default) or default)
        except (TypeError, ValueError):
            logging.warning(f"{key} 配置无效，已按 {default} 处理。")
            return float(default)
        if not np.isfinite(value) or value <= 0:
            logging.warning(f"{key} 配置无效，已按 {default} 处理。")
            return float(default)
        return value

    def apply_axis_ticks(self, ax, max_x, max_y, x_step_key, y_step_key,
                         x_default, y_default, start_x=0, start_y=0):
        x_step = self.get_positive_config_float(x_step_key, x_default)
        y_step = self.get_positive_config_float(y_step_key, y_default)
        start_x = min(float(start_x), float(max_x))
        start_y = min(float(start_y), float(max_y))
        end_x = np.ceil(float(max_x) / x_step) * x_step
        end_y = np.ceil(float(max_y) / y_step) * y_step

        ax.set_xticks(np.arange(start_x, end_x + x_step / 2, x_step))
        ax.set_yticks(np.arange(start_y, end_y + y_step / 2, y_step))
        ax.set_xlim(start_x, end_x)
        ax.set_ylim(start_y, end_y)
