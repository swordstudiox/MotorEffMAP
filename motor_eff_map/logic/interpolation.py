import numpy as np

from .config_values import get_non_negative_config_float


def validate_interpolation_points(points, message_name):
    unique_points = np.unique(points, axis=0)
    if len(unique_points) < 3:
        raise ValueError(f"插值失败：有效{message_name}点少于 3 个，无法生成 MAP。")
    if np.linalg.matrix_rank(unique_points - unique_points.mean(axis=0)) < 2:
        raise ValueError(f"插值失败：有效{message_name}点共线或分布退化，无法生成 MAP。")


def apply_start_cutoff(config, xi, yi, *z_arrays, y_cutoff=None):
    start_speed = get_non_negative_config_float(config, "StartSpeed", 0)
    if y_cutoff is None:
        y_cutoff = get_non_negative_config_float(config, "StartTorque", 0)
    cutoff_mask = (xi < start_speed) | (yi < y_cutoff)
    for z_arr in z_arrays:
        z_arr[cutoff_mask] = np.nan
    geo_mask = (~np.isnan(yi)) & (~cutoff_mask)
    return geo_mask

