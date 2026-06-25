import numpy as np


def get_config_text(config, key):
    return str(config.get(key, "") or "").strip().strip("'").strip('"')


def get_positive_config_float(config, key, default):
    try:
        value = float(config.get(key, default))
    except (TypeError, ValueError):
        raise ValueError(f"{key} 必须是大于 0 的数字")

    if not np.isfinite(value) or value <= 0:
        raise ValueError(f"{key} 必须是大于 0 的数字")
    return value


def get_non_negative_config_float(config, key, default=0):
    try:
        value = float(config.get(key, default) or default)
    except (TypeError, ValueError):
        raise ValueError(f"{key} 必须是大于等于 0 的数字")

    if not np.isfinite(value) or value < 0:
        raise ValueError(f"{key} 必须是大于等于 0 的数字")
    return value


def get_start_power_cutoff(config):
    start_power_text = get_config_text(config, "StartPower")
    if start_power_text:
        return get_non_negative_config_float(config, "StartPower", 0)

    start_speed = get_non_negative_config_float(config, "StartSpeed", 0)
    start_torque = get_non_negative_config_float(config, "StartTorque", 0)
    return start_speed * start_torque / 9550.0


def parse_step_string(step_str):
    """解析步长字符串，如 '10,20,30' 或 '10:10:90'。"""
    if ":" in step_str:
        try:
            parts = [float(x) for x in step_str.split(":")]
            if len(parts) == 2:
                return list(np.arange(parts[0], parts[1] + parts[1] / 1000.0, 1))
            if len(parts) == 3:
                return list(np.arange(parts[0], parts[2] + parts[1] / 1000.0, parts[1]))
        except Exception:
            pass

    try:
        return [float(x) for x in step_str.replace(";", " ").replace(",", " ").split()]
    except Exception:
        return []

