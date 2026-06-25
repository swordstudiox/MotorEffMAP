CONFIG_LABELS = {
    "VehicleCode": "车型代号",
    "Speed": "数据里的转速名称",
    "Torque": "数据里的扭矩名称",
    "P_Motor": "数据里的功率名称",
    "Eff_MCU": "数据里的控制器效率名称",
    "Eff_Motor": "数据里的电机效率名称",
    "Eff_SYS": "数据里的系统效率名称",
    "U_dc": "数据里的母线电压名称",
    "customUdc": "固定母线电压（填写后会覆盖U_dc）",
    "MCUMAP": "控制器效率MAP绘制",
    "MCUAreaRatioCalculation": "控制器效率区域占比计算",
    "MotorMAP": "电机效率MAP绘制",
    "MotorAreaRatioCalculation": "电机效率区域占比计算",
    "SYSMAP": "系统效率MAP绘制",
    "SYSAreaRatioCalculation": "系统效率区域占比计算",
    "SpeedPowerMAP": "转速-功率-效率MAP绘制",
    "LossMAP": "转速-扭矩-损耗MAP绘制",
    "ExternalCharacteristicPlot": "外特性曲线图绘制",
    "EffMAPStep": "效率等高线级别",
    "PowerMAPStep": "功率等高线级别",
    "LossMAPStep": "损耗等高线级别",
    "xstepSpeed": "转速轴刻度步长",
    "ystepTorque": "扭矩轴刻度步长",
    "ystepPower": "功率轴刻度步长",
    "StartSpeed": "起始转速",
    "StartTorque": "起始扭矩",
    "StartPower": "起始功率",
    "SpeedGrid": "转速网格步长",
    "TorqueGrid": "扭矩网格步长",
    "MaxGridPoints": "最大网格点数",
    "customSpeedDirection": "固定转向名称",
    "customMotionState": "固定工况状态名称",
}

SWITCH_CONFIG_KEYS = {
    "MCUMAP",
    "MCUAreaRatioCalculation",
    "MotorMAP",
    "MotorAreaRatioCalculation",
    "SYSMAP",
    "SYSAreaRatioCalculation",
    "SpeedPowerMAP",
    "LossMAP",
    "ExternalCharacteristicPlot",
}

EFFICIENCY_MAP_OUTPUTS = (
    ("Eff_MCU", "MCUMAP", "MCU", "控制器"),
    ("Eff_Motor", "MotorMAP", "Motor", "电机"),
    ("Eff_SYS", "SYSMAP", "SYS", "系统"),
)

EFFICIENCY_RATIO_OUTPUTS = (
    ("Eff_MCU", "MCUAreaRatioCalculation", "控制器"),
    ("Eff_Motor", "MotorAreaRatioCalculation", "电机"),
    ("Eff_SYS", "SYSAreaRatioCalculation", "系统"),
)

DEFAULT_CONFIG_VALUES = {
    "SpeedPowerMAP": "1",
    "LossMAP": "1",
    "ExternalCharacteristicPlot": "1",
    "LossMAPStep": "",
    "StartPower": "",
    "xstepSpeed": "500",
    "ystepTorque": "20",
    "ystepPower": "10",
}

EXPORT_FIGURE_SIZE = (9.84, 7.87)
EXPORT_ASPECT_RATIO = EXPORT_FIGURE_SIZE[0] / EXPORT_FIGURE_SIZE[1]

FIGURE_LAYOUT = {
    "left": 0.08,
    "bottom": 0.08,
    "right": 0.96,
    "top": 0.94,
}

DUAL_Y_FIGURE_LAYOUT = {
    "left": 0.08,
    "bottom": 0.08,
    "right": 0.91,
    "top": 0.94,
}

