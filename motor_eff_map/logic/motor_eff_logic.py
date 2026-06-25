import numpy as np
import pandas as pd
from scipy.interpolate import griddata, interp1d, PchipInterpolator
from scipy.spatial import QhullError
import logging

from motor_eff_map.logic.config_values import (
    get_config_text,
    get_non_negative_config_float,
    get_positive_config_float,
    get_start_power_cutoff,
    parse_step_string,
)
from motor_eff_map.logic.interpolation import (
    apply_start_cutoff,
    validate_interpolation_points,
)

class MotorEffLogic:
    EFFICIENCY_OUTPUT_SWITCHES = {
        'Eff_MCU': ('MCUMAP', 'MCUAreaRatioCalculation'),
        'Eff_Motor': ('MotorMAP', 'MotorAreaRatioCalculation'),
        'Eff_SYS': ('SYSMAP', 'SYSAreaRatioCalculation'),
    }

    def __init__(self, config):
        """
        初始化逻辑类
        config 是一个字典，包含从 INI 文件读取的键值对。
        例如 config['Speed'], config['Torque'] 等。
        """
        self.config = config
        self.raw_df = None       # 原始 DataFrame
        self.processed_df = None # 处理后的 DataFrame
        self.sheets_dict = {}
        self.current_file = None
        self.current_sheet = None
        self.f_edge_curve = None

        # 映射的数据列
        self.speed = None
        self.torque = None
        self.p_motor = None
        self.eff_mcu = None
        self.eff_motor = None
        self.eff_sys = None
        self.u_dc = None
        self.available_efficiency_types = []
        self.last_error = ""

    def _get_config_text(self, key):
        return get_config_text(self.config, key)

    def is_efficiency_output_enabled(self, eff_type):
        switch_keys = self.EFFICIENCY_OUTPUT_SWITCHES.get(eff_type, ())
        return any(str(self.config.get(key, '0')).strip() == '1' for key in switch_keys)

    def is_efficiency_configured(self, eff_type):
        return bool(self._get_config_text(eff_type))

    def has_efficiency_data(self, eff_type):
        df = self._get_valid_efficiency_df(eff_type)
        return df is not None and not df.empty

    def _get_valid_efficiency_df(self, eff_type):
        if self.processed_df is None or eff_type not in self.processed_df.columns:
            return None
        df = self.processed_df.dropna(subset=['Speed', 'Torque', 'P_Motor', eff_type]).copy()
        df[eff_type] = pd.to_numeric(df[eff_type], errors='coerce')
        df = df.dropna(subset=[eff_type])
        return df[(df[eff_type] >= 0) & (df[eff_type] < 100)]

    def load_data(self, file_path):
        """
        从 Excel 加载数据。
        返回: (Success, Message)
        """
        try:
            # 读取所有 sheet (返回字典 {sheet_name: df})
            self.sheets_dict = pd.read_excel(file_path, sheet_name=None)
            self.current_file = file_path
            return True, f"加载了 {len(self.sheets_dict)} 个工作表"
        except Exception as e:
            logging.error(f"加载数据出错: {e}")
            return False, str(e)

    def set_current_sheet(self, sheet_name):
        """设置当前要处理的工作表"""
        if hasattr(self, 'sheets_dict') and sheet_name in self.sheets_dict:
            self.raw_df = self.sheets_dict[sheet_name]
            self.current_sheet = sheet_name
            return True
        return False

    def filter_data(self):
        """
        根据 ini 配置的列名映射数据并存储。
        Determines direction and state.
        返回: (SpeedDirection, MotionState) 字符串
        """
        if self.raw_df is None:
            return None, None
        self.last_error = ""

        # 清理列名 (去除首尾空格)，并避免 strip 后出现重复列导致取列返回 DataFrame。
        stripped_columns = self.raw_df.columns.astype(str).str.strip()
        duplicate_columns = stripped_columns[stripped_columns.duplicated()].unique()
        if len(duplicate_columns) > 0:
            duplicate_text = "、".join(str(col) for col in duplicate_columns)
            self.last_error = f"Excel 列名去除首尾空格后存在重复列名：{duplicate_text}。请修改表头后重试。"
            logging.error(self.last_error)
            self.processed_df = None
            return None, None
        self.raw_df.columns = stripped_columns

        logging.info(f"Excel 中发现的列: {list(self.raw_df.columns)}")
        logging.info(f"配置键映射: Speed='{self.config.get('Speed')}', Torque='{self.config.get('Torque')}'")

        missing_columns = []

        # 为了安全获取列数据的辅助函数
        def get_col(name_key, required=True):
            col_name = self._get_config_text(name_key)

            if not col_name:
                msg = f"配置项 '{name_key}' 未填写。"
                if required:
                    missing_columns.append(msg)
                    logging.error(msg)
                else:
                    logging.info(f"{msg}跳过该数据项。")
                return None

            if col_name and col_name in self.raw_df.columns:
                series = pd.to_numeric(self.raw_df[col_name], errors='coerce')
                if series.notna().sum() == 0:
                    logging.warning(f"列 '{col_name}' (映射自 {name_key}) 全是 0 或 NaN。")
                return series
            else:
                msg = f"在 Excel 文件中未找到列 '{col_name}' (映射自 {name_key})。"
                if required:
                    missing_columns.append(msg)
                    logging.error(msg)
                else:
                    logging.warning(msg)
            return None

        speed_col = get_col('Speed')
        torque_col = get_col('Torque')
        p_motor_col = get_col('P_Motor')

        efficiency_columns = {}
        for eff_type in self.EFFICIENCY_OUTPUT_SWITCHES:
            if not self.is_efficiency_configured(eff_type):
                get_col(eff_type, required=False)
                continue
            required = self.is_efficiency_output_enabled(eff_type)
            col = get_col(eff_type, required=required)
            if col is not None:
                efficiency_columns[eff_type] = col

        # U_dc 逻辑: 优先使用 customUdc
        custom_udc = self.config.get('customUdc', '').strip()
        used_custom_udc = False
        if custom_udc:
            try:
                # customUdc 是代表电压的标量字符串
                val = float(custom_udc)
                # 创建一个常数序列
                self.u_dc = pd.Series(np.full(len(self.raw_df), val))
                logging.info(f"使用配置中的 customUdc: {val}")
                used_custom_udc = True
            except ValueError:
                logging.warning(f"customUdc '{custom_udc}' 不是有效数字。回退到使用数据列。")

        if not used_custom_udc:
             self.u_dc = get_col('U_dc')

        if not efficiency_columns:
            missing_columns.append("至少需要填写并匹配一个效率列配置项：Eff_MCU、Eff_Motor、Eff_SYS。")

        if missing_columns:
            self.last_error = "；".join(missing_columns)
            self.processed_df = None
            return None, None

        # 映射数据
        # 转速、扭矩、功率统一使用绝对值参与效率 MAP 计算
        self.speed = np.abs(speed_col)
        self.torque = np.abs(torque_col)
        self.p_motor = np.abs(p_motor_col)
        self.eff_mcu = efficiency_columns.get('Eff_MCU')
        self.eff_motor = efficiency_columns.get('Eff_Motor')
        self.eff_sys = efficiency_columns.get('Eff_SYS')
        self.available_efficiency_types = list(efficiency_columns.keys())

        # 基于原始值的均值确定方向/状态 (取绝对值之前)
        raw_speed_col = self.config.get('Speed', '')
        if raw_speed_col in self.raw_df.columns:
            mean_speed = pd.to_numeric(self.raw_df[raw_speed_col], errors='coerce').mean()
            if pd.isna(mean_speed):
                speed_direction = "未知"
            else:
                speed_direction = "正转" if mean_speed > 0 else "反转"
        else:
            speed_direction = "未知"

        raw_p_col = self.config.get('P_Motor', '')
        if raw_p_col in self.raw_df.columns:
            mean_p = pd.to_numeric(self.raw_df[raw_p_col], errors='coerce').mean()
            if pd.isna(mean_p):
                motion_state = "未知"
            else:
                motion_state = "电动" if mean_p > 0 else "发电"
        else:
            motion_state = "未知"

        custom_speed_direction = self._get_config_text('customSpeedDirection')
        if custom_speed_direction:
            speed_direction = custom_speed_direction

        custom_motion_state = self._get_config_text('customMotionState')
        if custom_motion_state:
            motion_state = custom_motion_state

        # 组合成 DataFrame 以便处理，只包含当前配置可用的效率列
        processed_data = {
            'Speed': self.speed,
            'Torque': self.torque,
            'P_Motor': self.p_motor,
            'U_dc': self.u_dc
        }
        processed_data.update(efficiency_columns)
        self.processed_df = pd.DataFrame(processed_data)

        return speed_direction, motion_state

    def normalization(self):
        """
        排序，分组转速，移除 NaN 和无效的效率值。
        处理流程:
        1. 按转速排序
        2. 转速分组 (如果转速差值 <= 6rpm 则合并为一组)
        3. 按转速、扭矩排序
        4. 移除 NaN
        5. 对当前可用的效率列过滤 [0, 100)
        """
        if self.processed_df is None:
            return pd.DataFrame()

        df = self.processed_df.copy()
        eff_cols = [col for col in self.available_efficiency_types if col in df.columns]
        core_cols = ['Speed', 'Torque', 'P_Motor', 'U_dc']
        df = df.dropna(subset=core_cols)

        # 1. 按转速排序
        df = df.sort_values(by='Speed')

        # 2. 转速分组：相邻转速差值 <= 6rpm 时合并并取平均值
        speeds = df['Speed'].values
        if len(speeds) > 0:
            new_speeds = speeds.copy()
            speed_cnt = speeds[0]
            count = 0

            for i in range(len(speeds) - 1):
                if abs(speeds[i] - speeds[i+1]) <= 6:
                    speed_cnt += speeds[i+1]
                    count += 1
                else:
                    avg_val = round(speed_cnt / (count + 1))
                    # 赋值给范围
                    new_speeds[i-count : i+1] = avg_val
                    count = 0
                    speed_cnt = speeds[i+1]

            # 处理最后一组
            avg_val = round(speed_cnt / (count + 1))
            new_speeds[len(speeds)-1-count : len(speeds)] = avg_val

            df['Speed'] = new_speeds

        # 3. 按转速、然后按扭矩排序
        df = df.sort_values(by=['Speed', 'Torque'])

        # 4. 效率列不在归一化阶段全局过滤。
        # 不同效率类型可能有不同缺失/无效点，必须在对应 MAP 计算中单独过滤，
        # 否则某个无关效率列会静默删除其他效率 MAP 的有效数据。

        self.processed_df = df
        return df

    def get_external_characteristics(self):
        """
        查找包络线 (最大扭矩) 曲线。
        """
        df = self.processed_df
        # 获取每个转速分组下的最大扭矩，作为外特性包络线采样点
        # normalization 中已经完成转速分组，这里直接按 Speed 聚合。

        max_curve = df.groupby('Speed')['Torque'].max().reset_index()
        max_curve.columns = ['Speed', 'Torque']

        if max_curve.empty:
            self.f_edge_curve = None
            return max_curve

        x = max_curve['Speed'].to_numpy(dtype=float)
        y = max_curve['Torque'].to_numpy(dtype=float)
        max_observed_torque = float(np.nanmax(y))
        upper_limit = max_observed_torque * 1.05

        if len(max_curve) == 1:
            constant_torque = float(y[0])

            def edge_curve(values):
                arr = np.asarray(values, dtype=float)
                return np.full_like(arr, constant_torque, dtype=float)
        elif len(max_curve) > 2:
            interpolator = PchipInterpolator(x, y, extrapolate=True)

            def edge_curve(values):
                return interpolator(values)
        else:
            interpolator = interp1d(x, y, kind='linear', fill_value="extrapolate")

            def edge_curve(values):
                return interpolator(values)

        def bounded_edge_curve(values):
            vals = np.asarray(edge_curve(values), dtype=float)
            vals = np.nan_to_num(vals, nan=0.0, posinf=upper_limit, neginf=0.0)
            return np.clip(vals, 0.0, upper_limit)

        self.f_edge_curve = bounded_edge_curve

        return max_curve

    def get_external_characteristics_data(self):
        """
        返回外特性原始采样点：每个转速分组下最大扭矩点，以及对应功率。
        """
        df = self.processed_df
        if df is None or df.empty:
            return pd.DataFrame(columns=['Speed', 'Torque', 'P_Motor'])

        curve = (
            df.sort_values(by=['Speed', 'Torque', 'P_Motor'])
            .groupby('Speed', as_index=False)
            .tail(1)[['Speed', 'Torque', 'P_Motor']]
            .copy()
        )
        curve = curve.sort_values(by='Speed').reset_index(drop=True)
        return curve

    def _build_edge_curve_from_df(self, df):
        max_curve = df.groupby('Speed')['Torque'].max().reset_index()
        max_curve.columns = ['Speed', 'Torque']

        if max_curve.empty:
            return None

        x = max_curve['Speed'].to_numpy(dtype=float)
        y = max_curve['Torque'].to_numpy(dtype=float)
        max_observed_torque = float(np.nanmax(y))
        upper_limit = max_observed_torque * 1.05

        if len(max_curve) == 1:
            constant_torque = float(y[0])

            def edge_curve(values):
                arr = np.asarray(values, dtype=float)
                return np.full_like(arr, constant_torque, dtype=float)
        elif len(max_curve) > 2:
            interpolator = PchipInterpolator(x, y, extrapolate=True)

            def edge_curve(values):
                return interpolator(values)
        else:
            interpolator = interp1d(x, y, kind='linear', fill_value="extrapolate")

            def edge_curve(values):
                return interpolator(values)

        def bounded_edge_curve(values):
            vals = np.asarray(edge_curve(values), dtype=float)
            vals = np.nan_to_num(vals, nan=0.0, posinf=upper_limit, neginf=0.0)
            return np.clip(vals, 0.0, upper_limit)

        return bounded_edge_curve

    def calculate_loss_values(self, power_kw, efficiency_percent):
        """
        根据输出功率(kW)和效率(%)计算损耗(W)。
        """
        power = pd.to_numeric(power_kw, errors='coerce').astype(float)
        eff = pd.to_numeric(efficiency_percent, errors='coerce').astype(float)
        valid = power.notna() & eff.notna() & (eff > 0) & (eff < 100)
        loss = pd.Series(np.nan, index=power.index, dtype=float)
        loss.loc[valid] = power.loc[valid] * 1000.0 * (100.0 / eff.loc[valid] - 1.0)
        return loss

    def _get_positive_config_float(self, key, default):
        return get_positive_config_float(self.config, key, default)

    def _get_non_negative_config_float(self, key, default=0):
        return get_non_negative_config_float(self.config, key, default)

    def get_start_power_cutoff(self):
        return get_start_power_cutoff(self.config)

    def _validate_interpolation_points(self, points, message_name):
        validate_interpolation_points(points, message_name)

    def _apply_start_cutoff(self, xi, yi, *z_arrays, y_cutoff=None):
        return apply_start_cutoff(self.config, xi, yi, *z_arrays, y_cutoff=y_cutoff)

    def process_map_data(self, eff_type='Eff_MCU', y_cutoff=None):
        """
        为等高线图生成网格数据。
        eff_type: 'Eff_MCU', 'Eff_Motor', 或 'Eff_SYS'
        """
        if self.processed_df is None or self.processed_df.empty:
            return None
        if eff_type not in self.processed_df.columns:
            raise ValueError(f"{eff_type} 未配置或没有有效数据，无法生成对应效率 MAP。")

        df = self._get_valid_efficiency_df(eff_type)
        if df is None or df.empty:
            raise ValueError(f"{eff_type} 没有可用于生成 MAP 的有效效率数据。")

        edge_curve = self._build_edge_curve_from_df(df)
        if edge_curve is None:
            raise ValueError("外特性包络线不可用，请先检查有效数据。")

        speed_grid_step = self._get_positive_config_float('SpeedGrid', 50)
        torque_grid_step = self._get_positive_config_float('TorqueGrid', 5)

        # 源点
        points = df[['Speed', 'Torque']].values
        eff_values = df[eff_type].values
        power_values = df['P_Motor'].values
        self._validate_interpolation_points(points, "转速/扭矩")

        max_speed = df['Speed'].max()
        if max_speed == 0: return None

        # 创建网格
        # 1. 定义转速轴：从 0 到最大转速生成等距网格
        n_speed_steps = int(max_speed / speed_grid_step) + 1
        xi_speed_axis = np.linspace(0, max_speed, n_speed_steps)

        # 2. 定义扭矩网格：先计算最大扭矩曲线，再逐列填充 Y 坐标。
        edge_torques = edge_curve(xi_speed_axis)

        # 确定所需的最大行数 (最大可能扭矩 / 步长 + 2 用于安全/边界)
        max_edge_torque = np.max(edge_torques)
        max_rows = int(np.ceil(max_edge_torque / torque_grid_step)) + 2

        # 初始化 NaNs 二维数组
        # 形状: (行, 列) -> (扭矩, 转速) 以匹配矩阵惯例?
        # 但是 meshgrid 通常给出 XI (行, 列) 形状相同。
        # 让我们使用 (N_Torque, N_Speed) 形状，这是图像绘图 (Y, X) 的标准。
        n_cols = len(xi_speed_axis)
        max_grid_points = int(self.config.get('MaxGridPoints', 5_000_000))
        if max_rows * n_cols > max_grid_points:
            raise ValueError(
                f"网格过大 ({max_rows} x {n_cols})，请增大 SpeedGrid/TorqueGrid 或检查包络线。"
            )

        XI = np.tile(xi_speed_axis, (max_rows, 1)) # X 坐标对每一行重复
        YI = np.full((max_rows, n_cols), np.nan)   # Y 坐标初始化为 NaN

        # 逐列填充 YI
        for col_idx in range(n_cols):
            # 如果最后一个点不是边缘，添加边缘。
            this_speed = xi_speed_axis[col_idx]
            this_max_torque = edge_torques[col_idx]

            # np.arange 不包含终止值，所以稍微放宽上界，再手动补齐边缘点
            # 0:step:Limit
            # np.arange(0, Limit + epsilon, step)
            current_torques = np.arange(0, this_max_torque + torque_grid_step/1000.0, torque_grid_step)
            current_torques = current_torques[current_torques <= this_max_torque] # 安全检查

            # 检查边缘包含
            if len(current_torques) > 0 and not np.isclose(current_torques[-1], this_max_torque):
                 current_torques = np.append(current_torques, this_max_torque)
            elif len(current_torques) == 0:
                 current_torques = np.array([0, this_max_torque]) if this_max_torque > 0 else np.array([0])

            # 填充到 YI
            n_pts = len(current_torques)
            if n_pts > max_rows:
                # 给定 max_rows 定义不应发生，但为了安全进行裁剪
                current_torques = current_torques[:max_rows]
                n_pts = max_rows

            YI[:n_pts, col_idx] = current_torques

            # XI 已经设置正确 (常数列)

        # 3. 插值
        # griddata 通常接受点列表。
        # 我们只需要在有效的 (非 NaN) YI 点进行插值。
        # 掩码：我们要评估的网格点
        eval_mask = ~np.isnan(YI)

        XI_valid = XI[eval_mask]
        YI_valid = YI[eval_mask]

        try:
            ZI_Eff_flat = griddata(points, eff_values, (XI_valid, YI_valid), method='linear')
            ZI_Power_flat = griddata(points, power_values, (XI_valid, YI_valid), method='linear')
        except QhullError as e:
            raise ValueError(f"插值失败：有效转速/扭矩点分布不足以生成二维 MAP。({e.__class__.__name__})") from e

        # 重构二维数组
        ZI_Eff = np.full(YI.shape, np.nan)
        ZI_Power = np.full(YI.shape, np.nan)

        ZI_Eff[eval_mask] = ZI_Eff_flat
        ZI_Power[eval_mask] = ZI_Power_flat

        # 4. 应用启动转速/启动扭矩 截止 (Cutoff)
        mask_valid_geo = self._apply_start_cutoff(XI, YI, ZI_Eff, ZI_Power, y_cutoff=y_cutoff)

        # 对于区域计算，"有效几何" 是从 0rpm、0Nm 开始的包络线内部网格点，
        # 再扣除启动转速/启动扭矩截止区域。
        # 因此 mask_valid_geo 正是 (~np.isnan(YI)) & (~cutoff_mask)

        # 注意: 如果点在数据凸包之外，griddata 可能会返回 nan。
        # 如果我们无法插值，我们是否应该不将其计为效率比率的有效区域？
        # 线性插值在凸包外会返回 NaN，特别是人工边界点未被三角剖分覆盖时。
        # 但是，标准逻辑将分母视为几何区域 (Geometry Area)。
        # 让我们坚持使用几何区域作为分母 (mask_valid_geo)。
        # 分子计算实际值。

        return XI, YI, ZI_Power, ZI_Eff, mask_valid_geo

    def process_speed_power_efficiency_map_data(self, eff_type='Eff_MCU'):
        """
        生成转速-功率-效率 MAP 网格。
        """
        df = self.processed_df
        if df is None or df.empty:
            return None
        if eff_type not in df.columns:
            raise ValueError(f"{eff_type} 未配置或没有有效数据，无法生成对应效率 MAP。")

        start_power = self.get_start_power_cutoff()
        res = self.process_map_data(eff_type, y_cutoff=0)
        if res is None:
            return None

        XI, YI_Torque, _, ZI_Eff, geo_mask = res
        YI_Power = XI * YI_Torque / 9550.0
        cutoff_mask = YI_Power < start_power
        ZI_Eff[cutoff_mask] = np.nan
        geo_mask = geo_mask & (~cutoff_mask)
        return XI, YI_Power, ZI_Eff, geo_mask

    def process_torque_loss_map_data(self, eff_type='Eff_MCU'):
        """
        生成转速-扭矩-损耗 MAP 网格，损耗单位为 W。
        """
        df = self.processed_df
        if df is None or df.empty:
            return None
        if eff_type not in df.columns:
            raise ValueError(f"{eff_type} 未配置或没有有效数据，无法生成对应损耗 MAP。")
        if self.f_edge_curve is None:
            raise ValueError("外特性包络线不可用，请先检查有效数据。")

        loss_values = self.calculate_loss_values(df['P_Motor'], df[eff_type])
        valid_df = df.assign(Loss=loss_values).dropna(subset=['Speed', 'Torque', 'Loss'])
        if valid_df.empty:
            raise ValueError(f"{eff_type} 没有可用于计算损耗的有效效率数据。")

        points = valid_df[['Speed', 'Torque']].values
        self._validate_interpolation_points(points, "转速/扭矩")

        base_res = self.process_map_data(eff_type)
        if base_res is None:
            return None

        XI, YI, _, _, geo_mask = base_res
        eval_mask = ~np.isnan(YI)
        try:
            ZI_Loss_flat = griddata(
                points,
                valid_df['Loss'].values,
                (XI[eval_mask], YI[eval_mask]),
                method='linear',
            )
        except QhullError as e:
            raise ValueError(f"插值失败：有效转速/扭矩点分布不足以生成二维 MAP。({e.__class__.__name__})") from e

        ZI_Loss = np.full(YI.shape, np.nan)
        ZI_Loss[eval_mask] = ZI_Loss_flat
        ZI_Loss[~geo_mask] = np.nan
        return XI, YI, ZI_Loss, geo_mask

    def _parse_step_string(self, step_str):
        return parse_step_string(step_str)

    def calculate_area_ratios(self, z_eff, geo_mask=None):
        """
        计算效率区域占比。
        z_eff: 效率值的二维数组 (或扁平化)
        geo_mask: 可选的有效几何区域布尔掩码 (分母)。
                  如果为 None，假设 z_eff 中的所有非 NaN 值都是有效区域。
        """
        eff_step_str = self.config.get('EffMAPStep', '90 85 80 70')
        eff_levels = sorted(self._parse_step_string(eff_step_str), reverse=True)

        # 分母逻辑更新
        # 如果提供了 geo_mask，我们使用它来定义总的 '运行区域'。
        # 这包括 griddata 可能失败 (NaN) 但在扭矩曲线内部的点。
        # 这里按包络线内部的有效几何网格点计数。
        if geo_mask is not None:
             if geo_mask.shape != z_eff.shape:
                 raise ValueError("geo_mask 与效率矩阵形状不一致，无法计算效率区域占比。")
             denominator = np.sum(geo_mask)
        else:
             valid_mask = ~np.isnan(z_eff)
             denominator = np.sum(valid_mask)

        results = []

        if denominator == 0:
            return []

        for level in eff_levels:
            # 统计 >= level 的点
            # 我们必须在 z_eff 比较中忽略 NaN
            with np.errstate(invalid='ignore'):
                if geo_mask is not None:
                    count = np.sum((z_eff >= level) & geo_mask)
                else:
                    count = np.sum((z_eff >= level))

            ratio = (count / denominator) * 100
            results.append({
                'Level': level,
                'Ratio': ratio
            })

        return results
