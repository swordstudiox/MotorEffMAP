import numpy as np
import pandas as pd
from scipy.interpolate import griddata, interp1d, UnivariateSpline
from scipy.spatial import ConvexHull
import logging

class MotorEffLogic:
    def __init__(self, config):
        """
        初始化逻辑类
        config 是一个字典，包含从 INI 文件读取的键值对。
        例如 config['Speed'], config['Torque'] 等。
        """
        self.config = config
        self.raw_df = None       # 原始 DataFrame
        self.processed_df = None # 处理后的 DataFrame
        
        # 映射的数据列
        self.speed = None
        self.torque = None
        self.p_motor = None
        self.eff_mcu = None
        self.eff_motor = None
        self.eff_sys = None
        self.u_dc = None

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
            
        # 清理列名 (去除首尾空格)
        self.raw_df.columns = self.raw_df.columns.str.strip()
        
        logging.info(f"Excel 中发现的列: {list(self.raw_df.columns)}")
        logging.info(f"配置键映射: Speed='{self.config.get('Speed')}', Torque='{self.config.get('Toqrue')}'")

        # 为了安全获取列数据的辅助函数
        def get_col(name_key):
            col_name = self.config.get(name_key, '')
            # 去除引号 (有些用户在 INI 中加了引号)
            if col_name: col_name = col_name.strip("'").strip('"')
            
            if col_name and col_name in self.raw_df.columns:
                series = pd.to_numeric(self.raw_df[col_name], errors='coerce').fillna(0)
                if series.abs().sum() == 0:
                    logging.warning(f"列 '{col_name}' (映射自 {name_key}) 全是 0 或 NaN。")
                return series
            else:
                logging.warning(f"在 Excel 文件中未找到列 '{col_name}' (映射自 {name_key})。")
            return np.zeros(len(self.raw_df))

        # 映射数据
        # 注意: MATLAB 代码对 Speed, Torque, P_Motor 使用了 abs()
        self.speed = np.abs(get_col('Speed'))
        self.torque = np.abs(get_col('Toqrue')) # 注意 MATLAB 代码中的 INI 键名拼写错误 'Toqrue'
        self.p_motor = np.abs(get_col('P_Motor'))
        
        self.eff_mcu = get_col('Eff_MCU')
        self.eff_motor = get_col('Eff_Motor')
        self.eff_sys = get_col('Eff_SYS')
        
        # U_dc 逻辑: 优先使用 customUdc
        custom_udc = self.config.get('customUdc', '').strip()
        used_custom_udc = False
        if custom_udc:
            try:
                # MATLAB 代码暗示 customUdc 是代表电压的标量字符串
                val = float(custom_udc)
                # 创建一个常数序列
                self.u_dc = pd.Series(np.full(len(self.raw_df), val))
                logging.info(f"使用配置中的 customUdc: {val}")
                used_custom_udc = True
            except ValueError:
                logging.warning(f"customUdc '{custom_udc}' 不是有效数字。回退到使用数据列。")
        
        if not used_custom_udc:
             self.u_dc = get_col('U_dc')

        # 基于原始值的均值确定方向/状态 (取绝对值之前)
        raw_speed_col = self.config.get('Speed', '')
        if raw_speed_col in self.raw_df.columns:
            mean_speed = self.raw_df[raw_speed_col].mean()
            speed_direction = "正转" if mean_speed > 0 else "反转"
        else:
            speed_direction = "未知"

        raw_p_col = self.config.get('P_Motor', '')
        if raw_p_col in self.raw_df.columns:
            mean_p = self.raw_df[raw_p_col].mean()
            motion_state = "电动" if mean_p > 0 else "发电"
        else:
            motion_state = "未知"

        # 组合成 DataFrame 以便处理
        self.processed_df = pd.DataFrame({
            'Speed': self.speed,
            'Torque': self.torque,
            'P_Motor': self.p_motor,
            'Eff_MCU': self.eff_mcu,
            'Eff_Motor': self.eff_motor,
            'Eff_SYS': self.eff_sys,
            'U_dc': self.u_dc
        })
        
        return speed_direction, motion_state

    def normalization(self):
        """
        排序，分组转速，移除 NaN 和无效的效率值。
        Matlab 逻辑:
        1. 按转速排序
        2. 转速分组 (如果转速差值 <= 6rpm 则合并为一组)
        3. 按转速、扭矩排序
        4. 移除 NaN
        5. 过滤效率值 [0, 100)
        """
        df = self.processed_df.copy()
        
        # 1. 按转速排序
        df = df.sort_values(by='Speed')
        
        # 2. 转速分组 (逻辑翻译)
        # MATLAB: 迭代并平均转速，如果差值 <= 6
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
        
        # 4. 移除 NaN
        df = df.dropna()
        
        # 5. 过滤效率
        df = df[ (df['Eff_MCU'] >= 0) & (df['Eff_MCU'] < 100) ]
        df = df[ (df['Eff_Motor'] >= 0) & (df['Eff_Motor'] < 100) ]
        df = df[ (df['Eff_SYS'] >= 0) & (df['Eff_SYS'] < 100) ]
        
        self.processed_df = df
        return df

    def get_external_characteristics(self):
        """
        查找包络线 (最大扭矩) 曲线。
        """
        df = self.processed_df
        # 获取每个唯一转速的最大扭矩
        # MATLAB 再次用循环和 'count' 来查找唯一转速组?
        # 既然我们在 normalization 中已经对转速进行了分组，我们可以直接按 'Speed' 分组。
        
        max_curve = df.groupby('Speed')['Torque'].max().reset_index()
        max_curve.columns = ['Speed', 'Torque']
        
        # 确定曲线拟合函数 (包络线)
        # MATLAB 使用: fit(SpeedEdge, TorqueEdge, 'smoothingspline')
        # 我们可以使用 UnivariateSpline 或 interp1d
        if len(max_curve) > 3:
            # MATLAB 使用 'smoothingspline'
            # UnivariateSpline(s=None) 估计平滑因子。
            # 为了更接近 MATLAB 的平滑包络线: 允许平滑 (不强制 s=0)。
            try:
                # 默认平滑
                self.f_edge_curve = UnivariateSpline(max_curve['Speed'], max_curve['Torque'], s=None, k=3)
            except:
                self.f_edge_curve = interp1d(max_curve['Speed'], max_curve['Torque'], kind='linear', fill_value="extrapolate")
        else:
             self.f_edge_curve = interp1d(max_curve['Speed'], max_curve['Torque'], kind='linear', fill_value="extrapolate")
             
        return max_curve

    def process_map_data(self, eff_type='Eff_MCU'):
        """
        为等高线图生成网格数据。
        eff_type: 'Eff_MCU', 'Eff_Motor', 或 'Eff_SYS'
        """
        df = self.processed_df
        
        speed_grid_step = float(self.config.get('SpeedGrid', 50))
        torque_grid_step = float(self.config.get('TorqueGrid', 5))
        
        # 源点
        points = df[['Speed', 'Torque']].values
        eff_values = df[eff_type].values
        power_values = df['P_Motor'].values
        
        max_speed = df['Speed'].max()
        if max_speed == 0: return None
        
        # 创建网格
        # 1. 定义转速轴 (匹配 MATLAB linspace)
        # MATLAB: SpeedRange=linspace(0,max(Speed),max(Speed)/SpeedGrid+1)
        n_speed_steps = int(max_speed / speed_grid_step) + 1
        xi_speed_axis = np.linspace(0, max_speed, n_speed_steps)
        
        # 2. 定义扭矩网格，模仿 MATLAB 的逐列逻辑
        # MATLAB 先计算 MaxTorque 曲线，然后对每个转速列填充 Y。
        edge_torques = self.f_edge_curve(xi_speed_axis)
        
        # 确定所需的最大行数 (最大可能扭矩 / 步长 + 2 用于安全/边界)
        max_edge_torque = np.max(edge_torques)
        max_rows = int(np.ceil(max_edge_torque / torque_grid_step)) + 2
        
        # 初始化 NaNs 二维数组
        # 形状: (行, 列) -> (扭矩, 转速) 以匹配矩阵惯例?
        # 但是 meshgrid 通常给出 XI (行, 列) 形状相同。
        # 让我们使用 (N_Torque, N_Speed) 形状，这是图像绘图 (Y, X) 的标准。
        n_cols = len(xi_speed_axis)
        
        XI = np.tile(xi_speed_axis, (max_rows, 1)) # X 坐标对每一行重复
        YI = np.full((max_rows, n_cols), np.nan)   # Y 坐标初始化为 NaN
        
        # 逐列填充 YI
        for col_idx in range(n_cols):
            # MATLAB: Torque_temp_Datas=0:TorqueGrid:f_EdgeCurve(SpeedRange(i));
            # 如果最后一个点不是边缘，添加边缘。
            this_speed = xi_speed_axis[col_idx]
            this_max_torque = edge_torques[col_idx]
            
            # Arange 不包含终止值，所以稍微大一点，但手动处理逻辑以匹配 MATLAB
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
        
        ZI_Eff_flat = griddata(points, eff_values, (XI_valid, YI_valid), method='linear')
        ZI_Power_flat = griddata(points, power_values, (XI_valid, YI_valid), method='linear')
        
        # 重构二维数组
        ZI_Eff = np.full(YI.shape, np.nan)
        ZI_Power = np.full(YI.shape, np.nan)
        
        ZI_Eff[eval_mask] = ZI_Eff_flat
        ZI_Power[eval_mask] = ZI_Power_flat
        
        # 4. 应用启动转速/启动扭矩 截止 (Cutoff)
        start_speed = float(self.config.get('StartSpeed', 0) or 0)
        start_torque = float(self.config.get('StartTorque', 0) or 0)
        
        cutoff_mask = (XI < start_speed) | (YI < start_torque)
        
        # 屏蔽截止区域 (设为 NaN)
        ZI_Eff[cutoff_mask] = np.nan
        ZI_Power[cutoff_mask] = np.nan
        
        # 对于区域计算，"有效几何" 确切地是 "YI 内部非 NaN" 减去截止区域。
        # 在 MATLAB 中，Area 是 'TorqueFull' (即我们的 YI 有效点) 中的点总和
        # 减去截止点。
        # 所以 mask_valid_geo 正是 (~np.isnan(YI)) & (~cutoff_mask)
        mask_valid_geo = (~np.isnan(YI)) & (~cutoff_mask)
        
        # 注意: 如果点在数据凸包之外，griddata 可能会返回 nan。
        # 如果我们无法插值，我们是否应该不将其计为效率比率的有效区域？
        # MATLAB fit(...) 可能会外推或返回 NaN? 'linearinterp' 通常创建三角剖分。
        # 如果三角剖分没有覆盖人工边界点 (MaxTorque)，会发生什么？
        # MATLAB 可能在那里也返回 NaN (如果在凸包外)。
        # 但是，标准逻辑将分母视为几何区域 (Geometry Area)。
        # 让我们坚持使用几何区域作为分母 (mask_valid_geo)。
        # 分子计算实际值。
        
        return XI, YI, ZI_Power, ZI_Eff, mask_valid_geo

    def _parse_step_string(self, step_str):
        """辅助函数: 解析步长字符串，如 '10,20,30' 或 '10:10:90'。"""
        if ':' in step_str:
            try:
                parts = [float(x) for x in step_str.split(':')]
                if len(parts) == 2:
                    return list(np.arange(parts[0], parts[1] + parts[1]/1000.0, 1)) # 假设步长 1? 不，默认 matlab 行为
                elif len(parts) == 3:
                     # start:step:end (包含)
                    return list(np.arange(parts[0], parts[2] + parts[1]/1000.0, parts[1]))
            except:
                pass # 回退
        
        # 默认空格/逗号分割
        try:
             return [float(x) for x in step_str.replace(';', ' ').replace(',', ' ').split()]
        except:
             return []

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
        # 这与 MATLAB 计数 TorqueFull 网格点相匹配。
        if geo_mask is not None:
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
                # 如果提供了 geo_mask，严格来说我们只关心 z_eff 值。
                # geo_mask 之外的点反正也是 NaN。
                count = np.sum((z_eff >= level))
            
            ratio = (count / denominator) * 100
            results.append({
                'Level': level,
                'Ratio': ratio
            })
            
        return results

