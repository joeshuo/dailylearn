###加载数据

import pandas as pd
import numpy as np
import lightgbm as lgb
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.metrics import mean_absolute_error, mean_squared_error, mean_absolute_percentage_error

# 设置绘图风格和中文显示
plt.style.use('seaborn-v0_8-whitegrid')
plt.rcParams['font.sans-serif'] = ['SimHei']  
plt.rcParams['axes.unicode_minus'] = False  

# --- 加载我们上一阶段准备好的最终数据 ---
try:
    # 请确保文件名与您保存的文件一致
    data_path = r"E:\A智网\电量预测数据\lightgbm模型\133个行业电量与天气对齐数据_0815.csv"
    df_panel = pd.read_csv(data_path, parse_dates=['date'], encoding='gbk')
    print("数据加载成功！")
    print(df_panel.head())

except FileNotFoundError:
    print(f"错误：文件未找到 {data_path}，请检查路径。")
    df_panel = pd.DataFrame()
except UnicodeDecodeError:
    # 如果'gbk'也不行，说明文件可能被意外地用GBK编码覆盖了
    print("使用 'gbk' 编码失败，这很奇怪。正在尝试 'utf-8' 作为备用方案...")
    try:
        df_panel = pd.read_csv(data_path, parse_dates=['date'], encoding='utf-8')
        print("数据加载成功！")
        print(df_panel.head())
    except Exception as e:
        print(f"尝试所有编码后依然加载失败: {e}")
        df_panel = pd.DataFrame()
except Exception as e:
    print(f"加载数据时发生未知错误: {e}")
    df_panel = pd.DataFrame()

# ==============================================================================
# 模块 1.5：加载完整天气数据源 (历史 + 未来预报)
# 目的：为特征工程模块提供一个独立的、包含所有天气信息的“上帝视角”数据源。
# ==============================================================================
try:
    print("\n--- 正在加载【完整】天气数据源 (包含历史与未来) ---")
    
    # 1. 定义天气全集文件的路径
    full_weather_path = r"E:\A智网\电量预测数据\lightgbm模型\湖北省每日温度特征.xlsx"
    
    # 2. 加载并将其转换为模型需要的长表格式
    df_weather_raw = pd.read_excel(full_weather_path)
    df_weather_full = df_weather_raw.set_index('温度特征').T
    df_weather_full.index.name = 'date'
    df_weather_full.index = pd.to_datetime(df_weather_full.index, format='%Y%m%d')
    df_weather_full.rename(columns={
        '日平均温度': 'temp_mean', '日最高温度': 'temp_max', '日最低温度': 'temp_min',
        '日温度标准差': 'temp_std', '日温差': 'temp_range'
    }, inplace=True)
    
    print("完整天气数据源加载并转换成功！")
    print(f"天气数据覆盖范围: 从 {df_weather_full.index.min().date()} 到 {df_weather_full.index.max().date()}")

except Exception as e:
    print(f"处理完整天气文件时出错: {e}")
    # 创建一个空对象，防止后续代码在找不到文件时报错
    df_weather_full = pd.DataFrame()


# ==============================================================================
# 模块二(V4 - 最终完美版)：在工作日数据上进行滚动窗口异常检测
# 目的：精准识别“噪音”异常，同时完全保护“节假日/周末效应”这一宝贵信号。
# ==============================================================================
from chinese_calendar import is_workday

def flag_outliers_on_workdays(df_panel, industry_col='行业名称', value_col='load_MWh', date_col='date', window_size=31, factor=2.5):
    """
    只在工作日数据上，使用滚动窗口IQR方法，进行异常检测并创建标记特征。
    """
    print("\n" + "="*50)
    print("      开始执行：工作日滚动窗口异常检测")
    print("="*50)
    
    df_flagged = df_panel.copy()
    # 1. 【核心新增】首先，识别出所有的工作日
    workday_mask = df_flagged[date_col].apply(is_workday)
    
    # 创建一个新列来存放异常标记，默认所有都不是异常
    df_flagged['is_outlier'] = 0
    
    industries = df_flagged[industry_col].unique()
    total_outliers_found = 0

    try:
        from tqdm import tqdm
        print("正在逐一分析各行业的工作日数据...")
        iterator = tqdm(industries)
    except ImportError:
        iterator = industries

    for industry in iterator:
        # 2. 【核心新增】只筛选出当前行业的“工作日”数据进行分析
        industry_workday_mask = (df_flagged[industry_col] == industry) & workday_mask
        series = df_flagged.loc[industry_workday_mask, value_col]
        
        if len(series) < 20: continue # 如果工作日数据太少，则跳过
            
        # 3. 在工作日数据上计算滚动的分位数和IQR
        rolling_q1 = series.rolling(window=window_size, center=True, min_periods=window_size//2).quantile(0.25)
        rolling_q3 = series.rolling(window=window_size, center=True, min_periods=window_size//2).quantile(0.75)
        rolling_iqr = rolling_q3 - rolling_q1
        
        # 4. 定义滚动的异常边界 (进一步放宽factor到2.5，更严格地定义异常)
        lower_bound = rolling_q1 - (rolling_iqr * factor)
        upper_bound = rolling_q3 + (rolling_iqr * factor)
        
        # 5. 找出超出局部边界的异常点
        outlier_mask = (series < lower_bound) | (series > upper_bound)
        outlier_indices = series[outlier_mask].index
        
        # 6. 标记异常
        df_flagged.loc[outlier_indices, 'is_outlier'] = 1
        total_outliers_found += outlier_mask.sum()

    print(f"\n处理完成！总共在【工作日】数据中标记了 {total_outliers_found} 个统计异常点。")
    print("="*50 + "\n")
    return df_flagged

# --- 调用最终版标记函数 ---
if 'df_panel' in locals() and not df_panel.empty:
    df_flagged = flag_outliers_on_workdays(df_panel)
    df_panel = df_flagged
    print("原始变量 'df_panel' 已被添加了'is_outlier'特征的数据覆盖。")
    
    if df_panel[df_panel['is_outlier'] == 1].shape[0] > 0:
        print("\n新特征预览 (部分被标记为异常的工作日):")
        print(df_panel[df_panel['is_outlier'] == 1].head())
    else:
        print("\n在新的检测方法下，工作日数据中未发现任何需要标记的异常点。")
else:
    print("错误：原始数据 'df_panel' 未加载，无法进行异常处理。")

# ==============================================================================
# 模块四 (全新)：先划分数据集，彻底杜绝数据泄露
# ==============================================================================
# 假设 df_panel (来自异常标记) 和 df_weather_full (天气全集) 已准备好

# 1. 定义最终测试集的开始日期
test_start_date = df_panel['date'].max() - pd.DateOffset(days=29)

# 2. 划分历史数据和未来天气“情报”
historical_df = df_panel[df_panel['date'] < test_start_date].copy()
test_df_base = df_panel[df_panel['date'] >= test_start_date].copy()

# 我们的“天气情报”就是完整的 df_weather_full
print(f"数据集已划分为：")
print(f"  - 历史数据 (用于训练和分层): {len(historical_df)} 行, 截止到 {historical_df['date'].max().date()}")
print(f"  - 最终测试集 (基底): {len(test_df_base)} 行, 从 {test_start_date.date()} 开始")

# ==============================================================================
# 模块五 (最终输出确认版): 先做特征工程，再做模型分层
# ==============================================================================
# 假设 historical_df (来自模块四) 已经存在
# 假设 df_weather_full (来自模块1.5) 已经存在

# --- 5a. 在纯净的历史数据上，进行特征工程 ---
print("\n--- 正在为【纯净历史数据】创建特征 (无泄露) ---")

#   我们使用一个【简化版】的特征工程函数
def create_features_without_tier(df_input, df_weather_source):
    df_proc = df_input.copy()
    df_proc = df_proc.sort_values(by=['行业名称', 'date']).reset_index(drop=True)
    # ... (所有特征计算代码) ...
    df_proc['month'] = df_proc['date'].dt.month
    df_proc['dayofweek'] = df_proc['date'].dt.dayofweek
    df_proc['dayofyear'] = df_proc['date'].dt.dayofyear
    df_proc['weekofyear'] = df_proc['date'].dt.isocalendar().week.astype(int)
    df_proc['is_holiday'] = df_proc['date'].apply(is_holiday).astype(int)
    df_proc['is_weekend_norm'] = (df_proc['date'].dt.dayofweek >= 5).astype(int)
    df_proc['is_adj_workday'] = df_proc.apply(lambda row: 1 if is_workday(row['date']) and row['dayofweek'] >= 5 else 0, axis=1)
    df_proc['is_offday'] = df_proc.apply(lambda row: 1 if row['is_holiday'] == 1 or (row['is_weekend_norm'] == 1 and row['is_adj_workday'] == 0) else 0, axis=1)
    lags = [1, 2, 7, 14]
    for lag in lags:
        df_proc[f'load_lag_{lag}'] = df_proc.groupby('行业名称')['load_MWh'].shift(lag)
    df_proc['rolling_mean_7'] = df_proc.groupby('行业名称')['load_MWh'].shift(1).rolling(window=7, min_periods=1).mean()
    df_proc['rolling_std_7'] = df_proc.groupby('行业名称')['load_MWh'].shift(1).rolling(window=7, min_periods=1).std()
    df_proc['load_diff_1_7'] = df_proc['load_lag_1'] - df_proc['load_lag_7']
    df_proc['load_ratio_1_roll7'] = (df_proc['load_lag_1'] / (df_proc['rolling_mean_7'] + 1e-6)) - 1
    df_proc.set_index('date', inplace=True)
    future_days = [1, 2, 3]
    weather_cols = ['temp_max', 'temp_min', 'temp_mean']
    for day in future_days:
        df_weather_shifted = df_weather_source[weather_cols].shift(-day)
        df_weather_shifted.columns = [f'{col}_future_d{day}' for col in weather_cols]
        df_proc = df_proc.join(df_weather_shifted)
    df_proc.reset_index(inplace=True)
    df_proc['行业名称'] = df_proc['行业名称'].astype('category')
    df_proc = df_proc.drop(['is_weekend_norm', 'is_adj_workday'], axis=1)
    return df_proc

df_featured_historical = create_features_without_tier(historical_df, df_weather_full)
df_featured_historical = df_featured_historical.dropna()

# --- 5b. 在【特征工程之后】，我们再来合并tier信息 ---
print("\n--- 正在为特征化后的历史数据添加tier信息 ---")

#   计算tier的映射关系 (这一步不变)
industry_avg_load = historical_df.groupby('行业名称')['load_MWh'].mean().to_frame('avg_load')
def assign_tier(avg_load):
    if avg_load >= 1000: return '大行业'
    elif avg_load < 100: return '小行业'
    else: return '中行业'
industry_avg_load['tier'] = industry_avg_load['avg_load'].apply(assign_tier)

#   将tier信息合并到我们刚刚创建的 df_featured_historical 中
df_featured_historical = pd.merge(df_featured_historical, industry_avg_load[['tier']], on='行业名称', how='left')
df_featured_historical['tier'].fillna('小行业', inplace=True)

# --- 5c. 【核心】现在，我们安全地进行数据集拆分，并创建 data_tiers_hist 字典 ---
print("\n--- 在添加tier后，进行数据集拆分 ---")
# 创建一个空字典
data_tiers_hist = {}

# 检查df_featured_historical是否已成功创建
if 'df_featured_historical' in locals() and not df_featured_historical.empty:
    data_tiers_hist['大行业'] = df_featured_historical[df_featured_historical['tier'] == '大行业'].copy()
    data_tiers_hist['中行业'] = df_featured_historical[df_featured_historical['tier'] == '中行业'].copy()
    data_tiers_hist['小行业'] = df_featured_historical[df_featured_historical['tier'] == '小行业'].copy()
    print("分层数据集已成功创建并存入 'data_tiers_hist' 字典！")
else:
    print("错误：特征工程未能成功生成 df_featured_historical。")

# ==============================================================================
# 模块六 (最终修正版): 超参数调优
# 输入: data_tiers_hist (来自模块五的字典)
# ==============================================================================
import optuna
import pandas as pd
import lightgbm as lgb
from sklearn.metrics import mean_absolute_error
import json

# 创建一个空字典，用于自动存储每个模型的最优参数
best_params_per_tier = {}

# 【核心修正】我们不再寻找df_large等变量，而是直接检查 data_tiers_hist 字典
if 'data_tiers_hist' in locals() and data_tiers_hist:
    
    print("\n" + "="*50)
    print("      开始为三个层级的模型进行自动化超参数调优")
    print("="*50)

    # 直接遍历我们从上一步得到的 data_tiers_hist 字典
    for tier_name, df_tier in data_tiers_hist.items():
        if df_tier.empty: 
            print(f"\n--- 跳过【{tier_name}】模型，因为该层级没有数据 ---")
            continue
            
        print(f"\n--- 正在为【{tier_name}】模型寻找最优参数 ---")
        
        # 1. 数据划分
        #    注意：这里的df_tier就是 data_tiers_hist['大行业'] 等
        val_date_start = df_tier['date'].max() - pd.DateOffset(days=60)
        val_date_end = df_tier['date'].max() - pd.DateOffset(days=31)
        train = df_tier[df_tier['date'] < val_date_start].copy()
        val = df_tier[(df_tier['date'] >= val_date_start) & (df_tier['date'] <= val_date_end)].copy()
        
        FEATURES = [col for col in df_tier.columns if col not in ['date', 'load_MWh', 'tier']]
        TARGET = 'load_MWh'
        X_train, y_train = train[FEATURES], train[TARGET]
        X_val, y_val = val[FEATURES], val[TARGET]

        # 2. 定义Optuna的目标函数 (内部逻辑不变)
        def objective(trial):
            # ... (objective函数内部完全不变) ...
            params = {
                'objective': 'regression_l1', 'metric': 'l1', 'n_estimators': 2000,
                'random_state': 42, 'n_jobs': -1, 'verbosity': -1,
                'learning_rate': trial.suggest_float('learning_rate', 0.01, 0.3, log=True),
                'num_leaves': trial.suggest_int('num_leaves', 20, 300),
                'max_depth': trial.suggest_int('max_depth', 3, 12),
                'min_child_samples': trial.suggest_int('min_child_samples', 5, 100),
                'feature_fraction': trial.suggest_float('feature_fraction', 0.4, 1.0),
                'bagging_fraction': trial.suggest_float('bagging_fraction', 0.4, 1.0),
                'bagging_freq': trial.suggest_int('bagging_freq', 1, 7),
                'lambda_l1': trial.suggest_float('lambda_l1', 1e-8, 10.0, log=True),
                'lambda_l2': trial.suggest_float('lambda_l2', 1e-8, 10.0, log=True),
            }
            model = lgb.LGBMRegressor(**params)
            model.fit(X_train, y_train,
                      eval_set=[(X_val, y_val)], eval_metric='mae',
                      callbacks=[lgb.early_stopping(100, verbose=False)],
                      categorical_feature=['行业名称'])
            preds = model.predict(X_val)
            mae = mean_absolute_error(y_val, preds)
            return mae

        # 3. 创建并运行Optuna研究
        study = optuna.create_study(direction='minimize') 
        study.optimize(objective, n_trials=50, show_progress_bar=True)
        
        # 4. 自动存储找到的最佳参数
        best_params_per_tier[tier_name] = study.best_params
        
        print(f"【{tier_name}】模型调优完成！")

    print("\n" + "="*50)
    print("      所有模型的超参数调优均已完成！")
    print("="*50)
    print("最终找到的各层级最优参数为：")
    import json
    print(json.dumps(best_params_per_tier, indent=4))

else:
    print("错误：未能找到'data_tiers_hist'字典。请确保【模块五】已成功运行。")