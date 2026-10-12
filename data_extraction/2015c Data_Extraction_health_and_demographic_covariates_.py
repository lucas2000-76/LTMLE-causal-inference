"""
Script Name: CHARLS_2015_Health_and_Demo_Covariates_Pipeline.py
Description:
    Processes CHARLS 2015 data to compute health behaviors and demographic covariates.

    Main Workflows:
    1. Health Behaviors: Computes drinking (L2_Drink), smoking status (L2_Smoke), and sleep duration (L2_Sleep_time).
    2. Demographic Covariates: Computes marital status (L2_Martial_Status) applying a specific mapping dictionary.
    3. Export (Data Emerge): Maps all computed variables into the final 'DATA.xlsx' using targeted multi-level index alignments.
"""

import os
import pandas as pd
import numpy as np

# =============================================================================
# 1. 环境设置 & 路径定义 / Environment Setup & Path Definition
# =============================================================================
pd.set_option('future.no_silent_downcasting', True)

base_dir = r"C:\Users\26627\Desktop"
source_path_health = os.path.join(base_dir, r"Data\2015\New_2015_Health_Status_and_Functioning_Full_Translated.xlsx")
source_path_demo = os.path.join(base_dir, r"Data\2015\2015_Demographic_Background.xlsx")
target_path = os.path.join(base_dir, "DATA.xlsx")

# ID 列定义 (全局通用) / Define core ID columns globally
col_ind_id = "Individual ID"
col_hh_id = "Household ID"
col_comm_id = "Community ID"
id_cols = [col_ind_id, col_hh_id, col_comm_id]


# =============================================================================
# 通用数据加载函数 / Universal Data Loading Function
# =============================================================================
def load_and_clean_data(file_path, skip_rows=0):
    """
    通用读取并清洗数据：强制将 ID 读取为字符串，跳过说明行，并清除尾部 .0 和空格
    Universal function to load data, skip explanatory rows, and strictly clean IDs as strings.
    """
    try:
        df_full = pd.read_excel(
            file_path,
            header=0,
            dtype={col_ind_id: str, col_hh_id: str, col_comm_id: str}
        )
        # 跳过说明行 / Skip explanatory rows
        df = df_full.iloc[skip_rows:].copy() if skip_rows > 0 else df_full.copy()

        # 清洗 ID / Clean IDs
        for col in id_cols:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

        df.columns = df.columns.astype(str).str.strip()
        print(f"   > [成功/Success] 读取成功 / Read successful. 样本量/Samples: {len(df)}")
        return df
    except Exception as e:
        print(f"   [错误/Error] 读取失败 / Read failed for {file_path}: {e}")
        return pd.DataFrame()


# ==============================================================================
# 第一部分：处理 Health 数据 (饮酒、吸烟、睡眠) / Part 1: Process Health Data
# ==============================================================================
print("\n" + "=" * 50)
print(f"1. 正在处理 Health 数据... / Processing Health data...")
df_health = load_and_clean_data(source_path_health, skip_rows=0)

# --- A. 处理 Drink (饮酒) ---
print("   -> 处理 L2_Drink (饮酒)...")
# 映射逻辑 / Mapping logic: 1->1, 2->1, 3->0
drink_map = {1: 1, 2: 1, 3: 0}
df_health['L2_Drink'] = df_health['Drink_2015'].map(drink_map)

# --- B. 处理 Smoke (吸烟) ---
print("   -> 处理 L2_Smoke (吸烟)...")


def calculate_smoke_level(row):
    """ 计算吸烟状态 (优先级: 数量 > 戒烟状态 > 吸烟史) / Calculate smoking status """
    try:
        s = float(row.get('Smoke', np.nan))
        q = float(row.get('Still Smoke or Quit', np.nan))
        c = float(row.get('The Number of Cigarettes You Consume', np.nan))
    except:
        return np.nan

    # 优先级 1: 吸烟数量 / Priority 1: Cigarette consumption
    if pd.notna(c):
        if c > 20:
            return 1
        elif 0 <= c <= 20:
            return 1

    # 优先级 2: 戒烟状态 / Priority 2: Quitting status
    if pd.notna(q):
        if q == 2:
            return 0
        elif q == 1:
            return 1

    # 优先级 3: 吸烟史 / Priority 3: Smoking history
    if pd.notna(s):
        if s == 2:
            return 0
        elif s == 1:
            return 1

    return np.nan


df_health['L2_Smoke'] = df_health.apply(calculate_smoke_level, axis=1)

# --- C. 处理 Sleep Time (睡眠时间) ---
print("   -> 处理 L2_Sleep_time (睡眠时间)...")
if 'Sleeping Time_night' in df_health.columns:
    night_sleep = pd.to_numeric(df_health['Sleeping Time_night'], errors='coerce')
else:
    print("   [警告/Warning] 未找到 'Sleeping Time_night'，晚间睡眠设为 NaN / Night sleep set to NaN")
    night_sleep = pd.Series(np.nan, index=df_health.index)

if 'During Last Month Time for a Nap after Lunch' in df_health.columns:
    nap_mins = pd.to_numeric(df_health['During Last Month Time for a Nap after Lunch'], errors='coerce')
    nap_hours = nap_mins / 60.0
else:
    print("   [警告/Warning] 未找到午睡列，午睡设为 0 / Nap time set to 0")
    nap_hours = pd.Series(np.nan, index=df_health.index)

# 两者相加 (使用 add 配合 fill_value=0，防止其中一项没填导致整体变 NaN)
# Sum both (use fill_value=0 to prevent NaN if one is missing)
df_health['L2_Sleep_time'] = night_sleep.add(nap_hours, fill_value=0)

# ==============================================================================
# 第二部分：处理 Demographic 数据 (婚姻状态) / Part 2: Process Demographic Data
# ==============================================================================
print("\n" + "-" * 50)
print(f"2. 正在处理 Demographic 数据... / Processing Demographic data...")
# 人口学数据前两行是说明行，跳过2行 / Skip 2 explanatory rows
df_demo = load_and_clean_data(source_path_demo, skip_rows=2)

print("   -> 处理 L2_Martial_Status (婚姻)...")
col_marital = "Martial_Status"

# 映射字典 / Mapping dictionary: 1,2,7 -> 0; 3,4,5,6 -> 1
marital_map = {1: 0, 2: 0, 7: 0, 3: 1, 4: 1, 5: 1, 6: 1}


def calculate_marital(row):
    """ 计算婚姻状态 / Calculate marital status """
    try:
        return marital_map.get(int(float(row.get(col_marital, np.nan))), np.nan)
    except:
        return np.nan


if col_marital in df_demo.columns:
    df_demo['L2_Martial_Status'] = df_demo.apply(calculate_marital, axis=1)
    print("   > L2_Martial_Status 计算完成 / Marital Status completed.")
else:
    print(f"   [严重错误/Error] 源文件中未找到列名 / Column not found: {col_marital}")
    df_demo['L2_Martial_Status'] = np.nan

# ==============================================================================
# 第三部分：分步索引对齐更新 DATA.xlsx / Part 3: Multi-step Data Emerge
# ==============================================================================
print("\n" + "=" * 50)
print("3. 开始执行流水线数据融合 (Data Emerge)... / Starting Pipeline Data Emerge...")

if not os.path.exists(target_path):
    print(f"   [错误/Error] 找不到目标文件 / Target file not found: {target_path}")
    exit()

print("   > 正在读取 DATA.xlsx (保护ID格式) / Reading DATA.xlsx with ID protection...")
df_target = load_and_clean_data(target_path, skip_rows=0)
if df_target.empty: exit()

# --- 策略 A: 更新 Health 变量 (Index: 3个 ID 联合对齐) ---
health_cols_to_update = ['L2_Drink', 'L2_Smoke', 'L2_Sleep_time']
print(f"   [A] 映射 Health 级别变量 (3-ID Index) / Mapping Health variables: {health_cols_to_update}")

df_target.set_index(id_cols, inplace=True)
df_health.set_index(id_cols, inplace=True)

for col in health_cols_to_update:
    if col in df_health.columns:
        if col not in df_target.columns: df_target[col] = np.nan
        df_target[col] = df_health[col]

df_target.reset_index(inplace=True)

# --- 策略 B: 更新 Demographic 变量 (Index: 2个 ID 联合对齐) ---
demo_cols_to_update = ['L2_Martial_Status']
print(f"   [B] 映射 Demographic 级别变量 (2-ID Index) / Mapping Demographic variables: {demo_cols_to_update}")

index_demo = [col_ind_id, col_hh_id]
df_target.set_index(index_demo, inplace=True)
df_demo.set_index(index_demo, inplace=True)

for col in demo_cols_to_update:
    if col in df_demo.columns:
        if col not in df_target.columns: df_target[col] = np.nan
        df_target[col] = df_demo[col]

df_target.reset_index(inplace=True)

# ==============================================================================
# 统计报告与保存 / Statistics Report and Save
# ==============================================================================
print("\n" + "-" * 50)
print(f"【最终融合统计报告 / Final Data Emerge Report】 (总样本数 / Total Samples: {len(df_target)})")

all_updated_cols = health_cols_to_update + demo_cols_to_update
for col in all_updated_cols:
    if col in df_target.columns:
        n_valid = df_target[col].notna().sum()
        n_missing = df_target[col].isna().sum()
        pct = (n_valid / len(df_target)) * 100 if len(df_target) > 0 else 0
        print(f"  [{col}]: 成功写入/Valid: {n_valid} ({pct:.1f}%) | 缺失/Missing: {n_missing}")
    else:
        print(f"  [{col}]: 未创建 / Not created")

print("-" * 50)
print(f"正在保存最终 DATA.xlsx ... / Saving final DATA.xlsx ...")
df_target.to_excel(target_path, index=False)
print("全部流水线任务安全完成！ / All pipeline tasks safely completed!")
print("=" * 50)