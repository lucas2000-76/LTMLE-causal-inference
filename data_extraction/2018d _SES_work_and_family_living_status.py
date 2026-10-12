"""
Script Name: CHARLS_2018_Work_Family_SES_Pipeline.py
Description:
    Processes CHARLS 2018 data to compute work, family, and socioeconomic covariates.

    Main Workflows:
    1. Family Living Status: Aggregates child contact frequency to determine household living arrangements (L3_Living_with_Child).
    2. Work Status: Classifies physical work engagements into agricultural and non-agricultural categories (L3_Work_status).
    3. Household SES: Computes a 4-level socioeconomic status index based on per-capita food expenditure and farming (L3_Household_SES).
    4. Export (Data Emerge): Maps all computed variables into the final 'DATA.xlsx' using targeted multi-level index alignments.
"""

import pandas as pd
import numpy as np
import os

# =============================================================================
# 0. 环境设置 & 路径定义 / Environment Setup & Path Definition
# =============================================================================
pd.set_option('future.no_silent_downcasting', True)

base_dir = r"C:\Users\26627\Desktop\Data\2018"
target_dir = r"C:\Users\26627\Desktop"

# 输入文件路径 (2018) / Input file paths
path_work = os.path.join(base_dir, "2018_Work_Retirement.xlsx")
path_family = os.path.join(base_dir, "2018_Family_Transfer.xlsx")
path_income = os.path.join(base_dir, "2018_Household_Income.xlsx")

# 输出文件路径 / Output file path
target_path = os.path.join(target_dir, "DATA.xlsx")

# ID 列定义 (全局通用) / Define core ID columns globally
col_ind_id = "Individual ID"
col_hh_id = "Household ID"
col_comm_id = "Community ID"
id_cols = [col_ind_id, col_hh_id, col_comm_id]


# =============================================================================
# 通用数据加载与清洗函数 / Universal Data Loading & Cleaning Function
# =============================================================================
def load_and_clean_data(file_path, skip_rows=0, is_family_data=False):
    """
    通用读取并清洗数据：强制将 ID 读取为字符串，跳过说明行，并清除尾部 .0 和空格
    Universal function to load data, skip explanatory rows, and strictly clean IDs as strings.
    """
    try:
        df = pd.read_excel(
            file_path, header=0,
            dtype={col_ind_id: str, col_hh_id: str, col_comm_id: str}
        )

        # 特殊处理 Family 数据切片逻辑 (保留原始逻辑) / Special slicing logic for Family data
        if is_family_data and len(df) > 0:
            df = df.iloc[2:].copy() if pd.to_numeric(df.index[0], errors='coerce') != 0 else df.copy()
        elif skip_rows > 0:
            df = df.iloc[skip_rows:].copy()

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
# 第一部分：处理 L3_Living_with_Child (Family Information) / Part 1: Living with Child
# ==============================================================================
print("=" * 50)
print(f"1. 正在处理家庭子女居住信息... / Processing Living with Child...")

df_family = load_and_clean_data(path_family, is_family_data=True)

# 动态查找所有子女见面频率列 / Dynamically find all child contact frequency columns
child_cols = [c for c in df_family.columns if "How Ofen Do You See Child" in str(c)]
print(f"   > 找到子女见面频率列 / Found contact columns: {len(child_cols)} 个")


def calc_living_status(row):
    """ 计算子女同住状态 (取见面最频繁的) / Calculate living status based on most frequent contact """
    vals = []
    for c in child_cols:
        val = pd.to_numeric(row[c], errors='coerce')
        if pd.notna(val) and val != 10:  # 过滤 NaN 和 10(其他) / Filter out NaN and 10
            vals.append(val)

    if not vals: return np.nan
    best_contact = min(vals)

    if best_contact == 1:
        return 0  # 0: 同住(Co-residing)
    elif best_contact <= 3:
        return 1  # 1: 紧密(Close Contact)
    else:
        return 2  # 2: 分离(Separated)


if child_cols and not df_family.empty:
    df_family['L3_Living_with_Child'] = df_family.apply(calc_living_status, axis=1)
    print("   > L3_Living_with_Child 计算完成 / Living with Child completed.")
else:
    print("   [警告/Warning] 未找到子女见面频率列，跳过计算 / Columns not found, skipped.")
    df_family['L3_Living_with_Child'] = np.nan

# ==============================================================================
# 第二部分：处理工作状态 (Work Retirement) / Part 2: Process Work Status
# ==============================================================================
print("\n" + "-" * 50)
print(f"2. 正在处理工作状态... / Processing Work Status...")

# 跳过前2行说明行 / Skip 2 explanatory rows
df_work = load_and_clean_data(path_work, skip_rows=2)

# 提取列名 (使用 L3 特有的列名) / Extract columns
col_agri = "Engage in Agricultural Work for at Least One Month Last Year"
col_other = "Work for at Least One Hour Last WEEK"


def calculate_work_status(row):
    """ 分类工作状态 (0=无业, 1=涉农, 2=纯非农) / Classify work status """
    try:
        agri = float(row.get(col_agri))
        raw_other = row.get(col_other)
        other = float(raw_other) if pd.notna(raw_other) else np.nan

        if agri == 1:
            if other == 1:
                return 1
            elif other == 2 or pd.isna(other):
                return 1
        elif agri == 2:
            if other == 1:
                return 2  # 纯非农 / Pure non-agricultural
            elif other == 2:
                return 0  # 无业 / Unemployed
    except:
        pass
    return np.nan


if col_agri in df_work.columns and col_other in df_work.columns:
    df_work['L3_Work_status'] = df_work.apply(calculate_work_status, axis=1)
    print(f"   > L3_Work_status 计算完成 / Work Status completed.")
else:
    print(f"   [警告/Warning] 未找到指定的工作列 / Specific work columns not found.")
    df_work['L3_Work_status'] = np.nan

# ==============================================================================
# 第三部分：处理 L3_Household_SES (Household Income) / Part 3: Household SES
# ==============================================================================
print("\n" + "-" * 50)
print(f"3. 正在处理家庭社会经济地位 (SES)... / Processing Household SES...")

df_income = load_and_clean_data(path_income, skip_rows=2)

col_diners = "Number of People Regularly Having Meals"
col_food_exp = "Expenditure on Food"
col_farm = "Produce Agricultural Products by Yourself"

if not df_income.empty and all(c in df_income.columns for c in [col_diners, col_food_exp, col_farm]):
    # 强制数值化 / Force numeric
    df_income[col_diners] = pd.to_numeric(df_income[col_diners], errors='coerce')
    df_income[col_food_exp] = pd.to_numeric(df_income[col_food_exp], errors='coerce')
    df_income[col_farm] = pd.to_numeric(df_income[col_farm], errors='coerce')

    # 计算人均食品支出及分位数 / Calculate per capita food expenditure and tertiles
    df_income['diners_fixed'] = np.where((df_income[col_diners].isna()) | (df_income[col_diners] <= 0), 1,
                                         df_income[col_diners])
    df_income['per_capita_food'] = df_income[col_food_exp] / df_income['diners_fixed']

    mask_valid = df_income['per_capita_food'].notna()
    df_income.loc[mask_valid, 'food_rank'] = df_income.loc[mask_valid, 'per_capita_food'].rank(method='first')
    df_income.loc[mask_valid, 'food_tertile'] = pd.qcut(df_income.loc[mask_valid, 'food_rank'], q=3,
                                                        labels=['T1', 'T2', 'T3'])


    def calculate_ses_4level(row):
        """ 计算 4 级 SES 指数 / Calculate 4-level SES index """
        if pd.isna(row.get(col_food_exp)) or pd.isna(row.get('per_capita_food')): return np.nan
        tertile = row.get('food_tertile')
        farm = row.get(col_farm)

        if pd.isna(tertile): return np.nan

        if tertile == 'T3':
            return 1
        elif tertile == 'T2' and farm == 1:
            return 2
        elif tertile == 'T2' and farm == 2:
            return 3
        elif tertile == 'T1' and farm == 1:
            return 3
        elif tertile == 'T1' and farm == 2:
            return 4
        return np.nan


    df_income['L3_Household_SES'] = df_income.apply(calculate_ses_4level, axis=1)
    print(f"   > L3_Household_SES 计算完成 / Household SES calculated.")
else:
    print("   [警告/Warning] Income 文件缺少所需列，跳过计算 / Missing required columns, skipped.")
    df_income['L3_Household_SES'] = np.nan

# ==============================================================================
# 第四部分：分步索引对齐更新 DATA.xlsx / Part 4: Multi-step Data Emerge
# ==============================================================================
print("\n" + "=" * 50)
print("4. 开始执行多表数据融合 (Data Emerge)... / Starting Pipeline Data Emerge...")

if not os.path.exists(target_path):
    print(f"   [错误/Error] 找不到目标文件 / Target file not found: {target_path}")
    exit()

print("   > 正在读取 DATA.xlsx (保护ID格式) / Reading DATA.xlsx with ID protection...")
df_target = load_and_clean_data(target_path, skip_rows=0)
if df_target.empty: exit()

# ------------------------------------------------------------------
# [A] 更新 L3_Living_with_Child (基于 Household ID)
# ------------------------------------------------------------------
print("\n   [A] 映射 L3_Living_with_Child (Household Level)")
if not df_family.empty and 'L3_Living_with_Child' in df_family.columns:
    family_grouped = df_family.groupby(col_hh_id)['L3_Living_with_Child'].min()

    if 'L3_Living_with_Child' not in df_target.columns: df_target['L3_Living_with_Child'] = np.nan

    df_target.set_index(col_hh_id, inplace=True)
    df_target['L3_Living_with_Child'] = family_grouped
    df_target.reset_index(inplace=True)

# ------------------------------------------------------------------
# [B] 更新 L3_Work_status (基于 Individual ID)
# ------------------------------------------------------------------
print("   [B] 映射 L3_Work_status (Individual Level)")
if not df_work.empty and 'L3_Work_status' in df_work.columns:
    if 'L3_Work_status' not in df_target.columns: df_target['L3_Work_status'] = np.nan

    df_target.set_index(col_ind_id, inplace=True)
    df_target['L3_Work_status'] = df_work.set_index(col_ind_id)['L3_Work_status']
    df_target.reset_index(inplace=True)

# ------------------------------------------------------------------
# [C] 更新 L3_Household_SES (基于 Household ID)
# ------------------------------------------------------------------
print("   [C] 映射 L3_Household_SES (Household Level)")
if not df_income.empty and 'L3_Household_SES' in df_income.columns:
    # 提取同一家庭的有效 SES 值（同户人结果一致，用 min 自动忽略 NaN）
    income_grouped = df_income.groupby(col_hh_id)['L3_Household_SES'].min()

    if 'L3_Household_SES' not in df_target.columns: df_target['L3_Household_SES'] = np.nan

    df_target.set_index(col_hh_id, inplace=True)
    df_target['L3_Household_SES'] = income_grouped
    df_target.reset_index(inplace=True)

# ==============================================================================
# 5. 统计报告与保存 / Statistics Report and Save
# ==============================================================================
print("\n" + "-" * 50)
print(f"【最终融合统计报告 / Final Data Emerge Report】 (总样本数 / Total Samples: {len(df_target)})")

final_check_cols = ['L3_Work_status', 'L3_Living_with_Child', 'L3_Household_SES']
for col in final_check_cols:
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