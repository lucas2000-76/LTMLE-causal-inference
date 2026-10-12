"""
Script Name: CHARLS_Lifestyle_and_SES_Covariates_Pipeline.py
Description:
    Processes CHARLS data to compute lifestyle, family, and multi-level socioeconomic covariates.

    Main Workflows:.
    1. BMI & Work Status: Computes BMI and classifies physical work engagements.
    2. Family Living Status: Aggregates child contact frequency.
    3. Household SES: Computes a 4-level SES index using per-capita food expenditure.
    4. Community SES: Computes a community-level SES index using 2011 infrastructure and income data.
    5. Export (Data Emerge): Precisely maps all computed variables into 'DATA.xlsx' using multi-level index and 6-digit Community ID alignments.
"""

import os
import pandas as pd
import numpy as np

# =============================================================================
# 1. 环境设置 & 路径定义 / Environment Setup & Path Definition
# =============================================================================
# 防崩溃设置 / Anti-crash settings
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"
os.environ["OMP_NUM_THREADS"] = "1"

pd.set_option('future.no_silent_downcasting', True)

base_dir_2013 = r"C:\Users\26627\Desktop\Data\2013"
base_dir_2011 = r"C:\Users\26627\Desktop\Data\2011"
target_dir = r"C:\Users\26627\Desktop"

# 2013 输入文件 / 2013 Input files
path_bio = os.path.join(base_dir_2013, "2013_Biomarker.xlsx")
path_work = os.path.join(base_dir_2013, "2013_Work_Retirement_and_Pension.xlsx")
path_family = os.path.join(base_dir_2013, "2013_Family_Information.xlsx")
path_income = os.path.join(base_dir_2013, "2013_Household_Income.xlsx")

# 2011 输入文件 / 2011 Input files
path_psu = os.path.join(base_dir_2011, "2011_psu.xlsx")
path_comm = os.path.join(base_dir_2011, "2011_community.xlsx")

# 输出文件 / Output files
target_path = os.path.join(target_dir, "DATA.xlsx")
output_dir_db = os.path.join(target_dir, "database")
output_file_comm = os.path.join(output_dir_db, "community.xlsx")

if not os.path.exists(output_dir_db):
    os.makedirs(output_dir_db)

# ID 列定义 (全局通用) / Define core ID columns globally
col_ind_id = "Individual ID"
col_hh_id = "Household ID"
col_comm_id = "Community ID"


# =============================================================================
# 通用数据加载函数 / Universal Data Loading Function
# =============================================================================
def load_and_clean_data(file_path, skip_rows):
    """
    通用读取并清洗数据：强制将 ID 读取为字符串，跳过说明行，并清除尾部 .0 和空格
    Universal function to load data, skip explanatory rows, and strictly clean IDs as strings.
    """
    try:
        df_full = pd.read_excel(
            file_path, header=0,
            dtype={col_ind_id: str, col_hh_id: str, col_comm_id: str}
        )
        df = df_full.iloc[skip_rows:].copy()

        # 清洗 ID / Clean IDs
        for col in [col_ind_id, col_hh_id, col_comm_id]:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

        df.columns = df.columns.astype(str).str.strip()
        print(f"   > [成功/Success] 读取成功 / Read successful. 样本量/Samples: {len(df)}")
        return df
    except Exception as e:
        print(f"   [错误/Error] 读取失败 / Read failed for {file_path}: {e}")
        return pd.DataFrame()


# ==============================================================================
# 第一部分：处理 BMI (2013_Biomarker.xlsx) / Part 1: Process BMI
# ==============================================================================
print("\n" + "=" * 50)
print(f"1. 正在处理 BMI 数据... / Processing BMI data...")
df_bio = load_and_clean_data(path_bio, skip_rows=2)


def calculate_bmi(row):
    """ 计算 BMI 并剔除极端异常值 / Calculate BMI and filter extreme outliers """
    try:
        h = pd.to_numeric(row.get("Height"), errors='coerce')
        w = pd.to_numeric(row.get("Weight Measurement"), errors='coerce')
        if pd.notna(h) and pd.notna(w) and 0 < h < 300 and w > 0:
            h_m = h / 100.0
            bmi = w / (h_m ** 2)
            if 10 < bmi < 40: return round(bmi, 2)
    except:
        pass
    return np.nan


if "Height" in df_bio.columns and "Weight Measurement" in df_bio.columns:
    df_bio['L1_BMI'] = df_bio.apply(calculate_bmi, axis=1)
    print("   > BMI 计算完成 / BMI calculation completed.")
else:
    df_bio['L1_BMI'] = np.nan

# ==============================================================================
# 第二部分：处理工作状态 (2013_Work...xlsx) / Part 2: Process Work Status
# ==============================================================================
print("\n" + "-" * 50)
print(f"2. 正在处理工作状态... / Processing Work Status...")
df_work = load_and_clean_data(path_work, skip_rows=2)

col_agri = "Engage in Agricultural Work for More Than 10 Days"
col_other = "Work for at Least One Hour Last WEEK"


def calculate_work_status(row):
    """ 分类工作状态 / Classify work status """
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
                return 2
            elif other == 2:
                return 0
    except:
        pass
    return np.nan


if col_agri in df_work.columns and col_other in df_work.columns:
    df_work['L1_Work_status'] = df_work.apply(calculate_work_status, axis=1)
    print("   > Work Status 计算完成 / Work Status completed.")
else:
    df_work['L1_Work_status'] = np.nan

# ==============================================================================
# 第三部分：处理家庭子女居住地 (Family_Information) / Part 3: Living with Child
# ==============================================================================
print("\n" + "-" * 50)
print(f"3. 正在处理家庭子女居住地... / Processing Living with Child...")
df_family = load_and_clean_data(path_family, skip_rows=3)

# 动态查找所有子女见面频率列 / Dynamically find all child contact frequency columns
child_cols = [c for c in df_family.columns if "How Ofen Do You See Child" in str(c)]


def calculate_living_status(row):
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


if child_cols:
    df_family['L1_Living_with_Child'] = df_family.apply(calculate_living_status, axis=1)
    print("   > L1_Living_with_Child 计算完成 / Living with Child completed.")
else:
    df_family['L1_Living_with_Child'] = np.nan

# ==============================================================================
# 第四部分：处理家庭社会经济地位 (Household Income) / Part 4: Household SES
# ==============================================================================
print("\n" + "-" * 50)
print(f"4. 正在处理家庭社会经济地位... / Processing Household SES...")
df_income = load_and_clean_data(path_income, skip_rows=2)

col_diners = "Number of People Regularly Having Meals"
col_food_exp = "Expenditure on Food"
col_farm = "Produce Agricultural Products by Yourself"

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
    if pd.isna(row[col_food_exp]) or pd.isna(row['per_capita_food']): return np.nan
    tertile, farm = row['food_tertile'], row[col_farm]
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


df_income['L1_Household_SES'] = df_income.apply(calculate_ses_4level, axis=1)
print("   > L1_Household_SES 4级决策计算完成 / Household SES calculated.")

# ==============================================================================
# 第五部分：处理社区级社会经济地位 (2011 PSU & Community) / Part 5: Community SES
# ==============================================================================
print("\n" + "-" * 50)
print(f"5. 正在计算社区社会经济地位 (2011数据)... / Processing 2011 Community SES...")

df_psu = load_and_clean_data(path_psu, skip_rows=2)
df_comm = load_and_clean_data(path_comm, skip_rows=2)

col_area = 'Area Type according to NBS'
col_inc_comm = 'Per-Capita Net Income of Your V/C'
col_road = 'Main Type of Roads'

# 提取并去重 / Extract and drop duplicates
df_part1 = df_psu[[col_comm_id, col_area]].drop_duplicates(subset=[col_comm_id])
df_part2 = df_comm[[col_comm_id, col_inc_comm, col_road]].drop_duplicates(subset=[col_comm_id])

# 合并社区数据 / Merge community data
df_merged_comm = pd.merge(df_part1, df_part2, on=col_comm_id, how='inner')


# A. 社区类型得分 / Area Score
def get_area_score(val):
    try:
        code = int(float(val))
    except:
        return np.nan
    if code == 111:
        return 3
    elif code in [220, 210]:
        return 1
    elif code in [112, 121, 122, 123, 210]:
        return 2
    else:
        return np.nan


df_merged_comm['score_area'] = df_merged_comm[col_area].apply(get_area_score)


# B. 道路得分 / Road Score
def get_road_score(val):
    try:
        code = int(float(val))
    except:
        return np.nan
    if code == 1:
        return 2
    else:
        return 1


df_merged_comm['score_road'] = df_merged_comm[col_road].apply(get_road_score)

# C. 收入得分 / Income Score
df_merged_comm[col_inc_comm] = pd.to_numeric(df_merged_comm[col_inc_comm], errors='coerce')
df_merged_comm['score_income'] = np.nan
mask_has_inc = df_merged_comm[col_inc_comm].notna()

if mask_has_inc.sum() > 0:
    try:
        df_merged_comm.loc[mask_has_inc, 'score_income'] = pd.qcut(
            df_merged_comm.loc[mask_has_inc, col_inc_comm], q=[0, 0.3, 0.7, 1.0], labels=[1, 2, 3]).astype(int)
    except ValueError:
        df_merged_comm.loc[mask_has_inc, 'score_income'] = pd.qcut(
            df_merged_comm.loc[mask_has_inc, col_inc_comm].rank(method='first'), q=[0, 0.3, 0.7, 1.0],
            labels=[1, 2, 3]).astype(int)

# D. 汇总社区 SES / Calculate total Community SES
df_merged_comm['community_level_ses'] = df_merged_comm['score_area'] + df_merged_comm['score_road'] + df_merged_comm[
    'score_income']

# 保存社区中间表 / Save intermediate community table
cols_to_save = [col_comm_id, col_area, 'score_area', col_road, 'score_road', col_inc_comm, 'score_income',
                'community_level_ses']
df_merged_comm[[c for c in cols_to_save if c in df_merged_comm.columns]].to_excel(output_file_comm, index=False)
print(f"   > 社区 SES 计算完成，已保存至 database / Community SES completed and saved.")

# ==============================================================================
# 第六部分：分步索引对齐更新 DATA.xlsx / Part 6: Multi-step Data Emerge
# ==============================================================================
print("\n" + "=" * 50)
print("6. 开始执行流水线数据融合 (Data Emerge)... / Starting Pipeline Data Emerge...")

if not os.path.exists(target_path):
    print(f"   [错误/Error] 找不到目标文件 / Target file not found: {target_path}")
    exit()

try:
    df_target = pd.read_excel(target_path, dtype={col_ind_id: str, col_hh_id: str, col_comm_id: str})
except Exception as e:
    print(f"   [警告/Warning] 尝试基础保护模式读取 / Fallback read mode: {e}")
    df_target = pd.read_excel(target_path, dtype={col_ind_id: str, col_hh_id: str})

# 清洗目标表 ID / Clean Target IDs
for col in [col_ind_id, col_hh_id, col_comm_id]:
    if col in df_target.columns:
        df_target[col] = df_target[col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

# --- 策略 A: 更新 BMI (Index: Individual ID) ---
print("   [A] 映射 Individual 级别变量 / Mapping Individual level variables")
df_target.set_index(col_ind_id, inplace=True)
df_target['L1_BMI'] = df_bio.set_index(col_ind_id)['L1_BMI']
df_target.reset_index(inplace=True)

# --- 策略 B: 更新子女居住 & 家庭SES (Index: Household ID) ---
print("   [B] 映射 Household 级别变量 / Mapping Household level variables")
if col_hh_id in df_family.columns and col_hh_id in df_income.columns:
    family_grouped = df_family.groupby(col_hh_id)['L1_Living_with_Child'].min()
    income_grouped = df_income.groupby(col_hh_id)['L1_Household_SES'].min()

    df_target.set_index(col_hh_id, inplace=True)
    df_target['L1_Living_with_Child'] = family_grouped
    df_target['L1_Household_SES'] = income_grouped
    df_target.reset_index(inplace=True)

# --- 策略 C: 更新工作状态 (Index: Individual ID + Household ID) ---
print("   [C] 映射 Dual-Index 级别变量 / Mapping Dual-Index level variables")
df_target.set_index([col_ind_id, col_hh_id], inplace=True)
df_target['L1_Work_status'] = df_work.set_index([col_ind_id, col_hh_id])['L1_Work_status']
df_target.reset_index(inplace=True)

# --- 策略 D: 更新社区 SES (Index: Community ID 后6位) ---
print("   [D] 映射 Community 级别变量 (基于后6位) / Mapping Community level variables (6-digit ID)")
if col_comm_id in df_target.columns:
    # 提取目标表和社区表 ID 的后 6 位进行匹配 / Map based on last 6 digits of Community ID
    target_id_short = df_target[col_comm_id].str[-6:]
    df_merged_comm['temp_short_id'] = df_merged_comm[col_comm_id].str[-6:]

    # 构建映射字典 / Create mapping dictionary
    ses_map = df_merged_comm.set_index('temp_short_id')['community_level_ses'].to_dict()
    df_target['V_community_level_ses'] = target_id_short.map(ses_map)

    # 单独更新 Area Score 到 DATA.xlsx 的 area_score 列
    area_score_map = df_merged_comm.set_index('temp_short_id')['score_area'].to_dict()
    df_target['area_score'] = target_id_short.map(area_score_map)
else:
    print("   [警告/Warning] DATA.xlsx 中找不到 Community ID，跳过社区级合并 / Missing Community ID.")
# ==============================================================================
# 统计报告与保存 / Statistics Report and Save
# ==============================================================================
print("\n" + "-" * 50)
print(f"【最终融合统计报告 / Final Data Emerge Report】 (总样本数 / Total Samples: {len(df_target)})")

check_cols = ['L1_BMI', 'L1_Work_status', 'L1_Living_with_Child', 'L1_Household_SES', 'V_community_level_ses']
for col in check_cols:
    if col in df_target.columns:
        n_valid = df_target[col].notna().sum()
        n_missing = df_target[col].isna().sum()
        pct = (n_valid / len(df_target)) * 100 if len(df_target) > 0 else 0
        print(f"  [{col}]: 成功匹配/Mapped: {n_valid} ({pct:.1f}%) | 缺失/Missing: {n_missing}")
    else:
        print(f"  [{col}]: 未创建 / Not created")

print("-" * 50)
print(f"正在保存最终 DATA.xlsx ... / Saving final DATA.xlsx ...")
df_target.to_excel(target_path, index=False)
print("全部流水线任务安全完成！ / All pipeline tasks safely completed!")
print("=" * 50)