"""
Script Name: CHARLS_Baseline_Covariates_Pipeline.py
Description:
    Processes CHARLS data to compute baseline demographic and health covariates.
    Updates these calculated covariates into the existing 'DATA.xlsx' using precise index alignment.

    *Enhanced Version*:
    - Gender uses Majority Vote with fallback priority: Health2015 > Demo_2013 > Health2013.
    - Age uses a sophisticated 3-way distance-based arbitration logic prioritizing 2015 data.

    Main Workflows:
    1. Read Data: Loads 2013 Health, 2013 Demo, and 2015 Health datasets.
    2. Health Covariates: Computes Drink, Smoke, Sleep time, and applies Gender voting.
    3. Demographic Covariates: Computes Education, Marital Status, and robust 3-way Age.
    4. Export (Data Update): Precisely maps derived covariates into the existing 'DATA.xlsx'
       using MultiIndex ('Individual ID', 'Household ID') alignment, ensuring type safety.
"""

import os
import pandas as pd
import numpy as np
from datetime import datetime
import warnings

# =============================================================================
# 1. 环境设置 / Environment Setup
# =============================================================================
warnings.filterwarnings("ignore")
pd.set_option('future.no_silent_downcasting', True)

# =============================================================================
# 2. 定义文件路径 / Define File Paths
# =============================================================================
base_dir_2013 = r"C:\Users\26627\Desktop\Data\2013"
base_dir_2015 = r"C:\Users\26627\Desktop\Data\2015"
target_dir = r"C:\Users\26627\Desktop"

source_path_health = os.path.join(base_dir_2013, "New_2013_Health_Status_and_Functioning_Full_Translated.xlsx")
source_path_demo = os.path.join(base_dir_2013, "2013_Demographic_Background.xlsx")
source_path_health_2015 = os.path.join(base_dir_2015, "New_2015_Health_Status_and_Functioning_Full_Translated.xlsx")
target_path = os.path.join(target_dir, "DATA.xlsx")

id_cols = ['Individual ID', 'Household ID', 'Community ID']
index_cols = ['Individual ID', 'Household ID']

# 全局参考时间 / Global reference date for baseline age
TARGET_DATE = datetime(2013, 7, 1)

# =============================================================================
# 3. 读取数据与标准化 / Read and Standardize Data
# =============================================================================
print("=" * 50)
print("1. 正在读取并标准化数据... / Reading and standardizing data...")


def clean_ids(df):
    """ 清洗 ID 格式 (去空格，去 .0 后缀) / Clean ID formats (remove spaces and .0 suffix) """
    for col in id_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    return df


try:
    dtype_dict = {col: str for col in id_cols}

    print("   -> 读取 2013 Health 数据... / Reading 2013 Health data...")
    df_health = pd.read_excel(source_path_health, header=0, dtype=dtype_dict)
    df_health = clean_ids(df_health)

    print("   -> 读取 2013 Demo 数据... / Reading 2013 Demo data...")
    df_demo_full = pd.read_excel(source_path_demo, header=0, dtype=dtype_dict)
    df_demo = df_demo_full.iloc[2:].copy()
    df_demo = clean_ids(df_demo)

    print("   -> 读取 2015 Health 数据 (用于交叉验证)... / Reading 2015 Health data (for cross-validation)...")
    df_health_2015_full = pd.read_excel(source_path_health_2015, header=0, dtype=dtype_dict)
    df_health_2015 = df_health_2015_full.iloc[2:].copy()
    df_health_2015 = clean_ids(df_health_2015)

    if os.path.exists(target_path):
        print(f"   -> 读取目标文件... / Reading target file: DATA.xlsx")
        df_target = pd.read_excel(target_path, dtype=dtype_dict)
        df_target = clean_ids(df_target)
    else:
        print(f"   [错误/Error] 找不到目标文件 / Target file not found: {target_path}")
        exit()

except Exception as e:
    print(f"   [错误/Error] 读取失败 / Read failed: {e}")
    exit()

# =============================================================================
# 4. 提取 2015 验证特征 (Age & Gender) / Extract 2015 Validation Features
# =============================================================================
print("-" * 50)
print("2. 处理 2015 年数据源交叉验证指标... / Processing 2015 cross-validation metrics...")


def calculate_age_base(y, m, d):
    """ 基础年龄计算函数 (支持复用) / Base age calculation function (reusable) """
    if pd.isna(y) or y <= 0: return np.nan
    fm = int(m) if pd.notna(m) and m > 0 else 6
    fd = int(d) if pd.notna(d) and d > 0 else 15

    try:
        b_date = datetime(int(y), fm, fd)
    except ValueError:
        try:
            b_date = datetime(int(y), fm, 1)
        except ValueError:
            return np.nan

    age = (TARGET_DATE - b_date).days / 365.25
    return round(age, 1) if 0 <= age <= 120 else np.nan


def calculate_age_2015(row):
    """ 基于 2015 年规则计算基线(2013)年龄 / Calculate baseline (2013) age based on 2015 rules """
    choice_val = row.get('When were Your Born', row.get('When were You Born', np.nan))
    try:
        choice = float(choice_val)
    except:
        choice = np.nan

    if choice == 2:
        y_col, m_col, d_col = 'When were You Born: Year', 'When were You Born: Month', 'When were You Born: Day'
    else:
        y_col, m_col, d_col = 'Birth Time on ID Card or Hukou:Year', 'Birth Time on ID Card or Hukou:Month', 'Birth Time on ID Card or Hukou:Day'

    y = pd.to_numeric(row.get(y_col, np.nan), errors='coerce')
    m = pd.to_numeric(row.get(m_col, np.nan), errors='coerce')
    d = pd.to_numeric(row.get(d_col, np.nan), errors='coerce')

    return calculate_age_base(y, m, d)


# 提取 2015 提取出的验证数据 / Extract validation data from 2015
df_2015_cv = df_health_2015[id_cols].copy()
df_2015_cv['Age_2015'] = df_health_2015.apply(calculate_age_2015, axis=1)
df_2015_cv['Gender_2015'] = pd.to_numeric(df_health_2015.get('Gender', np.nan), errors='coerce')

# =============================================================================
# 5. 健康协变量处理 (Health Covariates)
# =============================================================================
print("-" * 50)
print("3. 计算健康相关协变量... / Computing health covariates...")

subset_health = df_health[
    id_cols + ['Gender of R', 'Drink_2013', 'Smoke', 'Still Smoke or Quit', 'The Number of Cigarettes You Consume',
               'Sleeping Time_night', 'During Last Month Time for a Nap after Lunch']].copy()
subset_demo = df_demo[
    id_cols + ['Interviewer Record Gender']].copy() if 'Interviewer Record Gender' in df_demo.columns else df_demo[
    id_cols].copy()

df_health_features = pd.merge(subset_health, subset_demo, on=id_cols, how='left')
df_health_features = pd.merge(df_health_features, df_2015_cv[id_cols + ['Gender_2015']], on=id_cols, how='left')

# --- A. 处理 Gender (多数投票制 / Majority Vote) ---
print("   -> 处理 Gender (规则: 多数举手 -> 兜底)... / Processing Gender (Majority vote with fallback)...")
if 'Interviewer Record Gender' in df_health_features.columns:
    df_health_features.rename(columns={'Interviewer Record Gender': 'Gender_Demo'}, inplace=True)
else:
    df_health_features['Gender_Demo'] = np.nan

df_health_features.rename(columns={'Gender of R': 'Gender_Health'}, inplace=True)
df_health_features['Gender_Demo'] = pd.to_numeric(df_health_features['Gender_Demo'], errors='coerce')
df_health_features['Gender_Health'] = pd.to_numeric(df_health_features['Gender_Health'], errors='coerce')


def majority_vote_gender(row):
    """ 执行性别多数举手与优先级兜底原则 / Execute gender majority vote and priority fallback """
    genders = {
        'H15': row['Gender_2015'],
        'Demo': row['Gender_Demo'],
        'H13': row['Gender_Health']
    }

    # 严格遵循指定的优先级顺序 / Strict priority order
    priority_order = ['H15', 'Demo', 'H13']

    # 按优先级提取有效性别 / Extract valid genders in priority order
    valid_genders = [genders[k] for k in priority_order if pd.notna(genders[k]) and genders[k] in [1, 2]]

    if not valid_genders:
        return np.nan

    counts = pd.Series(valid_genders).value_counts()

    # 规则1: 有明确的多数派 (采纳多数意见) / Rule 1: Clear majority wins
    if len(counts) == 1 or counts.iloc[0] > counts.iloc[1]:
        return counts.index[0]

    # 规则2: 平票或只有一个有效值，按高优顺序兜底 / Rule 2: Tie-breaker using priority order
    return valid_genders[0]


df_health_features['V_Gender'] = df_health_features.apply(majority_vote_gender, axis=1)

# --- B. 处理 Drink (饮酒) / Drink ---
print("   -> 处理 Drink... / Processing Drink...")
df_health_features['L1_Drink'] = df_health_features['Drink_2013'].map({1: 1, 2: 1, 3: 0})

# --- C. 处理 Smoke (吸烟) / Smoke ---
print("   -> 处理 Smoke... / Processing Smoke...")


def calculate_smoke_level(row):
    """ 计算吸烟状态 / Calculate smoking status """
    try:
        s = float(row['Smoke']) if pd.notna(row['Smoke']) else np.nan
        q = float(row['Still Smoke or Quit']) if pd.notna(row['Still Smoke or Quit']) else np.nan
        c = float(row['The Number of Cigarettes You Consume']) if pd.notna(
            row['The Number of Cigarettes You Consume']) else np.nan
    except:
        return np.nan

    if pd.notna(c): return 1 if c >= 0 else np.nan
    if pd.notna(q): return 0 if q == 2 else 1 if q == 1 else np.nan
    if pd.notna(s): return 0 if s == 2 else 1 if s == 1 else np.nan
    return np.nan


df_health_features['L1_Smoke'] = df_health_features.apply(calculate_smoke_level, axis=1)

# --- D. 处理 Sleep Time (睡眠时间) / Sleep Time ---
print("   -> 处理 Sleep Time... / Processing Sleep Time...")
night_sleep = pd.to_numeric(df_health_features['Sleeping Time_night'], errors='coerce')
nap_hours = pd.to_numeric(df_health_features['During Last Month Time for a Nap after Lunch'], errors='coerce') / 60.0
df_health_features['L1_Sleep_time'] = night_sleep.add(nap_hours, fill_value=0)

# =============================================================================
# 6. 人口学协变量处理 (Demographic Covariates)
# =============================================================================
print("-" * 50)
print("4. 计算人口学相关协变量... / Computing demographic covariates...")

df_demo = pd.merge(df_demo, df_2015_cv[id_cols + ['Age_2015']], on=id_cols, how='left')

# --- A. 计算学历 (V_Edu) / Education ---
print("   -> 计算学历 (V_Edu)... / Calculating Education (V_Edu)...")
edu_map = {1: 0, 2: 1, 3: 1, 4: 2, 5: 3, 6: 4, 7: 4, 8: 4, 9: 4, 10: 4, 11: 4}


def calculate_education(row):
    """ 计算最高学历 / Calculate highest education level """
    try:
        val_is_right = float(row.get('Is The Record of Highest Education Level Right', np.nan))
    except:
        val_is_right = np.nan
    raw_edu = np.nan

    if val_is_right == 1:
        raw_edu = row.get('The Highest level of Education Attained', np.nan)
    elif val_is_right == 2:
        try:
            val_change = float(row.get('Have Your Highest Level of Education Changed If So, What Is It', np.nan))
        except:
            val_change = np.nan
        if val_change == 12:
            raw_edu = row.get('The Highest level of Education Attained', np.nan)
        elif pd.notna(val_change):
            raw_edu = val_change
        else:
            raw_edu = row.get('The Highest level of Education Attained Now', np.nan)
    elif pd.isna(val_is_right):
        raw_edu = row.get('The Highest level of Education Attained Now', np.nan)

    try:
        return edu_map.get(int(float(raw_edu)), np.nan) if pd.notna(raw_edu) else np.nan
    except:
        return np.nan


df_demo['V_Edu'] = df_demo.apply(calculate_education, axis=1)

# --- B. 计算年龄 (L1_Age: 动态三方距离裁决) / Age (3-way distance-based arbitration) ---
print("   -> 计算年龄 (L1_Age - 执行三方距离动态裁决)... / Calculating Age (3-way arbitration)...")


def calculate_final_age(row):
    """ 基于三方距离和优先级的年龄终极裁决 / Final age arbitration based on 3-way distance and priority """
    # 提取 2013_A (记录生日) / Extract 2013_A (Recorded birthday)
    y_a = pd.to_numeric(row.get('When Were You Born: Year', np.nan), errors='coerce')
    m_a = pd.to_numeric(row.get('When Were You Born: Month', np.nan), errors='coerce')
    d_a = pd.to_numeric(row.get('When Were You Born: Day', np.nan), errors='coerce')
    age_13a = calculate_age_base(y_a, m_a, d_a)

    # 提取 2013_B (更改后生日) / Extract 2013_B (Changed birthday)
    y_b = pd.to_numeric(row.get('When were You Born:Year1', np.nan), errors='coerce')
    m_b = pd.to_numeric(row.get('When were You Born:Month1', np.nan), errors='coerce')
    d_b = pd.to_numeric(row.get('When were You Born:Day1', np.nan), errors='coerce')
    age_13b = calculate_age_base(y_b, m_b, d_b)

    # 提取 2015 验证年龄 / Extract 2015 validation age
    age_15 = row.get('Age_2015', np.nan)

    # 计算原 2013 逻辑兜底值 / Calculate original 2013 fallback value
    try:
        is_right = float(row.get('We Record Your Birth Time is [Birth Time], Is It Right', np.nan))
    except:
        is_right = np.nan

    age_orig = age_13a if is_right == 1 else age_13b
    # 防止严格映射导致的 NaN，取两者间有效的那个 / Prevent NaN from strict mapping, take the valid one
    if pd.isna(age_orig):
        age_orig = age_13a if pd.notna(age_13a) else age_13b

    # 聚合所有有效年龄值 / Aggregate all valid age values
    ages = {'15': age_15, '13a': age_13a, '13b': age_13b}
    valid_ages = {k: v for k, v in ages.items() if pd.notna(v)}

    # 规则 1 & 2: 如果存在 3 个有效值，按距离裁决 / Rule 1 & 2: 3 valid values, arbitrate by distance
    if len(valid_ages) == 3:
        dist_13a_13b = abs(valid_ages['13a'] - valid_ages['13b'])
        dist_15_13a = abs(valid_ages['15'] - valid_ages['13a'])
        dist_15_13b = abs(valid_ages['15'] - valid_ages['13b'])

        # 如果 2013 两个数值更近，采用原 2013 逻辑 / If 2013 values are closer, use 2013 logic
        if dist_13a_13b < min(dist_15_13a, dist_15_13b):
            return age_orig
        else:
            # 否则以 2015 为准 / Otherwise, prioritize 2015
            return valid_ages['15']

    # 规则 3: 只有 2 个有效值，只要 2015 存在就以其为主 / Rule 3: 2 valid values, prioritize 2015 if present
    elif len(valid_ages) == 2:
        if '15' in valid_ages:
            return valid_ages['15']
        else:
            return age_orig

    # 规则 4: 只有 1 个有效值，直接使用 / Rule 4: 1 valid value, use it directly
    elif len(valid_ages) == 1:
        return list(valid_ages.values())[0]

    return np.nan


df_demo['L1_Age'] = df_demo.apply(calculate_final_age, axis=1)

# --- C. 计算婚姻 (L1_Martial_Status) / Marital Status ---
print("   -> 计算婚姻 (L1_Martial_Status)... / Calculating Marital Status (L1_Martial_Status)...")


def calculate_marital(row):
    """ 计算婚姻状态 / Calculate marital status """
    try:
        return {1: 0, 2: 0, 7: 0, 3: 1, 4: 1, 5: 1, 6: 1}.get(int(float(row.get('Martial Status', np.nan))), np.nan)
    except:
        return np.nan


df_demo['L1_Martial_Status'] = df_demo.apply(calculate_marital, axis=1)

# =============================================================================
# 7. 保存结果回 Demo 源文件 / Save Results back to Demo Source File
# =============================================================================
print("-" * 50)
print("5. 正在保存结果回 Demo 源文件... / Saving results back to Demo source file...")

df_demo_full.loc[df_demo.index, 'V_Edu'] = df_demo['V_Edu']
df_demo_full.loc[df_demo.index, 'L1_Age'] = df_demo['L1_Age']
df_demo_full.loc[df_demo.index, 'L1_Martial_Status'] = df_demo['L1_Martial_Status']

df_demo_full.to_excel(source_path_demo, index=False)
print(f"   > [成功/Success] 源文件更新完成 / Source file updated: {source_path_demo}")

# =============================================================================
# 8. 基于索引精准更新 DATA.xlsx (Data Emerge) / Precise Update via Index Alignment
# =============================================================================
print("\n" + "=" * 50)
print("6. 开始执行索引对齐更新 / Starting Index Alignment Update (DATA.xlsx) ...")

health_update_cols = ['V_Gender', 'L1_Drink', 'L1_Smoke', 'L1_Sleep_time']
demo_update_cols = ['V_Edu', 'L1_Age', 'L1_Martial_Status']
all_update_cols = health_update_cols + demo_update_cols

print(f"   > 准备映射列 / Preparing to map columns: {all_update_cols}")

# 设置多重索引 / Set MultiIndex
df_health_indexed = df_health_features.drop_duplicates(subset=index_cols).set_index(index_cols)
df_demo_indexed = df_demo.drop_duplicates(subset=index_cols).set_index(index_cols)
df_target_indexed = df_target.set_index(index_cols)

# 统一的安全对齐与赋值逻辑 / Unified safe alignment and assignment logic
for col in all_update_cols:
    # 强制初始化目标列格式为 float64 以防止类型警告 / Initialize target col to float64
    if col not in df_target_indexed.columns:
        df_target_indexed[col] = pd.Series(np.nan, index=df_target_indexed.index, dtype='float64')

# 映射 Health 协变量 / Map Health covariates
for col in health_update_cols:
    if col in df_health_indexed.columns:
        valid_source = df_health_indexed[col].dropna().astype('float64')
        matched_idx = df_target_indexed.index.intersection(valid_source.index)
        if len(matched_idx) > 0:
            df_target_indexed.loc[matched_idx, col] = valid_source.loc[matched_idx]

# 映射 Demo 协变量 / Map Demo covariates
for col in demo_update_cols:
    if col in df_demo_indexed.columns:
        valid_source = df_demo_indexed[col].dropna().astype('float64')
        matched_idx = df_target_indexed.index.intersection(valid_source.index)
        if len(matched_idx) > 0:
            df_target_indexed.loc[matched_idx, col] = valid_source.loc[matched_idx]

# 重置索引 / Reset index
df_target_final = df_target_indexed.reset_index()

# =============================================================================
# 9. 打印更新统计报告并保存 / Print Statistics and Save
# =============================================================================
print("-" * 50)
print("【最终统计报告 / Final Statistical Report】(DATA.xlsx)")
print(f"总样本数 / Total samples: {len(df_target_final)}")

for col in all_update_cols:
    n_valid = df_target_final[col].notna().sum()
    n_missing = df_target_final[col].isna().sum()
    print(f"  [{col}]: 有效/Valid: {n_valid} | 缺失/Missing: {n_missing}")

print("-" * 50)
print(f"正在保存最终 DATA.xlsx ... / Saving final DATA.xlsx ...")
df_target_final.to_excel(target_path, index=False)
print("全部任务完美竣工！ / All tasks completed flawlessly!")
print("=" * 50)