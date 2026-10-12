"""
Script Name: CHARLS_Symmetric_PA_Paradox_Pipeline.py
Description:
    Processes CHARLS data for "Physical Activity Paradox" research. / 处理 CHARLS 数据以用于“体力活动悖论”研究。

    Main Workflows / 主要工作流:
    1. Clean & Validate: Corrects 2013 baseline data using 2015 longitudinal records. / 清洗与验证：利用 2015 年纵向记录修正 2013 年基线数据。
    2. Outcomes & Covariates: Computes CVD/Diabetes events (Y1) and comorbidity scores (L1). / 结局与协变量：计算心血管/糖尿病事件 (Y1) 及共病分数 (L1)。
    3. PA Calculation: Flags completely empty PA samples (keeps them with NaN group), calculates PA duration. METs have been completely removed. / PA 计算：标记完全为空的 PA 样本 (保留并归为 NaN 组)，计算 PA 时长。已彻底移除 METs。
    4. Independent Grouping (4-Level): Classifies OPA and LTPA into 4 levels based on duration without any upper limit truncation. / 独立分组 (四分类)：根据时长对 OPA 和 LTPA 独立进行 4 分类，不再设定截断上限。
       - OPA:
         0: Weighted MVPA == 0 (无 MVPA，含仅 LPA / no MVPA, incl. only LPA)
         1: 0 < Weighted MVPA < 75
         2: 75 <= Weighted MVPA < 150
         3: 150 <= Weighted MVPA < 600
         4: Weighted MVPA >= 600
       - LTPA:
         0: No LTPA (Weighted MVPA == 0 & LPA == 0) / 无参加任何 LTPA
         1: Only LPA (Weighted MVPA == 0 & LPA > 0) / 仅参加 LPA
         2: 0 < Weighted MVPA < 150
         3: Weighted MVPA >= 150
    5. Export (Data Update): Directly updates calculated variables (Groups only) into the existing 'DATA.xlsx' by matching 'Individual ID', strictly preserving existing rows without adding new IDs. / 导出 (数据更新)：通过匹配 'Individual ID' 直接将计算变量（仅分组）更新至现有的 'DATA.xlsx'，严格保留现有行，不新增 ID。
"""

import pandas as pd
import numpy as np
import os

# =============================================================================
# 0. 配置路径 / Configuration & Paths
# =============================================================================
pd.set_option('future.no_silent_downcasting', True)

base_dir = r"C:\Users\26627\Desktop\Data\2013"
base_dir_2015 = r"C:\Users\26627\Desktop\Data\2015"
target_dir = r"C:\Users\26627\Desktop"

input_file = os.path.join(base_dir, "2013_Health_Status_and_Functioning_Full_Translated.xlsx")
output_file = os.path.join(base_dir, "New_2013_Health_Status_and_Functioning_Full_Translated.xlsx")
check_file = os.path.join(base_dir, "Check_Calculation_Detail.xlsx")
target_path = os.path.join(target_dir, "DATA.xlsx")

# 定义核心 ID 列 (必须锁定为字符串保留前导零) / Define core ID columns strictly as strings preserving leading zeros
id_cols = ['Individual ID', 'Household ID', 'Community ID']


def clean_ids(df):
    """ 通用 ID 清洗函数 (去空格，去 .0) / Universal ID cleaning function (remove spaces and .0) """
    for col in id_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    return df


# =============================================================================
# 1. 读取原始数据 / Read Raw Data
# =============================================================================
print("=" * 50)
print(f"1. 正在读取原始文件 / Reading original file: {input_file}")

try:
    df = pd.read_excel(input_file, header=0, dtype={col: str for col in id_cols})
    df = df.iloc[3:].reset_index(drop=True)  # 去除说明行 / Remove explanatory rows
    df.columns = df.columns.astype(str).str.strip()  # 先清理列名空格 / Clean column name spaces first
    df = clean_ids(df)
    print(f"   > [成功/Success] 数据读取成功。原始样本量 / Data read successfully. Original sample size: {len(df)}")
except Exception as e:
    print(f"   [错误/Error] 读取错误 / Read error: {e}")
    exit()

# =============================================================================
# 1.05 二次验证 (利用2015年数据修正2013年慢性病) / Second Validation (Using 2015 data to correct 2013 chronic diseases)
# =============================================================================
print("-" * 50)
print("1.05 正在加载 2015 年数据进行慢性病二次验证... / Loading 2015 data for chronic disease validation...")
file_2015 = os.path.join(base_dir_2015, "2015_Health_Status_and_Functioning_Full_Translated.xlsx")

try:
    df_2015 = pd.read_excel(file_2015, header=0, dtype={col: str for col in id_cols})
    df_2015 = df_2015.iloc[3:].reset_index(drop=True)
    df_2015.columns = df_2015.columns.astype(str).str.strip()
    df_2015 = clean_ids(df_2015)

    # 慢性病列定义 / Chronic disease columns definition
    disease_cols = [
        'Chronic Disease_Hypertension', 'Chronic Disease_Dyslipidemia',
        'Chronic Disease_Diabetes or high blood sugar', 'Chronic Disease_ Cancer',
        'Chronic Disease_Chronic lung diseases', 'Chronic Disease_Liver disease',
        'Chronic Disease_coronary heart disease', 'Chronic Disease_Stroke',
        'Chronic Disease_Kidney disease', 'Chronic Disease_Stomach or other digestive disease',
        'Chronic Disease_Emotional, nervous, or psychiatric problems', 'Chronic Disease',
        'Chronic Disease_Arthritis or rheumatism', 'Chronic Disease_Asthma'
    ]

    cols_2015_to_keep = id_cols.copy()
    validation_map = {}
    for c in disease_cols:
        c_2015 = f"Compared to ZIWTime{c}"
        if c_2015 in df_2015.columns:
            cols_2015_to_keep.append(c_2015)
            validation_map[c] = c_2015

    df_2015_val = df_2015[cols_2015_to_keep].copy()
    df = df.merge(df_2015_val, on=id_cols, how='left')

    # 执行验证和倒置修正 / Execute validation and value flipping
    for c, c_2015 in validation_map.items():
        if c in df.columns:
            df[c_2015] = pd.to_numeric(df[c_2015], errors='coerce')
            df[c] = pd.to_numeric(df[c], errors='coerce')

            # 如果 2015 年回答为 2 (错误) / If 2015 response is 2 (wrong)
            mask_wrong = df[c_2015] == 2
            mask_1 = mask_wrong & (df[c] == 1)
            mask_2 = mask_wrong & (df[c] == 2)

            # 翻转数值 / Flip values
            df.loc[mask_1, c] = 2
            df.loc[mask_2, c] = 1

            n_flipped = mask_1.sum() + mask_2.sum()
            n_skipped = df[c_2015].isna().sum()
            print(f"   > [{c}]: 修正了/Corrected {n_flipped}, 跳过/Skipped {n_skipped}.")

    df.drop(columns=list(validation_map.values()), inplace=True)
    print("   > [成功/Success] 慢性病二次验证完成！ / Validation completed!")
except Exception as e:
    print(f"   [警告/Warning] 处理2015验证数据出错: {e}. 跳过 / Error processing 2015 validation data. Skipped.")

# =============================================================================
# 1.1 计算结局事件 (Y1_Event) / Calculate Outcome Events (Y1_Event)
# =============================================================================
print("-" * 50)
print("1.1 计算结局事件 (Y1_Event)... / Calculating Outcome Events...")


def map_outcome_strict(val):
    """ 严格映射结局：1 为发生(1)，2 为未发生(0) / Strict outcome mapping: 1->1, 2->0 """
    try:
        v = float(val)
        return 1 if v == 1 else (0 if v == 2 else np.nan)
    except:
        return np.nan


# 心血管疾病结局 / CVD Outcome
col_heart = "Chronic Disease_coronary heart disease"
col_stroke = "Chronic Disease_Stroke"
if col_heart in df.columns and col_stroke in df.columns:
    df['Y1_Event_CVD'] = pd.concat([df[col_heart].apply(map_outcome_strict),
                                    df[col_stroke].apply(map_outcome_strict)], axis=1).max(axis=1)
else:
    df['Y1_Event_CVD'] = np.nan

# 糖尿病结局 / Diabetes Outcome
col_diab = "Chronic Disease_Diabetes or high blood sugar"
if col_diab in df.columns:
    df['Y1_Event_Diabetes'] = df[col_diab].apply(map_outcome_strict)
else:
    df['Y1_Event_Diabetes'] = np.nan

# =============================================================================
# 2. 疾病剔除逻辑 / Disease Exclusion Criteria
# =============================================================================
print("-" * 50)
print("2. 执行疾病排除 (Exclusion Criteria)... / Executing Disease Exclusion...")

exclusion_cols = [
     "Chronic Disease_ Cancer",
    "Chronic Disease_Kidney disease", "Physical Disabilities",
    "Brain Damage/Mental Retardation", "Vision Problem"
]

for col in exclusion_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')
        n_drop = df[df[col] == 1].shape[0]
        df = df[df[col] != 1].copy()
        print(f"   - 剔除/Dropped '{col}': {n_drop} 个样本/samples, 剩余/Remaining: {len(df)}")

# =============================================================================
# 2.1 计算两类共病分数 / Calculate Comorbidity Scores
# =============================================================================
print("-" * 50)
print("2.1 计算两类共病分数 (L1_disease_scores)... / Calculating Comorbidity Scores...")


def calculate_comorbidity_score(df_target, disease_list, score_col_name):
    """ 计算共病分数 (忽略NaN但若全为NaN则返回NaN) / Calculate score (ignore NaN, return NaN if all NaN) """
    valid_cols = [c for c in disease_list if c in df_target.columns]
    if not valid_cols:
        df_target[score_col_name] = np.nan
        return
    df_temp = df_target[valid_cols].apply(pd.to_numeric, errors='coerce')
    all_nan_mask_score = df_temp.isna().all(axis=1)

    df_calc = df_temp.copy()
    for c in valid_cols:
        df_calc[c] = np.where(df_calc[c] == 1, 1, 0)

    df_target[score_col_name] = df_calc.sum(axis=1)
    df_target.loc[all_nan_mask_score, score_col_name] = np.nan


list_for_cvd = [
    "Chronic Disease_Diabetes or high blood sugar", "Chronic Disease_Hypertension",
    "Chronic Disease_Dyslipidemia", "Chronic Disease_Chronic lung diseases",
    "Chronic Disease_Liver disease", "Chronic Disease_Stomach or other digestive disease",
    "Chronic Disease_Emotional, nervous, or psychiatric problems", "Chronic Disease_Arthritis or rheumatism",
    "Chronic Disease_Asthma", "Chronic Disease_ Cancer", "Chronic Disease_Kidney disease"
]
list_for_diabetes = [
    "Chronic Disease_Hypertension", "Chronic Disease_Dyslipidemia",
    "Chronic Disease_Chronic lung diseases", "Chronic Disease_Liver disease",
    "Chronic Disease_Stomach or other digestive disease", "Chronic Disease_Emotional, nervous, or psychiatric problems",
    "Chronic Disease_Arthritis or rheumatism", "Chronic Disease_Asthma", "Chronic Disease_ Cancer",
    "Chronic Disease_coronary heart disease", "Chronic Disease_Stroke", "Chronic Disease_Kidney disease"
]
calculate_comorbidity_score(df, list_for_cvd, 'L1_disease_scores_CVD_event')
calculate_comorbidity_score(df, list_for_diabetes, 'L1_disease_scores_Diabetes_event')

# =============================================================================
# 3. 标记 PA 全空数据 (不删除) / Flag Empty PA Data (Do NOT drop)
# =============================================================================
print("-" * 50)
print("3. 标记 PA 全空数据 (保留在库中)... / Flagging completely empty PA data...")
time_cols = ['Time_VPA1', 'Time_MPA1', 'Time_LPA1', 'Time_VPA2', 'Time_MPA2', 'Time_LPA2', 'Time_VPA3', 'Time_MPA3',
             'Time_LPA3']
valid_time_cols = [c for c in time_cols if c in df.columns]

for col in valid_time_cols:
    df[col] = pd.to_numeric(df[col], errors='coerce')

# 找出全空样本并生成掩码列 (不执行物理剔除) / Create mask for empty PA samples (no physical dropping)
df['is_empty_pa'] = df[valid_time_cols].isna().all(axis=1)
empty_count = df['is_empty_pa'].sum()
print(
    f"   > [策略更新/Strategy] 发现 PA 数据全空样本 {empty_count} 个。已保留原数据，将在分组时赋予 NaN。 / {empty_count} empty PA samples found. Kept for NaN grouping.")

# =============================================================================
# 4. 计算 PA 时长 / Calculate PA Duration (METs completely removed, Truncation removed)
# =============================================================================
print("-" * 50)
print("4. 计算身体活动时长 (移除截断)... / Calculating PA duration (truncation removed)...")


def calculate_daily_minutes(row, type_prefix):
    """ 计算每日时长 (基于问卷跳转逻辑) / Calculate daily minutes based on skip logic """
    t1 = row.get(f'{type_prefix}1')
    t2 = row.get(f'{type_prefix}2')
    t3 = row.get(f'{type_prefix}3')
    if pd.isna(t1): return 0

    if t1 == 1:  # < 2h
        if t2 == 1:
            return 20
        elif t2 == 2:
            return 75
        else:
            return 30
    elif t1 == 2:  # >= 2h
        if t3 == 1:
            return 180
        elif t3 == 2:
            return 240
        else:
            return 120
    return 0


activity_types = ['VPA', 'MPA', 'LPA']
days_cols_map = {'VPA': 'VPA-Days', 'MPA': 'MPA-Days', 'LPA': 'LPA-Days'}

for act in activity_types:
    df[f'{act}_Daily_Mins'] = df.apply(lambda row: calculate_daily_minutes(row, f'Time_{act}'), axis=1)

    days_col = days_cols_map[act]
    if days_col in df.columns:
        df[days_col] = pd.to_numeric(df[days_col], errors='coerce').fillna(0)
    else:
        df[days_col] = 0

    # 计算每周总时长，不再执行截断 / Calculate total weekly duration, NO truncation applied
    df[f'{act}_Weekly_Mins'] = df[f'{act}_Daily_Mins'] * df[days_col]

for c in ['purpose of VPA', 'purpose of MPA', 'purpose of LPA']:
    if c in df.columns:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

# =============================================================================
# 5. 构建 OPA / LTPA 独立暴露分组 / Construct OPA/LTPA Independent Groupings
# =============================================================================
print("-" * 50)
print("5. 计算独立 OPA 与 LTPA 四分类分组结果... / Calculating independent 4-level OPA/LTPA groups...")

# -----------------------------------------------------------------------------
# A. OPA 五分类评估与分组 / OPA 5-level Evaluation and Grouping
# -----------------------------------------------------------------------------
opa_purposes = [1, 4]

# 提取职业性 VPA, MPA, LPA 周时长 / Extract occupational VPA, MPA, LPA weekly minutes
opa_vpa_weekly = np.where(df['purpose of VPA'].isin(opa_purposes), df['VPA_Weekly_Mins'], 0)
opa_mpa_weekly = np.where(df['purpose of MPA'].isin(opa_purposes), df['MPA_Weekly_Mins'], 0)
opa_lpa_weekly = np.where(df['purpose of LPA'].isin(opa_purposes), df['LPA_Weekly_Mins'], 0)

# 计算 OPA 加权 MVPA / Calculate OPA weighted MVPA (1 min VPA = 2 mins MPA)
opa_weighted_mvpa = opa_vpa_weekly * 2 + opa_mpa_weekly

# OPA 五分类逻辑 / OPA 5-level logic
# 4: Weighted MVPA >= 600
# 3: 150 <= Weighted MVPA < 600
# 2: 75 <= Weighted MVPA < 150
# 1: 0 < Weighted MVPA < 75
# 0: Weighted MVPA == 0 (无 MVPA，含仅 LPA / no MVPA, incl. only LPA)
cond_opa_4 = opa_weighted_mvpa >= 600
cond_opa_3 = (opa_weighted_mvpa >= 150) & (opa_weighted_mvpa < 600)
cond_opa_2 = (opa_weighted_mvpa >= 75) & (opa_weighted_mvpa < 150)
cond_opa_1 = (opa_weighted_mvpa > 0) & (opa_weighted_mvpa < 75)
cond_opa_0 = (opa_weighted_mvpa == 0)

df['A1_OPA_Group'] = np.select(
    [cond_opa_4, cond_opa_3, cond_opa_2, cond_opa_1, cond_opa_0],
    [4.0, 3.0, 2.0, 1.0, 0.0],
    default=np.nan
)
df.loc[df['is_empty_pa'], 'A1_OPA_Group'] = np.nan # 全空样本赋 NaN / Assign NaN for empty samples

# -----------------------------------------------------------------------------
# B. LTPA 四分类评估与分组 / LTPA 4-level Evaluation and Grouping
# -----------------------------------------------------------------------------
ltpa_purposes = [2, 3]

# 提取休闲性 VPA, MPA, LPA 周时长 / Extract leisure-time VPA, MPA, LPA weekly minutes
ltpa_vpa_weekly = np.where(df['purpose of VPA'].isin(ltpa_purposes), df['VPA_Weekly_Mins'], 0)
ltpa_mpa_weekly = np.where(df['purpose of MPA'].isin(ltpa_purposes), df['MPA_Weekly_Mins'], 0)
ltpa_lpa_weekly = np.where(df['purpose of LPA'].isin(ltpa_purposes), df['LPA_Weekly_Mins'], 0)

# 计算 LTPA 加权 MVPA / Calculate LTPA weighted MVPA (1 min VPA = 2 mins MPA)
ltpa_weighted_mvpa = ltpa_vpa_weekly * 2 + ltpa_mpa_weekly

# LTPA 四分类逻辑 / LTPA 4-level logic
# 3: 加权 MVPA >= 150 / Weighted MVPA >= 150
# 2: 0 < 加权 MVPA < 150 / 0 < Weighted MVPA < 150
# 1: 仅参加 LPA / Only LPA (Weighted MVPA == 0 & LPA > 0)
# 0: 无参加任何 LTPA / No LTPA (Weighted MVPA == 0 & LPA == 0)
cond_ltpa_3 = ltpa_weighted_mvpa >= 150
cond_ltpa_2 = (ltpa_weighted_mvpa > 0) & (ltpa_weighted_mvpa < 150)
cond_ltpa_1 = (ltpa_weighted_mvpa == 0) & (ltpa_lpa_weekly > 0)
cond_ltpa_0 = (ltpa_weighted_mvpa == 0) & (ltpa_lpa_weekly == 0)

df['A1_LTPA_Group'] = np.select(
    [cond_ltpa_3, cond_ltpa_2, cond_ltpa_1, cond_ltpa_0],
    [3.0, 2.0, 1.0, 0.0],
    default=np.nan
)
df.loc[df['is_empty_pa'], 'A1_LTPA_Group'] = np.nan # 全空样本赋 NaN / Assign NaN for empty samples

# 打印各组人数 / Print exposure group sizes
print(f"\n   > [统计/Stat] 最终 OPA 五分类暴露分布 / Final OPA 5-level Distribution:")
print(df['A1_OPA_Group'].value_counts(dropna=False).sort_index())
print(f"\n   > [统计/Stat] 最终 LTPA 四分类暴露分布 / Final LTPA 4-level Distribution:")
print(df['A1_LTPA_Group'].value_counts(dropna=False).sort_index())

# =============================================================================
# 6. 保存详细核查表 / Save Intermediate Check Files
# =============================================================================
print("-" * 50)
print("6. 保存核查表与中间文件... / Saving check file and intermediate data...")

check_cols = [
    'Individual ID', 'Household ID', 'Community ID', 'is_empty_pa',
    'A1_OPA_Group', 'A1_LTPA_Group',
    'L1_disease_scores_CVD_event', 'L1_disease_scores_Diabetes_event',
    'Y1_Event_CVD', 'Y1_Event_Diabetes'
]
df_check = df[[c for c in check_cols if c in df.columns]].copy()
df_check.to_excel(check_file, index=False)
df.to_excel(output_file, index=False)
print(f"   > [成功/Success] 全量文件已保存 / Full dataset saved.")

# =============================================================================
# 7. 更新目标文件 DATA.xlsx (Data Update) / Update existing DATA.xlsx
# =============================================================================
print("\n" + "=" * 50)
print("7. 开始基于已有 ID 更新 DATA.xlsx (Data Update)... / Starting DATA.xlsx update based on existing IDs...")

if not os.path.exists(target_path):
    print(f"   [错误/Error] 未找到目标文件，无法进行更新 / Target file not found: {target_path}")
    exit()

# --- 7.1 读取现有的 DATA.xlsx / Read existing DATA.xlsx ---
print(f"   > 步骤 7.1: 读取目标文件... / Reading target file...")
try:
    # 强制锁定 ID 列为 str 类型防科学计数法 / Force ID columns to string to prevent scientific notation
    df_target = pd.read_excel(target_path, dtype={col: str for col in id_cols})

    # 清理列名潜在的空格 / Clean potential spaces in column names
    df_target.columns = df_target.columns.astype(str).str.strip()

    # 统一调用清洗函数 / Call universal cleaning function
    df_target = clean_ids(df_target)

    print(f"   > [成功/Success] 成功加载 DATA.xlsx，当前样本量 / Current sample size: {len(df_target)}")
except Exception as e:
    print(f"   [错误/Error] 读取 DATA.xlsx 失败 / Failed to read DATA.xlsx: {e}")
    exit()

# --- 7.2 以目标文件原有 ID 为基础进行更新 / Update based on existing IDs ---
print("   > 步骤 7.2: 以现存的 Individual ID 为基准精确覆盖并同步删除... / Accurately mapping data and syncing drops...")
merge_key = 'Individual ID'

if merge_key not in df.columns or merge_key not in df_target.columns:
    print(f"   [错误/Error] 缺失 '{merge_key}' 列，无法对齐 / Missing merge key.")
    exit()

# 构建索引映射以防止行列错乱 / Index-based mapping to prevent misalignment
df_source_idx = df.drop_duplicates(subset=[merge_key]).set_index(merge_key)
df_target_idx = df_target.set_index(merge_key)

# ============================================================================
initial_target_len = len(df_target_idx)
df_target_idx = df_target_idx.loc[df_target_idx.index.intersection(df_source_idx.index)]
print(f"   > [同步剔除/Sync Drop] 从 DATA.xlsx 中彻底删除了被排除的样本数: {initial_target_len - len(df_target_idx)}")
# ============================================================================

# 需要更新到 DATA.xlsx 的列名 / Columns to update
update_cols = [
    'L1_disease_scores_CVD_event', 'L1_disease_scores_Diabetes_event',
    'Y1_Event_CVD', 'Y1_Event_Diabetes',
    'A1_OPA_Group', 'A1_LTPA_Group'
]

for col in update_cols:
    # 强制初始化或对齐目标列数据格式为 float64 / Safely initialize target column to float64
    if col not in df_target_idx.columns:
        df_target_idx[col] = pd.Series(np.nan, index=df_target_idx.index, dtype='float64')

    if col in df_source_idx.columns:
        # 提取源列非空数据，统一精度以消除 pandas 警告 / Extract valid source data and cast dtype
        valid_source = df_source_idx[col].dropna().astype('float64')

        # 提取共同的 ID 交集 / Find intersection of IDs
        matched_idx = df_target_idx.index.intersection(valid_source.index)

        # 显式赋值 / Explicit assignment
        if len(matched_idx) > 0:
            df_target_idx.loc[matched_idx, col] = valid_source.loc[matched_idx]

        print(f"     > 成功映射列 / Successfully mapped column: {col} (有效更新数/Updated: {len(matched_idx)})")

# 恢复索引结构 / Restore original index structure
df_target_final = df_target_idx.reset_index()

# 保存 / Save modifications
df_target_final.to_excel(target_path, index=False)

print("-" * 50)
print(f"全部流水线任务安全完成！ DATA.xlsx 已成功更新。 / All tasks completed. DATA.xlsx is updated.")
print(f" -> 已严格保护原始表格结构，并同步彻底删除了疾病剔除的样本 / Original table structure protected, excluded samples permanently dropped.")
print(f" -> 最终目标表剩余总行数 / Final Target table row count: {len(df_target_final)}")
print("=" * 50)
