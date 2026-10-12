"""
Script Name: CHARLS_2015_Wave2_Pipeline.py
Description:
    Processes CHARLS 2015 wave data for longitudinal health outcomes and physical activity. / 处理 CHARLS 2015 波段数据，用于纵向健康结局和体力活动分析。

    Main Workflows / 主要工作流:
    1. Clean & Validate: Reads 2015 baseline data and corrects chronic disease statuses using 2018 longitudinal records. / 清洗与验证：读取 2015 年基线数据，并利用 2018 年纵向记录修正慢性病状态。
    2. Outcomes & Covariates: Computes CVD/Diabetes events (Y2) and comorbidity scores (L2). / 结局与协变量：计算心血管/糖尿病事件 (Y2) 及共病分数 (L2)。
    3. PA Calculation: Flags completely empty PA samples (keeps them with NaN group) and calculates PA duration. METs have been completely removed. / PA 计算：标记完全为空的 PA 样本 (保留并归为 NaN 组)，计算 PA 时长。已彻底移除 METs。
    4. Independent Grouping (4-Level): Classifies OPA and LTPA independently into 4 levels based on duration without any upper limit truncation. / 独立分组 (四分类)：根据时长对 OPA 和 LTPA 独立进行 4 分类，不再设定截断上限。
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
    5. Export (Data Emerge): Maps all derived variables into the final 'DATA.xlsx' using 'Individual ID' for downstream modeling. / 导出 (数据合并)：通过 'Individual ID' 将所有衍生变量映射到最终的 'DATA.xlsx' 中，用于下游建模。
"""

import pandas as pd
import numpy as np
import os
import traceback

# =============================================================================
# 0. 配置路径 / Configuration & Paths
# =============================================================================
pd.set_option('future.no_silent_downcasting', True)

base_dir_2015 = r"C:\Users\26627\Desktop\Data\2015"
base_dir_2018 = r"C:\Users\26627\Desktop\Data\2018"
target_dir = r"C:\Users\26627\Desktop"

input_file = os.path.join(base_dir_2015, "2015_Health_Status_and_Functioning_Full_Translated.xlsx")
output_file = os.path.join(base_dir_2015, "New_2015_Health_Status_and_Functioning_Full_Translated.xlsx")
check_file = os.path.join(base_dir_2015, "Check_Calculation_Detail_2015.xlsx")
target_path = os.path.join(target_dir, "DATA.xlsx")

# 定义 ID 列 (关键：必须锁定为字符串) / Define core ID columns strictly as strings
id_cols = ['Individual ID', 'Household ID', 'Community ID']

def clean_ids(df):
    """ 通用 ID 清洗函数 (去空格，去 .0) / Universal ID cleaning function (remove spaces and .0) """
    for col in id_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
    return df


# =============================================================================
# 1. 读取 2015 原始数据 / Read 2015 Raw Data
# =============================================================================
print("=" * 50)
print(f"1. 正在读取原始文件 / Reading original file: {input_file}")

try:
    df = pd.read_excel(input_file, header=0, dtype={col: str for col in id_cols})
    df = df.iloc[3:].reset_index(drop=True)  # 去除前3行说明行 / Skip 3 explanatory rows
    df = clean_ids(df)
    df.columns = df.columns.astype(str).str.strip()

    print(f"   > [成功/Success] 数据读取成功。原始样本量 / Data read successfully. Original sample size: {len(df)}")
except Exception as e:
    print(f"   [错误/Error] 读取错误 / Read error: {e}")
    exit()


# =============================================================================
# 1.05 二次验证 (利用2018年数据修正2015年慢性病) / Second Validation using 2018 Data
# =============================================================================
print("-" * 50)
print("1.05 正在加载 2018 年数据进行慢性病二次验证... / Loading 2018 data for chronic disease validation...")
file_2018 = os.path.join(base_dir_2018, "2018_Health_Status_and_Functioning_Full_Translated.xlsx")

try:
    df_2018 = pd.read_excel(file_2018, header=0, dtype={col: str for col in id_cols})
    df_2018 = df_2018.iloc[3:].reset_index(drop=True)
    df_2018 = clean_ids(df_2018)
    df_2018.columns = df_2018.columns.astype(str).str.strip()

    # 定义需要验证的14种慢性病列名 / Define 14 chronic diseases for validation
    disease_cols = [
        'Chronic Disease_Hypertension', 'Chronic Disease_Dyslipidemia',
        'Chronic Disease_Diabetes or high blood sugar', 'Chronic Disease_ Cancer',
        'Chronic Disease_Chronic lung diseases', 'Chronic Disease_Liver disease',
        'Chronic Disease_coronary heart disease', 'Chronic Disease_Stroke',
        'Chronic Disease_Kidney disease', 'Chronic Disease_Stomach or other digestive disease',
        'Chronic Disease_Emotional, nervous, or psychiatric problems', 'Chronic Disease',
        'Chronic Disease_Arthritis or rheumatism', 'Chronic Disease_Asthma'
    ]

    cols_2018_to_keep = id_cols.copy()
    validation_map = {}
    rename_dict = {}

    for c in disease_cols:
        c_2018 = f"Compared to ZIWTime{c}"
        if c_2018 in df_2018.columns:
            cols_2018_to_keep.append(c_2018)
            # 加专属后缀防冲突 / Add exclusive suffix to prevent merge conflicts
            new_col_name = f"{c_2018}_2018val"
            validation_map[c] = new_col_name
            rename_dict[c_2018] = new_col_name

    df_2018_val = df_2018[cols_2018_to_keep].copy()
    df_2018_val.rename(columns=rename_dict, inplace=True)

    # 左连接合并 / Left merge validation data
    df = df.merge(df_2018_val, on=id_cols, how='left')

    # 开始执行验证与修正 / Execute validation and correction
    for c, c_2018_val in validation_map.items():
        if c in df.columns and c_2018_val in df.columns:
            df[c_2018_val] = pd.to_numeric(df[c_2018_val], errors='coerce')
            df[c] = pd.to_numeric(df[c], errors='coerce')

            # 99: 从未患有此疾病 -> 修正 2015 为 2 (未患病) / 99: Never had it -> Correct 2015 to 2
            mask_99 = df[c_2018_val] == 99
            # 1,2,3: 确认患有 -> 修正 2015 为 1 (患病) / 1,2,3: Confirmed had it -> Correct 2015 to 1
            mask_123 = df[c_2018_val].isin([1, 2, 3])

            n_flipped_to_2 = (mask_99 & (df[c] != 2)).sum()
            n_flipped_to_1 = (mask_123 & (df[c] != 1)).sum()

            df.loc[mask_99, c] = 2
            df.loc[mask_123, c] = 1
            n_skipped = df[c_2018_val].isna().sum()

            print(f"   > [{c}]: 修正了/Corrected {n_flipped_to_2} 人为未患病(2), 确认了/Confirmed {n_flipped_to_1} 人为患病(1), 跳过/Skipped {n_skipped}.")

    # 清除辅助列 / Drop temporary validation columns
    df.drop(columns=list(validation_map.values()), inplace=True)
    print("   > [成功/Success] 慢性病二次验证与修正全部完成！ / Chronic disease validation completed!")

except Exception as e:
    print(f"   [警告/Warning] 处理2018年验证数据时出错 / Validation error: {type(e).__name__} - {e}")
    print(traceback.format_exc())


# =============================================================================
# 2. 计算结局事件 (Outcome Events - Y2) / Calculate Outcome Events (Y2)
# =============================================================================
print("-" * 50)
print("2. 计算结局事件 (Y2_Event)... / Calculating Outcome Events (Y2)...")

def map_outcome_strict(val):
    """ 严格映射结局：1 为发生(1)，2 为未发生(0)，其余为 NaN
        Strict mapping: 1 -> occurred(1), 2 -> not occurred(0), others -> NaN """
    try:
        v = float(val)
        if v == 1: return 1
        elif v == 2: return 0
        else: return np.nan
    except:
        return np.nan

# CVD (心血管疾病 / Cardiovascular Disease)
col_heart = "Chronic Disease_coronary heart disease"
col_stroke = "Chronic Disease_Stroke"
if col_heart in df.columns and col_stroke in df.columns:
    s_heart = df[col_heart].apply(map_outcome_strict)
    s_stroke = df[col_stroke].apply(map_outcome_strict)
    df['Y2_Event_CVD'] = pd.concat([s_heart, s_stroke], axis=1).max(axis=1)
else:
    df['Y2_Event_CVD'] = np.nan

# Diabetes (糖尿病 / Diabetes)
col_diab = "Chronic Disease_Diabetes or high blood sugar"
if col_diab in df.columns:
    df['Y2_Event_Diabetes'] = df[col_diab].apply(map_outcome_strict)
else:
    df['Y2_Event_Diabetes'] = np.nan

print(f"   > [统计/Stat] Y2_Event_CVD 非空样本数 / Non-null CVD events: {df['Y2_Event_CVD'].notna().sum()}")
print(f"   > [统计/Stat] Y2_Event_Diabetes 非空样本数 / Non-null Diabetes events: {df['Y2_Event_Diabetes'].notna().sum()}")


# =============================================================================
# 3. 计算共病分数 (L2 Disease Scores) / Calculate Comorbidity Scores
# =============================================================================
print("-" * 50)
print("3. 计算两类共病分数... / Calculating Comorbidity Scores (L2)...")

def calculate_comorbidity_score(df_target, disease_list, score_col_name):
    """ 计算共病得分 / Calculate comorbidity score """
    valid_cols = [c for c in disease_list if c in df_target.columns]
    if not valid_cols:
        df_target[score_col_name] = np.nan
        return
    df_temp = df_target[valid_cols].apply(pd.to_numeric, errors='coerce')
    all_nan_mask = df_temp.isna().all(axis=1)

    df_calc = df_temp.copy()
    for c in valid_cols:
        df_calc[c] = np.where(df_calc[c] == 1, 1, 0)

    df_target[score_col_name] = df_calc.sum(axis=1)
    df_target.loc[all_nan_mask, score_col_name] = np.nan

# CVD 共病列表 / List for CVD
list_for_cvd = [
    "Chronic Disease_Diabetes or high blood sugar", "Chronic Disease_Hypertension",
    "Chronic Disease_Dyslipidemia", "Chronic Disease_Chronic lung diseases",
    "Chronic Disease_Liver disease", "Chronic Disease_Stomach or other digestive disease",
    "Chronic Disease_Emotional, nervous, or psychiatric problems", "Chronic Disease_Arthritis or rheumatism",
    "Chronic Disease_Asthma", "Chronic Disease_ Cancer", "Chronic Disease_Kidney disease"
]
calculate_comorbidity_score(df, list_for_cvd, 'L2_disease_scores_CVD_event')

# Diabetes 共病列表 / List for Diabetes
list_for_diabetes = [
    "Chronic Disease_Hypertension", "Chronic Disease_Dyslipidemia",
    "Chronic Disease_Chronic lung diseases", "Chronic Disease_Liver disease",
    "Chronic Disease_Stomach or other digestive disease", "Chronic Disease_Emotional, nervous, or psychiatric problems",
    "Chronic Disease_Arthritis or rheumatism", "Chronic Disease_Asthma", "Chronic Disease_ Cancer",
    "Chronic Disease_coronary heart disease", "Chronic Disease_Stroke", "Chronic Disease_Kidney disease"
]
calculate_comorbidity_score(df, list_for_diabetes, 'L2_disease_scores_Diabetes_event')

print(f"   > [统计/Stat] L2_disease_scores_CVD_event 非空样本数 / Non-null CVD comorbidity scores: {df['L2_disease_scores_CVD_event'].notna().sum()}")


# =============================================================================
# 4. 标记 PA 全空数据 (不删除) / Flag Empty PA Data (Do NOT drop)
# =============================================================================
print("-" * 50)
print("4. 标记 PA 全空数据 (保留在库中)... / Flagging completely empty PA data...")

time_cols = ['Time_VPA1', 'Time_MPA1', 'Time_LPA1', 'Time_VPA2', 'Time_MPA2', 'Time_LPA2', 'Time_VPA3', 'Time_MPA3', 'Time_LPA3']
valid_time_cols = [c for c in time_cols if c in df.columns]

if any(c not in df.columns for c in time_cols):
    print("   [错误/Error] 缺少核心 PA 列 / Missing core PA columns.")
    exit()

for col in valid_time_cols:
    df[col] = pd.to_numeric(df[col], errors='coerce')

# 找出全空样本并生成掩码列 (不执行物理剔除) / Create mask for empty PA samples (no dropping)
df['is_empty_pa'] = df[valid_time_cols].isna().all(axis=1)
empty_count = df['is_empty_pa'].sum()
print(f"   > [策略更新/Strategy] 发现 PA 数据全空样本 {empty_count} 个。已保留原数据，将在分组时赋予 NaN。 / Found {empty_count} completely empty PA samples. Kept for NaN grouping.")


# =============================================================================
# 5. 计算 PA 时长 / Calculate PA Duration (METs completely removed)
# =============================================================================
print("-" * 50)
print("5. 计算身体活动时长 (移除截断限制)... / Calculating PA duration (Truncation removed)...")

def calculate_daily_minutes(row, type_prefix):
    """ 计算每日时长 (基于问卷跳转逻辑) / Calculate daily minutes based on skip logic """
    t1 = row.get(f'{type_prefix}1')
    t2 = row.get(f'{type_prefix}2')
    t3 = row.get(f'{type_prefix}3')
    if pd.isna(t1): return 0

    if t1 == 1:  # < 2h
        if t2 == 1: return 20
        elif t2 == 2: return 75
        else: return 30  # 插补 / Impute
    elif t1 == 2:  # >= 2h
        if t3 == 1: return 180
        elif t3 == 2: return 240
        else: return 120  # 插补 / Impute
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

    # 移除原本的 .clip(upper=2400) 截断，直接保留原始每周时长 / Removed .clip(upper=2400), keep raw weekly minutes
    df[f'{act}_Weekly_Mins'] = df[f'{act}_Daily_Mins'] * df[days_col]

for c in ['purpose of VPA', 'purpose of MPA', 'purpose of LPA']:
    if c in df.columns:
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)


# =============================================================================
# 6. 构建 OPA / LTPA 独立暴露分组 (四分类) / Construct Independent 4-level Groupings
# =============================================================================
print("-" * 50)
print("6. 计算独立 OPA 与 LTPA 四分类分组结果... / Calculating independent 4-level OPA/LTPA groups...")

# -----------------------------------------------------------------------------
# A. OPA 五分类评估与分组 / OPA 5-level Evaluation and Grouping
# -----------------------------------------------------------------------------
opa_purposes = [1, 4]

# 提取职业性 VPA, MPA, LPA 周时长 / Extract occupational weekly minutes
opa_vpa_weekly = np.where(df['purpose of VPA'].isin(opa_purposes), df['VPA_Weekly_Mins'], 0)
opa_mpa_weekly = np.where(df['purpose of MPA'].isin(opa_purposes), df['MPA_Weekly_Mins'], 0)
opa_lpa_weekly = np.where(df['purpose of LPA'].isin(opa_purposes), df['LPA_Weekly_Mins'], 0)

# 计算 OPA 加权 MVPA (1分钟VPA = 2分钟MPA) / Calculate weighted MVPA (1m VPA = 2m MPA)
opa_weighted_mvpa = opa_vpa_weekly * 2 + opa_mpa_weekly

# OPA 五分类逻辑 / OPA 5-level logic
# 4: Weighted MVPA >= 600
# 3: 150 <= Weighted MVPA < 600
# 2: 75 <= Weighted MVPA < 150
# 1: 0 < Weighted MVPA < 75
# 0: Weighted MVPA == 0 (无 MVPA，含仅 LPA / no MVPA, incl. only LPA)
cond_opa = [
    (opa_weighted_mvpa >= 600),
    (opa_weighted_mvpa >= 150) & (opa_weighted_mvpa < 600),
    (opa_weighted_mvpa >= 75) & (opa_weighted_mvpa < 150),
    (opa_weighted_mvpa > 0) & (opa_weighted_mvpa < 75),
    (opa_weighted_mvpa == 0)
]
choice_opa = [4.0, 3.0, 2.0, 1.0, 0.0]

df['A2_OPA_Group'] = np.select(cond_opa, choice_opa, default=np.nan)
df.loc[df['is_empty_pa'], 'A2_OPA_Group'] = np.nan # 全空样本赋 NaN / Assign NaN for empty samples

# -----------------------------------------------------------------------------
# B. LTPA 四分类评估与分组 / LTPA 4-level Evaluation and Grouping
# -----------------------------------------------------------------------------
ltpa_purposes = [2, 3]

# 提取休闲性 VPA, MPA, LPA 周时长 / Extract leisure-time weekly minutes
ltpa_vpa_weekly = np.where(df['purpose of VPA'].isin(ltpa_purposes), df['VPA_Weekly_Mins'], 0)
ltpa_mpa_weekly = np.where(df['purpose of MPA'].isin(ltpa_purposes), df['MPA_Weekly_Mins'], 0)
ltpa_lpa_weekly = np.where(df['purpose of LPA'].isin(ltpa_purposes), df['LPA_Weekly_Mins'], 0)

# 计算 LTPA 加权 MVPA / Calculate weighted MVPA
ltpa_weighted_mvpa = ltpa_vpa_weekly * 2 + ltpa_mpa_weekly

# LTPA 四分类逻辑 / LTPA 4-level logic
# 0: 无参加任何 LTPA (Weighted MVPA == 0 & LPA == 0)
# 1: 仅参加 LPA (Weighted MVPA == 0 & LPA > 0)
# 2: 0 < 加权 MVPA < 150
# 3: 加权 MVPA >= 150
cond_ltpa = [
    (ltpa_weighted_mvpa >= 150),
    (ltpa_weighted_mvpa > 0) & (ltpa_weighted_mvpa < 150),
    (ltpa_weighted_mvpa == 0) & (ltpa_lpa_weekly > 0),
    (ltpa_weighted_mvpa == 0) & (ltpa_lpa_weekly == 0)
]
choice_ltpa = [3.0, 2.0, 1.0, 0.0]

df['A2_LTPA_Group'] = np.select(cond_ltpa, choice_ltpa, default=np.nan)
df.loc[df['is_empty_pa'], 'A2_LTPA_Group'] = np.nan # 全空样本赋 NaN / Assign NaN for empty samples

# 打印各个暴露组的人数分布 / Print the distribution of each exposure group
print(f"\n   > [统计/Stat] 最终 OPA 五分类暴露分布 / Final OPA 5-level Distribution:")
print(df['A2_OPA_Group'].value_counts(dropna=False).sort_index())
print(f"\n   > [统计/Stat] 最终 LTPA 四分类暴露分布 / Final LTPA 4-level Distribution:")
print(df['A2_LTPA_Group'].value_counts(dropna=False).sort_index())


# =============================================================================
# 7. 保存中间文件 / Save Intermediate Files
# =============================================================================
print("-" * 50)
print("7. 保存核查表与中间文件... / Saving check file and intermediate data...")

check_cols = [
    'Individual ID', 'Household ID', 'Community ID', 'is_empty_pa',
    'L2_disease_scores_CVD_event', 'L2_disease_scores_Diabetes_event',
    'Y2_Event_CVD', 'Y2_Event_Diabetes',
    'A2_OPA_Group', 'A2_LTPA_Group'
]
valid_check_cols = [c for c in check_cols if c in df.columns]

df_check = df[valid_check_cols].copy()
df_check.to_excel(check_file, index=False)
print(f"   > [成功/Success] 核查表已保存 / Check file saved: {check_file}")

df.to_excel(output_file, index=False)
print(f"   > [成功/Success] 全量文件已保存 / Full dataset saved: {output_file}")


# =============================================================================
# 8. 更新 DATA.xlsx (Data Update) / Update DATA.xlsx
# =============================================================================
print("\n" + "=" * 50)
print("8. 开始更新 DATA.xlsx (Data Update)... / Starting DATA.xlsx Update...")

if not os.path.exists(target_path):
    print(f"   [错误/Error] 找不到目标文件 / Target file not found: {target_path}")
    exit()

print("   > 正在读取 DATA.xlsx (保护ID格式) / Reading DATA.xlsx with ID protection...")
try:
    df_target = pd.read_excel(target_path, dtype={col: str for col in id_cols})
    df_target = clean_ids(df_target)
    print(f"   > 读取成功，现有样本/Existing samples: {len(df_target)}")
except Exception as e:
    print(f"   [错误/Error] 读取失败 / Read failed: {e}")
    exit()

merge_index_key = 'Individual ID'

if merge_index_key not in df.columns or merge_index_key not in df_target.columns:
    print(f"   [错误/Error] '{merge_index_key}' 列缺失，无法对齐 / Missing merge key.")
    exit()

# 检查重复并设置单索引 / Check duplicates and set Individual ID as index
if df[merge_index_key].duplicated().any():
    print("   [警告/Warning] 源数据 ID 重复，保留首个 / Source ID duplicates found, keeping first.")
    df_source_indexed = df.drop_duplicates(subset=[merge_index_key]).set_index(merge_index_key)
else:
    df_source_indexed = df.set_index(merge_index_key)

if df_target[merge_index_key].duplicated().any():
    print("   [警告/Warning] 目标数据 ID 重复，保留首个 / Target ID duplicates found, keeping first.")
    df_target_indexed = df_target.drop_duplicates(subset=[merge_index_key]).set_index(merge_index_key)
else:
    df_target_indexed = df_target.set_index(merge_index_key)

# 映射字典 (源列名 : 目标列名) / Mapping dictionary
update_map = {
    'L2_disease_scores_CVD_event': 'L2_disease_scores_CVD_event',
    'L2_disease_scores_Diabetes_event': 'L2_disease_scores_Diabetes_event',
    'Y2_Event_CVD': 'Y2_Event_CVD',
    'Y2_Event_Diabetes': 'Y2_Event_Diabetes',
    'A2_OPA_Group': 'A2_OPA_Group',     # 更新独立 OPA 五分类分组 / Update independent 5-level OPA group
    'A2_LTPA_Group': 'A2_LTPA_Group'    # 更新独立 LTPA 四分类分组 / Update independent 4-level LTPA group
}

print("   > 正在执行数据映射与对齐 / Executing data mapping and alignment...")

for src_col, tgt_col in update_map.items():
    if src_col in df_source_indexed.columns:
        if tgt_col not in df_target_indexed.columns:
            df_target_indexed[tgt_col] = np.nan

        # 索引对齐赋值 (源表中包含 NaN 组的全空 PA 样本将直接更新为 NaN，保持目标表结构不变)
        # Assignment by Index (Updates valid rows, empty PA gets mapped as NaN)
        df_target_indexed[tgt_col] = df_source_indexed[src_col]

        common_ids = df_source_indexed.index.intersection(df_target_indexed.index)
        matched_valid_count = df_target_indexed.loc[common_ids, tgt_col].notna().sum()

        print(f"     [{tgt_col}]: 匹配交集/Matched: {len(common_ids)} | 成功写入/Valid mapped: {matched_valid_count}")
    else:
        print(f"     [警告/Warning] 源数据未算出 '{src_col}'，跳过映射 / Column missing in source.")

# 恢复索引并保存 / Reset index and save
df_target_final = df_target_indexed.reset_index()

print("-" * 50)
print(f"正在保存最终 DATA.xlsx ... / Saving final DATA.xlsx...")
df_target_final.to_excel(target_path, index=False)
print("全部流水线任务安全完成！ / All pipeline tasks safely completed!")
print("=" * 50)
