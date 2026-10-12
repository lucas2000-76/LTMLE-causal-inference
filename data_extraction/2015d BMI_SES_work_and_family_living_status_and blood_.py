"""
Script Name: CHARLS_2015_Lifestyle_SES_and_Bio_Covariates_Pipeline.py
Description:
    Processes CHARLS 2015 data to compute lifestyle, family, SES covariates, and biomarker changes.

    Main Workflows:
    1. BMI & Work Status: Computes BMI and classifies physical work engagements.
    2. Family Living Status: Aggregates child contact frequency to determine household living arrangements.
    3. Household SES: Computes a 4-level socioeconomic status index based on food expenditure.
    4. Biomarkers: Computes TyG index and synchronizes blood markers (HbA1c, CRP, LDL, HDL).
    5. Export & Change Calc: Maps all variables into 'DATA.xlsx' and computes longitudinal changes (L2 - Baseline) for biomarkers.
"""

import pandas as pd
import numpy as np
import os

# =============================================================================
# 0. 环境设置 & 路径定义 / Environment Setup & Path Definition
# =============================================================================
pd.set_option('future.no_silent_downcasting', True)

base_dir = r"C:\Users\26627\Desktop\Data\2015"
target_dir = r"C:\Users\26627\Desktop"

# 输入文件路径 / Input file paths
path_work = os.path.join(base_dir, "2015_Work_Retirement_and_Pension.xlsx")
path_family = os.path.join(base_dir, "2015_Family_Information.xlsx")
path_bio_health = os.path.join(base_dir, "2015_Biomarker.xlsx")
path_income = os.path.join(base_dir, "2015_Household_Income.xlsx")
path_blood = os.path.join(base_dir, "2015_Blood.xlsx")

# 输出文件路径 / Output file path
target_path = os.path.join(target_dir, "DATA.xlsx")

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
            file_path, header=0,
            dtype={col_ind_id: str, col_hh_id: str, col_comm_id: str}
        )
        df = df_full.iloc[skip_rows:].copy() if skip_rows > 0 else df_full.copy()

        # 清洗 ID / Clean IDs
        for col in id_cols:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

        df.columns = df.columns.astype(str).str.strip()
        print(f"   > [成功/Success] 读取成功 / Read successful. 样本量/Samples: {len(df)}")
        return df, df_full  # 返回清洗后的 df 和原始 df_full (以备写回 / for writing back)
    except Exception as e:
        print(f"   [错误/Error] 读取失败 / Read failed for {file_path}: {e}")
        return pd.DataFrame(), pd.DataFrame()


# ==============================================================================
# 第一部分：处理 L2_Living_with_Child (Family Information) / Part 1: Living with Child
# ==============================================================================
print("=" * 50)
print(f"1. 正在处理家庭子女居住信息... / Processing Living with Child...")
df_family, _ = load_and_clean_data(path_family, skip_rows=2)

# 动态查找所有子女见面频率列 / Dynamically find all child contact frequency columns
child_cols = [c for c in df_family.columns if "How Ofen Do You See Child" in str(c)]
print(f"   > 找到子女见面频率列 / Found contact columns: {len(child_cols)} 个/count")

def calc_living_status(row):
    """ 计算子女同住状态 (取见面最频繁的) / Calculate living status based on most frequent contact """
    vals = []
    for c in child_cols:
        val = pd.to_numeric(row[c], errors='coerce')
        if pd.notna(val) and val != 10: # 过滤 NaN 和 10(其他) / Filter out NaN and 10
            vals.append(val)

    if not vals: return np.nan
    best_contact = min(vals)

    # 0: 同住(Co-residing), 1: 紧密(Close Contact), 2: 分离(Separated)
    if best_contact == 1: return 0
    elif best_contact <= 3: return 1
    else: return 2

if child_cols:
    df_family['L2_Living_with_Child'] = df_family.apply(calc_living_status, axis=1)
    print("   > L2_Living_with_Child 计算完成 / Living with Child completed.")
else:
    df_family['L2_Living_with_Child'] = np.nan


# ==============================================================================
# 第二部分：处理工作状态 (Work Retirement) / Part 2: Process Work Status
# ==============================================================================
print("\n" + "-" * 50)
print(f"2. 正在处理工作状态... / Processing Work Status...")
df_work, _ = load_and_clean_data(path_work, skip_rows=2)

col_agri = "Engage in Agricultural Work for More Than 10 Days"
col_other = "Work for at Least One Hour Last WEEK"

def calculate_work_status(row):
    """ 分类工作状态 / Classify work status """
    try:
        agri = float(row.get(col_agri))
        raw_other = row.get(col_other)
        other = float(raw_other) if pd.notna(raw_other) else np.nan

        if agri == 1:
            if other == 1: return 1
            elif other == 2 or pd.isna(other): return 1
        elif agri == 2:
            if other == 1: return 2
            elif other == 2: return 0
    except: pass
    return np.nan

if col_agri in df_work.columns and col_other in df_work.columns:
    df_work['L2_Work_status'] = df_work.apply(calculate_work_status, axis=1)
    print("   > Work Status 计算完成 / Work Status completed.")
else:
    df_work['L2_Work_status'] = np.nan


# ==============================================================================
# 第三部分：处理 BMI (Biomarker) / Part 3: Process BMI
# ==============================================================================
print("\n" + "-" * 50)
print(f"3. 正在处理 BMI 数据... / Processing BMI data...")
df_bio, _ = load_and_clean_data(path_bio_health, skip_rows=2)

col_height = "Height"
col_weight = "Weight Measurement"

def calculate_bmi(row):
    """ 计算 BMI 并剔除极端异常值 / Calculate BMI and filter extreme outliers """
    try:
        h = pd.to_numeric(row.get(col_height), errors='coerce')
        w = pd.to_numeric(row.get(col_weight), errors='coerce')
        if pd.notna(h) and pd.notna(w) and 0 < h < 300 and w > 0:
            h_m = h / 100.0
            bmi = w / (h_m ** 2)
            if 10 < bmi < 40: return round(bmi, 2)
    except: pass
    return np.nan

if col_height in df_bio.columns and col_weight in df_bio.columns:
    df_bio['L2_BMI'] = df_bio.apply(calculate_bmi, axis=1)
    print("   > BMI 计算完成 / BMI calculation completed.")
else:
    df_bio['L2_BMI'] = np.nan


# ==============================================================================
# 第四部分：处理 L2_Household_SES (Household Income) / Part 4: Household SES
# ==============================================================================
print("\n" + "-" * 50)
print(f"4. 正在处理家庭社会经济地位 (SES)... / Processing Household SES...")
df_income, _ = load_and_clean_data(path_income, skip_rows=2)

col_diners = "Number of People Regularly Having Meals"
col_food_exp = "Expenditure on Food"
col_farm = "Produce Agricultural Products by Yourself"

df_income[col_diners] = pd.to_numeric(df_income[col_diners], errors='coerce')
df_income[col_food_exp] = pd.to_numeric(df_income[col_food_exp], errors='coerce')
df_income[col_farm] = pd.to_numeric(df_income[col_farm], errors='coerce')

# 计算人均食品支出及分位数 / Calculate per capita food expenditure and tertiles
df_income['diners_fixed'] = np.where((df_income[col_diners].isna()) | (df_income[col_diners] <= 0), 1, df_income[col_diners])
df_income['per_capita_food'] = df_income[col_food_exp] / df_income['diners_fixed']

mask_valid = df_income['per_capita_food'].notna()
df_income.loc[mask_valid, 'food_rank'] = df_income.loc[mask_valid, 'per_capita_food'].rank(method='first')
df_income.loc[mask_valid, 'food_tertile'] = pd.qcut(df_income.loc[mask_valid, 'food_rank'], q=3, labels=['T1', 'T2', 'T3'])

def calculate_ses_4level(row):
    """ 计算 4 级 SES 指数 / Calculate 4-level SES index """
    if pd.isna(row.get(col_food_exp)) or pd.isna(row.get('per_capita_food')): return np.nan
    tertile, farm = row.get('food_tertile'), row.get(col_farm)
    if pd.isna(tertile): return np.nan

    if tertile == 'T3': return 1
    elif tertile == 'T2' and farm == 1: return 2
    elif tertile == 'T2' and farm == 2: return 3
    elif tertile == 'T1' and farm == 1: return 3
    elif tertile == 'T1' and farm == 2: return 4
    return np.nan

df_income['L2_Household_SES'] = df_income.apply(calculate_ses_4level, axis=1)
print("   > L2_Household_SES 4级决策计算完成 / Household SES calculated.")


# ==============================================================================
# 第五部分：计算 2015 血液指标与 TyG / Part 5: Biomarkers, TyG Index & HDL
# ==============================================================================
print("\n" + "-" * 50)
print(f"5. 正在处理血液指标与计算 TyG... / Processing Blood markers and TyG...")
df_blood, df_blood_full = load_and_clean_data(path_blood, skip_rows=2)

# 定义血检源列名与目标列名 / Define blood source and target column names
col_glu = "Glucose (mg/dl)"
col_tg = "Triglycerides (mg/dl)"
col_hdl_src = "HDL Cholesterol (mg/dl)"

col_tyg = "L2_Bio_TyG"
col_hdl_tgt = "L2_Bio_HDL_Pre"

blood_extract_cols = ['L2_Bio_HbA1c_Pre', col_tyg, 'L2_Bio_CRP_Pre', 'L2_Bio_LDL_Pre', col_hdl_tgt]

val_glu = pd.to_numeric(df_blood[col_glu], errors='coerce')
val_tg = pd.to_numeric(df_blood[col_tg], errors='coerce')

# 计算公式: ln[ (Tg * Glu) / 2 ] / Calculate TyG Index
df_blood[col_tyg] = np.log((val_tg * val_glu) / 2)
df_blood[col_tyg] = df_blood[col_tyg].replace([np.inf, -np.inf], np.nan)

# 提取并映射 HDL 胆固醇 / Extract and map HDL Cholesterol
if col_hdl_src in df_blood.columns:
    df_blood[col_hdl_tgt] = pd.to_numeric(df_blood[col_hdl_src], errors='coerce')
else:
    df_blood[col_hdl_tgt] = np.nan

print(f"   > L2 TyG & HDL 处理完成。有效数量 / TyG & HDL processed. Valid TyG count: {df_blood[col_tyg].notna().sum()}")

# 将计算结果保存回源文件 / Save results back to source file
if not df_blood_full.empty:
    for c in [col_tyg, col_hdl_tgt]:
        if c not in df_blood_full.columns: df_blood_full[c] = np.nan
        df_blood_full.loc[df_blood.index, c] = df_blood[c]
    try:
        df_blood_full.to_excel(path_blood, index=False)
        print(f"   > [成功/Success] 源文件更新保存成功 / Source file updated: {path_blood}")
    except: pass

# 提取子集并确保列存在 / Extract subset and ensure columns exist
for col in blood_extract_cols:
    if col not in df_blood.columns: df_blood[col] = np.nan


# ==============================================================================
# 第六部分：分步索引对齐更新与差值计算 / Part 6: Multi-step Data Emerge & Change Calculation
# ==============================================================================
print("\n" + "=" * 50)
print("6. 开始执行流水线数据融合及变动计算... / Starting Data Emerge & Change calculation...")

if not os.path.exists(target_path):
    print(f"   [错误/Error] 找不到目标文件 / Target file not found: {target_path}")
    exit()

print("   > 读取 DATA.xlsx (保护ID格式) / Reading DATA.xlsx with ID protection...")
df_target, _ = load_and_clean_data(target_path, skip_rows=0)
if df_target.empty: exit()

# --- 策略 A: 更新个体级别变量 (BMI, Work, Blood) / Strategy A: Individual Level ---
print("   [A] 映射 Individual 级别变量 / Mapping Individual level variables")
df_target.set_index(col_ind_id, inplace=True)

# 1. 映射 BMI / Map BMI
if not df_bio.empty and 'L2_BMI' in df_bio.columns:
    df_target['L2_BMI'] = df_bio.set_index(col_ind_id)['L2_BMI']

# 2. 映射 工作状态 / Map Work Status
if not df_work.empty and 'L2_Work_status' in df_work.columns:
    df_target['L2_Work_status'] = df_work.set_index(col_ind_id)['L2_Work_status']

# 3. 映射 Blood 变量并计算纵向差值 / Map Blood variables and calc longitudinal changes
if not df_blood.empty:
    df_blood_indexed = df_blood.set_index(col_ind_id)
    for col in blood_extract_cols:
        if col not in df_target.columns: df_target[col] = np.nan
        df_target[col] = df_blood_indexed[col]

    # 定义变动公式: 差值 = 2015(L2) - 2011(V_Baseline) / Change = L2 - Baseline
    change_mappings = [
        ('L2_Bio_HbA1c_change', 'L2_Bio_HbA1c_Pre', 'V_Bio_HbA1c_Pre'),
        ('L2_Bio_TyG_change', 'L2_Bio_TyG', 'V_Bio_TyG'),
        ('L2_Bio_CRP_Pre_change', 'L2_Bio_CRP_Pre', 'V_Bio_CRP_Pre'),
        ('L2_Bio_LDL_Pre_change', 'L2_Bio_LDL_Pre', 'V_Bio_LDL_Pre'),
        ('L2_Bio_HDL_Pre_change', 'L2_Bio_HDL_Pre', 'V_Bio_HDL_Pre') # 新增 HDL 差值计算 / Added HDL change
    ]

    print("   > 正在计算纵向差值 / Calculating longitudinal changes...")
    for change_col, l2_col, v_col in change_mappings:
        if l2_col in df_target.columns and v_col in df_target.columns:
            val_l2 = pd.to_numeric(df_target[l2_col], errors='coerce')
            val_v = pd.to_numeric(df_target[v_col], errors='coerce')
            df_target[change_col] = val_l2 - val_v

df_target.reset_index(inplace=True)

# --- 策略 B: 更新家庭级别变量 (Living_with_Child, SES) / Strategy B: Household Level ---
print("   [B] 映射 Household 级别变量 / Mapping Household level variables")
df_target.set_index(col_hh_id, inplace=True)

if not df_family.empty and 'L2_Living_with_Child' in df_family.columns:
    df_target['L2_Living_with_Child'] = df_family.groupby(col_hh_id)['L2_Living_with_Child'].min()

if not df_income.empty and 'L2_Household_SES' in df_income.columns:
    df_target['L2_Household_SES'] = df_income.groupby(col_hh_id)['L2_Household_SES'].min()

df_target.reset_index(inplace=True)


# ==============================================================================
# 统计报告与保存 / Statistics Report and Save
# ==============================================================================
print("\n" + "-" * 50)
print(f"【最终融合统计报告 / Final Data Emerge Report】 (总样本数 / Total Samples: {len(df_target)})")

check_cols = ['L2_BMI', 'L2_Work_status', 'L2_Living_with_Child', 'L2_Household_SES',
              'L2_Bio_TyG', 'L2_Bio_HDL_Pre', 'L2_Bio_HbA1c_change', 'L2_Bio_TyG_change']

for col in check_cols:
    if col in df_target.columns:
        n_valid = df_target[col].notna().sum()
        print(f"  [{col}]: 成功写入/Valid: {n_valid}")
    else:
        print(f"  [{col}]: 未创建 / Not created")

print("-" * 50)
print(f"正在保存最终 DATA.xlsx ... / Saving final DATA.xlsx ...")
df_target.to_excel(target_path, index=False)
print("全部流水线任务安全完成！ / All pipeline tasks safely completed!")
print("=" * 50)