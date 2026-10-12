"""
Script Name: CHARLS_Blood_Data_Pipeline.py
Description: Calculates blood biomarkers and safely updates the main dataset.

What it does:
1. Calculate: Computes the TyG Index and HDL from the 2011 blood data.
2. Fix IDs: Upgrades old 11-digit IDs into standard 12-digit IDs.
3. Safe Update: Adds the new results into 'DATA.xlsx' accurately, without messing up any existing rows.
"""

import os
import pandas as pd
import numpy as np

# ==========================================
# 1. 环境设置与路径 / Environment Setup and Paths
# ==========================================
pd.set_option('future.no_silent_downcasting', True)

base_dir = r"C:\Users\26627\Desktop"
# 输入文件路径 / Source file path
source_path_health = os.path.join(base_dir, r"Data\2011\Blood_20140429.xlsx")
# 目标文件路径 / Target file path
target_path = os.path.join(base_dir, "DATA.xlsx")

# 列名定义 / Column name definitions
col_ind_id_source = "Individual ID1"  # 源文件中的旧ID (11位) / Old ID in source (11 digits)
col_ind_id_target = "Individual ID"  # 目标文件及新ID (12位) / Target ID (12 digits)
col_hh_id = "Household ID"
col_comm_id = "Community ID"

# ==============================================================================
# 2. 读取源文件并预处理 / Read Source File and Preprocessing
# ==============================================================================
print("-" * 40)
print(f"1. 正在读取源文件: {source_path_health} / Reading source file")

try:
    # 读取源文件 / Read source file (keep full structure for later saving)
    df_source = pd.read_excel(
        source_path_health,
        header=0,
        dtype={col_ind_id_source: str, col_hh_id: str}
    )

    # 切片：跳过前两行说明行 / Slice: skip first two metadata rows for calculation
    df_data = df_source.iloc[2:].copy()

    # 清洗源ID / Clean source IDs (remove spaces and .0)
    if col_ind_id_source in df_data.columns:
        df_data[col_ind_id_source] = df_data[col_ind_id_source].astype(str).str.strip().str.replace(r'\.0$', '',
                                                                                                    regex=True)

    print(f"   > 读取成功。有效计算样本量: {len(df_data)} / Read successful. Sample size: {len(df_data)}")

except Exception as e:
    print(f"读取错误 / Read Error: {e}")
    exit()

# ==============================================================================
# 3. 计算指标 (TyG & HDL) / Calculate Metrics (TyG & HDL)
# ==============================================================================
print("\n" + "=" * 40)
print("2. 计算指标 ... / Calculating Metrics ...")

# 定义列名 / Define column names
col_glu = "Glucose (mg/dl)"
col_tg = "Triglycerides (mg/dl)"
col_hdl_src = "Hdl Cholesterol (mg/dl)"  # 源文件HDL列 / Source HDL column
col_tyg = "V_Bio_TyG"
col_hdl_tgt = "V_Bio_HDL_Pre"  # 目标文件HDL列 / Target HDL column

# 1. 计算 TyG 指数 / Calculate TyG Index
val_glu = pd.to_numeric(df_data[col_glu], errors='coerce')
val_tg = pd.to_numeric(df_data[col_tg], errors='coerce')

# 公式: ln[ (Tg * Glu) / 2 ] / Formula: ln[ (Tg * Glu) / 2 ]
df_data[col_tyg] = np.log((val_tg * val_glu) / 2)
df_data[col_tyg] = df_data[col_tyg].replace([np.inf, -np.inf], np.nan)

# 2. 映射 HDL Cholesterol / Map HDL Cholesterol
df_data[col_hdl_tgt] = pd.to_numeric(df_data[col_hdl_src], errors='coerce')

print(f"   > TyG 计算完成。/ TyG calculation completed.")
print(f"   > HDL 映射完成 ({col_hdl_tgt})。/ HDL mapping completed.")

# ==============================================================================
# 4. ID 转换 (11位 -> 12位) / ID Transformation (11 -> 12 digits)
# ==============================================================================
print("\n" + "=" * 40)
print("3. 执行 ID 转换 (插入 '0') ... / Performing ID conversion (insert '0') ...")


def transform_id(old_id):
    """将11位ID转换为12位 / Converts 11-digit ID to 12-digit"""
    s = str(old_id).strip()
    if len(s) == 11:
        return s[:9] + '0' + s[9:]
    return s


if col_ind_id_source in df_data.columns:
    df_data[col_ind_id_target] = df_data[col_ind_id_source].apply(transform_id)
    print(f"   > ID 转换完成。示例: {df_data[col_ind_id_source].iloc[0]} -> {df_data[col_ind_id_target].iloc[0]}")
else:
    print(f"错误：源文件中找不到列 {col_ind_id_source} / Error: Column not found")
    exit()

# ==============================================================================
# 5. 更新计算结果至源文件 / Saving Results Back to Source File
# ==============================================================================
print("\n" + "=" * 40)
print(f"3.5. 正在将计算结果保存回源文件: {source_path_health}")

# 1. 回填 TyG 和 HDL / Backfill TyG and HDL
for c in [col_tyg, col_hdl_tgt]:
    if c not in df_source.columns:
        df_source[c] = np.nan
    df_source.loc[df_data.index, c] = df_data[c]

# 2. 回填新的 12位 ID / Backfill new 12-digit ID
if col_ind_id_target not in df_source.columns:
    df_source[col_ind_id_target] = np.nan

# 强制转换为 object 以避免类型冲突 / Force to object to avoid type conflicts
df_source[col_ind_id_target] = df_source[col_ind_id_target].astype(object)
df_source.loc[df_data.index, col_ind_id_target] = df_data[col_ind_id_target]

# 3. 保存覆盖原文件 / Save and overwrite source file
try:
    df_source.to_excel(source_path_health, index=False)
    print(f"   > 源文件保存成功！已包含新列 {col_tyg}, {col_hdl_tgt} 和 {col_ind_id_target}")
except Exception as e:
    print(f"   [警告] 源文件保存失败: {e} / [Warning] Source save failed")

# ==============================================================================
# 6. 准备更新子集 (用于 DATA.xlsx) / Prepare Subset for DATA.xlsx
# ==============================================================================
# 定义需要提取到 DATA.xlsx 的列 / Columns to extract to DATA.xlsx
cols_to_extract = ['V_Bio_LDL_Pre', 'V_Bio_CRP_Pre', 'V_Bio_HbA1c_Pre', col_hdl_tgt]
update_cols = cols_to_extract + [col_tyg]

# 确保所有目标列在 df_data 中存在 / Ensure target columns exist in df_data
for col in update_cols:
    if col not in df_data.columns:
        df_data[col] = np.nan

# 提取并去重 / Extract and remove duplicates
df_source_ready = df_data[[col_ind_id_target] + update_cols].copy()
duplicate_count = df_source_ready.duplicated(subset=[col_ind_id_target]).sum()
if duplicate_count > 0:
    print(f"\n   [注意] 发现重复 ID: {duplicate_count} 个 / Duplicates found: {duplicate_count}")

df_source_ready = df_source_ready.drop_duplicates(subset=[col_ind_id_target])
df_source_indexed = df_source_ready.set_index(col_ind_id_target)

# ==============================================================================
# 7. 更新目标文件 (DATA.xlsx) / Update Target File (DATA.xlsx)
# ==============================================================================
print("\n" + "=" * 40)
print(f"4. 开始更新目标文件 (DATA.xlsx) ... / Updating {target_path} ...")

if not os.path.exists(target_path):
    print(f"错误：找不到目标文件 {target_path}")
    exit()

# 1. 读取目标文件 / Read target file
try:
    df_target = pd.read_excel(
        target_path,
        dtype={col_ind_id_target: str, col_hh_id: str, col_comm_id: str}
    )
except Exception:
    df_target = pd.read_excel(target_path, dtype={col_ind_id_target: str})

# 2. 清洗 ID / Clean IDs
if col_ind_id_target in df_target.columns:
    df_target[col_ind_id_target] = df_target[col_ind_id_target].astype(str).str.strip().str.replace(r'\.0$', '',
                                                                                                    regex=True)

# 3. 执行对齐更新 / Execute alignment update
df_target_indexed = df_target.set_index(col_ind_id_target)

# --- 优化后的替换代码 ---
print(f"   > 正在更新列: {update_cols} / Updating columns")
for col in update_cols:
    # 1. 安全初始化：明确指定 dtype 为 float64，避免后续类型冲突
    if col not in df_target_indexed.columns:
        df_target_indexed[col] = pd.Series(np.nan, index=df_target_indexed.index, dtype='float64')

    if col in df_source_indexed.columns:
        # 2. 提取源列中的非空数据，并强制统一精度为 float64
        valid_source = df_source_indexed[col].dropna().astype('float64')

        # 3. 获取目标文件和源文件共同存在的 ID 索引 (Intersection)
        matched_idx = df_target_indexed.index.intersection(valid_source.index)

        # 4. 使用 .loc 进行显式赋值，彻底弃用引发警告的 .update()
        if len(matched_idx) > 0:
            df_target_indexed.loc[matched_idx, col] = valid_source.loc[matched_idx]
# 4. 结果重置与保存 / Reset index and save
df_target_final = df_target_indexed.reset_index()

matched_ids = df_target_indexed.index.intersection(df_source_indexed.index)
print(f"   > 成功匹配更新的样本数: {len(matched_ids)} / Successfully matched samples: {len(matched_ids)}")

print("-" * 30)
print(f"正在保存 DATA.xlsx ... / Saving DATA.xlsx ...")
df_target_final.to_excel(target_path, index=False)
print("全部任务完成！ / All tasks completed!")