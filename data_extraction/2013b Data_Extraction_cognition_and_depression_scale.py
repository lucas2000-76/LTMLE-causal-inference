"""
Script Name: CHARLS_Mental_Health_Covariates_Pipeline.py
Description:
    Processes CHARLS data to compute mental health and cognitive covariates.
    Main Workflows:
    1. Cognitive Functioning: Computes cognitive scores across multiple dimensions (orientation, math, drawing, memory). Applies a validity threshold (requires >= 5 valid answers).
    2. Depression Scale (CES-D-10): Computes depression scores using reverse-scoring logic. Applies a validity threshold (requires >= 8 valid answers).
    3. Export (Data Emerge): Maps all variables into the final 'DATA.xlsx' for downstream modeling.
"""

import pandas as pd
import numpy as np
import os

# =============================================================================
# 配置路径 / Configuration & Paths
# =============================================================================
base_dir = r"C:\Users\26627\Desktop\Data\2013"
input_file = os.path.join(base_dir, "New_2013_Health_Status_and_Functioning_Full_Translated.xlsx")
# 覆盖保存源文件 / Overwrite original file
output_file = input_file

# =============================================================================
# 1. 读取数据 / Read Data
# =============================================================================
print("=" * 50)
print(f"1. 正在读取文件 / Reading file: {input_file}")

# 定义 ID 列 (必须与 Excel 列名完全一致) / Define ID columns (Must match Excel exactly)
id_cols = ['Individual ID', 'Household ID', 'Community ID']

try:
    # 【核心修改】读取源文件时，必须强制指定 ID 为字符串！
    # [Core Fix] Force ID columns to be read as strings to prevent losing leading zeros (e.g., "0101" -> 101)
    df = pd.read_excel(
        input_file,
        dtype={col: str for col in id_cols}
    )

    # 【额外保险】再次清洗 ID (防止有 .0 或 空格)
    # [Extra Insurance] Clean IDs again to remove spaces or trailing ".0"
    for col in id_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.replace(r'\.0$', '', regex=True)

except FileNotFoundError:
    print(f"   [错误/Error] 找不到文件 / File not found: {input_file}")
    exit()
except Exception as e:
    # 如果指定 dtype 失败(比如列名不对)，尝试普通读取但打印警告
    # Fallback to normal read if dtype fails, but print a warning
    print(f"   [警告/Warning] 带格式读取失败 / Formatted read failed ({e}).")
    print("   > 尝试普通读取（ID前导零可能会丢失）/ Attempting normal read (Leading zeros may be lost)...")
    df = pd.read_excel(input_file)

# 清洗列名 / Clean column names
df.columns = df.columns.astype(str).str.strip()
print(f"   > [统计/Stat] 样本量 / Sample size: {len(df)}")

# =============================================================================
# 2. 计算认知能力 (L1_Cognitive_functioning) / Calculate Cognitive Functioning
# =============================================================================
print("-" * 50)
print("2. 计算认知能力量表... / Calculating Cognitive Functioning Score...")

# 收集所有认知相关的列 / Collect all cognitive-related columns
cog_cols_all = []

# A. 日期 / Orientation (Date)
date_cols = ['Checking Date_1', 'Checking Date_2', 'Checking Date_3']
cog_cols_all.extend([c for c in date_cols if c in df.columns])

# B. 星期季节 / Orientation (Day & Season)
week_cols = ['Checking Day of Week', 'Checking Season']
cog_cols_all.extend([c for c in week_cols if c in df.columns])

# C. 减法 / Math (Serial 7s)
math_cols = ['100 Minus 7_1', 'Minus 7_2', 'Minus 7_3', 'Minus 7_4', 'Minus 7_5']
cog_cols_all.extend([c for c in math_cols if c in df.columns])

# D. 画图 / Drawing
if 'Draw Picture' in df.columns:
    cog_cols_all.append('Draw Picture')

# E. 词汇回忆 / Word Recall
word_cols = [f'Recalled Words_{i}' for i in range(11, 21)] + [f'Recall Words_{i}' for i in range(1, 11)]
cog_cols_all.extend([c for c in word_cols if c in df.columns])

# 检查缺失的列 / Check for missing columns
missing_cog_cols = set(date_cols + week_cols + math_cols + word_cols + ['Draw Picture']) - set(df.columns)
if missing_cog_cols:
    print(f"   [提示/Info] 数据集中缺失以下认知列 / Missing cognitive columns: {missing_cog_cols}")

# 初始化得分 (全为0) / Initialize scores to 0
cog_score = pd.Series(0, index=df.index)

# 计算逻辑：有回答且正确加1 / Scoring Logic: Add 1 for each correct answer
# ------------------------------------
# A. 日期 (1,2,3) / Date
for col, target in zip(['Checking Date_1', 'Checking Date_2', 'Checking Date_3'], [1, 2, 3]):
    if col in df.columns:
        cog_score += np.where(pd.to_numeric(df[col], errors='coerce') == target, 1, 0)

# B. 星期季节 (1) / Day & Season
for col in ['Checking Day of Week', 'Checking Season']:
    if col in df.columns:
        cog_score += np.where(pd.to_numeric(df[col], errors='coerce') == 1, 1, 0)

# C. 减法 (93, 86...) / Math
math_targets = [93, 86, 79, 72, 65]
for col, target in zip(math_cols, math_targets):
    if col in df.columns:
        cog_score += np.where(pd.to_numeric(df[col], errors='coerce') == target, 1, 0)

# D. 画图 (1) / Drawing
if 'Draw Picture' in df.columns:
    cog_score += np.where(pd.to_numeric(df['Draw Picture'], errors='coerce') == 1, 1, 0)

# E. 词汇回忆 (1-10) / Word Recall
word_targets = list(range(1, 11)) + list(range(1, 11))
for col, target in zip(word_cols, word_targets):
    if col in df.columns:
        val = pd.to_numeric(df[col], errors='coerce')
        cog_score += np.where(val == target, 1, 0)

# 【核心修改】统计有效回答数量：有数据的题目数量少于 5 个，则总分记为 NaN
# [Core Logic] Count valid answers: If fewer than 5 valid answers, score is NaN
valid_answers_count = df[cog_cols_all].notna().sum(axis=1)
invalid_mask = valid_answers_count < 5

# 将无效样本的分数设为 NaN / Set invalid samples to NaN
cog_score = cog_score.astype(float)
cog_score[invalid_mask] = np.nan

# 写入认知分 / Assign to DataFrame
df['L1_Cognitive_functioning'] = cog_score

valid_cog_count = df['L1_Cognitive_functioning'].notna().sum()
print(f"   > [成功/Success] 认知分计算完毕 / Cognitive score calculated.")
print(f"     - 有效样本 / Valid samples (>= 5 answers): {valid_cog_count}")
print(f"     - 无效样本 / Invalid samples (< 5 answers, set to NaN): {invalid_mask.sum()}")
print(f"     - 平均分 / Mean: {df['L1_Cognitive_functioning'].mean():.2f}")

# =============================================================================
# 3. 抑郁量表 (L1_Depression_scale) / Calculate Depression Scale (CES-D-10)
# =============================================================================
print("-" * 50)
print("3. 处理 CES-D-10 抑郁量表... / Processing CES-D-10 Depression Scale...")

neg_cols = [
    "Bothered by Things", "Had Trouble Keeping Mind", "Felt Depressed",
    "I Felt Everything I Did Was An Effort.", "I Felt Fearful.",
    "My Sleep Was Restless.", "I Felt Lonely.", "I Could Not Get"
]
pos_cols = ["I Felt Hopeful about The Future.", "I Was Happy."]
all_cesd_cols = neg_cols + pos_cols

if any(c not in df.columns for c in all_cesd_cols):
    print("   [错误/Error] 缺少 CES-D 核心列 / Missing core CES-D columns.")
else:
    # --- A. 映射分数 / Map Scores ---
    df_cesd = df[all_cesd_cols].copy()
    for col in all_cesd_cols:
        df_cesd[col] = pd.to_numeric(df_cesd[col], errors='coerce')

    # 反向计分 (消极题) / Scoring for negative questions
    for col in neg_cols:
        df_cesd[col] = df_cesd[col].map({1: 0, 2: 1, 3: 2, 4: 3})

    # 正向计分 (积极题) / Reverse scoring for positive questions
    for col in pos_cols:
        df_cesd[col] = df_cesd[col].map({1: 3, 2: 2, 3: 1, 4: 0})

    # --- B. 统计有效回答数 / Count valid answers ---
    valid_cesd_count = df_cesd.count(axis=1)

    # --- C. 计算总分 / Calculate raw sum ---
    raw_cesd_sum = df_cesd.sum(axis=1, min_count=1)

    # --- D. 应用筛选规则 / Apply validity rule ---
    # 如果有效回答数 < 8，强制设为 NaN / Set to NaN if < 8 valid answers
    final_cesd_score = raw_cesd_sum.where(valid_cesd_count >= 8, np.nan)

    df['L1_Depression_scale'] = final_cesd_score

    # 统计信息 / Statistics
    n_valid_dep = df['L1_Depression_scale'].notna().sum()
    n_dropped_dep = len(df) - n_valid_dep

    print(f"   > [成功/Success] 抑郁得分计算完毕 / Depression score calculated.")
    print(f"     - 有效样本 / Valid samples (>= 8 answers): {n_valid_dep}")
    print(f"     - 无效样本 / Invalid samples (< 8 answers, set to NaN): {n_dropped_dep}")
    print(f"     - 平均分 / Mean: {df['L1_Depression_scale'].mean():.2f}")

# =============================================================================
# 4. 保存计算结果 / Save Intermediate Results
# =============================================================================
print("-" * 50)
print("4. 保存计算中间结果 / Saving calculated results...")
print(f"   > 正在保存至 / Saving to: {output_file}")
df.to_excel(output_file, index=False)
print("   > [成功/Success] 保存完毕 / Save complete.")

# =================================================================================================
# 第二部分：精准更新 DATA.xlsx (Data Emerge) / Part II: Data Emerge to DATA.xlsx
# =================================================================================================
print("\n" + "=" * 50)
print("开始执行第二部分：精准更新 DATA.xlsx / Starting Part II: Precisely updating DATA.xlsx")

target_dir = r"C:\Users\26627\Desktop"
target_path = os.path.join(target_dir, "DATA.xlsx")

# 1. 检查目标文件 / Check target file existence
if not os.path.exists(target_path):
    print(f"   [错误/Error] 目标文件不存在 / Target file does not exist: {target_path}")
    exit()

print("   > 正在读取 DATA.xlsx ... / Reading DATA.xlsx ...")

# ==============================================================================
# 【核心修改 1】读取时强制指定 ID 列为字符串 / Force ID columns to be strings
# ==============================================================================
# 构建 dtype 字典: {'Individual ID': str, 'Household ID': str, ...}
# Builds dtype dict to strictly preserve leading zeros during read
try:
    df_target = pd.read_excel(target_path, dtype={col: str for col in id_cols})
    print(
        f"   > [成功/Success] 读取成功 (已强制 ID 为文本) / Read successful (IDs forced as text). 行数/Rows: {len(df_target)}")
except ValueError:
    # 容错：万一文件里没有这些列，先普通读取 / Fallback to normal read if error occurs
    print("   [警告/Warning] 读取时指定 dtype 失败，尝试普通读取 / dtype parsing failed, attempting normal read...")
    df_target = pd.read_excel(target_path)

# 2. 定义要更新的列 / Columns to update
update_cols = ['L1_Cognitive_functioning', 'L1_Depression_scale']

# 3. 执行更新逻辑 / Execute update logic
print("   > 正在依据 ID 键值更新数据 / Updating data based on ID mapping...")

# 检查 df_target 是否包含我们要更新的列，如果没有则新建
# Check if target columns exist, create them as NaN if they don't
for col in update_cols:
    if col not in df_target.columns:
        print(f"     - 目标文件中未找到 '{col}'，已新建此列 / Column created.")
        df_target[col] = np.nan

# 【核心逻辑】使用 set_index 进行精准映射 / [Core Logic] Use index matching for precise data emerge

# 1. 将目标表设置为 ID 索引 / Set target dataframe index to IDs
df_target_indexed = df_target.set_index(id_cols)

# 2. 将源数据表设置为 ID 索引 / Set source dataframe index to IDs
df_source_indexed = df.set_index(id_cols)

# 3. 逐列更新 / Update column by column
for col in update_cols:
    if col in df_source_indexed.columns:
        print(f"     - 正在映射 / Mapping column: {col}")
        # Pandas 的索引对齐功能会自动匹配相同的 ID
        # Index alignment automatically matches identical IDs
        df_target_indexed[col] = df_source_indexed[col]

# 4. 重置索引，恢复原始 DataFrame 结构 / Reset index to restore original structure
df_target_final = df_target_indexed.reset_index()

# =============================================================================
# 打印更新统计报告 / Print Update Statistics
# =============================================================================
print("\n📊 更新统计报告 / Update Statistical Report:")

total_target_rows = len(df_target_indexed)
matched_ids = df_target_indexed.index.intersection(df_source_indexed.index)
updated_count = len(matched_ids)

print(f"   1. 原文件总行数 / Total rows in target: {total_target_rows}")
print(f"   2. 成功匹配并更新 / Successfully matched & updated: {updated_count} 行/rows")
print(f"   3. 匹配覆盖率 / Mapping Coverage: {updated_count / total_target_rows:.1%}")

# 5. 保存 / Save
print("-" * 50)
print(f"正在保存最终结果至 / Saving final results to: {target_path}")
# 保存时不需要特殊处理，内存中已经是字符串 / No special handling needed for saving string IDs
df_target_final.to_excel(target_path, index=False)

# 6. 验证 / Verification
print("-" * 50)
print("处理完成 / Process Completed.")
print(f"最终总样本数 / Final sample size: {len(df_target_final)}")

for col in update_cols:
    print(f"   > '{col}' 缺失值数量 / Missing values: {df_target_final[col].isna().sum()}")
print("=" * 50)