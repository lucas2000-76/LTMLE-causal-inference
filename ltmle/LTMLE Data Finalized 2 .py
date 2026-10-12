# -*- coding: utf-8 -*-
"""
Script Summary (Unified EDA, Missingness Analysis & Original Data Preservation for Triple Cohorts):
This script performs exploratory data analysis, calculates true/crude missing rates with aligned A-nodes,
and generates UpSet plots for missing patterns.
Crucially, it now sequentially processes three independent datasets (Base, CVD, and Diabetes)
in a single automated pipeline while PRESERVING 100% ORIGINAL DATA SCALES. /
本脚本执行探索性数据分析，计算对齐 A 节点的真实/粗略缺失率，并生成缺失模式的 UpSet 图。
最关键的是，它现在可以在一个自动化流水线中，顺序处理三个独立的数据集（基础版、CVD队列、Diabetes队列），同时 100% 保留原始数据尺度。

Recent Updates:
- Multi-Cohort Processing: Iteratively applies identical analysis logic to three files. / 多队列处理：循环对三个文件应用相同的分析逻辑。
- Dynamic Output: Automatically tags Excel reports and UpSet plots with cohort labels (e.g., '_CVD', '_Diabetes'). / 动态输出：自动为 Excel 报表和 UpSet 图打上队列标签（例如：'_CVD', '_Diabetes'）。
- Variable Deduplication: Added robust dictionary-based deduplication to prevent groupby errors. / 变量去重：添加了强大的基于字典的去重逻辑，防止分组报错。
- RAW DATA PRESERVATION: All clipping (e.g., sleep time), log transformations (e.g., CRP), and Winsorization (e.g., BMI, biomarkers) have been completely REMOVED. The script now outputs the exact original raw scales without any standardization or cleaning. / 原始数据保留：所有截断（如睡眠时间）、对数转换（如 CRP）和缩尾处理（如 BMI、生物标志物）均已彻底移除。脚本现在输出绝对原始的尺度，不进行任何标准化或清洗。
- Automatically saves the data in-place back to their respective files with intact IDs. / 自动将数据原地保存回各自的文件中，并保持 ID 完整。
"""

import pandas as pd
import numpy as np
import os
import re
import warnings
import matplotlib.pyplot as plt
from upsetplot import plot as upset_plot

# ==========================================
# 0. 全局设置 / Global Settings
# ==========================================
warnings.filterwarnings("ignore")
warnings.filterwarnings("ignore", category=FutureWarning, module="upsetplot")
pd.set_option('future.no_silent_downcasting', True)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)

# ==========================================
# 1. 设定路径与多文件清单 / Setup Paths & Multi-file List
# ==========================================
base_dir = r"C:\Users\26627\Desktop"
result_dir = os.path.join(base_dir, "result")

if not os.path.exists(result_dir):
    os.makedirs(result_dir)

# 定义三个目标文件的路径 / Define paths for the three target files
final_output_path = os.path.join(base_dir, "DATA_Final_Cleaned.xlsx")
cvd_output_path   = os.path.join(base_dir, "DATA_CVD_Final_Cleaned.xlsx")
diab_output_path  = os.path.join(base_dir, "DATA_diabetes_Final_Cleaned.xlsx")

# 创建处理任务清单 / Create processing task list
file_tasks = [
    (final_output_path, "Base"),
    (cvd_output_path,   "CVD"),
    (diab_output_path,  "Diabetes")
]

# ==========================================
# 2. 精确定义分类变量与连续变量 / Define Categorical and Numeric Variables
# ==========================================
# 拆分单维度的 A_Group 为 OPA 和 LTPA 的 Group / Split single A_Group into OPA and LTPA Groups
categorical_cols = [
    'V_Gender', 'V_Edu', 'V_community_level_ses',
    'L1_disease_scores_CVD_event', 'L1_Household_SES', 'L1_Martial_Status', 'L1_Smoke', 'L1_Drink', 'L1_Work_status', 'L1_Living_with_Child',
    'A1_OPA_Group', 'A1_LTPA_Group', 'Y1_Event_CVD','L1_disease_scores_Diabetes_event','Y1_Event_Diabetes',
    'L2_disease_scores_CVD_event', 'L2_Household_SES', 'L2_Martial_Status', 'L2_Smoke', 'L2_Drink', 'L2_Work_status', 'L2_Living_with_Child',
    'A2_OPA_Group', 'A2_LTPA_Group', 'C2', 'Y2_Event_CVD','L2_disease_scores_Diabetes_event','Y2_Event_Diabetes',
    'L3_disease_scores_CVD_event', 'L3_Household_SES', 'L3_Martial_Status', 'L3_Smoke', 'L3_Drink', 'L3_Work_status', 'L3_Living_with_Child',
    'A3_OPA_Group', 'A3_LTPA_Group', 'C3', 'Y3_Event_CVD','L3_disease_scores_Diabetes_event','Y3_Event_Diabetes',
    'L4_disease_scores_CVD_event', 'L4_Household_SES', 'L4_Martial_Status', 'L4_Smoke', 'L4_Drink', 'L4_Work_status', 'L4_Living_with_Child',
    'A4_OPA_Group', 'A4_LTPA_Group', 'Y4_Event_CVD','L4_disease_scores_Diabetes_event','Y4_Event_Diabetes'
]


numeric_cols = [
    'V_Bio_HbA1c_Pre', 'V_Bio_TyG', 'V_Bio_CRP_Pre', 'V_Bio_LDL_Pre','V_Bio_HDL_Pre',
    'L1_Age', 'L1_Sleep_time', 'L1_BMI', 'L1_Cognitive_functioning', 'L1_Depression_scale',
    'L2_Age', 'L2_Sleep_time', 'L2_BMI', 'L2_Bio_HbA1c_Pre', 'L2_Bio_TyG', 'L2_Bio_CRP_Pre', 'L2_Bio_LDL_Pre', 'L2_Bio_HDL_Pre',
    'L2_Cognitive_functioning', 'L2_Depression_scale',
    'L3_Age', 'L3_Sleep_time', 'L3_Cognitive_functioning', 'L3_Depression_scale',
    'L4_Age', 'L4_Sleep_time', 'L4_Cognitive_functioning', 'L4_Depression_scale',
]

# 强制去重，防止重复导致 groupby 报错 / Force deduplication to prevent groupby errors from duplicated variables
all_target_cols = list(dict.fromkeys(categorical_cols + numeric_cols))


# ==========================================
# 3. 启动多文件处理大循环 / Start Multi-file Processing Loop
# ==========================================
for file_path, label in file_tasks:
    # 优化点：删除了循环内重复的 import 语句 / Optimization: Removed redundant imports inside the loop

    print("\n" + "=" * 80)
    print(f"🚀 正在处理队列: [{label}] / Processing Cohort: [{label}]")
    print(f"📁 文件路径 / File Path: {file_path}")
    print("=" * 80)

    # --- 3.1 读取数据 / Read Data ---
    print(f"\n步骤 1: 正在读取数据 ... / Step 1: Reading data...")
    try:
        # 在读取阶段直接指定 ID 列为 str 类型，从源头防止前导 0 丢失
        id_columns = ['Individual ID', 'Household ID', 'Community ID']
        dtype_dict = {col: str for col in id_columns}

        # 传入 dtype 参数，即使某些列不存在，Pandas 也会自动忽略并不报错
        df = pd.read_excel(file_path, dtype=dtype_dict)
        total_samples = len(df)
        print(f"   > 读取成功！数据集总行数 (样本数) / Read successful! Total sample size: {total_samples}")
    except Exception as e:
        print(f"   [致命错误/Fatal Error] 读取失败 / Read failed: {e}")
        continue  # 若读取失败则跳过当前文件进入下一个 / Skip to next file if read fails

    # --- 3.2 检查数据框中缺失的预设变量 / Validation ---
    print("\n步骤 2: 🚨 检查数据框中缺失的预设变量 / Step 2: Validating expected variables in dataframe 🚨")
    missing_cat = [c for c in categorical_cols if c not in df.columns]
    missing_num = [c for c in numeric_cols if c not in df.columns]

    if missing_cat:
        print(f"❌ 缺失的分类变量 / Missing categorical variables ({len(missing_cat)}): {missing_cat}")
    else:
        print("✅ 分类变量: 无缺失，全部存在于数据集中。 / Categorical variables: All expected variables exist.")

    if missing_num:
        print(f"❌ 缺失的连续变量 / Missing numeric variables ({len(missing_num)}): {missing_num}")
    else:
        print("✅ 连续变量: 无缺失，全部存在于数据集中。 / Numeric variables: All expected variables exist.")

    # --- 3.3 计算缺失率 / Missing Rates Calculation ---
    print("\n步骤 3: 正在计算条目缺失率矩阵 (True vs Crude) / Step 3: Calculating Item Missing Rates (True vs Crude)")

    active_masks = {
        'Wave 1': pd.Series(True, index=df.index),
        'Wave 2': df['C1'] == 1 if 'C1' in df.columns else pd.Series(True, index=df.index),
        'Wave 3': df['C2'] == 1 if 'C2' in df.columns else pd.Series(True, index=df.index),
        'Wave 4': df['C3'] == 1 if 'C3' in df.columns else pd.Series(True, index=df.index)
    }

    def get_clean_base(col):
        """提取基础变量名，将双维组统一归类 / Extract base variable name, grouping dual domains unified"""
        if col in ['A1_OPA_Group', 'A2_OPA_Group', 'A3_OPA_Group', 'A4_OPA_Group']:
            return 'Treatment_A_OPA_Group'
        if col in ['A1_LTPA_Group', 'A2_LTPA_Group', 'A3_LTPA_Group', 'A4_LTPA_Group']:
            return 'Treatment_A_LTPA_Group'
        if col.startswith('V_'): return col
        if re.match(r'^C\d$', col): return 'Censoring_Node'
        m = re.match(r'^[A-Z]\d_(.*)', col)
        if m: return m.group(1)
        return col

    base_vars = sorted(
        list(set([get_clean_base(c) for c in all_target_cols if c in df.columns and not re.match(r'^C\d$', c)])))

    true_missing_records, crude_missing_records = [], []

    for base in base_vars:
        true_row, crude_row = {'变量名': base}, {'变量名': base}

        for wave_idx, wave_name in enumerate(['Wave 1', 'Wave 2', 'Wave 3', 'Wave 4'], start=1):
            target_col = None

            if base.startswith('V_'):
                target_col = base if wave_idx == 1 else None
            elif base == 'Treatment_A_OPA_Group':
                expected_a = f'A{wave_idx}_OPA_Group'
                if expected_a in df.columns: target_col = expected_a
            elif base == 'Treatment_A_LTPA_Group':
                expected_a = f'A{wave_idx}_LTPA_Group'
                if expected_a in df.columns: target_col = expected_a

                if expected_a in df.columns: target_col = expected_a
            else:
                candidates = [f"{prefix}{wave_idx}_{base}" for prefix in ['L', 'Y']]
                target_col = next((c for c in candidates if c in df.columns), None)

            if target_col:
                mask = active_masks[wave_name]
                active_count = mask.sum()

                if active_count > 0:
                    missing_count = df.loc[mask, target_col].isna().sum()
                    rate = (missing_count / active_count) * 100
                    true_row[wave_name] = f"{rate:.2f}% ({missing_count}/{active_count})"
                else:
                    true_row[wave_name] = "N/A"

                missing_crude = df[target_col].isna().sum()
                rate_crude = (missing_crude / total_samples) * 100
                crude_row[wave_name] = f"{rate_crude:.2f}% ({missing_crude}/{total_samples})"
            else:
                true_row[wave_name] = "-"
                crude_row[wave_name] = "-"

        true_missing_records.append(true_row)
        crude_missing_records.append(crude_row)

    df_true_missing = pd.DataFrame(true_missing_records)
    true_missing_path = os.path.join(result_dir, f"True_Item_Missing_Rates_{label}.xlsx")
    df_true_missing.to_excel(true_missing_path, index=False)
    print(f"   > [表1/Table 1] 排除删失的【真实缺失率矩阵】已保存至 / True Missing Matrix saved to: {true_missing_path}")

    df_crude_missing = pd.DataFrame(crude_missing_records)
    crude_missing_path = os.path.join(result_dir, f"Crude_Total_Missing_Rates_{label}.xlsx")
    df_crude_missing.to_excel(crude_missing_path, index=False)
    print(f"   > [表2/Table 2] 包含删失的【粗略缺失率矩阵】已保存至 / Crude Missing Matrix saved to: {crude_missing_path}")

    # --- 3.4 生成 UpSet 缺失值可视化分布图 / UpSet Plots for Missingness ---
    print("\n步骤 4: 正在绘制 Missing UpSet Plots (Top 8 缺失模式) / Step 4: Generating UpSet Plots (Top 8)")

    def plot_cleaned_upset(missing_bool_df, save_name, title_prefix=""):
        """绘制剔除全满变量后的 UpSet 图 / Plot UpSet dropping fully observed columns"""
        missing_bool_df = missing_bool_df.loc[:, missing_bool_df.sum() > 0]
        if missing_bool_df.empty:
            print(f"   > [{title_prefix}] 暂无缺失数据，无法生成图表。 / No missing data, cannot generate plot.")
            return

        upset_data = missing_bool_df.groupby(missing_bool_df.columns.tolist()).size()
        top8_data = upset_data.sort_values(ascending=False).head(8)

        idx_df = top8_data.index.to_frame(index=False)
        active_cols = idx_df.columns[idx_df.any(axis=0)]

        if len(active_cols) == 0:
            print(f"   > [{title_prefix}] 顶级模式均为全量完整，无法绘制 UpSet 图。 / Top patterns are fully complete, cannot plot.")
            return

        new_index = pd.MultiIndex.from_frame(idx_df[active_cols])
        top8_data.index = new_index

        fig = plt.figure(figsize=(12, 6))
        upset_plot(top8_data, fig=fig, show_counts=True, element_size=45)
        plt.title(f"{label} - {title_prefix} Top 8 Missing Patterns", fontsize=14, pad=20)

        save_path = os.path.join(result_dir, f"{save_name}_{label}.png")
        plt.savefig(save_path, format='png', bbox_inches='tight', dpi=300)
        plt.close()
        print(f"   > [图/Plot] {title_prefix} 已剔除无用变量并保存至 / Saved clean plot to: {save_path}")

    try:
        exist_cols = [c for c in all_target_cols if c in df.columns]
        plot_cleaned_upset(df[exist_cols].isna(), save_name="Missing_UpSet_Plot_Crude_Top8", title_prefix="Crude")

        true_missing_bool = pd.DataFrame(False, index=df.index, columns=exist_cols)
        for col in exist_cols:
            wave_name = 'Wave 2' if '2_' in col or col == 'C2' else \
                'Wave 3' if '3_' in col or col == 'C3' else \
                    'Wave 4' if '4_' in col or col == 'C4' else 'Wave 1'
            true_missing_bool[col] = df[col].isna() & active_masks[wave_name]

        plot_cleaned_upset(true_missing_bool, save_name="Missing_UpSet_Plot_True_Top8", title_prefix="True")

    except Exception as e:
        print(f"   [警告/Warning] UpSet 图生成失败 / UpSet plot generation failed: {e}")

    # --- 3.5 分类变量分布情况 / Categorical Distribution ---
    print("\n步骤 5: 分类变量分布情况 / Step 5: Categorical Variables Distribution")

    cat_groups = {}
    for col in [c for c in categorical_cols if c in df.columns]:
        cat_groups.setdefault(get_clean_base(col), []).append(col)

    for base, cols in cat_groups.items():
        print(f"\n>>>>>>>> 同类变量组 / Categorical Group: {base} <<<<<<<<")
        for col in sorted(cols):
            summary_df = pd.DataFrame({
                '频数 (N) / Count': df[col].value_counts(dropna=False),
                '占比 (%) / Percent': (df[col].value_counts(dropna=False, normalize=True) * 100).round(2)
            })
            print(f"\n[ 变量/Variable: {col} ]\n{summary_df}")

    # --- 3.6 连续变量：分布情况与参考 / Numeric Distribution & References ---
    print("\n步骤 6: 连续变量 数据分布情况 / Step 6: Numeric Variables Distribution")

    reference_knowledge = {
        'Age': '一般中老年队列 (如CHARLS) 年龄多在 45 岁以上。 / Generally >45 years old in aging cohorts.',
        'Sleep_time': '成人正常健康睡眠时间通常推荐为 7 - 9 小时/天。 / Recommended sleep is 7-9 hours/day.',
        'BMI': '中国成人BMI标准: 正常 18.5-23.9, 超重 24.0-27.9, 肥胖 ≥28.0。 / Chinese BMI thresholds.',
        'Bio_CRP_Pre': 'C-反应蛋白 (CRP): 一般参考值 < 3 mg/L, >3 mg/L 提示系统性炎症。 / CRP >3 mg/L implies systemic inflammation.',
        'Bio_LDL_Pre': '低密度脂蛋白 (LDL-C): 理想值通常 < 3.4 mmol/L (130 mg/dL)。 / Ideal LDL-C is < 3.4 mmol/L.',
        'Bio_HbA1c_Pre': '糖化血红蛋白 (HbA1c): 正常值 < 5.7%，≥ 6.5% 诊断糖尿病。 / HbA1c ≥ 6.5% indicates diabetes.',
        'Bio_TyG': 'TyG 指数: 胰岛素抵抗的替代指标，通常分布在 7.5 - 10 之间。 / TyG index naturally distributes around 7.5 - 10.',
        'Cognitive_functioning': '认知功能总分: 分数越高代表认知能力越好。 / Higher scores indicate better cognitive function.',
        'Depression_scale': '抑郁量表: 得分越高提示显著抑郁症状。 / Higher scores indicate significant depressive symptoms.',
        'Treatment_A_OPA_Mets': '职业性体力活动连续 MET 值。 / Continuous METs for Occupational PA.',
        'Treatment_A_LTPA_Mets': '休闲性体力活动连续 MET 值。 / Continuous METs for Leisure-Time PA.'
    }

    num_groups = {}
    for col in [c for c in numeric_cols if c in df.columns]:
        num_groups.setdefault(get_clean_base(col), []).append(col)

    for base, cols in num_groups.items():
        print(f"\n>>>>>>>> 同类连续变量组 / Numeric Group: {base} <<<<<<<<")
        for col in sorted(cols):
            desc = df[col].describe()
            print(f"\n[ 变量/Variable: {col} ]")
            print(f" -> 均值/Mean: {desc['mean']:.3f}, 标准差/Std: {desc['std']:.3f}")
            print(f" -> Min: {desc['min']:.3f} | Max: {desc['max']:.3f}")
            print(f" -> 四分位分布/Quartiles: Q1={desc['25%']:.3f}, Median={desc['50%']:.3f}, Q3={desc['75%']:.3f}")

            if base in reference_knowledge:
                print(f" 💡 [医学知识参考/Medical Ref]: {reference_knowledge[base]}")

    # --- 3.7 连续变量保持原始尺度 (不进行清洗) / Keep Original Scale for Continuous Variables (No Cleaning) ---
    print("\n步骤 7: 检查并确认保持原始变量尺度 / Step 7: Verifying preservation of original variable scales")

    for col in [c for c in numeric_cols if 'Sleep_time' in c and c in df.columns]:
        print(f"   > [{col}] 保持原始尺度，未执行截断 / Kept original scale, did not execute clipping.")

    for col in [c for c in numeric_cols if 'CRP' in c and c in df.columns]:
        print(f"   > [{col}] 保持原始尺度，未执行对数转换 / Kept original scale, did not execute Log transformation.")

    win_targets = ['BMI', 'LDL', 'HbA1c', 'TyG', 'HDL']
    for col in numeric_cols:
        if col in df.columns and any(target in col for target in win_targets):
            print(f"   > [{col}] 保持原始尺度，未执行缩尾 / Kept original scale, did not execute Winsorization.")

    print("   > [其他连续变量/Other Variables] 年龄与心理认知量表均保持原始真实分布 / Age and cognitive scales kept original.")

    # --- 3.8 原地保存数据 / Save Data In-Place ---
    print(f"\n步骤 8: 正在原地保存数据 / Step 8: Saving data in-place ...")

    # 即便读取时已经是字符串，依然保留此清洗逻辑作为冗余防御，防止脏数据影响
    id_save_cols = ['Individual ID', 'Household ID', 'Community ID']
    for col in id_save_cols:
        if col in df.columns:
            # 强制将列转换为字符串，并清洗掉可能在原始 Excel 单元格中残留的纯文本 '.0'
            df[col] = df[col].astype(str).str.replace(r'\.0$', '', regex=True)

    try:
        # 保存时 Pandas 默认会将纯数字的字符串写入为文本格式，从而在 Excel 中保留先导零
        df.to_excel(file_path, index=False)
        print(f"   > [成功/Success] 数据已原地覆盖保存至 / Data successfully saved to: {file_path}")
    except Exception as e:
        print(f"   [错误/Error] 保存失败 / Save failed: {e}")

print("\n" + "=" * 80)
print("🎉 全部任务完毕！所有三个独立队列的分析均已完成，且全部保持绝对原始数据尺度。 / All EDA tasks completed for all three cohorts with absolute original data scales preserved!")