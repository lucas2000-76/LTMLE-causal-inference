"""
Script Summary (Unified Pipeline: LTMLE Preparation, Outcome Cleaning, Disease Cohort Split & Strict Censoring):
This script processes the longitudinal 'DATA.xlsx' dataset through a comprehensive pipeline for causal inference.

Main Workflows:
--- PHASE 1: Read Data / 数据读取 ---
1. Read Data & ID Padding: Loads data safely and ensures 12-digit standard IDs. / 读取数据并保护 12 位 ID 的先导零。

--- PHASE 2: Censoring Generation / 删失节点生成 ---
2. Strict C Nodes Generation: Dynamically generates C1, C2, C3 based on the genuine presence of subsequent wave data. / 基于后续波次原始数据的真实存在情况动态生成 C1, C2, C3。
3. Column Reordering: Moves C nodes immediately after their corresponding A nodes (LTPA). / 将 C 节点移动到对应的 A 节点 (LTPA) 之后。

--- PHASE 3: Intermittent Fix & Age Correction / 间歇性缺失修正与年龄修正 ---
4. Intermittent Missingness Fix: Identifies individuals who missed a wave but returned later, forcing bridging C nodes to 1. / 识别缺席某一波次但随后回归的个体，强制其桥接 C 节点为 1。
5. Age Correction & Exclusion: Calculates L2_Age, L3_Age, L4_Age based on C nodes. Excludes samples with L1_Age < 44. / 基于 C 节点计算后续年龄，并剔除基线 L1_Age < 44 的样本。

--- PHASE 4: Outcome Cleaning / 结局变量深度清洗 ---
6. Outcome Backward Fill: Cross-wave backfills missing outcome states (Y) to rescue missing records before strict exclusion. / 在严格剔除前，安全地进行结局变量跨波次反向填补 (以后填前)，挽救记录。

--- PHASE 5: Outcome Intersection Drop & Disease-Specific Split / 结局交集剔除与特定疾病新发队列拆分 ---
7. Intersection Drop & Disease Baseline Exclusion: For each of the three output paths (Base, CVD-incident, Diabetes-incident),
   independently removes samples missing all Y nodes for that disease (or both diseases for Base), then additionally drops
   baseline prevalent cases for disease-specific cohorts. / 对三个输出流分别独立剔除该疾病结局在所有波次全部缺失的样本（Base 文件剔除两种疾病交集全缺失），再分别剔除基线患病人群，生成初步的新发队列数据。

--- PHASE 6: Strict Exclusion on Split Cohorts / 在特定疾病队列上独立执行严格缺失剔除 ---
8. Missing Sample Exclusion (Dual Intervention): Applied INDEPENDENTLY to the three files (Base, CVD, Diabetes) generated in Phase 5. Evaluates missing patterns for OPA and LTPA. Enforces >=60% L-covariate completeness for intermittent cases and checks subsequent A/Y nodes. Retains ONLY IF BOTH patterns are compliant. Final clean data is re-saved to the respective files. / 在基础、CVD 和糖尿病三个队列上分别独立执行缺失剔除。独立评估 OPA 和 LTPA，间歇性缺失需满足 >=60% 同期协变量并检查后续 A/Y 节点，双侧均合规方可保留。最终将极致纯净的数据覆盖保存回三个文件。
"""

import pandas as pd
import numpy as np
import os
import sys
import warnings

# ==========================================
# 0. 全局设置与辅助函数 / Global Settings & Helpers
# ==========================================
warnings.filterwarnings("ignore")
pd.set_option('future.no_silent_downcasting', True)
pd.set_option('display.max_columns', None)
pd.set_option('display.width', 1000)

base_dir = r"C:\Users\26627\Desktop"
target_path = os.path.join(base_dir, "DATA.xlsx")

# 三个输出文件的路径 / Paths for the three output files
final_output_path = os.path.join(base_dir, "DATA_Final_Cleaned.xlsx")
cvd_output_path   = os.path.join(base_dir, "DATA_CVD_Final_Cleaned.xlsx")
diab_output_path  = os.path.join(base_dir, "DATA_diabetes_Final_Cleaned.xlsx")

ID_COL = 'Individual ID'
ID_LEN = 12
DISEASES = ["CVD", "Diabetes"]
id_cols = [ID_COL, 'Household ID', 'Community ID']

def section(title: str) -> None:
    """打印带分隔线的节标题 / Print a section header with separator."""
    print(f"\n{'=' * 50}\n{title}")

def pad_id(series: pd.Series) -> pd.Series:
    """将 ID 列统一转为干净的 12 位字符串，安全处理 NaN / Pad ID column to 12-digit strings, safely handling NaN."""
    return (
        series.astype(str)
              .str.replace(r"\.0$", "", regex=True)
              .apply(lambda x: x.zfill(ID_LEN) if x != "nan" else np.nan)
    )

# 定义各波次的变量清单 / Define variable lists for each wave
wave1_expected = [
    'L1_disease_scores_CVD_event', 'L1_Martial_Status', 'L1_Household_SES',
    'L1_Age', 'L1_Sleep_time', 'L1_BMI', 'L1_Smoke', 'L1_Drink',
    'L1_Work_status', 'L1_Living_with_Child', 'L1_Cognitive_functioning',
    'L1_Depression_scale', 'A1_OPA_Group', 'A1_LTPA_Group',  'Y1_Event_CVD',
    'L1_disease_scores_Diabetes_event', 'Y1_Event_Diabetes'
]
wave2_expected = [
    'L2_disease_scores_CVD_event', 'L2_Household_SES', 'L2_Martial_Status',
    'L2_Age', 'L2_Sleep_time', 'L2_BMI', 'L2_Bio_CRP_Pre', 'L2_Bio_LDL_Pre',
    'L2_Bio_HDL_Pre', 'L2_Bio_HbA1c_Pre', 'L2_Bio_TyG', 'L2_Smoke', 'L2_Drink',
    'L2_Work_status', 'L2_Living_with_Child', 'L2_Cognitive_functioning',
    'L2_Depression_scale', 'A2_OPA_Group', 'A2_LTPA_Group', 'Y2_Event_CVD',
    'L2_disease_scores_Diabetes_event', 'Y2_Event_Diabetes'
]
wave3_expected = [
    'L3_disease_scores_CVD_event', 'L3_Household_SES', 'L3_Martial_Status',
    'L3_Age', 'L3_Sleep_time', 'L3_Smoke', 'L3_Drink', 'L3_Work_status',
    'L3_Living_with_Child', 'L3_Cognitive_functioning', 'L3_Depression_scale',
    'A3_OPA_Group', 'A3_LTPA_Group', 'Y3_Event_CVD', 'L3_disease_scores_Diabetes_event', 'Y3_Event_Diabetes'
]
wave4_expected = [
    'L4_disease_scores_CVD_event', 'L4_Household_SES', 'L4_Martial_Status',
    'L4_Age', 'L4_Sleep_time', 'L4_Smoke', 'L4_Drink', 'L4_Work_status',
    'L4_Living_with_Child', 'L4_Cognitive_functioning', 'L4_Depression_scale',
    'A4_OPA_Group', 'A4_LTPA_Group', 'Y4_Event_CVD', 'L4_disease_scores_Diabetes_event', 'Y4_Event_Diabetes'
]


# ==========================================
# PHASE 1: 读取数据 / Phase 1: Read Data
# ==========================================
section(f"步骤 1: 读取数据 / Step 1: Read Data\n   > {target_path}")

try:
    df = pd.read_excel(
        target_path,
        dtype={ID_COL: str, "Household ID": str, "Community ID": str},
    )
except Exception as e:
    sys.exit(f"   [致命错误/Fatal Error] 读取失败 / Read failed: {e}")

# 修复先导零 (Excel 可能丢弃) / Restore leading zeros potentially stripped by Excel
for col in id_cols:
    if col in df.columns:
        df[col] = pad_id(df[col])

print(f"   > 原始样本量 / Original sample size: {len(df)}")
print(f"   > [检查/Check] ID 列类型 / ID column type: {df[ID_COL].dtype} (应为 object/string / Should be object/string)")


# ==========================================
# PHASE 2: 生成 C 节点、重排 / Phase 2: Censoring Nodes Generation
# ==========================================
section("第一阶段: 删失节点生成 (PHASE 2) / Phase 2: Censoring Nodes Generation")

print("步骤 2: 正在生成随访删失节点 (C) / Step 2: Generating Censoring Nodes (C) ...")
wave2_cols = [c for c in wave2_expected if c in df.columns]
wave3_cols = [c for c in wave3_expected if c in df.columns]
wave4_cols = [c for c in wave4_expected if c in df.columns]

# 初始生成 C 节点：检查下一波是否有任意数据 / Initial C node generation: Check if any data exists in the next wave
if wave2_cols: df['C1'] = df[wave2_cols].notna().any(axis=1).astype(int)
if wave3_cols: df['C2'] = df[wave3_cols].notna().any(axis=1).astype(int)
if wave4_cols: df['C3'] = df[wave4_cols].notna().any(axis=1).astype(int)

# 确保无 C4 列 / Ensure no C4 column exists
if 'C4' in df.columns:
    df = df.drop(columns=['C4'])
print("   > 初始删失节点 C1, C2, C3 生成完毕。 / Initial censoring nodes C1, C2, C3 generated.")

print("\n步骤 3: 正在重排变量顺序 (将 C 放在 A 之后) / Step 3: Reordering C after A ...")
def move_c_after_a(columns, c_node, a_node):
    """ 将 C 节点移动到对应的 A 节点之后 / Move C node right after the corresponding A node """
    if c_node in columns and a_node in columns:
        columns.remove(c_node)
        columns.insert(columns.index(a_node) + 1, c_node)
    return columns

cols_list = df.columns.tolist()
# 放到每波最后一个 A (也就是 LTPA) 之后 / Move after the last A node of each wave (LTPA)
cols_list = move_c_after_a(cols_list, 'C1', 'A1_LTPA_Group')
cols_list = move_c_after_a(cols_list, 'C2', 'A2_LTPA_Group')
cols_list = move_c_after_a(cols_list, 'C3', 'A3_LTPA_Group')
df = df[cols_list]
print("   > 列重排完成。 / Column reordering completed.")


# ==========================================
# PHASE 3: 间歇性缺失修正与年龄计算 / Phase 3: Intermittent Fix & Age Calculation
# ==========================================
section("第二阶段: 间歇性缺失修正与特征对齐 (PHASE 3) / Phase 3: Intermittent Fix & Alignment")

print("步骤 4: 🚀 统计间歇性缺失并提前强制设为未删失(1) / Step 4: Forcing intermittent missing C nodes to 1 ...")
# 查找离开后又回来的间歇性缺失者 / Identify intermittently missing subjects who missed a wave but returned later
intermittent_w2_missing = ((df['C1'] == 0) & (df[wave3_cols + wave4_cols].notna().any(axis=1)))
intermittent_w3_missing = ((df['C2'] == 0) & (df[wave4_cols].notna().any(axis=1)))

is_intermittent = intermittent_w2_missing | intermittent_w3_missing
total_intermittent = is_intermittent.sum()

print(f"   > [数据探查/Data Explorer] 发现间歇性缺失 (离开后又回归) 的情况：总计有 {total_intermittent} 人。 / Found {total_intermittent} intermittent cases.")
if total_intermittent > 0:
    print(f"      - 第2波没来，但后面回来了 / Missed W2 but returned: {intermittent_w2_missing.sum()} 人/people")
    print(f"      - 第3波没来，但后面回来了 / Missed W3 but returned: {intermittent_w3_missing.sum()} 人/people")
    # 强制修正 C 节点为 1，确保后续年龄计算有依有据 / Force C nodes to 1 to ensure valid subsequent age calculation
    df.loc[intermittent_w2_missing, 'C1'] = 1
    df.loc[intermittent_w3_missing, 'C2'] = 1
    print(f"   ℹ️ 已将这 {total_intermittent} 人的桥接删失状态提前修正为 1 (未删失)。 / Forced C=1 for these bridging cases.")
else:
    print("   ✅ 完美！没有间歇性缺失现象。 / Perfect! No intermittent missingness.")

print("\n步骤 5: 正在根据最终 C 节点修正年龄 / Step 5: Correcting Age based on finalized C Nodes ...")
if 'L1_Age' in df.columns:
    df['L1_Age'] = pd.to_numeric(df['L1_Age'], errors='coerce')
    # 只有 C=1（明确存活在队，包括通过间歇挽救回来的）才给算年龄 / Only calculate age if C=1 (explicitly alive in cohort, including rescued intermittent cases)
    if 'C1' in df.columns: df['L2_Age'] = np.where(df['C1'] == 1, df['L1_Age'] + 2, np.nan)
    if 'C2' in df.columns: df['L3_Age'] = np.where(df['C2'] == 1, df['L1_Age'] + 4, np.nan)
    if 'C3' in df.columns: df['L4_Age'] = np.where(df['C3'] == 1, df['L1_Age'] + 7, np.nan)
    print("   > 年龄修正完毕。 / Age correction completed.")

# --- 剔除基线年龄 < 44 / Drop baseline age < 44 ---
if "L1_Age" in df.columns:
    mask_young = df["L1_Age"] < 44
    print(f"   > [清洗/Cleaning] 剔除 L1_Age < 44 样本 {mask_young.sum()} 条 / Dropping {mask_young.sum()} samples with Age < 44.")
    df = df[~mask_young].copy()


# ==========================================
# PHASE 4: 结局清洗 (反向填补) / Phase 4: Outcome Cleaning (Backfill)
# ==========================================
section("第四阶段: 结局变量深度清洗 (PHASE 4) / Phase 4: Outcome Variable Deep Cleaning")

print("步骤 6: 结局变量反向填充 (以后填前) / Step 6: Backward-fill outcome variables")
# 后面的波次可以一次填充前面的所有缺失，而前面的不能填充后面的。
# Later waves simultaneously backfill all missing values in earlier waves; no forward filling.
for disease in DISEASES:
    y_cols = [f"Y{w}_Event_{disease}" for w in range(1, 5) if f"Y{w}_Event_{disease}" in df.columns]
    if not y_cols:
        continue

    df[y_cols] = df[y_cols].apply(pd.to_numeric, errors="coerce")
    orig = df[y_cols].copy()

    # bfill 沿 axis=1 从 Y4 → Y1 方向传播，天然支持跨波次一次性填充。
    # bfill along axis=1 propagates from Y4 → Y1, naturally supporting cross-wave propagation.
    df[y_cols] = df[y_cols].bfill(axis=1)

    filled_mask = orig.isna() & df[y_cols].notna()
    n_filled = filled_mask.any(axis=1).sum()
    print(f"   > [{disease}] 反向填充完成，{n_filled} 个样本得到跨波次填补。"
          f" / Backward fill done, {n_filled} samples backfilled.")

    if n_filled > 0:
        print(f"     [被反向填充样本预览 Top 10 / Backfilled sample preview Top 10]:")
        preview = pd.concat(
            [orig.loc[filled_mask.any(axis=1)].head(10).add_suffix("_Orig"),
             df[y_cols].loc[filled_mask.any(axis=1)].head(10).add_suffix("_New")],
            axis=1,
        ).sort_index(axis=1)
        print(preview.to_string())

# 将 Individual ID 设为索引，使其完全脱离业务计算逻辑。
# Set Individual ID as index to fully isolate it from business computation logic.
df.set_index(ID_COL, inplace=True)
# 剔除无用的 Household ID 和 Community ID / Drop useless Household ID and Community ID
df.drop(columns=[c for c in ["Household ID", "Community ID"] if c in df.columns], inplace=True)


# ==========================================
# PHASE 5: 结局交集剔除与最终拆分导出 / Phase 5: Outcome Intersection Drop & Disease-Specific Export
# ==========================================
section("第五阶段: 结局交集剔除与特定疾病基线剔除及初步保存 (PHASE 5) / Phase 5: Outcome Intersection Drop & Initial Save")

# 将 ID 从索引恢复回普通列，以备输出 / Restore ID from index back to a column for output
df = df.reset_index()

# 预先计算各疾病的结局列 / Pre-compute outcome column lists for each disease
y_cvd_cols  = [c for c in df.columns if c.startswith("Y") and "CVD"      in c]
y_diab_cols = [c for c in df.columns if c.startswith("Y") and "Diabetes" in c]

try:
    # ---------------------------------------------------------
    # 数据流 1：基础纯净版 / Data Stream 1: Cleaned Base
    # 步骤：剔除 CVD 与 Diabetes 结局均全缺失的交集样本。
    # Step: Drop samples where BOTH CVD and Diabetes outcomes are entirely missing.
    # ---------------------------------------------------------
    print("\n--- 队列 1: 基础合并队列 / Cohort 1: Base combined cohort ---")
    df_base = df.copy()

    # 剔除 CVD 全缺失样本 / Drop samples with all CVD outcomes missing
    mask_cvd_all_missing = df_base[y_cvd_cols].isna().all(axis=1) if y_cvd_cols else pd.Series(False, index=df_base.index)
    n_before = len(df_base)
    n_drop   = mask_cvd_all_missing.sum()
    df_base  = df_base[~mask_cvd_all_missing].copy()
    print(f"   > [清洗/Cleaning] CVD 结局全缺失剔除 / CVD all-outcome-missing drop: "
          f"总数/Total={n_before}, 删除/Dropped={n_drop}, 剩余/Remaining={len(df_base)}")

    # 剔除 Diabetes 全缺失样本 / Drop samples with all Diabetes outcomes missing
    mask_diab_all_missing = df_base[y_diab_cols].isna().all(axis=1) if y_diab_cols else pd.Series(False, index=df_base.index)
    n_before = len(df_base)
    n_drop   = mask_diab_all_missing.sum()
    df_base  = df_base[~mask_diab_all_missing].copy()
    print(f"   > [清洗/Cleaning] Diabetes 结局全缺失剔除 / Diabetes all-outcome-missing drop: "
          f"总数/Total={n_before}, 删除/Dropped={n_drop}, 剩余/Remaining={len(df_base)}")

    df_base.to_excel(final_output_path, index=False)
    print(f"   > [保存/Save] 基础合并文件已初步生成 / Base combined file initially generated: DATA_Final_Cleaned.xlsx (N={len(df_base)})")

    # ---------------------------------------------------------
    # 数据流 2：CVD 新发队列 / Data Stream 2: Incident CVD Cohort
    # 步骤：在基础版基础上，先剔除 CVD 结局全缺失样本，再剔除基线 CVD 患病样本。
    # Step: Starting from base df, first drop samples with all CVD outcomes missing,
    #       then drop baseline CVD prevalent cases.
    # ---------------------------------------------------------
    print("\n--- 队列 2: CVD 新发队列 / Cohort 2: CVD incident cohort ---")
    df_cvd = df.copy()

    # 剔除 CVD 结局全缺失样本 / Drop samples with all CVD outcomes missing
    mask_cvd_all_missing = df_cvd[y_cvd_cols].isna().all(axis=1) if y_cvd_cols else pd.Series(False, index=df_cvd.index)
    n_before = len(df_cvd)
    n_drop   = mask_cvd_all_missing.sum()
    df_cvd   = df_cvd[~mask_cvd_all_missing].copy()
    print(f"   > [清洗/Cleaning] CVD 结局全缺失剔除 / CVD all-outcome-missing drop: "
          f"总数/Total={n_before}, 删除/Dropped={n_drop}, 剩余/Remaining={len(df_cvd)}")

    # 剔除基线 CVD 患病样本 / Drop baseline CVD prevalent cases
    if 'Y1_Event_CVD' in df_cvd.columns:
        mask_baseline_cvd = df_cvd['Y1_Event_CVD'] == 1
        n_before = len(df_cvd)
        n_drop   = mask_baseline_cvd.sum()
        df_cvd   = df_cvd[~mask_baseline_cvd].copy()
        print(f"   > [清洗/Cleaning] 基线 CVD 患病剔除 / Baseline CVD prevalent drop: "
              f"总数/Total={n_before}, 删除/Dropped={n_drop}, 剩余/Remaining={len(df_cvd)}")

    df_cvd.to_excel(cvd_output_path, index=False)
    print(f"   > [保存/Save] CVD 特定文件已初步生成 / CVD-specific file initially generated: DATA_CVD_Final_Cleaned.xlsx (N={len(df_cvd)})")

    # ---------------------------------------------------------
    # 数据流 3：Diabetes 新发队列 / Data Stream 3: Incident Diabetes Cohort
    # 步骤：在基础版基础上，先剔除 Diabetes 结局全缺失样本，再剔除基线 Diabetes 患病样本。
    # Step: Starting from base df, first drop samples with all Diabetes outcomes missing,
    #       then drop baseline Diabetes prevalent cases.
    # ---------------------------------------------------------
    print("\n--- 队列 3: Diabetes 新发队列 / Cohort 3: Diabetes incident cohort ---")
    df_diab = df.copy()

    # 剔除 Diabetes 结局全缺失样本 / Drop samples with all Diabetes outcomes missing
    mask_diab_all_missing = df_diab[y_diab_cols].isna().all(axis=1) if y_diab_cols else pd.Series(False, index=df_diab.index)
    n_before = len(df_diab)
    n_drop   = mask_diab_all_missing.sum()
    df_diab  = df_diab[~mask_diab_all_missing].copy()
    print(f"   > [清洗/Cleaning] Diabetes 结局全缺失剔除 / Diabetes all-outcome-missing drop: "
          f"总数/Total={n_before}, 删除/Dropped={n_drop}, 剩余/Remaining={len(df_diab)}")

    # 剔除基线 Diabetes 患病样本 / Drop baseline Diabetes prevalent cases
    if 'Y1_Event_Diabetes' in df_diab.columns:
        mask_baseline_diab = df_diab['Y1_Event_Diabetes'] == 1
        n_before = len(df_diab)
        n_drop   = mask_baseline_diab.sum()
        df_diab  = df_diab[~mask_baseline_diab].copy()
        print(f"   > [清洗/Cleaning] 基线 Diabetes 患病剔除 / Baseline Diabetes prevalent drop: "
              f"总数/Total={n_before}, 删除/Dropped={n_drop}, 剩余/Remaining={len(df_diab)}")

    df_diab.to_excel(diab_output_path, index=False)
    print(f"   > [保存/Save] Diabetes 特定文件已初步生成 / Diabetes-specific file initially generated: DATA_diabetes_Final_Cleaned.xlsx (N={len(df_diab)})")

    print(f"\n   > [验证/Verified] 队列拆分完成，准备进入第六阶段进行独立严格剔除。 / Cohort split complete, ready for independent strict exclusion in Phase 6.")

except Exception as e:
    print(f"   [错误/Error] 初步保存失败 / Initial save failed: {e}")


# ==========================================
# PHASE 6: 严格缺失剔除 (在三个独立文件上分别执行) / Phase 6: Strict Exclusion on Independent Files
# ==========================================
section("第六阶段: 缺失样本物理剔除 (双干预模式) (PHASE 6) / Phase 6: Dual Intervention Missing Sample Exclusion")

print("步骤 7: 正在针对三个输出文件独立执行不符合插补要求的残缺样本剔除 / Step 7: Executing missing samples deletion independently on the three output cohorts ...")

def execute_strict_exclusion(df_target: pd.DataFrame, dataset_name: str, save_path: str) -> pd.DataFrame:
    """
    针对特定的数据集执行严格的协变量与随访节点缺失剔除，并覆盖保存。
    Executes strict covariate and follow-up node exclusion on a specific dataset and overwrites the file.
    """
    print(f"\n" + "=" * 40)
    print(f"🚀 开始审查与过滤: [{dataset_name}] / Starting review & filtering: [{dataset_name}]")
    print("=" * 40)

    # 确保 OPA 和 LTPA 节点存在 / Ensure OPA and LTPA nodes exist
    opa_nodes = ['A1_OPA_Group', 'A2_OPA_Group', 'A3_OPA_Group', 'A4_OPA_Group']
    ltpa_nodes = ['A1_LTPA_Group', 'A2_LTPA_Group', 'A3_LTPA_Group', 'A4_LTPA_Group']

    for a in opa_nodes + ltpa_nodes:
        if a not in df_target.columns:
            df_target[a] = np.nan

    # 构建每个样本的 A 节点缺失模式字符串 (如 '1100') / Build A node pattern string (e.g., '1100')
    opa_pattern = (df_target['A1_OPA_Group'].notna().astype(int).astype(str) +
                   df_target['A2_OPA_Group'].notna().astype(int).astype(str) +
                   df_target['A3_OPA_Group'].notna().astype(int).astype(str) +
                   df_target['A4_OPA_Group'].notna().astype(int).astype(str))

    ltpa_pattern = (df_target['A1_LTPA_Group'].notna().astype(int).astype(str) +
                    df_target['A2_LTPA_Group'].notna().astype(int).astype(str) +
                    df_target['A3_LTPA_Group'].notna().astype(int).astype(str) +
                    df_target['A4_LTPA_Group'].notna().astype(int).astype(str))

    # 提取 L 协变量 / Extract L covariates specific to this dataset
    w1_l_cols = [c for c in wave1_expected if c.startswith('L') and c in df_target.columns]
    w2_l_cols = [c for c in wave2_expected if c.startswith('L') and c in df_target.columns]
    w3_l_cols = [c for c in wave3_expected if c.startswith('L') and c in df_target.columns]

    # 计算并打印间歇缺失的严格条件参数 / Calculate and print strict condition parameters for intermittent cases
    print("   > [统计/Stats] 高质量间歇缺失者的变量完备性阈值要求 (>=60% 同期协变量) / Completeness thresholds for intermittent cases:")
    w1_req = int(np.ceil(len(w1_l_cols) * 0.6))
    w2_req = int(np.ceil(len(w2_l_cols) * 0.6))
    w3_req = int(np.ceil(len(w3_l_cols) * 0.6))

    print(f"      - Wave 1: 包含所有L/A/Y变量共 {len([c for c in wave1_expected if c in df_target.columns])} 个, 其中 L 协变量 {len(w1_l_cols)} 个, 至少需满足 {w1_req} 个 (>=60%). / Total vars: {len([c for c in wave1_expected if c in df_target.columns])}, L covars: {len(w1_l_cols)}, Min required: {w1_req}")
    print(f"      - Wave 2: 包含所有L/A/Y变量共 {len([c for c in wave2_expected if c in df_target.columns])} 个, 其中 L 协变量 {len(w2_l_cols)} 个, 至少需满足 {w2_req} 个 (>=60%). / Total vars: {len([c for c in wave2_expected if c in df_target.columns])}, L covars: {len(w2_l_cols)}, Min required: {w2_req}")
    print(f"      - Wave 3: 包含所有L/A/Y变量共 {len([c for c in wave3_expected if c in df_target.columns])} 个, 其中 L 协变量 {len(w3_l_cols)} 个, 至少需满足 {w3_req} 个 (>=60%). / Total vars: {len([c for c in wave3_expected if c in df_target.columns])}, L covars: {len(w3_l_cols)}, Min required: {w3_req}")

    def check_subsequent_ay(target_df, wave, domain):
        """ 检查后续波次是否存在指定的 A 和 Y 节点 / Check if subsequent A and Y nodes exist """
        a_col = f'A{wave}_{domain}_Group'
        y_cvd, y_diab = f'Y{wave}_Event_CVD', f'Y{wave}_Event_Diabetes'
        m = pd.Series(True, index=target_df.index)
        if a_col in target_df.columns: m &= target_df[a_col].notna()
        if y_cvd in target_df.columns: m &= target_df[y_cvd].notna()
        if y_diab in target_df.columns: m &= target_df[y_diab].notna()
        return m

    def get_valid_mask(target_df, pattern, domain):
        """ 计算指定干预域的合法保留掩码 / Calculate valid retention mask for specific intervention domain """
        # 真实删失或全勤验证 / Verify true dropouts or full completers
        mask_1111 = pattern == '1111'
        mask_1110 = (pattern == '1110') & (target_df['C3'] == 0 if 'C3' in target_df.columns else False)
        mask_1100 = (pattern == '1100') & (target_df['C2'] == 0 if 'C2' in target_df.columns else False)
        mask_1000 = (pattern == '1000') & (target_df['C1'] == 0 if 'C1' in target_df.columns else False)
        mask_censored_or_full = mask_1111 | mask_1110 | mask_1100 | mask_1000

        # 保护人群B: 满足严苛条件的高质量间歇缺失者 / Group B: High-quality intermittent missingness
        mask_0111 = (pattern == '0111') & (target_df[w1_l_cols].notna().sum(axis=1) >= w1_req) & check_subsequent_ay(target_df, 2, domain)
        mask_1011 = (pattern == '1011') & (target_df[w2_l_cols].notna().sum(axis=1) >= w2_req) & check_subsequent_ay(target_df, 3, domain)
        mask_1101 = (pattern == '1101') & (target_df[w3_l_cols].notna().sum(axis=1) >= w3_req) & check_subsequent_ay(target_df, 4, domain)

        return mask_censored_or_full | mask_0111 | mask_1011 | mask_1101

    # 获取两个干预维度的独立合法性掩码 / Get independent validity masks for both domains
    valid_OPA = get_valid_mask(df_target, opa_pattern, 'OPA')
    valid_LTPA = get_valid_mask(df_target, ltpa_pattern, 'LTPA')

    # 计算仅有一个模式符合条件的样本数 / Calculate samples where only one pattern is valid
    only_one_valid_count = (valid_OPA ^ valid_LTPA).sum()
    print(f"\n   > [审查/Review] 发现仅有单侧干预模式满足保留要求的样本数 (将被淘汰): {only_one_valid_count} / Found {only_one_valid_count} samples with only ONE valid intervention pattern (will be excluded).")

    # 执行终极过滤：必须直通两个干预模式都满足要求才保留 / Execute ultimate filtering: Retain ONLY if BOTH OPA and LTPA patterns are valid
    final_keep_mask = valid_OPA & valid_LTPA
    df_cleaned = df_target[final_keep_mask].copy()

    filtered_size = len(df_cleaned)
    original_size_before_exclusion = len(valid_OPA)
    print(f"   > 执行严格剔除前的样本量 / N before strict exclusion : {original_size_before_exclusion}")
    print(f"   > 筛选后保留 / Valid N                       : {filtered_size}")
    print(f"   > 物理删除数 / Dropped N                     : {original_size_before_exclusion - filtered_size}")
    print("   > [成功/Success] 已彻底剔除该队列中所有不符合双通道插补要求的残缺样本！ / Dropped all non-compliant samples failing dual-channel imputation requirements for this cohort!")

    total_n = len(df_cleaned)
    print("-" * 30)
    print(f"[{dataset_name}] 当前基线 (Wave 1) 最终样本量 / Final Baseline Size: {total_n} (默认 100% 在队 / 100% default)")
    if 'C1' in df_cleaned.columns:
        print(f"C1 (衔接 W1->W2) 缺失人数 / Missing in C1: {total_n - df_cleaned['C1'].sum()} (W2保留率/Retention: {df_cleaned['C1'].mean():.2%})")
    if 'C2' in df_cleaned.columns:
        print(f"C2 (衔接 W2->W3) 缺失人数 / Missing in C2: {total_n - df_cleaned['C2'].sum()} (W3保留率/Retention: {df_cleaned['C2'].mean():.2%})")
    if 'C3' in df_cleaned.columns:
        print(f"C3 (衔接 W3->W4) 缺失人数 / Missing in C3: {total_n - df_cleaned['C3'].sum()} (W4保留率/Retention: {df_cleaned['C3'].mean():.2%})")

    # 覆盖保存最终极致纯净的文件 / Overwrite save the ultimately clean file
    try:
        df_cleaned.to_excel(save_path, index=False)
        print(f"\n   > [保存/Save] {dataset_name} 已完成最终物理剔除并更新至: {save_path} / Finished filtering and successfully updated the path.")
    except Exception as e:
        print(f"\n   [错误/Error] {dataset_name} 更新保存失败 / Update save failed: {e}")

    return df_cleaned


# --- 依次将第六阶段过滤法则应用到三个生成的文件上 / Sequentially apply Phase 6 exclusion logic to the three generated datasets ---
df_base = execute_strict_exclusion(df_base, "基础合并队列 (Base)", final_output_path)
df_cvd  = execute_strict_exclusion(df_cvd,  "CVD 新发队列 (CVD)",  cvd_output_path)
df_diab = execute_strict_exclusion(df_diab, "Diabetes 新发队列 (Diabetes)", diab_output_path)

print("\n" + "=" * 50)
print("🎯 全部流水线任务完美竣工！所有特定队列已独立完成深层物理过滤并最终落盘。 / All pipeline tasks perfectly completed! Deep physical filtering executed independently on all specific cohorts and finally saved.")