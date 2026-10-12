"""
Script Summary (Independent Disease Cohorts Imputation):
This script performs Multiple Imputation by Chained Equations (MICE) using a mixed Random Forest engine
for two independent disease-specific cohorts: CVD and Diabetes. It preserves the original raw data scales
and focuses on imputing physical activity group variables and other covariates, with METs variables removed. /
脚本摘要（独立疾病队列插补）：
此脚本针对 CVD 和 Diabetes 两个独立疾病队列，使用混合随机森林引擎执行多重插补（MICE）。
脚本保留原始数据尺度，重点插补体力活动分组变量及其他协变量，并已移除 METs 变量。

Key Updates & Workflows / 核心更新与工作流:
1. Forward Absorbing State / 前向吸收态:
   - Once a disease event occurs (Y=1), all subsequent waves are strictly forced to 1. /
     一旦确诊疾病事件（Y=1），所有后续波次将被严格强制设定为 1。
   - After enforcing the absorbing state, the script prints cumulative event counts by wave for each disease file. /
     前向吸收态处理完成后，脚本会分别打印每个疾病文件中各波次结局事件的累积人数。

2. CVD Area Score Distribution / CVD 队列 area_score 分布:
   - For DATA_CVD_Final_Cleaned.xlsx, the script prints counts and proportions of area_score values 1, 2, and 3. /
     对 DATA_CVD_Final_Cleaned.xlsx，脚本会打印 area_score 中 1、2、3 的人数和比例。

3. Independent Cohort Imputation / 独立队列单轨插补:
   - Executes imputation independently for CVD and Diabetes incident cohort files. /
     分别针对 CVD 和 Diabetes 的新发队列文件独立执行插补流水线。

4. Robust Imputation & Censoring Rules / 稳健插补与删失规则:
   - Enforces LTMLE monotonic missingness rules through spatial-temporal shift correction for C1, C2, and C3. /
     通过对 C1、C2、C3 节点进行时空偏移修正，强制执行 LTMLE 单调缺失规则。

5. Engineering Polish / 工程优化:
   - Integrated garbage collection (gc.collect()) to reduce RAM pressure during repeated forest-based imputation. /
     集成垃圾回收机制（gc.collect()），降低多轮森林插补过程中的内存压力。
   - Added categorical bounding limits to ensure rounded predictions remain within observed category ranges. /
     添加分类变量边界限制，确保四舍五入后的预测值保留在观测到的类别范围内。

6. ID Isolation & Intentional Drop / ID 隔离与计划性丢弃:
   - Reads "Individual ID" strictly as string upfront, isolates it, and intentionally drops it to reduce analytical redundancy in subsequent models. /
     提前将 "Individual ID" 严格读取为字符串进行隔离，并在后续分析中计划性丢弃以减少模型冗余。

7. Exact Categorical Restoration / 分类变量精确还原:
   - Creates a mapping dictionary prior to encoding and maps integer codes back to exact original categorical labels post-imputation. /
     在编码前建立映射字典，并在插补完成后将整数编码精确映射回原始的类别标签。

8. Cumulative Incidence Risk Output / 累计发病风险输出:
   - Calculates and prints the absolute cumulative incidence risk across the entire follow-up period for the Diabetes cohort. /
     计算并打印糖尿病队列在整个随访期间的绝对累计发病风险。
"""

import os
import sys
import gc
import warnings

import numpy as np
import pandas as pd
from sklearn.base import BaseEstimator, RegressorMixin
from sklearn.ensemble import RandomForestClassifier, RandomForestRegressor
from sklearn.experimental import enable_iterative_imputer  # noqa: F401
from sklearn.impute import IterativeImputer

# ==========================================
# 0. 全局设置 / Global Settings
# ==========================================
warnings.filterwarnings("ignore")
pd.set_option("future.no_silent_downcasting", True)
pd.set_option("display.max_columns", None)
pd.set_option("display.width", 1000)

BASE_DIR = r"C:\Users\26627\Desktop"
OUTPUT_DIR = os.path.join(BASE_DIR, "database")
ID_COL = "Individual ID"  # 定义 ID 列常量 / Define ID column constant

# 分别定义两种疾病的输入文件路径 / Define input file paths for both diseases separately
CVD_INPUT_PATH = os.path.join(BASE_DIR, "DATA_CVD_Final_Cleaned.xlsx")
DIAB_INPUT_PATH = os.path.join(BASE_DIR, "DATA_diabetes_Final_Cleaned.xlsx")

os.makedirs(OUTPUT_DIR, exist_ok=True)

M_DATASETS = 30  # 多重插补数据集数量 / Number of multiply-imputed datasets

# 将 4 个波次的体力活动分组变量全部列出 / List all physical activity group variables across 4 waves
ALL_GROUP_COLS = [f"A{w}_OPA_Group" for w in range(1, 5)] + [f"A{w}_LTPA_Group" for w in range(1, 5)]


# ==========================================
# 辅助函数 / Helper Functions
# ==========================================

def section(title: str) -> None:
    """打印带分隔线的节标题 / Print a section header with separator."""
    print(f"\n{'=' * 50}\n{title}")


def print_outcome_cumulative_counts(df: pd.DataFrame, disease_name: str) -> None:
    """
    打印前向吸收态处理后的各波次结局累积人数。
    Print cumulative outcome event counts by wave after enforcing the forward absorbing state.
    """
    y_cols = [f"Y{w}_Event_{disease_name}" for w in range(1, 5)]

    print(f"\n   > [{disease_name}] 前向吸收态后结局累积人数 / "
          f"Cumulative event counts after forward absorbing state:")

    total_n = len(df)
    for col in y_cols:
        if col in df.columns:
            event_count = (pd.to_numeric(df[col], errors="coerce") == 1).sum()
            event_pct = event_count / total_n * 100 if total_n > 0 else np.nan
            print(f"      - {col}: {event_count}/{total_n} ({event_pct:.2f}%)")
        else:
            print(f"      - {col}: 未找到该列 / Column not found")


def print_cumulative_risk(df: pd.DataFrame, disease_name: str) -> None:
    """
    计算并打印 Diabetes 队列的整个随访期累计发病风险。
    Compute and print the cumulative incidence risk across the entire follow-up period for the Diabetes cohort.
    """
    if disease_name != "Diabetes":
        return

    y_cols = [f"Y{w}_Event_{disease_name}" for w in range(1, 5) if f"Y{w}_Event_{disease_name}" in df.columns]
    if not y_cols:
        print("   > [警告/Warning] Diabetes 结局列未找到 / No outcome columns found.")
        return

    # 只要任意一波发生事件（Y=1），即记为发病 / If an event (Y=1) occurs in any wave, flag as incident case
    has_event = (df[y_cols] == 1).any(axis=1)
    total_events = has_event.sum()
    total_n = len(df)

    risk_pct = (total_events / total_n) * 100 if total_n > 0 else 0

    print(f"\n   > [Diabetes] 随访期间累计发病风险 / Cumulative incidence risk: "
          f"{total_events}/{total_n} ({risk_pct:.2f}%)")


def print_area_score_distribution(df: pd.DataFrame, disease_name: str) -> None:
    """
    打印 CVD 队列中 area_score 取值为 1、2、3 的人数和比例。
    Print counts and proportions of area_score values 1, 2, and 3 in the CVD cohort.
    """
    if disease_name != "CVD":
        return

    print("\n   > [CVD] area_score 分布统计 / area_score distribution:")

    if "area_score" not in df.columns:
        print("      - 未找到 area_score 列 / Column area_score not found.")
        return

    area_series = pd.to_numeric(df["area_score"], errors="coerce")
    total_n = len(df)
    nonmissing_n = area_series.notna().sum()

    for value in [1, 2, 3]:
        count = (area_series == value).sum()
        pct_total = count / total_n * 100 if total_n > 0 else np.nan
        pct_nonmissing = count / nonmissing_n * 100 if nonmissing_n > 0 else np.nan

        print(
            f"      - area_score = {value}: "
            f"{count}/{total_n} ({pct_total:.2f}% of total; "
            f"{pct_nonmissing:.2f}% of non-missing)"
        )


# ==========================================
# 预定义导出列顺序 / Predefine Ordered Columns for Export
# ==========================================
# 尽管 ID 将被丢弃，但保留在基础变量列表中，动态导出时会自动跳过它 /
# ID is kept in BASE_VARS, but will be safely ignored during dynamic export if stripped
BASE_VARS = [ID_COL, "V_Gender", "V_Edu", "V_community_level_ses"]

GROUP_CVD_COLS_ORDERED = BASE_VARS + [
    "V_Bio_HbA1c_Pre", "V_Bio_CRP_Pre", "V_Bio_TyG", "V_Bio_LDL_Pre", "V_Bio_HDL_Pre",
    "L1_disease_scores_CVD_event", "L1_Martial_Status", "L1_Household_SES", "L1_Age",
    "L1_Sleep_time", "L1_BMI", "L1_Smoke", "L1_Drink", "L1_Work_status",
    "L1_Living_with_Child", "L1_Cognitive_functioning", "L1_Depression_scale",
    "A1_OPA_Group", "A1_LTPA_Group", "C1", "Y1_Event_CVD",
    "L2_disease_scores_CVD_event", "L2_Martial_Status", "L2_Household_SES", "L2_Age",
    "L2_Sleep_time", "L2_BMI", "L2_Bio_CRP_Pre", "L2_Bio_LDL_Pre", "L2_Bio_HDL_Pre",
    "L2_Bio_HbA1c_Pre", "L2_Bio_TyG", "L2_Smoke", "L2_Drink", "L2_Work_status",
    "L2_Living_with_Child", "L2_Cognitive_functioning", "L2_Depression_scale",
    "A2_OPA_Group", "A2_LTPA_Group", "C2", "Y2_Event_CVD",
    "L3_disease_scores_CVD_event", "L3_Martial_Status", "L3_Household_SES", "L3_Age",
    "L3_Sleep_time", "L3_Smoke", "L3_Drink", "L3_Work_status",
    "L3_Living_with_Child", "L3_Cognitive_functioning", "L3_Depression_scale",
    "A3_OPA_Group", "A3_LTPA_Group", "C3", "Y3_Event_CVD",
    "L4_disease_scores_CVD_event", "L4_Martial_Status", "L4_Household_SES", "L4_Age",
    "L4_Sleep_time", "L4_Smoke", "L4_Drink", "L4_Work_status",
    "L4_Living_with_Child", "L4_Cognitive_functioning", "L4_Depression_scale",
    "A4_OPA_Group", "A4_LTPA_Group", "Y4_Event_CVD",
]

GROUP_DIAB_COLS_ORDERED = BASE_VARS + [
    "V_Bio_HbA1c_Pre", "V_Bio_CRP_Pre", "V_Bio_TyG", "V_Bio_HDL_Pre","V_Bio_LDL_Pre",
    "L1_disease_scores_Diabetes_event", "L1_Martial_Status", "L1_Household_SES", "L1_Age",
    "L1_Sleep_time", "L1_BMI", "L1_Smoke", "L1_Drink", "L1_Work_status",
    "L1_Living_with_Child", "L1_Cognitive_functioning", "L1_Depression_scale",
    "A1_OPA_Group", "A1_LTPA_Group", "C1", "Y1_Event_Diabetes",
    "L2_disease_scores_Diabetes_event", "L2_Martial_Status", "L2_Household_SES", "L2_Age",
    "L2_Sleep_time", "L2_BMI", "L2_Bio_CRP_Pre", "L2_Bio_LDL_Pre", "L2_Bio_HDL_Pre",
    "L2_Bio_HbA1c_Pre", "L2_Bio_TyG", "L2_Smoke", "L2_Drink", "L2_Work_status",
    "L2_Living_with_Child", "L2_Cognitive_functioning", "L2_Depression_scale",
    "A2_OPA_Group", "A2_LTPA_Group", "C2", "Y2_Event_Diabetes",
    "L3_disease_scores_Diabetes_event", "L3_Martial_Status", "L3_Household_SES", "L3_Age",
    "L3_Sleep_time", "L3_Smoke", "L3_Drink", "L3_Work_status",
    "L3_Living_with_Child", "L3_Cognitive_functioning", "L3_Depression_scale",
    "A3_OPA_Group", "A3_LTPA_Group", "C3", "Y3_Event_Diabetes",
    "L4_disease_scores_Diabetes_event", "L4_Martial_Status", "L4_Household_SES", "L4_Age",
    "L4_Sleep_time", "L4_Smoke", "L4_Drink", "L4_Work_status",
    "L4_Living_with_Child", "L4_Cognitive_functioning", "L4_Depression_scale",
    "A4_OPA_Group", "A4_LTPA_Group", "Y4_Event_Diabetes",
]

CATEGORICAL_COLS = [
    "V_Gender", "V_Edu", "V_community_level_ses",
    "L1_disease_scores_CVD_event", "L1_disease_scores_Diabetes_event",
    "L1_Household_SES", "L1_Martial_Status", "L1_Smoke", "L1_Drink",
    "L1_Work_status", "L1_Living_with_Child",
    "A1_OPA_Group", "A1_LTPA_Group", "Y1_Event_CVD", "Y1_Event_Diabetes",

    "L2_disease_scores_CVD_event", "L2_disease_scores_Diabetes_event",
    "L2_Household_SES", "L2_Martial_Status", "L2_Smoke", "L2_Drink",
    "L2_Work_status", "L2_Living_with_Child",
    "A2_OPA_Group", "A2_LTPA_Group", "Y2_Event_CVD", "Y2_Event_Diabetes",

    "L3_disease_scores_CVD_event", "L3_disease_scores_Diabetes_event",
    "L3_Household_SES", "L3_Martial_Status", "L3_Smoke", "L3_Drink",
    "L3_Work_status", "L3_Living_with_Child",
    "A3_OPA_Group", "A3_LTPA_Group", "Y3_Event_CVD", "Y3_Event_Diabetes",

    "L4_disease_scores_CVD_event", "L4_disease_scores_Diabetes_event",
    "L4_Household_SES", "L4_Martial_Status", "L4_Smoke", "L4_Drink",
    "L4_Work_status", "L4_Living_with_Child",
    "A4_OPA_Group", "A4_LTPA_Group", "Y4_Event_CVD", "Y4_Event_Diabetes",
]

NUMERIC_COLS = [
    "V_Bio_HbA1c_Pre", "V_Bio_TyG", "V_Bio_CRP_Pre", "V_Bio_LDL_Pre",
    "L1_Age", "L1_Sleep_time", "L1_BMI", "L1_Cognitive_functioning", "L1_Depression_scale",
    "L2_Age", "L2_Sleep_time", "L2_BMI",
    "L2_Bio_HbA1c_Pre", "L2_Bio_TyG", "L2_Bio_CRP_Pre", "L2_Bio_LDL_Pre",
    "L2_Cognitive_functioning", "L2_Depression_scale",
    "L3_Age", "L3_Sleep_time", "L3_Cognitive_functioning", "L3_Depression_scale",
    "L4_Age", "L4_Sleep_time", "L4_Cognitive_functioning", "L4_Depression_scale"
]


# ==========================================
# 混合随机森林智能路由包装类 / Mixed Random Forest Intelligent Wrapper
# ==========================================
class MixedRandomForest(BaseEstimator, RegressorMixin):
    """
    智能包装类：若目标变量为分类变量 (全整数且种类 ≤ 10)，自动召唤 RandomForestClassifier；
    否则召唤 RandomForestRegressor，解决 IterativeImputer 的单一估计器限制。

    Intelligent wrapper: Automatically routes categorical targets (all-integer, ≤ 10 unique values)
    to RandomForestClassifier, and continuous targets to RandomForestRegressor,
    solving IterativeImputer's single-estimator limitation.
    """

    def __init__(self, n_estimators=50, max_features=0.6, max_depth=50,
                 min_samples_leaf=3, random_state=None):
        self.n_estimators = n_estimators
        self.max_features = max_features
        self.max_depth = max_depth
        self.min_samples_leaf = min_samples_leaf
        self.random_state = random_state
        self.model_ = None

    def fit(self, X, y):
        valid_y = y[~np.isnan(y)]
        is_categorical = np.all(valid_y == np.round(valid_y)) and len(np.unique(valid_y)) <= 10

        ModelClass = RandomForestClassifier if is_categorical else RandomForestRegressor
        self.model_ = ModelClass(
            n_estimators=self.n_estimators,
            max_features=self.max_features,
            max_depth=self.max_depth,
            min_samples_leaf=self.min_samples_leaf,
            random_state=self.random_state,
            n_jobs=-1,
        )
        self.model_.fit(X, y.astype(int) if is_categorical else y)
        return self

    def predict(self, X):
        return self.model_.predict(X)


# ==========================================
# 辅助：删失掩码时空偏移修正 / Helper: Censoring mask spatial-temporal shift correction
# ==========================================
def apply_censoring_masks(df_imp: pd.DataFrame) -> pd.DataFrame:
    """
    按照 LTMLE 单调缺失规则，强制清空因删失导致不应存在的后续波次数据。
    Enforces LTMLE monotonic missingness rules: strictly clears subsequent-wave data that should not exist due to censoring.
    """
    wave_cols = {
        w: [c for c in df_imp.columns if f"L{w}_" in c or f"A{w}_" in c or f"Y{w}_" in c]
        for w in range(2, 5)
    }
    censoring_rules = [
        ("C1", wave_cols[2] + wave_cols[3] + wave_cols[4] + ["C2", "C3"]),
        ("C2", wave_cols[3] + wave_cols[4] + ["C3"]),
        ("C3", wave_cols[4]),
    ]
    for c_node, cols_to_clear in censoring_rules:
        if c_node in df_imp.columns:
            mask = df_imp[c_node] == 0
            valid_cols = [c for c in cols_to_clear if c in df_imp.columns]
            if mask.any() and valid_cols:
                df_imp.loc[mask, valid_cols] = np.nan
    return df_imp


# ==========================================
# 辅助：执行单次插补并还原 / Helper: Run a single imputation pass and restore columns
# ==========================================
def run_imputation(df_numeric: pd.DataFrame, random_seed: int,
                   track_label: str, cat_cols_list: list, const_cols_dict: dict) -> pd.DataFrame:
    """
    对给定的数值矩阵执行一次 MICE-RF 插补，并还原常量列。
    Run one MICE-RF imputation pass on the given numeric matrix and restore constant columns.
    """
    df_impute_input = df_numeric.copy()

    # 0. 记录分类变量的极值，防越界 / Record categorical min/max for bounds clipping
    cat_bounds = {}
    for c in cat_cols_list:
        if c in df_impute_input.columns:
            cat_bounds[c] = (df_impute_input[c].min(), df_impute_input[c].max())

    # 1. 核心 MICE 插补 / Core MICE Imputation
    imputer = IterativeImputer(
        estimator=MixedRandomForest(
            n_estimators=60,
            max_features=0.6,
            max_depth=50,
            min_samples_leaf=5,
            random_state=random_seed,
        ),
        max_iter=8,
        tol=0.01,  # 允许适当的收敛容忍度 / Allow reasonable convergence tolerance
        random_state=random_seed,
        initial_strategy="mean",
    )

    imputed_arr = imputer.fit_transform(df_impute_input)
    imputed_df = pd.DataFrame(imputed_arr, columns=df_impute_input.columns, index=df_impute_input.index)

    # 2. 后处理：确保分类变量严格取整数且不越界 / Post-processing: Round and clip categorical columns strictly
    for col in cat_cols_list:
        if col in imputed_df.columns:
            min_val, max_val = cat_bounds[col]
            imputed_df[col] = imputed_df[col].round().clip(lower=min_val, upper=max_val)

    # 应用删失掩码修正 / Apply censoring mask correction
    imputed_df = apply_censoring_masks(imputed_df)

    # 还原常量列 / Restore constant columns
    for col_name, col_val in const_cols_dict.items():
        imputed_df[col_name] = col_val

    print(f"({track_label}: 迭代 {imputer.n_iter_} 次 / Iter {imputer.n_iter_})", end=" ", flush=True)
    return imputed_df


# ==========================================
# 核心业务管道：按特定疾病单独处理 / Core Pipeline: Process Specific Disease Independently
# ==========================================
def process_disease_pipeline(disease_name: str, input_path: str, group_cols_ordered: list):
    """
    针对特定的疾病执行数据读取、清洗与多重插补的完整流水线。
    Executes the full pipeline of data reading, cleaning, and multiple imputation for a specific disease.
    """
    section(f"开始处理 {disease_name} 数据流 / Starting pipeline for {disease_name}\n   > 文件/File: {input_path}")

    # ==========================================
    # 1. 读取数据 / 1. Read Data
    # ==========================================
    try:
        # 强制将 ID_COL 提前读取为字符串格式 / Force reading ID_COL as string upfront
        df = pd.read_excel(input_path, dtype={ID_COL: str})
    except Exception as e:
        sys.exit(f"   [致命错误/Fatal Error] 读取失败 / Read failed: {e}")

    print(f"   > 原始样本量 / Original sample size: {len(df)}")

    # 打印 CVD 队列 area_score 分布 / Print area_score distribution for the CVD cohort
    print_area_score_distribution(df, disease_name)

    # 1.1 ID列隔离保护与丢弃 / 1.1 Isolate and intentionally drop ID Column
    if ID_COL in df.columns:
        id_series = df.pop(ID_COL)  # 剥离 ID，不再进入后续插补分析 / Strip ID completely
        # 按照用户要求，此处刻意不再拼回原表以减小后续 LTMLE 冗余 /
        # Per user request, intentionally NOT re-attached later to reduce LTMLE redundancy.
    else:
        id_series = None
        print(f"   > [警告/Warning] 未找到 '{ID_COL}' 列 / Column '{ID_COL}' not found.")

    # ==========================================
    # 1.5 结局变量前向吸收态 / 1.5 Outcome Forward Absorbing State
    # ==========================================
    print(f"\n   > 正在处理结局变量疾病吸收态 / Processing absorbing state for outcomes...")
    y_cols = [f"Y{w}_Event_{disease_name}" for w in range(1, 5) if f"Y{w}_Event_{disease_name}" in df.columns]
    if y_cols:
        orig = df[y_cols].copy()
        df[y_cols] = df[y_cols].mask((df[y_cols] == 1).cummax(axis=1), 1)
        n_changed = ((orig != 1) & (df[y_cols] == 1)).sum().sum()
        print(f"   > [{disease_name}] 吸收态强制完成，{n_changed} 个后续状态更正为 1。"
              f" / Absorbing state enforced, {n_changed} subsequent states corrected to 1.")

        # 打印前向吸收态处理后的结局累积人数 / Print cumulative event counts after absorbing-state enforcement
        print_outcome_cumulative_counts(df, disease_name)
        # 计算并打印 Diabetes 队列累计发病风险 / Compute and print Diabetes cumulative risk
        print_cumulative_risk(df, disease_name)
    else:
        print(f"   > [警告/Warning] 未找到 {disease_name} 的结局变量列 / "
              f"No outcome columns found for {disease_name}.")

    # ==========================================
    # 2. 深度清洗 / 2. Deep Cleaning
    # ==========================================
    df.drop(columns=[c for c in ["Household ID", "Community ID"] if c in df.columns], inplace=True)

    constant_cols = {
        col: (df[col].dropna().iloc[0] if df[col].dropna().size > 0 else np.nan)
        for col in df.columns if df[col].nunique(dropna=True) <= 1
    }
    if constant_cols:
        df.drop(columns=list(constant_cols), inplace=True)
        print(f"   > [清洗/Cleaning] 移除常量列 {len(constant_cols)} 个 (插补后还原) / "
              f"Removed {len(constant_cols)} constant cols (to restore post-imputation).")

    sparse_cols = df.columns[df.isnull().mean() > 0.95].tolist()
    if sparse_cols:
        df.drop(columns=sparse_cols, inplace=True)
        print(f"   > [清洗/Cleaning] 剔除极度稀疏列 (>95% 缺失) / Dropped sparse cols (>95% missing): {sparse_cols}")

    # ==========================================
    # 3. 变量类型确认与编码 / 3. Variable Typing and Encoding
    # ==========================================
    cat_cols = [c for c in CATEGORICAL_COLS if c in df.columns]
    num_cols = [c for c in NUMERIC_COLS if c in df.columns]

    # V_Gender 原始编码 1=男, 2=女 → 插补前统一减 1，转为 0=男, 1=女（与 R 分析脚本 HDL 分箱等逻辑一致）
    # V_Gender raw coding 1=Male, 2=Female -> subtract 1 before imputation (0=Male, 1=Female)
    if "V_Gender" in df.columns:
        df["V_Gender"] = pd.to_numeric(df["V_Gender"], errors="coerce") - 1

    for col in num_cols:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    # 建立映射字典的密码本 / Build mapping dictionary for categorical restoration
    cat_mappings = {}
    df_impute_ready = df.copy()

    for col in cat_cols:
        # 【极简修复】：转为字符串 -> 用正则抹除结尾的 '.0' -> 还原 nan
        clean_series = df[col].astype(str).replace(r'\.0$', '', regex=True).replace('nan', np.nan)
        category_series = pd.Categorical(clean_series)

        # 记录映射关系：例如 {0: '1', 1: 'High'} / Record mappings
        cat_mappings[col] = {i: cat for i, cat in enumerate(category_series.categories)}

        # 转换为整数编码，缺失值设为 NaN / Convert to integer codes, setting missing values to NaN
        df_impute_ready[col] = category_series.codes
        df_impute_ready[col] = df_impute_ready[col].replace(-1, np.nan)

    # ==========================================
    # 4. 运行插补并导出 / 4. Run Imputation & Export
    # ==========================================
    print(f"\n   > 开始执行 {disease_name} MICE-RF 多重插补 / Starting MICE-RF Imputation for {disease_name}")

    for i in range(1, M_DATASETS + 1):
        print(f"     [Dataset {i:02d}/{M_DATASETS}] 正在插补 / Imputing ... ", end="", flush=True)

        try:
            # 执行插补 / Run imputation on the dataset
            imputed_group = run_imputation(
                df_impute_ready, random_seed=i + 5000, track_label="Main",
                cat_cols_list=cat_cols, const_cols_dict=constant_cols
            )
            # 还原分类变量 / Restore categorical variables using mappings
            for col in cat_cols:
                if col in imputed_group.columns:
                    # 1. 对整列执行向量化的四舍五入 (例如将 1.999 变为 2.0)
                    rounded_series = np.round(imputed_group[col])

                    # 2. 向量化转为 Pandas 特有的容忍缺失值的整数类型 (Int64)
                    # 这一步去掉了小数点，完美契合你的字典键类型
                    int_keys = rounded_series.astype("Int64")

                    # 3. 直接通过 map 映射字典，Pandas 底层 C 代码执行，速度极快
                    imputed_group[col] = int_keys.map(cat_mappings[col])

            # 按照预定义的列顺序提取并保存 (由于 ID 被刻意丢弃，此处动态过滤避免报错) /
            # Extract columns in predefined order (dynamically filtering out dropped ID)
            g_cols = [c for c in group_cols_ordered if c in imputed_group.columns]
            imputed_group[g_cols].to_csv(
                os.path.join(OUTPUT_DIR, f"Group_DATA_LTMLE_Ready_{disease_name}_{i}.csv"),
                index=False, encoding="utf-8-sig"
            )

            print(f"| 已保存 {disease_name} 数据集 / Saved {disease_name} dataset")

            # 内存回收防御，防止多轮深层森林后 RAM 溢出 / Memory GC to prevent RAM overflow
            del imputed_group
            gc.collect()

        except Exception as e:
            print(f"\n       [错误/Error]: {e}")


# ==========================================
# 🚀 启动流水线 / Launch Pipeline
# ==========================================

# 1. 处理 CVD 队列 / Process CVD Cohort
process_disease_pipeline("CVD", CVD_INPUT_PATH, GROUP_CVD_COLS_ORDERED)

# 2. 处理 Diabetes 队列 / Process Diabetes Cohort
process_disease_pipeline("Diabetes", DIAB_INPUT_PATH, GROUP_DIAB_COLS_ORDERED)

print("\n" + "=" * 50)
print(f"🎯 全部流水线任务完美竣工！疾病独立队列插补完成。")
print(f"All pipeline tasks perfectly completed! Independent disease cohort imputation finished.")