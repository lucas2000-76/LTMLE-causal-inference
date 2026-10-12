# ==============================================================================
# Script Name: TE_and_ATE_category_new_SA.R (Post-hoc sensitivity analyses; gbounds 0.05/0.95; LTPA dichotomized 0 vs 1-3)
# Description:
#     Post-hoc sensitivity analyses for the Physical Activity Paradox study.
#     本脚本是 TE_and_ATE_category_new.R 的事后敏感性分析版本，不重复运行主分析。
#     This script is the post-hoc sensitivity-analysis variant of TE_and_ATE_category_new.R;
#     it does NOT rerun the plain main analysis.
#
#     它复用与主分析完全相同的 LTMLE/MSM 流程、参数、预处理与 OPA 退出范式，
#     It reuses the exact same LTMLE/MSM pipeline, parameters, preprocessing, and OPA withdrawal regimes as the main analysis.
#
#     唯一区别：在读取每个插补数据集后，删除某一类协变量（非累积），共 3 个情景：
#     The only difference: after reading each imputed dataset, one covariate group is excluded (non-cumulative), across 3 scenarios:
#       1) objective biomarkers / 客观生物标志物
#       2) socioeconomic indicators / 社会经济指标
#       3) mental health & cognitive measures / 心理健康与认知指标
#
#     输出：Main_Pooled_MSM_All_Results_group_SA_<scenario>.xlsx（每个情景一个文件）。
#     Output: Main_Pooled_MSM_All_Results_group_SA_<scenario>.xlsx (one file per scenario).
#
#     干预策略 / Intervention Regimes:
#     针对职业体力活动 (OPA) 采用“退出范式” (Withdrawal Paradigm):
#     "0000: Always Low Exposure / 始终轻体力 (Ref)",
#     "1000: Withdrawal at Wave 2 / 第2波起脱离重体力",
#     "1100: Withdrawal at Wave 3 / 第3波起脱离重体力",
#     "1110: Withdrawal at Wave 4 / 第4波起脱离重体力",
#     "1111: Sustained High Exposure / 持续高暴露"
#
# [Architecture]:
#     采用并行计算架构，结合 foreach 多核并发与分离的数据网络映射。
#     Uses a parallel computing architecture combining foreach-based multicore execution
#     with isolated data-network mapping across imputed datasets.
# ==============================================================================

# ==============================================================================
# 1. 环境初始化与包管理 / Environment Initialization & Package Management
# ==============================================================================
# 清空环境变量和内存 / Clear environment variables and run garbage collection
rm(list = ls()); gc()

# 加载核心包 / Load core packages
libs <- c("ltmle", "SuperLearner", "fastDummies", "dplyr", "stringr", "readr", "openxlsx",
          "parallel", "doParallel", "foreach", "doRNG", "caret", "gtools", "ranger", "gam", "Amelia")

# 检查并安装缺失的包 / Check and install missing packages
missing_libs <- !libs %in% installed.packages()
if (any(missing_libs)) install.packages(libs[missing_libs])
invisible(lapply(libs, library, character.only = TRUE))

# 设置基础工作路径 / Set base working directory
setwd("C:/Users/26627/Desktop/database/")
set.seed(395702)

# 创建并指定桌面输出文件夹 / Create and specify output directory on Desktop
base_dir <- "C:/Users/26627/Desktop"
result_dir <- file.path(base_dir, "result")
if(!dir.exists(result_dir)) {
  dir.create(result_dir, recursive = TRUE)
  message(paste(">>> 已成功在桌面创建文件夹 / Directory successfully created at:", result_dir))
}

# 开启多核并行计算，预留部分核心给系统 / Enable multicore parallel computing while reserving cores for the system
num_cores <- max(1, detectCores(logical = FALSE) - 6)
cl <- makeCluster(num_cores)
registerDoParallel(cl)
message(paste0(">>> 并行集群已启动，分配核心数 / Parallel cluster started, allocated cores: ", num_cores))

# ==============================================================================
# 2. 参数设定与核心函数外部封装 / Parameters & Core Functions External Setup
# ==============================================================================
# 设定插补数据集总数 / Set total number of imputed datasets
M_datasets <- 30

# 追踪主分析的相关指标 / Track metrics for main analysis (新增RD)
cols_to_track <- c(
  "Risk_R1", "Risk_R2", "Risk_R3", "Risk_R4", "Risk_R5",
  "logRR_R2", "logRR_R3", "logRR_R4", "logRR_R5",
  "RD_R2", "RD_R3", "RD_R4", "RD_R5"  # 新增：风险差指标
)

# --- 定义算法库 / Define algorithm library ---
# 封装快速版随机森林以适配并行运算 / Wrap fast random forest for parallel computation
SL.ranger.fast <- function(Y, X, newX, family, obsWeights, id, ...) {
  SuperLearner::SL.ranger(Y = Y, X = X, newX = newX, family = family,
                          obsWeights = obsWeights, id = id,
                          num.trees = 200, num.threads = 1, ...)
}

# 设置 SuperLearner 预测算法库 / Set SuperLearner algorithm library
SLlib_ML <- list(
  Q = c("SL.mean", "SL.glm", "SL.gam"),
  g = c("SL.mean", "SL.glm", "SL.gam", "SL.ranger.fast")
)

# --- Delta Method 提取函数 / Delta Method Extraction Function (更新包含 RD) ---
# 从拟合的 MSM 模型中提取风险、logRR、RD 与标准误 / Extract risks, logRRs, RDs and standard errors from fitted MSM
extract_msm_results <- function(fit, res_q, res_se, i_idx, n_obs) {
  expit <- function(x) { 1 / (1 + exp(-x)) }
  
  bm <- fit$beta
  ic_m <- fit$IC
  
  # 对照组的基准发病率 / Baseline incidence
  p0 <- expit(bm[1])
  res_q[i_idx, "Risk_R1"] <- p0
  
  # Risk_R1 的 SE
  IC_p0 <- p0 * (1 - p0) * ic_m[, 1]
  res_se[i_idx, "Risk_R1"] <- sqrt(var(IC_p0) / n_obs)
  
  for(x in 2:5) {
    pX <- expit(bm[1] + bm[x])
    res_q[i_idx, paste0("Risk_R", x)] <- pX
    
    # ======== 计算 logRR 及其 SE ========
    logRR <- log(pX / p0)
    res_q[i_idx, paste0("logRR_R", x)] <- logRR
    
    grad_b1_RR <- p0 - pX
    grad_bx_RR <- 1 - pX
    IC_RR <- grad_b1_RR * ic_m[, 1] + grad_bx_RR * ic_m[, x]
    res_se[i_idx, paste0("logRR_R", x)] <- sqrt(var(IC_RR) / n_obs)
    
    # ======== 计算 RD (Risk Difference) 及其 SE ========
    RD <- pX - p0
    res_q[i_idx, paste0("RD_R", x)] <- RD
    
    # RD 的 Delta Method 偏导数链式法则 / Partial derivatives for RD
    grad_b1_RD <- pX * (1 - pX) - p0 * (1 - p0)
    grad_bx_RD <- pX * (1 - pX)
    
    IC_RD <- grad_b1_RD * ic_m[, 1] + grad_bx_RD * ic_m[, x]
    res_se[i_idx, paste0("RD_R", x)] <- sqrt(var(IC_RD) / n_obs)
  }
  
  return(list(q = res_q, se = res_se))
}

# ==============================================================================
# 3. 开启双结局大循环 / Start Double-Outcome Main Loop
# ==============================================================================
target_diseases <- c("CVD", "Diabetes")

# ==============================================================================
# 2.5 敏感性分析方案定义 / Post hoc Sensitivity Analysis Scenario Definitions
# ==============================================================================
# 注意：各情景【非累积】——每个情景只删除本情景对应的变量，
#       其余变量与分析流程与原主分析（TE_and_ATE_category_new.R）完全一致。
# Non-cumulative: each scenario excludes ONLY its own covariate set.
sa_scenarios <- list(
  list(
    name    = "objective_biomarkers",
    label   = "Excluding objective biomarkers",
    exclude = c("V_Bio_HbA1c_Pre", "V_Bio_CRP_Pre", "V_Bio_TyG", "V_Bio_LDL_Pre", "V_Bio_HDL_Pre",
                "L2_Bio_CRP_Pre", "L2_Bio_LDL_Pre", "L2_Bio_HDL_Pre", "L2_Bio_HbA1c_Pre", "L2_Bio_TyG")
  ),
  list(
    name    = "socioeconomic",
    label   = "Excluding socioeconomic indicators (community-level & household SES)",
    exclude = c("V_community_level_ses", "L1_Household_SES", "L2_Household_SES",
                "L3_Household_SES", "L4_Household_SES")
  ),
  list(
    name    = "mental_health_cognitive",
    label   = "Excluding mental health and cognitive measures",
    exclude = c("L1_Cognitive_functioning", "L1_Depression_scale",
                "L2_Cognitive_functioning", "L2_Depression_scale",
                "L3_Cognitive_functioning", "L3_Depression_scale",
                "L4_Cognitive_functioning", "L4_Depression_scale")
  )
)

# 预检：确认待删除变量在 CVD / Diabetes 数据中是否存在 / Pre-flight existence check
preflight_check <- function() {
  for (disease in target_diseases) {
    f0 <- paste0("Group_DATA_LTMLE_Ready_", disease, "_1.csv")
    if (!file.exists(f0)) {
      message(sprintf("[Pre-flight] File not found, skipped: %s", f0))
      next
    }
    hdr <- names(read_csv(f0, n_max = 0, show_col_types = FALSE))
    for (sc in sa_scenarios) {
      present <- intersect(sc$exclude, hdr)
      missing <- setdiff(sc$exclude, hdr)
      message(sprintf("[Pre-flight | %s | %s] present %d/%d, missing %d",
                      disease, sc$name, length(present), length(sc$exclude), length(missing)))
      if (length(missing) > 0) {
        message("    !! MISSING: ", paste(missing, collapse = ", "))
      }
    }
  }
}
preflight_check()

# ==============================================================================
# 2.6 敏感性分析运行函数 / SA Runner
#     与原主分析使用完全相同的参数、预处理与 LTMLE/MSM 流程，
#     仅在读取每个插补数据集后删除 exclude_vars 中存在的列。
# ==============================================================================
run_sa_scenario <- function(scenario_name, scenario_label, exclude_vars, scenario_idx) {

  # 每个情景独立设置种子，保证可复现（不同情景使用不同种子）
  set.seed(395702 + scenario_idx)

  message("\n>>> Starting Sensitivity Analysis: ", scenario_label,
          " | excluded covariates: ", length(exclude_vars))

  # 初始化该情景的 Excel 工作簿 / Initialize scenario workbook
  final_wb <- createWorkbook()


for (disease in target_diseases) {
  
  cat(sprintf("\n\n=======================================================================================\n"))
  cat(sprintf(" 🚀 正在启动 [%s] 的分析流程... / Launching analysis pipeline for [%s]... \n", disease, disease))
  cat(sprintf("=======================================================================================\n"))
  
  # 为当前疾病创建 Excel Sheet（仅主表） / Create Excel sheet for the current disease (main analysis only)
  addWorksheet(final_wb, sheetName = disease)
  
  current_excel_row <- 1
  
  # ==============================================================================
  # 3.1 foreach 核心并行计算 / Core Parallel Computation Using foreach
  # ==============================================================================
  results_list <- foreach(i = 1:M_datasets,
                          .packages = libs,
                          .export = c("extract_msm_results", "SL.ranger.fast", "SLlib_ML", "cols_to_track", "disease", "exclude_vars"),
                          .errorhandling = 'pass') %dorng% {
                            
                            # 注册随机森林函数，防止多核环境中函数丢失 / Register random forest wrapper to avoid loss in worker environments
                            assign("SL.ranger.fast", SL.ranger.fast, envir = .GlobalEnv)
                            
                            # 载入单一插补数据集 / Load one imputed dataset
                            input_file <- paste0("Group_DATA_LTMLE_Ready_", disease, "_", i, ".csv")
                            if(!file.exists(input_file)) return(NULL)
                            df <- read_csv(input_file, show_col_types = FALSE)
                            # ===== SA: drop this scenario's covariates (non-cumulative) =====
                            if (length(exclude_vars) > 0) {
                              drop_cols <- intersect(exclude_vars, colnames(df))
                              if (length(drop_cols) > 0) {
                                df <- df[, setdiff(colnames(df), drop_cols)]
                              }
                            }
                            
                            # 构建 3 种 OPA 二分类阈值与 LTPA 节点 / Build three binary OPA thresholds and LTPA nodes
                            for(w in 1:4) {
                              opa_col <- paste0("A", w, "_OPA_Group")
                              ltpa_col <- paste0("A", w, "_LTPA_Group")
                              
                              if(opa_col %in% colnames(df)) {
                                v_opa <- as.numeric(as.character(df[[opa_col]]))
                                df[[paste0(opa_col, "_T1")]] <- ifelse(v_opa >= 1, 1, 0)
                                df[[paste0(opa_col, "_T2")]] <- ifelse(v_opa >= 2, 1, 0)
                                df[[paste0(opa_col, "_T3")]] <- ifelse(v_opa >= 3, 1, 0)
                                df[[paste0(opa_col, "_T4")]] <- ifelse(v_opa >= 4, 1, 0)
                                df[[opa_col]] <- as.factor(df[[opa_col]])
                              }
                              
                              if(ltpa_col %in% colnames(df)) {
                                v_ltpa <- as.numeric(as.character(df[[ltpa_col]]))
                                
                                # 主分析：LTPA 二分类，0 = 无 LTPA；原始 1,2,3 合并为 1 = 有任何 LTPA / Main: binary LTPA, 0 = no LTPA; original 1,2,3 combined as 1 = any LTPA
                                df[[paste0(ltpa_col, "_Factor")]] <- factor(
                                  as.numeric(v_ltpa >= 1),
                                  levels = c(0, 1),
                                  labels = c("No_LTPA", "Any_LTPA")
                                )
                                
                                
                                df[[ltpa_col]] <- as.factor(df[[ltpa_col]])
                              }
                            }
                            
                            binned_vars <- c()
                            
                            # ==============================================================================
                            # 连续变量临床指南分箱 / Binning Continuous Variables Using Clinical Cutoffs
                            # ==============================================================================
                            for(j in 1:4) {
                              # 睡眠时间 U 型分箱 / U-shaped sleep duration binning
                              sleep_col <- paste0("L", j, "_Sleep_time")
                              if(sleep_col %in% colnames(df)) {
                                val <- as.numeric(as.character(df[[sleep_col]]))
                                df[[sleep_col]] <- cut(val, breaks = c(-Inf, 6, 8.00001, Inf), right = FALSE, labels = c("Short", "Normal", "Long"))
                                df[[sleep_col]] <- relevel(df[[sleep_col]], ref = "Normal")
                                binned_vars <- c(binned_vars, sleep_col)
                              }
                              
                              # 抑郁量表分箱 / Depression scale binning
                              dep_col <- paste0("L", j, "_Depression_scale")
                              if(dep_col %in% colnames(df)) {
                                val <- as.numeric(as.character(df[[dep_col]]))
                                df[[dep_col]] <- as.factor(ifelse(val >= 10, "High_Depression", "Low_Depression"))
                                df[[dep_col]] <- relevel(df[[dep_col]], ref = "Low_Depression")
                                binned_vars <- c(binned_vars, dep_col)
                              }
                            }
                            
                            # 糖化血红蛋白分箱 / HbA1c binning
                            hba1c_cols <- grep("HbA1c_Pre", colnames(df), value = TRUE)
                            for(hba1c_col in hba1c_cols) {
                              if(hba1c_col %in% colnames(df)) {
                                df[[hba1c_col]] <- cut(as.numeric(as.character(df[[hba1c_col]])), breaks = c(-Inf, 5.7, Inf), right = FALSE, labels = c("Normal", "Prediabetes"))
                                df[[hba1c_col]] <- relevel(df[[hba1c_col]], ref = "Normal")
                                binned_vars <- c(binned_vars, hba1c_col)
                              }
                            }
                            
                            # 低密度脂蛋白分箱 / LDL-C binning
                            ldl_cols <- grep("LDL_Pre", colnames(df), value = TRUE)
                            for(ldl_col in ldl_cols) {
                              if(ldl_col %in% colnames(df)) {
                                df[[ldl_col]] <- cut(as.numeric(as.character(df[[ldl_col]])), breaks = c(-Inf, 100, 130, Inf), right = FALSE, labels = c("Ideal", "Borderline", "High"))
                                df[[ldl_col]] <- relevel(df[[ldl_col]], ref = "Ideal")
                                binned_vars <- c(binned_vars, ldl_col)
                              }
                            }
                            
                            # 高密度脂蛋白依据性别特异性分箱 / HDL-C binning using sex-specific cutoffs
                            hdl_cols <- grep("HDL_Pre", colnames(df), value = TRUE)
                            if(length(hdl_cols) > 0 && "V_Gender" %in% colnames(df)) {
                              gender <- as.numeric(as.character(df[["V_Gender"]]))
                              for(hdl_col in hdl_cols) {
                                if(hdl_col %in% colnames(df)) {
                                  val <- as.numeric(as.character(df[[hdl_col]]))
                                  hdl_factor <- rep(NA, length(val))
                                  hdl_factor[gender == 0 & val >= 40] <- "LowRisk"
                                  hdl_factor[gender == 0 & val < 40] <- "Abnormal"
                                  hdl_factor[gender == 1 & val >= 50] <- "LowRisk"
                                  hdl_factor[gender == 1 & val < 50] <- "Abnormal"
                                  df[[hdl_col]] <- relevel(as.factor(hdl_factor), ref = "LowRisk")
                                  binned_vars <- c(binned_vars, hdl_col)
                                }
                              }
                            }
                            
                            # 共病分数分箱 / Comorbidity score binning
                            score_vars <- paste0(paste0("L", 1:4), "_disease_scores_", disease, "_event")
                            if(score_vars[1] %in% colnames(df)) {
                              df[[score_vars[1]]] <- cut(as.numeric(as.character(df[[score_vars[1]]])), breaks=c(-Inf, 0, Inf), labels=c("Score0", "Score1Plus"))
                            }
                            for(col in score_vars[2:4]) {
                              if(col %in% colnames(df)) df[[col]] <- cut(as.numeric(as.character(df[[col]])), breaks=c(-Inf, 0, 1, Inf), labels=c("Score0", "Score1", "Score2Plus"))
                            }
                            
                            # 分类变量因子化与提取 / Factorize and extract categorical variables
                            cat_vars_base <- c('V_Gender', 'V_Edu',
                                               'L1_Martial_Status', 'L2_Martial_Status', 'L3_Martial_Status', 'L4_Martial_Status',
                                               'L1_Smoke', 'L2_Smoke', 'L3_Smoke', 'L4_Smoke',
                                               'L1_Drink', 'L2_Drink', 'L3_Drink', 'L4_Drink',
                                               'L1_Household_SES', 'L2_Household_SES', 'L3_Household_SES', 'L4_Household_SES',
                                               'L1_Work_status', 'L2_Work_status', 'L3_Work_status', 'L4_Work_status',
                                               'L1_Living_with_Child', 'L2_Living_with_Child', 'L3_Living_with_Child', 'L4_Living_with_Child')
                            
                            vars_to_factor <- cat_vars_base[cat_vars_base %in% colnames(df)]
                            df[vars_to_factor] <- lapply(df[vars_to_factor], as.factor)
                            
# 变量 Dummy 化，包括主分析中的 LTPA 因子变量 / Dummy encoding including factorized LTPA variables used in the main analysis
                            
                            all_factors <- c(vars_to_factor, score_vars, binned_vars,
                                             paste0("A", 1:4, "_OPA_Group"),
                                             paste0("A", 1:4, "_LTPA_Group"),
                                             paste0("A", 1:4, "_LTPA_Group_Factor"))
                            
                            all_factors <- all_factors[all_factors %in% colnames(df)]
                            df <- dummy_cols(df, select_columns = all_factors, remove_first_dummy = TRUE, remove_selected_columns = TRUE, ignore_na = TRUE)
                            
                            # ==============================================================================
                            # 近零方差检查与变量移除 / Near-Zero Variance Check and Variable Pruning
                            # ==============================================================================
                            y_nodes <- paste0("Y", 1:4, "_Event_", disease)
                            all_cols <- colnames(df)
                            pa_related_cols <- all_cols[str_detect(all_cols, "OPA_Group|LTPA_Group")]
                            
                            check_cols <- setdiff(all_cols, c(paste0("C", 1:4), y_nodes, pa_related_cols))
                            check_cols <- check_cols[check_cols %in% colnames(df)]
                            nzv_metrics <- nearZeroVar(df[, check_cols], saveMetrics = TRUE)
                            problematic_vars <- rownames(nzv_metrics)[nzv_metrics$nzv == TRUE]
                            
                            # 执行近零方差变量切除 / Remove near-zero variance variables
                            if(length(problematic_vars) > 0) df <- df[, !colnames(df) %in% problematic_vars]
                            
                            # 结局不可逆推导 / Carry-forward logic for irreversible outcomes
                            for(col in c(paste0("C", 1:4), y_nodes)) {
                              if(col %in% colnames(df)) df[[col]] <- as.numeric(as.character(df[[col]]))
                            }
                            
                            if(all(y_nodes[1:2] %in% colnames(df))) df[[y_nodes[2]]] <- ifelse(df[[y_nodes[1]]] == 1, 1, df[[y_nodes[2]]])
                            if(all(y_nodes[2:3] %in% colnames(df))) df[[y_nodes[3]]] <- ifelse(df[[y_nodes[2]]] == 1, 1, df[[y_nodes[3]]])
                            if(all(y_nodes[3:4] %in% colnames(df))) df[[y_nodes[4]]] <- ifelse(df[[y_nodes[3]]] == 1, 1, df[[y_nodes[4]]])
                            
                            all_cols <- colnames(df)
                            v_nodes  <- all_cols[str_detect(all_cols, "^V_")]
                            
                            extract_l_base <- function(wave) { all_cols[str_detect(all_cols, paste0("^L", wave, "_"))] }
                            n_obs <- nrow(df)
                            
                            configs <- list(
                              list(name = "OPA_T1", target = "OPA", thresh = 1),
                              list(name = "OPA_T2", target = "OPA", thresh = 2),
                              list(name = "OPA_T3", target = "OPA", thresh = 3),
                              list(name = "OPA_T4", target = "OPA", thresh = 4)
                            )
                            
                            node_results <- list()
                            
                            for(cfg in configs) {
                              
                              # ============================================================================
                              # [阶段 1] 主分析构建 / Stage 1: Main Analysis Pipeline
                              # ============================================================================
                              # OPA 退出范式：0000, 1000, 1100, 1110, 1111 / OPA withdrawal regimes: 0000, 1000, 1100, 1110, 1111
                              regimes_array_main <- array(dim = c(n_obs, 4, 5))
                              regimes_array_main[, , 1] <- matrix(rep(c(0, 0, 0, 0), each = n_obs), nrow = n_obs)
                              regimes_array_main[, , 2] <- matrix(rep(c(1, 0, 0, 0), each = n_obs), nrow = n_obs)
                              regimes_array_main[, , 3] <- matrix(rep(c(1, 1, 0, 0), each = n_obs), nrow = n_obs)
                              regimes_array_main[, , 4] <- matrix(rep(c(1, 1, 1, 0), each = n_obs), nrow = n_obs)
                              regimes_array_main[, , 5] <- matrix(rep(c(1, 1, 1, 1), each = n_obs), nrow = n_obs)
                              
                              # 构建 MSM summary measures / Construct MSM summary measures
                              sum_measures <- array(dim = c(5, 4, 1))
                              dimnames(sum_measures)[[2]] <- c("R_1000", "R_1100", "R_1110", "R_1111")
                              sum_measures[1, , 1] <- c(0, 0, 0, 0)
                              sum_measures[2, , 1] <- c(1, 0, 0, 0)
                              sum_measures[3, , 1] <- c(0, 1, 0, 0)
                              sum_measures[4, , 1] <- c(0, 0, 1, 0)
                              sum_measures[5, , 1] <- c(0, 0, 0, 1)
                              
                              msm_formula <- "Y ~ R_1000 + R_1100 + R_1110 + R_1111"
                              
                              # 动态组装主分析列序列，顺序为 Lw, Aw, Cw, Yw / Dynamically assemble main-analysis columns in the order Lw, Aw, Cw, Yw
                              ordered_cols_main <- c(v_nodes)
                              for(w in 1:4) {
                                l_base_w <- extract_l_base(w)
                                y_w <- y_nodes[w]
                                c_w <- paste0("C", w)
                                
                                # 保留二分类 LTPA 作为同波混杂控制变量 / Retain binary LTPA as a same-wave confounding control variable
                                a_conf_w <- sort(
                                  grep(
                                    paste0("^A", w, "_LTPA_Group_Factor_"),
                                    colnames(df),
                                    value = TRUE
                                  )
                                )
                                a_trt_w <- paste0("A", w, "_OPA_Group_T", cfg$thresh)
                                
                                ordered_cols_main <- c(ordered_cols_main, l_base_w, a_trt_w, a_conf_w, c_w, y_w)
                              }
                              
                              ordered_cols_main <- ordered_cols_main[ordered_cols_main %in% colnames(df)]
                              data_main <- df[, ordered_cols_main]
                              
                              Anodes_main <- paste0("A", 1:4, "_OPA_Group_T", cfg$thresh)
                              Anodes_main <- Anodes_main[Anodes_main %in% colnames(data_main)]
                              
                              # 检查主分析 Anodes 完整性 / Check completeness of main-analysis Anodes
                              if(length(Anodes_main) != 4) stop("Main Anodes are incomplete.")
                              
                              Ynodes <- y_nodes[y_nodes %in% colnames(data_main)]
                              # ==== 已修正 Cnodes：仅保留 C1, C2, C3 ====
                              Cnodes <- c("C1", "C2", "C3")[c("C1", "C2", "C3") %in% colnames(data_main)]
                              
                              Lnodes_main <- setdiff(colnames(data_main), c(v_nodes, Anodes_main, Cnodes, Ynodes))
                              
                              # 运行主分析模型获取总体平均效应 / Run main model to estimate overall average effects
                              fit_msm_main <- suppressWarnings(ltmleMSM(data = data_main, Anodes = Anodes_main, Cnodes = Cnodes,
                                                                        Ynodes = Ynodes, Lnodes = Lnodes_main, survivalOutcome = TRUE,
                                                                        regimes = regimes_array_main,
                                                                        summary.measures = sum_measures,
                                                                        working.msm = msm_formula,
                                                                        final.Ynodes = tail(Ynodes, 1), SL.library = SLlib_ML,
                                         variance.method = "tmle", gbounds = c(0.05, 0.95), SL.cvControl = list(V = 5)))
                              
                              q_res_tmp <- matrix(NA, nrow=1, ncol=length(cols_to_track)); colnames(q_res_tmp) <- cols_to_track
                              se_res_tmp <- matrix(NA, nrow=1, ncol=length(cols_to_track)); colnames(se_res_tmp) <- cols_to_track
                              
                              res_extracted <- extract_msm_results(fit_msm_main, q_res_tmp, se_res_tmp, 1, n_obs)
                              node_results[[cfg$name]] <- list(q = res_extracted$q, se = res_extracted$se)
                              
                              
                            }
                            
                            # 内存回收 / Garbage collection
                            gc()
                            
                            # 保存被切除的近零方差变量名单 / Save names of removed near-zero variance variables
                            node_results$dropped <- problematic_vars
                            return(node_results)
                          }
  
  # ==============================================================================
  # 3.2 提取结果与整理 / Extract Results and Log NZV Variables
  # ==============================================================================
  valid_results <- Filter(function(x) !is.null(x) && !inherits(x, "error"), results_list)
  
  if(length(valid_results) == 0) {
    message("❌ 致命错误：所有数据集运行失败。具体报错信息如下：/ FATAL ERROR: All datasets failed. Error msg:")
    print(results_list[[1]])
    stop("程序已终止。 / Program terminated.")
  }
  
  all_dropped <- unique(unlist(lapply(valid_results, `[[`, "dropped")))
  
  if(length(all_dropped) > 0) {
    cat(sprintf("\n>>> [%s 预处理汇总] 本次 LTMLE 分析共自动切除了以下稀疏/近零方差变量：\n", disease))
    cat(sprintf(">>> [%s Preprocessing Summary] The following sparse/NZV variables were automatically dropped:\n", disease))
    cat("    ", paste(all_dropped, collapse = ", "), "\n")
  } else {
    message(sprintf("\n>>> [%s 预处理汇总] 所有数据集质量良好，未触发稀疏变量切除。 / [%s Preprocessing Summary] All datasets are of good quality.", disease, disease))
  }
  
  # ==============================================================================
  # 4. Rubin's Rules 合并与结果写入 Excel / Combine Using Rubin's Rules and Write to Excel
  # ==============================================================================
  message(sprintf("\n>>> 正在应用 Rubin's Rules 汇总并组装 [%s] 的结果矩阵... / Assembling result matrices for [%s]...", disease, disease))
  
  # 按结果名称逐个过滤有效插补结果 / Filter valid results (更新以包含 RD 的 SE 验证)
  get_valid_by_name <- function(valid_results, raw_name) {
    required_q_cols <- cols_to_track
    # 验证 SE 是否完整包含 RR 和 RD
    required_se_cols <- c("Risk_R1", "logRR_R2", "logRR_R3", "logRR_R4", "logRR_R5", 
                          "RD_R2", "RD_R3", "RD_R4", "RD_R5") 
    
    Filter(function(x) {
      !is.null(x[[raw_name]]) &&
        !is.null(x[[raw_name]]$q) &&
        !is.null(x[[raw_name]]$se) &&
        all(required_q_cols %in% colnames(x[[raw_name]]$q)) &&
        all(required_se_cols %in% colnames(x[[raw_name]]$se)) &&
        all(is.finite(x[[raw_name]]$q[, required_q_cols])) &&
        all(is.finite(x[[raw_name]]$se[, required_se_cols]))
    }, valid_results)
  }
  
  # ------------------------------------------------------------------------------
  # 统一表格生成函数，加入 RD 指标 / Unified table generator with RD
  # ------------------------------------------------------------------------------
  generate_table_and_write <- function(q_res, se_res, exp_name, sheet_target, current_row_ref) {
    pooled <- mi.meld(q = q_res, se = se_res)
    fq <- pooled$q.mi[1, ]
    fse <- pooled$se.mi[1, ]
    
    # logRR 格式化
    format_ci_p_log <- function(est, se) {
      if (is.na(se) || se == 0) return("-")
      l <- exp(est - 1.96*se)
      h <- exp(est + 1.96*se)
      v <- exp(est)
      p <- 2 * (1 - pnorm(abs(est / se)))
      str_p <- ifelse(p < 0.001, "<0.001", sprintf("%.3f", p))
      return(sprintf("%.2f [%.2f, %.2f], P=%s", v, l, h, str_p))
    }
    
    # 风险差 (RD) 格式化 (乘以 100 转为百分比直观展示)
    format_ci_p_rd <- function(est, se) {
      if (is.na(se) || se == 0) return("-")
      l <- est - 1.96*se
      h <- est + 1.96*se
      p <- 2 * (1 - pnorm(abs(est / se)))
      str_p <- ifelse(p < 0.001, "<0.001", sprintf("%.3f", p))
      return(sprintf("%.2f%% [%.2f%%, %.2f%%], P=%s", est*100, l*100, h*100, str_p))
    }
    
    format_rate <- function(est) { sprintf("%.1f%%", est * 100) }
    
    # 构建输出表格 (增加了一列 RD)
    df_tab <- data.frame(
      `OPA 暴露轨迹 / OPA Trajectory` = c("0000: Always Low Exposure / 始终轻体力 (Ref)",
                                      "1000: Withdrawal at Wave 2 / 第2波起脱离重体力",
                                      "1100: Withdrawal at Wave 3 / 第3波起脱离重体力",
                                      "1110: Withdrawal at Wave 4 / 第4波起脱离重体力",
                                      "1111: Sustained High Exposure / 持续高暴露"),
      `因果相对风险 / Causal Effect RR (95% CI, P)` = c(
        paste0("1.00 (Ref)\n(Risk: ", format_rate(fq["Risk_R1"]), ")"),
        format_ci_p_log(fq["logRR_R2"], fse["logRR_R2"]),
        format_ci_p_log(fq["logRR_R3"], fse["logRR_R3"]),
        format_ci_p_log(fq["logRR_R4"], fse["logRR_R4"]),
        format_ci_p_log(fq["logRR_R5"], fse["logRR_R5"])
      ),
      `因果风险差 / Causal Risk Difference (95% CI, P)` = c(
        "0.00% (Ref)",
        format_ci_p_rd(fq["RD_R2"], fse["RD_R2"]),
        format_ci_p_rd(fq["RD_R3"], fse["RD_R3"]),
        format_ci_p_rd(fq["RD_R4"], fse["RD_R4"]),
        format_ci_p_rd(fq["RD_R5"], fse["RD_R5"])
      ),
      check.names = FALSE
    )
    
    # 写入标题并设置样式 / Write title and apply style
    title_str <- sprintf("=== [%s] %s ===", disease, exp_name)
    writeData(final_wb, sheet = sheet_target, x = title_str, startRow = current_row_ref, startCol = 1)
    title_style <- createStyle(textDecoration = "bold", fontColour = "blue")
    addStyle(final_wb, sheet = sheet_target, style = title_style, rows = current_row_ref, cols = 1)
    
    # 写入主体与表头 / Write body and header
    header_style <- createStyle(textDecoration = "bold", border = "Bottom")
    writeData(final_wb, sheet = sheet_target, x = df_tab, startRow = current_row_ref + 1, startCol = 1, headerStyle = header_style)
    
    # 返回下一段输出的起始行 / Return next starting row
    return(current_row_ref + nrow(df_tab) + 4)
  }
  
  # ==============================================================================
  # OPA 主分析输出，输出 T1-T4 / Output Main OPA Effects for T1-T4
  # ==============================================================================
  exp_names_main <- list(
    OPA_T3 = "Primary Analysis: : OPA MVPA >= 150 min/week (Levels 3-4 vs. 0-2)",
    OPA_T1 = "Sensitivity Analysis 1: Any OPA MVPA (MVPA > 0) vs. none (Level 0)",
    OPA_T2 = "Sensitivity Analysis 2: OPA MVPA >= 75 min/week (Levels 2-4 vs. 0-1)",
    OPA_T4 = "Sensitivity Analysis 3: OPA MVPA >= 600 min/week (Level 4 vs. 0-3)"
  )
  
  for(raw_name in names(exp_names_main)) {
    valid_sub <- get_valid_by_name(valid_results, raw_name)
    
    if(length(valid_sub) == 0) {
      message(sprintf("⚠️ [%s - %s] 无有效插补结果，跳过。 / No valid imputed result; skipped.", disease, raw_name))
      next
    }
    
    message(sprintf(">>> [%s - %s] 用于 Rubin 合并的有效插补数据集数：%d / Number of valid imputed datasets used for Rubin pooling: %d",
                    disease, raw_name, length(valid_sub), length(valid_sub)))
    
    q_res  <- do.call(rbind, lapply(valid_sub, function(x) x[[raw_name]]$q))
    se_res <- do.call(rbind, lapply(valid_sub, function(x) x[[raw_name]]$se))
    
    current_excel_row <- generate_table_and_write(q_res, se_res, exp_names_main[[raw_name]], disease, current_excel_row)
  }
  
  # 为主表添加脚注 / Add footnotes to main sheet
  footnote_main <- data.frame(X1 = c(
    "表注 (Footnotes):",
    "此表展示 OPA 的总体平均主效应；LTPA 已二分类（0=无 LTPA，1=有任何 LTPA，原始 1/2/3 合并）并作为同波时变混杂变量进行调整，未设定 LTPA 反事实轨迹。 / This table presents the overall average main effect of OPA; LTPA is dichotomized (0 = no LTPA, 1 = any LTPA, original 1/2/3 combined) and adjusted as a same-wave time-varying confounder, without setting counterfactual LTPA trajectories."
  ))
  writeData(final_wb, sheet = disease, x = footnote_main, startRow = current_excel_row, startCol = 1, colNames = FALSE)
  # 额外注明本次敏感性分析所剔除的变量 / Note excluded covariates for this SA
  sa_note <- data.frame(X1 = sprintf(
    "Sensitivity analysis: excluded covariates (n=%d): %s",
    length(exclude_vars), paste(exclude_vars, collapse = ", ")
  ))
  writeData(final_wb, sheet = disease, x = sa_note, startRow = current_excel_row + 2, startCol = 1, colNames = FALSE)
  
  
}

  # 保存该情景结果 / Save scenario results
  final_excel_path <- file.path(result_dir, paste0("Main_Pooled_MSM_All_Results_group_SA_", scenario_name, ".xlsx"))
  saveWorkbook(final_wb, final_excel_path, overwrite = TRUE)

  message("\n>>> 敏感性分析完成并保存 / Sensitivity analysis saved: ", scenario_label)
  message(">>> ", final_excel_path)

  invisible(final_excel_path)
}

# ==============================================================================
# 4. 依次执行全部敏感性分析 / Run all SA scenarios sequentially
# ==============================================================================
for (s_idx in seq_along(sa_scenarios)) {
  run_sa_scenario(sa_scenarios[[s_idx]]$name,
                  sa_scenarios[[s_idx]]$label,
                  sa_scenarios[[s_idx]]$exclude,
                  s_idx)
}

# 关闭并行集群 / Stop parallel cluster
stopCluster(cl)

message("\n>>> 全部敏感性分析已完成。 / All sensitivity analyses completed.")
