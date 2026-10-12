# ==============================================================================
# Script Name: EValue_analysis_new.R
# Description:
#     E-value 敏感性分析 / E-value analyses for unmeasured confounding
#     针对主分析 (Main) 及 3 个 post hoc 敏感性分析（objective biomarkers /
#     socioeconomic / mental health & cognitive）的合并 LTMLE/MSM 结果，
#     计算每个 OPA 暴露轨迹因果 RR 的 E-value（点估计 + 置信区间边界）。
#
#     E-value = 未测混杂因素为完全解释观察到的关联，所需同时与暴露和结局
#     关联的最弱强度（RR 尺度）。参考 VanderWeele & Ding (2017), Ann Intern Med.
#
#     输入（C:/Users/26627/Desktop/result/）：
#       Main_Pooled_MSM_All_Results_group_Lucas.xlsx
#       Main_Pooled_MSM_All_Results_group_SA_objective_biomarkers.xlsx
#       Main_Pooled_MSM_All_Results_group_SA_socioeconomic.xlsx
#       Main_Pooled_MSM_All_Results_group_SA_mental_health_cognitive.xlsx
#       （不存在的文件会自动跳过并提示）
#     输出（同目录）：
#       EValue_Main_and_SA_Pooled_MSM_All_Results_group_new.xlsx
#       EValue_<情景>_Pooled_MSM_All_Results_group_new.csv
# ==============================================================================

# 1. 环境初始化 / Setup
rm(list = ls()); gc()

libs <- c("openxlsx", "stringr", "dplyr", "readr")
missing_libs <- !libs %in% installed.packages()
if (any(missing_libs)) install.packages(libs[missing_libs])
invisible(lapply(libs, library, character.only = TRUE))

result_dir <- "C:/Users/26627/Desktop/result"

# ==============================================================================
# 2. E-value 函数（与 EValue::evalues.RR 等价，无需额外依赖）
# ==============================================================================
# est/lo/hi: 风险比点估计与 95% CI 边界（RR 尺度）
# 公式：RR >= 1 时 E = RR + sqrt(RR*(RR-1))；RR < 1 时对 1/RR 计算。
#       CI 边界取最靠近无效值 1 的一侧；若 CI 包含 1，则 CI 边界 E-value = 1。
evalue_rr <- function(est, lo, hi) {
  if (any(is.na(c(est, lo, hi)))) return(c(point = NA_real_, ci = NA_real_))
  e_val <- function(rr) {
    rr <- max(rr, 1 / rr)
    rr + sqrt(rr * (rr - 1))
  }
  eval_point <- e_val(est)
  if (est >= 1) {
    lim <- lo
    eval_ci <- if (lim > 1) e_val(lim) else 1
  } else {
    lim <- hi
    eval_ci <- if (lim < 1) e_val(lim) else 1
  }
  c(point = eval_point, ci = eval_ci)
}

# ==============================================================================
# 3. 解析结果 Excel（每个 sheet 一种疾病；行结构：标题 / 表头 / 5 个 regime）
# ==============================================================================
parse_workbook <- function(file) {
  wb <- loadWorkbook(file)
  sheets <- names(wb)
  out <- list()
  for (s in sheets) {
    d <- read.xlsx(wb, sheet = s, colNames = FALSE)
    cur_disease  <- NA_character_
    cur_threshold <- NA_character_
    for (r in seq_len(nrow(d))) {
      c1 <- as.character(d[r, 1]); c2 <- as.character(d[r, 2])
      if (is.na(c1)) next
      # 标题行，例如：=== [CVD] Primary Analysis: OPA wMVPA >= 75 min/week ... ===
      m_title <- str_match(c1, "^===\\s*\\[([^]]+)\\]\\s*(.+?)\\s*===$")
      if (!is.na(m_title[1, 1])) {
        cur_disease  <- m_title[1, 2]
        cur_threshold <- m_title[1, 3]
        next
      }
      if (is.na(c2)) next
      # RR 单元格，例如：0.80 [0.73, 0.88], P=<0.001
      m_rr <- str_match(c2, "^([0-9]+\\.[0-9]+)\\s*\\[\\s*([0-9]+\\.[0-9]+)\\s*,\\s*([0-9]+\\.[0-9]+)\\s*\\]")
      if (is.na(m_rr[1, 1])) next
      out[[length(out) + 1]] <- data.frame(
        Disease   = cur_disease,
        Threshold = cur_threshold,
        Regime    = c1,
        RR_est    = as.numeric(m_rr[1, 2]),
        RR_lo     = as.numeric(m_rr[1, 3]),
        RR_hi     = as.numeric(m_rr[1, 4]),
        RR_label  = sprintf("%s [%s, %s]", m_rr[1, 2], m_rr[1, 3], m_rr[1, 4]),
        stringsAsFactors = FALSE
      )
    }
  }
  if (length(out) == 0) return(NULL)
  do.call(rbind, out)
}

# ==============================================================================
# 4. 情景定义：主分析 + 3 个敏感性分析（非累积）
# ==============================================================================
scenarios <- list(
  list(name = "Main",                        file = "Main_Pooled_MSM_All_Results_group_Lucas.xlsx"),
  list(name = "SA_objective_biomarkers",     file = "Main_Pooled_MSM_All_Results_group_SA_objective_biomarkers.xlsx"),
  list(name = "SA_socioeconomic",            file = "Main_Pooled_MSM_All_Results_group_SA_socioeconomic.xlsx"),
  list(name = "SA_mental_health_cognitive",  file = "Main_Pooled_MSM_All_Results_group_SA_mental_health_cognitive.xlsx")
)

final_wb <- createWorkbook()

for (sc in scenarios) {
  file <- file.path(result_dir, sc$file)
  if (!file.exists(file)) {
    message(sprintf("[%s] 结果文件不存在，跳过 / File not found, skipped: %s", sc$name, file))
    next
  }
  message(">>> 正在计算 E-value / Computing E-values: ", sc$name, " -> ", basename(file))

  res <- tryCatch(parse_workbook(file), error = function(e) {
    message("    ⚠️ 读取/解析失败，跳过 / Failed to parse: ", conditionMessage(e))
    NULL
  })
  if (is.null(res) || nrow(res) == 0) {
    message("    ⚠️ 未能从文件中解析出 RR 结果，跳过 / No RR estimates parsed; skipped.")
    next
  }

  # 计算 E-value（点估计 + CI 边界）
  ev <- t(apply(res[, c("RR_est", "RR_lo", "RR_hi")], 1, function(v) evalue_rr(v[1], v[2], v[3])))
  res$Evalue_point <- round(ev[, "point"], 2)
  res$Evalue_ci    <- round(ev[, "ci"], 2)

  tab <- res[, c("Disease", "Threshold", "Regime", "RR_label", "Evalue_point", "Evalue_ci")]
  colnames(tab) <- c("Disease / 疾病",
                     "OPA Threshold / OPA 阈值",
                     "OPA Trajectory / OPA 暴露轨迹",
                     "Observed RR [95% CI] / 观察到的 RR",
                     "Point Est. E-Value / 点估计 E-Value",
                     "CI Limit E-Value / 置信区间边界 E-Value")

  # CSV 输出（UTF-8）
  csv_path <- file.path(result_dir, paste0("EValue_", sc$name, "_Pooled_MSM_All_Results_group_Lucas.csv"))
  write_csv(tab, csv_path)
  message("    CSV : ", csv_path)

  # Excel 输出（追加 sheet）
  addWorksheet(final_wb, sheetName = sc$name)
  writeData(final_wb, sheet = sc$name, x = sprintf("=== E-value Analysis: %s ===", sc$name), startRow = 1, startCol = 1)
  header_style <- createStyle(textDecoration = "bold", border = "Bottom")
  writeData(final_wb, sheet = sc$name, x = tab, startRow = 3, startCol = 1, headerStyle = header_style)
  footnote <- data.frame(X1 = c(
    "E-value: 未测混杂因素为完全解释观察到的 RR（及其 CI 边界）所需的最弱关联强度（RR 尺度）。",
    "E-value 越大，结论对未测混杂越稳健；CI 包含 1 时 CI 边界 E-value 记为 1.00（无法排除无效）。",
    "计算与 EValue::evalues.RR 等价 (VanderWeele & Ding 2017)。"
  ))
  writeData(final_wb, sheet = sc$name, x = footnote, startRow = nrow(tab) + 6, startCol = 1, colNames = FALSE)
}

# ==============================================================================
# 5. 保存汇总 Excel / Save summary workbook
# ==============================================================================
excel_path <- file.path(result_dir, "EValue_Main_and_SA_Pooled_MSM_All_Results_group_new.xlsx")
if (length(names(final_wb)) == 0) {
  message("\n⚠️ 没有可用的结果，未生成 Excel 文件 / No results available; no Excel file written.")
} else {
  saveWorkbook(final_wb, excel_path, overwrite = TRUE)
  message("\n>>> E-value 分析完成 / All E-value analyses saved:")
  message(">>> ", excel_path)
}
