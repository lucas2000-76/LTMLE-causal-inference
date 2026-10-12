# LTMLE Causal Inference

Analysis code for a longitudinal study using the China Health and Retirement Longitudinal Study (CHARLS). The workflow extracts and harmonises wave-specific variables, prepares longitudinal datasets for targeted maximum likelihood estimation (LTMLE), and performs missingness analysis, multiple imputation, causal-effect estimation, and E-value sensitivity analyses.

> **Note:** This repository contains analytical code only. The underlying CHARLS data are not included because they are available only under the CHARLS data-access agreement.

## Repository structure

```
.
├── data_extraction/   # Python: wave-specific data extraction and covariate construction (2013/2015/2018/2020)
├── ltmle/             # Python: LTMLE dataset preparation, missingness analysis, and imputation
├── r_analysis/        # R: causal-effect estimation and E-value sensitivity analyses
├── LICENSE
├── .gitignore
└── README.md
```

---

## Python — data preparation

The `data_extraction/` folder contains one pipeline per wave (2013, 2015, 2018, 2020). Each wave has scripts that:

- **Exposure & comorbidity scores** — compute cardiovascular disease (CVD)/diabetes outcomes and comorbidity scores, and classify physical activity (OPA/LTPA) groups.
- **Cognition & depression scale** — compute cognitive scores and CES-D-10 depression scores.
- **Health & demographic covariates** — compute smoking, drinking, sleep, marital status, and other covariates.
- **SES / work / family living status** — compute BMI, work status, family living arrangements, and household socioeconomic status (SES).
- **Blood** — compute blood biomarkers (e.g., TyG index, HDL).

All extraction scripts read the translated CHARLS wave files and update a central `DATA.xlsx` dataset.

### LTMLE preparation scripts

| File | Purpose |
|------|---------|
| `LTMLE Data Finalized 1.py` | Unified pipeline for LTMLE preparation: reads and pads IDs, generates censoring nodes (C1–C3), fixes intermittent missingness, corrects age, cleans outcomes, and splits disease-specific incident cohorts. |
| `LTMLE Data Finalized 2 .py` | Exploratory data analysis and missingness analysis for the Base, CVD, and Diabetes cohorts; computes missing rates and generates UpSet plots while preserving original data scales. |
| `LTMLE Data Finalized 3 .py` | Multiple imputation by chained equations (MICE) with a mixed random-forest engine for the CVD and Diabetes cohorts, enforcing absorbing states and LTMLE monotone-missingness rules. |

---

## R — causal inference analyses

The `r_analysis/` folder contains the R scripts that estimate causal effects from the datasets prepared by the Python `ltmle/` pipeline.

| File | Purpose |
|------|---------|
| `TE_and_ATE_category_new.R` | Main causal-effect analysis. Uses LTMLE/MSM with a SuperLearner ensemble and parallel `foreach` computing to estimate the pooled causal effect of occupational physical activity (OPA) trajectories on CVD and diabetes, under a withdrawal paradigm (`0000/1000/1100/1110/1111`). |
| `TE_and_ATE_category_new_SA.R` | Main analysis plus post-hoc sensitivity analyses (objective biomarkers, socioeconomic, and mental health & cognitive scenarios). |
| `EValue_analysis_new.R` | Computes E-values for unmeasured confounding from the pooled LTMLE/MSM results (main + sensitivity scenarios), following VanderWeele & Ding (2017). |

### R dependencies

`ltmle`, `SuperLearner`, `fastDummies`, `dplyr`, `stringr`, `readr`, `openxlsx`, `parallel`, `doParallel`, `foreach`, `doRNG`, `caret`, `gtools`, `ranger`, `gam`, `Amelia` (E-value script additionally requires `openxlsx`, `stringr`, `dplyr`, `readr`).

---

## Data availability

The raw CHARLS data are not distributed in this repository. Access to CHARLS data can be requested at <https://charls.charlsdata.com/>.

## License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file.
