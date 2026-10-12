# LTMLE Causal Inference

Analysis code for a longitudinal study using the China Health and Retirement Longitudinal Study (CHARLS). The workflow extracts and harmonises wave-specific variables, prepares longitudinal datasets for targeted maximum likelihood estimation (LTMLE), and performs missingness analysis and multiple imputation.

> **Note:** This repository contains analytical code only. The underlying CHARLS data are not included because they are available only under the CHARLS data-access agreement.

## Repository structure

```
.
├── data_extraction/   # Wave-specific data extraction and covariate construction (2013/2015/2018/2020)
├── ltmle/             # LTMLE dataset preparation, missingness analysis, and imputation
└── README.md
```

## Data extraction scripts

The `data_extraction/` folder contains one pipeline per wave (2013, 2015, 2018, 2020). Each wave has scripts that:

- **Exposure & comorbidity scores** — compute cardiovascular disease (CVD)/diabetes outcomes and comorbidity scores, and classify physical activity (OPA/LTPA) groups.
- **Cognition & depression scale** — compute cognitive scores and CES-D-10 depression scores.
- **Health & demographic covariates** — compute smoking, drinking, sleep, marital status, and other covariates.
- **SES / work / family living status** — compute BMI, work status, family living arrangements, and household socioeconomic status (SES).
- **Blood** — compute blood biomarkers (e.g., TyG index, HDL).

All extraction scripts read the translated CHARLS wave files and update a central `DATA.xlsx` dataset.

## LTMLE preparation scripts

| File | Purpose |
|------|---------|
| `LTMLE Data Finalized 1.py` | Unified pipeline for LTMLE preparation: reads and pads IDs, generates censoring nodes (C1–C3), fixes intermittent missingness, corrects age, cleans outcomes, and splits disease-specific incident cohorts. |
| `LTMLE Data Finalized 2 .py` | Exploratory data analysis and missingness analysis for the Base, CVD, and Diabetes cohorts; computes missing rates and generates UpSet plots while preserving original data scales. |
| `LTMLE Data Finalized 3 .py` | Multiple imputation by chained equations (MICE) with a mixed random-forest engine for the CVD and Diabetes cohorts, enforcing absorbing states and LTMLE monotone-missingness rules. |

## Data availability

The raw CHARLS data are not distributed in this repository. Access to CHARLS data can be requested at <https://charls.charlsdata.com/>.

## Dependencies

- Python 3
- pandas, numpy, openpyxl, matplotlib
- scikit-learn (imputation)
- upsetplot (missingness plots)
