# Applied Machine Learning: Predicting Term-Deposit Subscription

Udacity / Woolf Master's in AI – Capstone Project 3: *Applied Machine Learning: Model Design, Training, and Performance Evaluation*

**Author:** Mohammad Afsharnejad ([@Sirvanmesh](https://github.com/Sirvanmesh))

## Problem

A Portuguese bank sold term deposits through phone campaigns, and most calls ended in "no". This project builds a model that estimates, **before a call is made**, how likely a client is to subscribe, so the call list can be prioritised.

- **Problem type:** Supervised learning – binary classification (target `y`: subscribed yes/no)
- **Main model:** class-weighted random forest (tuned with 5-fold cross-validation on PR-AUC), compared with a majority-class dummy and a logistic regression baseline
- **Key test-set results:** PR-AUC 0.454 (baseline 0.117), ROC-AUC 0.804, precision 0.432 and recall 0.553 at the tuned threshold

## Dataset

- **Name:** Bank Marketing (`bank-full.csv`: 45,211 rows, 16 features plus target)
- **Source:** UCI Machine Learning Repository – <https://archive.ics.uci.edu/dataset/222/bank+marketing>
- **License:** CC BY 4.0
- **Citation (APA 7):** Moro, S., Rita, P., & Cortez, P. (2014). *Bank marketing* [Dataset]. UCI Machine Learning Repository. https://doi.org/10.24432/C5K306

The call-duration column is dropped because it is only known after the call (target leakage).

## Setup

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS / Linux
source .venv/bin/activate

pip install -r requirements.txt
```

## Run

- Open and run `modeling.ipynb` (or execute it headless: `jupyter nbconvert --to notebook --execute --inplace modeling.ipynb`). It writes the figures and `metrics.json`.
- `python build_report.py` rebuilds the report from `metrics.json`. It converts the report to PDF with Microsoft Word (docx2pdf), falling back to LibreOffice.

## Outputs

| File | Contents |
|---|---|
| `module_summary.pdf` | ML analysis report (APA 7) |
| `modeling.ipynb` | Full workflow with saved outputs |
| `figures/` | Class balance, confusion matrices, ROC/PR curves, feature importance, age-group check |
| `metrics.json` | All final metrics used in the report |
| `bank-full.csv` | Dataset |
| `requirements.txt` | Pinned environment (`pip freeze`) |
