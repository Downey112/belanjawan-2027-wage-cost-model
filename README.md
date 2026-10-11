# Belanjawan 2027 Fresh Graduate Wage Cost Model 🇲🇾📊

[![Excel](https://img.shields.io/badge/Model-Microsoft_Excel-217346?style=flat&logo=microsoftexcel&logoColor=white)](https://office.com)
[![Python Generator](https://img.shields.io/badge/Generator-openpyxl-3776AB?style=flat&logo=python&logoColor=white)](https://openpyxl.readthedocs.io/)
[![Policy Benchmark](https://img.shields.io/badge/Policy-Budget_2027_Benchmarks-C0392B?style=flat)](https://budget.mof.gov.my)

A dynamic, formula-driven Excel model calculating fully loaded employer costs and net employee take-home pay across the three **Belanjawan 2027 wage benchmarks**: the national minimum wage, the graduate wage floor, and the GLC/GLIC living wage.

---

## 📌 Benchmark Comparison (Monthly RM)

| Metric | Minimum Wage (Jun 2027) | Fresh Graduate Floor | GLC / GLIC Living Wage |
|---|:---:|:---:|:---:|
| **Gross Wage** | **RM 2,000.00** | **RM 2,500.00** | **RM 3,400.00** |
| Employer Statutory On-Costs (~15.95%) | RM 319.00 | RM 398.75 | RM 542.30 |
| **Total Employer Fully Loaded Cost** | **RM 2,319.00** | **RM 2,898.75** | **RM 3,942.30** |
| Employee Statutory Deductions (11.7%) | RM 234.00 | RM 292.50 | RM 397.80 |
| **Net Employee Take-Home Pay** | **RM 1,766.00** | **RM 2,207.50** | **RM 3,002.20** |
| Monthly Tax Deduction (PCB / MTD) | RM 0.00 | RM 0.00 | RM 0.00 |

*Employer on-costs model statutory employer EPF (12–13%), SOCSO, EIS, and the HRD Corp training levy.*

---

## 📁 Repository Files & Structure

* **`Belanjawan2027_Wage_Cost_Model.xlsx`** — Master formula-driven workbook.
* **`build_model.py`** — Python generator script that compiles the complete workbook programmatically using `openpyxl`.

---

## 📑 Workbook Sheets & Features

1. **`Overview`** — Executive headline figures, color-coded input legends, and model version notes.
2. **`Assumptions`** — Centralized rate parameter tables (EPF, SOCSO caps, EIS rates, tax bands, personal reliefs). Modifying any cell dynamically ripples through the entire workbook.
3. **`Benchmark Comparison`** — Side-by-side pay slip breakdowns, statutory employer contributions, and net take-home calculators.
4. **`Headcount Budget`** — Scenario planner modeling annual fiscal exposure for corporate hiring plans transitioning headcount to the new wage floors.
5. **`Wage Ladder`** — Dynamic salary progression ladder from RM2,000 to RM6,000 with interactive visualization charts.

---

## ⚙️ How to Rebuild the Model

```bash
# 1. Clone repository
git clone https://github.com/Downey112/belanjawan-2027-wage-cost-model.git
cd belanjawan-2027-wage-cost-model

# 2. Install dependencies & compile
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
pip install openpyxl

# 3. Generate the Excel workbook
python build_model.py Belanjawan2027_Wage_Cost_Model.xlsx
```

---

## 🔍 Assumptions & Methodology

- Wage benchmarks reflect the Budget 2027 policy announcements.
- Statutory contributions model flat baseline statutory percentages (EPF, SOCSO, EIS, HRDF); minor variations may exist under specific gazetted band schedules.
- Net take-home pay models standard individual tax reliefs and basic MTD (PCB) exemption thresholds.
- Designed as an analytical policy planning tool.
