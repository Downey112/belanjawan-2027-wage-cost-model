# Belanjawan 2027 Fresh Graduate Wage Cost Model

An Excel model of what a fresh graduate hire costs a Malaysian employer, and what the
employee takes home, at the three Belanjawan 2027 wage benchmarks.

| Per month (RM) | Minimum wage | Fresh graduate floor | GLC / GLIC living wage |
|---|---|---|---|
| Gross wage | 2,000.00 | 2,500.00 | 3,400.00 |
| Employer on-costs (15.95%) | 319.00 | 398.75 | 542.30 |
| Fully loaded cost | 2,319.00 | 2,898.75 | 3,942.30 |
| Net take-home pay | 1,766.00 | 2,207.50 | 3,002.20 |
| PCB / MTD | 0 | 0 | 0 |

## Files

- `Belanjawan2027_Wage_Cost_Model.xlsx` - the model. Open it in Excel.
- `build_model.py` - the script that generates the workbook.

## Sheets

- **Overview** - headline figures, sheet guide and colour legend.
- **Assumptions** - every rate, cap, relief and tax band, each with its source. Change a blue value and the whole model updates.
- **Benchmark Comparison** - side-by-side payslip and employer cost for each benchmark, plus a custom wage.
- **Headcount Budget** - annual cost of a hiring plan and of moving staff up a benchmark. The headcounts are example values.
- **Wage Ladder** - cost and take-home from RM2,000 to RM6,000, with a chart.

## Rebuild

```
pip install openpyxl
python build_model.py Belanjawan2027_Wage_Cost_Model.xlsx
```

## Assumptions and limits

- Wage benchmarks and statutory rates come from a Belanjawan 2027 brief and have not been
  checked against gazetted schedules. Rows shaded orange on the Assumptions sheet are the
  ones to verify first (SOCSO caps, the RM12,000 individual relief, and the tax bands,
  EPF relief cap and rebate limit, which use current statutory values).
- Contributions use flat percentages rather than the official banded contribution tables,
  so real payslips can differ by a few ringgit.
- PCB is annual tax divided by 12 for a single resident with no other income or reliefs.
- Illustrative only; not payroll or tax advice.
