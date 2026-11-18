# Work Time Log Generator

A Python script that builds a fully formula-driven Excel workbook for tracking
daily work: tasks, hours, status and live summaries. Everything in the output
is a real Excel formula, so the workbook keeps recalculating as you edit it.

> All data in this repository is **fictional sample data**.

## Features

- **Time Log sheet**: 500 ready rows, auto Log ID / Day / Week / Month / Hours
- **Midnight-safe hours**: `MOD(End - Start, 1) * 24`, so 10 PM to 1 AM = 3.00 h
- **Dropdowns** for category, status, assigned-by, source, verified (editable on the `Lists` sheet); team field suggests values but accepts free text
- **Validation**: dates on/after the tracking start date, times between 12:00 AM and 11:59 PM
- **Conditional formatting**: yellow for missing start/end, green/amber/red status colours, "To confirm" highlight
- **Summary sheet**: entries, hours and % by category, status, month (12) and week (26)
- **Filters and frozen panes** for easy navigation

## Quick start

```bash
pip install -r requirements.txt
python build_work_time_log.py                      # workbook with sample data
python build_work_time_log.py --no-sample          # blank log
python build_work_time_log.py --owner "Jane Doe" --org "Acme Corp | Ops" -o my_log.xlsx
```

Options: `-o/--output`, `--owner`, `--org`, `--start-date YYYY-MM-DD`, `--no-sample`.

## Workbook layout

| Sheet | Purpose |
|---|---|
| How to Use | Instructions, colour legend, example row |
| Time Log | One row per task (blue = input, black = formula) |
| Summary | Live tables driven by the log; `C7` is the tracking start date |
| Lists | Dropdown values you can edit |

A pre-built example is included: `Work_Time_Log_Sample.xlsx`.

## Tech

Python 3, [openpyxl](https://openpyxl.readthedocs.io/): formulas, data
validation, conditional formatting, styled tables.

## Ideas for extension

- Charts on the Summary sheet
- Multi-person workbook with a combined team summary
- CSV/Google Sheets import and export
- Weekly report PDF generated from the log
