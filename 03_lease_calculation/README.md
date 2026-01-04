# Lease Calculation
租赁测算

Purpose: compute lease amortization schedules from Excel input.
目的：根据 Excel 输入生成租赁摊销测算表。

Quick use:
快速使用：
1) Rename your Excel to `input.xlsx` and place it here.
1) 将 Excel 重命名为 `input.xlsx` 并放到此目录。
2) Double-click `run.bat` (or run `python lease_calc.py`).
2) 双击 `run.bat`（或运行 `python lease_calc.py`）。
3) Output is saved as `output.xlsx`.
3) 输出为 `output.xlsx`。

Input workbook:
输入工作簿：
- Sheet: `Leases` (or specify `--sheet`)
- 工作表：`Leases`（或使用 `--sheet` 指定）
- Required columns:
  必填列：
  `contract_id`, `lease_start`, `lease_end`, `payment_amount`,
  `payment_frequency` (M/Q/A), `discount_rate` (annual), `payment_timing` (begin/end)
- Optional columns: `currency`
- 可选列：`currency`
- Also supports common headers like `合同编号`, `起始日期`, `结束日期`, `租金`, `折现率`.
- 也支持常见表头，如 `合同编号`、`起始日期`、`结束日期`、`租金`、`折现率`。

Usage:
用法：
- `python lease_calc.py --input input.xlsx --output output.xlsx`

Output:
输出：
- `Summary` sheet with contract totals.
- `Summary` 工作表包含合同汇总。
- `Schedule` sheet with period-by-period details.
- `Schedule` 工作表包含逐期明细。

Notes:
备注：
- If `discount_rate` > 1, it is treated as a percent (e.g., 5 = 5%).
- 若 `discount_rate` > 1，将被视为百分比（例如 5 表示 5%）。
- `payment_frequency` 支持 M/Q/A 或 月/季/年。
