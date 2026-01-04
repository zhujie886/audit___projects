# Bank Deposit Interest
银行存款利息测算

Purpose: calculate bank deposit interest from Excel input.
目的：根据 Excel 输入计算银行存款利息。

Quick use:
快速使用：
1) Rename your Excel to `input.xlsx` and place it here.
1) 将 Excel 重命名为 `input.xlsx` 并放到此目录。
2) Double-click `run.bat` (or run `python bank_interest.py`).
2) 双击 `run.bat`（或运行 `python bank_interest.py`）。
3) Output is saved as `output.xlsx`.
3) 输出为 `output.xlsx`。

Input workbook:
输入工作簿：
- Sheet: `Deposits` (or specify `--sheet`)
- 工作表：`Deposits`（或使用 `--sheet` 指定）
- Required columns: `principal`, `annual_rate`
- 必填列：`principal`, `annual_rate`
- Required either: `days` or (`start_date` and `end_date`)
- 必须提供：`days` 或（`start_date` 和 `end_date`）
- Optional columns: `account`, `day_count`
- 可选列：`account`, `day_count`
- Also supports common headers like `本金`, `年利率`, `天数`, `起息日`, `到期日`.
- 也支持常见表头，如 `本金`、`年利率`、`天数`、`起息日`、`到期日`。

Usage:
用法：
- `python bank_interest.py --input input.xlsx --output output.xlsx`

Output:
输出：
- Adds `days_calc`, `interest`, and `maturity_amount` columns.
- 新增列：`days_calc`, `interest`, `maturity_amount`

Notes:
备注：
- If `annual_rate` > 1, it is treated as a percent (e.g., 3.5 = 3.5%).
- 若 `annual_rate` > 1，将被视为百分比（例如 3.5 表示 3.5%）。
