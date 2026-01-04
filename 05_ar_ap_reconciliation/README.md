# AR/AP Reconciliation
应收应付勾稽

Purpose: reconcile parties across AR/AP/Other AR/Other AP.
目的：对往来单位的应收/应付/其他应收/其他应付进行勾稽。

Quick use:
快速使用：
1) Rename your Excel to `input.xlsx` and place it here.
1) 将 Excel 重命名为 `input.xlsx` 并放到此目录。
2) Double-click `run.bat` (or run `python reconcile_parties.py`).
2) 双击 `run.bat`（或运行 `python reconcile_parties.py`）。
3) Output is saved as `output.xlsx`.
3) 输出为 `output.xlsx`。

Input workbook:
输入工作簿：
- Sheets: `AR`, `AP`, `OtherAR`, `OtherAP` (override with CLI options)
- 工作表：`AR`, `AP`, `OtherAR`, `OtherAP`（可用参数覆盖）
- Columns: `party` (or `counterparty`/`vendor`/`customer`/`name`), `amount` (or `balance`)
- 列：`party`（或 `counterparty`/`vendor`/`customer`/`name`），`amount`（或 `balance`）
- Auto-detects sheet names like `应收账款`, `应付账款`, `其他应收款`, `其他应付款`.
- 会自动识别 `应收账款`、`应付账款`、`其他应收款`、`其他应付款` 等工作表。
- If only one sheet exists, it classifies by `科目名称` (e.g., 应收账款/应付账款).
- 若只有一个工作表，会根据 `科目名称` 自动分类（如 应收账款/应付账款）。

Usage:
用法：
- `python reconcile_parties.py --input input.xlsx --output output.xlsx`

Options:
选项：
- `--ar-sheet`, `--ap-sheet`, `--other-ar-sheet`, `--other-ap-sheet`

Output:
输出：
- `Summary` sheet with totals by party.
- `Summary` 工作表按往来单位汇总。
- `Issues` sheet listing parties with both receivable and payable balances.
- `Issues` 工作表列示同时存在应收与应付余额的单位。
- `Unclassified` sheet lists rows that could not be classified.
- `Unclassified` 工作表列示无法自动分类的明细行。

Notes:
备注：
- Use consistent sign conventions in the input (e.g., receivable positive, payable positive).
- 输入中请保持符号一致（例如：应收为正、应付为正）。
- Single-sheet mode also recognizes common codes like 1122/2202/1221/2241.
- 单表模式也会识别常见科目编码如 1122/2202/1221/2241。
