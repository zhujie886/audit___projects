# Excel Rounding
Excel 四舍五入

Purpose: round numeric values to a fixed number of decimals.
目的：将数值按固定小数位数进行四舍五入。

Quick use:
快速使用：
1) Rename your Excel to `input.xlsx` and place it here.
1) 将 Excel 重命名为 `input.xlsx` 并放到此目录。
2) Double-click `run.bat` (or run `python round_excel.py`).
2) 双击 `run.bat`（或运行 `python round_excel.py`）。
3) Output is saved as `output.xlsx`.
3) 输出为 `output.xlsx`。

Usage:
用法：
- `python round_excel.py --input input.xlsx --output output.xlsx --decimals 2`

Options:
选项：
- `--sheets Sheet1,Sheet2` (default: all)
- `--sheets Sheet1,Sheet2`（默认：全部）
- `--columns A,C` or `--columns Amount,Tax`
- `--columns A,C` 或 `--columns Amount,Tax`
- `--header-row 1`
- `--header-row 1`（表头行，默认 1）

Notes:
备注：
- Formula cells are skipped.
- 公式单元格会被跳过。
