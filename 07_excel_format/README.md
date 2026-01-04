# Excel Financial Format
Excel 财务格式

Purpose: apply financial-style formatting to Excel files.
目的：对 Excel 文件应用财务风格格式。

Quick use:
快速使用：
1) Rename your Excel to `input.xlsx` and place it here.
1) 将 Excel 重命名为 `input.xlsx` 并放到此目录。
2) Double-click `run.bat` (or run `python format_excel.py`).
2) 双击 `run.bat`（或运行 `python format_excel.py`）。
3) Output is saved as `output.xlsx`.
3) 输出为 `output.xlsx`。

Usage:
用法：
- `python format_excel.py --input input.xlsx --output output.xlsx`

Options:
选项：
- `--sheets Sheet1,Sheet2` (default: all)
- `--sheets Sheet1,Sheet2`（默认：全部）
- `--header-row 1`
- `--header-row 1`（表头行，默认 1）
- `--scan-rows 20` (rows used to detect numeric/date columns)
- `--scan-rows 20`（用于识别数值/日期列的扫描行数）

What it does:
功能说明：
- Formats numeric columns with accounting-style formats.
- 数值列应用会计格式。
- Formats date columns as `yyyy-mm-dd`.
- 日期列格式化为 `yyyy-mm-dd`。
- Styles the header row and freezes panes.
- 设置表头样式并冻结窗格。
- Adjusts column widths.
- 自动调整列宽。
