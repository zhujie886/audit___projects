# Confirmation Letter Generator
函证生成器

Purpose: generate confirmation letters (Word) from an Excel list.
目的：根据 Excel 列表生成 Word 函证。

Quick use:
快速使用：
1) Rename your Excel to `input.xlsx` and place it here.
1) 将 Excel 重命名为 `input.xlsx` 并放到此目录。
2) Double-click `run.bat` (or run `python generate_confirmations.py`).
2) 双击 `run.bat`（或运行 `python generate_confirmations.py`）。
3) Output is saved to `output` folder.
3) 输出保存到 `output` 文件夹。

Input workbook:
输入工作簿：
- Sheet: first sheet (or specify `--sheet`)
- 工作表：第一个工作表（或使用 `--sheet` 指定）
- Required columns: `party_name`, `amount`, `balance_date`
- 必填列：`party_name`, `amount`, `balance_date`
- Optional columns: `address`, `contact`, `currency`, `remarks`
- 可选列：`address`, `contact`, `currency`, `remarks`
- Also supports common headers like `往来单位`, `余额`, `截止日期`, `币种`.
- 也支持常见表头，如 `往来单位`、`余额`、`截止日期`、`币种`。

Usage:
用法：
- `python generate_confirmations.py --input input.xlsx --output output`
- `python generate_confirmations.py --make-template template.docx`
- `python generate_confirmations.py --input input.xlsx --output output --template template.docx`

Output:
输出：
- One `.docx` file per party in the output folder.
- 在输出文件夹中为每个往来单位生成一个 `.docx` 文件。
- `index.xlsx` summary listing all generated files.
- `index.xlsx` 汇总表列出所有生成文件。

Notes:
备注：
- Template placeholders use `{{key}}`, e.g., `{{party_name}}`, `{{amount_formatted}}`.
- 模板占位符使用 `{{key}}`，例如 `{{party_name}}`、`{{amount_formatted}}`。
- If using a Word template, keep placeholders in a single run (avoid mixed formatting).
- 使用 Word 模板时，尽量保证占位符在同一段落/同一格式内（避免被拆分）。
