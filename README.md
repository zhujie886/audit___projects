# Audit Mini-Programs for Excel (Python)
用于 Excel 的审计小程序（Python）

This repo contains standalone Python scripts that read and write `.xlsx` files.
本仓库包含可读写 `.xlsx` 文件的独立 Python 脚本。
Each program lives in its own folder with a README and usage examples.
每个程序位于独立文件夹，包含 README 与用法示例。

Quick start:
快速开始：
1) Install Python 3.9+.
1) 安装 Python 3.9+。
2) Install dependencies:
2) 安装依赖：
   `pip install -r requirements.txt`
   运行：`pip install -r requirements.txt`

One-click use (Kingdee/UFIDA exports):
一键使用（适配金蝶/用友导出）：
1) Copy a program folder to any location.
1) 复制任意一个程序文件夹到本地。
2) Rename your exported Excel to `input.xlsx` and place it in the folder.
2) 将导出的 Excel 重命名为 `input.xlsx` 并放入该文件夹。
3) Double-click `run.bat` (or run `python xxx.py` in that folder).
3) 双击 `run.bat`（或在该文件夹运行 `python xxx.py`）。
4) Output is saved as `output.xlsx` or `output` folder.
4) 输出为 `output.xlsx` 或 `output` 文件夹。

Programs:
程序列表：
- `01_rename_files`: rename files in a folder based on an Excel mapping list.
- `01_rename_files`：根据 Excel 映射表批量重命名文件夹内的文件。
- `02_confirmation_letters`: generate confirmation letters (Word) from an Excel list.
- `02_confirmation_letters`：根据 Excel 列表生成 Word 函证。
- `03_lease_calculation`: compute lease amortization schedules.
- `03_lease_calculation`：生成租赁摊销测算表。
- `04_bank_interest`: calculate bank deposit interest.
- `04_bank_interest`：计算银行存款利息。
- `05_ar_ap_reconciliation`: reconcile AR/AP/Other AR/AP by counterparty.
- `05_ar_ap_reconciliation`：按往来单位勾稽应收/应付/其他应收/其他应付。
- `06_financial_statements`: compute BS/IS/CF from trial balance + mapping, with checks.
- `06_financial_statements`：根据试算平衡表与映射生成资产负债表/利润表/现金流量表，并含校验。
- `07_excel_format`: apply financial-style formatting to Excel files.
- `07_excel_format`：对 Excel 应用财务格式。
- `08_excel_rounding`: round numeric values in Excel files.
- `08_excel_rounding`：对 Excel 数值进行四舍五入。

Each folder README documents required columns, sheet names, and output details.
每个文件夹的 README 说明必填列、工作表名称和输出细节。
