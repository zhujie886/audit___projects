# Financial Statements + Checks
财务报表测算与校验

Purpose: compute BS/IS/CF from a trial balance and mapping, with validation checks.
目的：根据试算平衡表与映射生成资产负债表/利润表/现金流量表，并进行校验。

Quick use (auto mode):
快速使用（自动模式）：
1) Export a trial balance from Kingdee/UFIDA and rename it to `input.xlsx`.
1) 从金蝶/用友导出试算平衡表并重命名为 `input.xlsx`。
2) Put it in this folder and double-click `run.bat`.
2) 放入本目录并双击 `run.bat`。
3) Output is saved as `output.xlsx`.
3) 输出为 `output.xlsx`。

Input workbook sheets:
输入工作簿工作表：
- `TB_Current`: `account_code`, `ending_balance` (account_name optional)
- `TB_Current`：`account_code`, `ending_balance`（`account_name` 可选）
- `Mapping`: `statement` (BS/IS/CF), `section`, `line_item`, `account_code`, `sign` (optional, default 1)
- `Mapping`：`statement`（BS/IS/CF），`section`, `line_item`, `account_code`, `sign`（可选，默认 1）
- `Parameters` (optional): two-column key/value pairs, e.g. `cash_begin`, `cash_end`, `tolerance`
- `Parameters`（可选）：两列表头键值对，例如 `cash_begin`, `cash_end`, `tolerance`
- If `Mapping` is missing, the script auto-classifies by 科目类型/科目名称/科目编码.
- 若缺少 `Mapping`，程序会按 科目类型/科目名称/科目编码 自动分类。
- It auto-detects TB sheets such as `科目余额表` or `试算平衡表`.
- 会自动识别 `科目余额表`、`试算平衡表` 等工作表。

Usage:
用法：
- `python financial_statements.py --input input.xlsx --output output.xlsx`

Output:
输出：
- `BS`, `IS`, `CF` sheets with line items.
- `BS`, `IS`, `CF` 工作表包含明细行。
- `Checks` sheet with errors and warnings.
- `Checks` 工作表列示错误与警告。
- `Unclassified` (auto mode) lists accounts not mapped to BS/IS.
- `Unclassified`（自动模式）列示未能归类的科目。

Checks:
校验项：
- Balance sheet equation (Assets = Liabilities + Equity).
- 资产负债表平衡关系（资产 = 负债 + 权益）。
- Cash flow net change vs. `cash_end - cash_begin` (if provided).
- 现金流净变动与 `cash_end - cash_begin` 比较（若提供）。
- Missing account codes in mapping.
- 映射表中缺失的科目代码。
- Auto mode reads `余额方向` / 借贷余额 to compute signed balances.
- 自动模式会读取 `余额方向` 或借贷余额计算方向。
