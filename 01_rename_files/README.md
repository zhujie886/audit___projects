# Rename Files from Excel
从 Excel 批量重命名文件

Purpose: rename files in a target folder using an Excel mapping list.
目的：使用 Excel 映射表批量重命名目标文件夹内的文件。

Quick use:
快速使用：
1) Put files to rename into the `files` folder.
1) 将要重命名的文件放入 `files` 文件夹。
2) Rename your mapping Excel to `input.xlsx` and place it here.
2) 将映射表 Excel 重命名为 `input.xlsx` 并放到此目录。
3) Double-click `run.bat` (or run `python rename_files.py`).
3) 双击 `run.bat`（或运行 `python rename_files.py`）。

Input workbook:
输入工作簿：
- Sheet: first sheet (or specify `--sheet`)
- 工作表：第一个工作表（或使用 `--sheet` 指定）
- Required columns (header row 1): `old_name`, `new_name`
- 必填列（第 1 行表头）：`old_name`, `new_name`
- Also supports common headers like `原文件名`, `新文件名`.
- 也支持常见表头，如 `原文件名`、`新文件名`。

Usage:
用法：
- `python rename_files.py --input input.xlsx --folder files`
- `python rename_files.py --input input.xlsx --folder files --dry-run`
- `python rename_files.py --input input.xlsx --folder files --overwrite`

Output:
输出：
- Files are renamed in the target folder.
- 在目标文件夹内完成文件重命名。

Notes:
备注：
- Uses a two-phase temp rename to avoid collisions.
- 使用两阶段临时重命名，避免重名冲突。
- Only files are renamed; directories are rejected.
- 仅重命名文件；目录将被拒绝。
