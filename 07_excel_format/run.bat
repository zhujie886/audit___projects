@echo off
setlocal
cd /d "%~dp0"
python format_excel.py --input input.xlsx --output output.xlsx
pause
