@echo off
setlocal
cd /d "%~dp0"
python financial_statements.py --input input.xlsx --output output.xlsx
pause
