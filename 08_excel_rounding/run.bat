@echo off
setlocal
cd /d "%~dp0"
python round_excel.py --input input.xlsx --output output.xlsx --decimals 2
pause
