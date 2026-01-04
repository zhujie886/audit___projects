@echo off
setlocal
cd /d "%~dp0"
python lease_calc.py --input input.xlsx --output output.xlsx
pause
