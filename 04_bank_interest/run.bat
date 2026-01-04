@echo off
setlocal
cd /d "%~dp0"
python bank_interest.py --input input.xlsx --output output.xlsx
pause
