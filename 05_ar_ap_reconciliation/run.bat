@echo off
setlocal
cd /d "%~dp0"
python reconcile_parties.py --input input.xlsx --output output.xlsx
pause
