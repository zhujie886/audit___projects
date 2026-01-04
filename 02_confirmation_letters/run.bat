@echo off
setlocal
cd /d "%~dp0"
python generate_confirmations.py --input input.xlsx --output output
pause
