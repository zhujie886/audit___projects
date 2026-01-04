@echo off
setlocal
cd /d "%~dp0"
python rename_files.py --input input.xlsx --folder files
pause
