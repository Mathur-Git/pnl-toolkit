@echo off
rem For statement PDFs dropped into Daily-statements by hand: reads the new ones, updates
rem PnL Record.xlsx and runs the monthly check. See README.md.
cd /d "%~dp0.."
python tools\update_record.py
echo.
pause
