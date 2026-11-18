@echo off
rem For statement PDFs dropped into Daily-statements by hand: reads the new ones, updates
rem PnL Record.xlsx and runs the monthly check. See README.md.
cd /d "%~dp0.."
rem the py launcher when installed: on some PCs "python" is only the Microsoft Store stub
set "PY=python"
where py >nul 2>&1 && set "PY=py -3"
%PY% tools\update_record.py
echo.
pause
