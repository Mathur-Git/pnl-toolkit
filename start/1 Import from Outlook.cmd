@echo off
rem Step 1 + 2 in one: takes the newest Outlook export (.pst) in Outlook-export, adds its daily and
rem monthly statement PDFs (never deletes or overwrites), then updates PnL Record.xlsx and runs the
rem monthly check. See README.md.
cd /d "%~dp0.."
rem the py launcher when installed: on some PCs "python" is only the Microsoft Store stub
set "PY=python"
where py >nul 2>&1 && set "PY=py -3"
%PY% tools\outlook_import.py
if errorlevel 1 goto end
echo.
%PY% tools\update_record.py
:end
echo.
pause
