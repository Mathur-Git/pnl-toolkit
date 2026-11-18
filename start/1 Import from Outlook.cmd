@echo off
rem Step 1 + 2 in one: takes the newest Outlook export (.pst) in Outlook-export, adds its daily and
rem monthly statement PDFs (never deletes or overwrites), then updates PnL Record.xlsx and runs the
rem monthly check. See README.md.
cd /d "%~dp0.."
python tools\outlook_import.py
if errorlevel 1 goto end
echo.
python tools\update_record.py
:end
echo.
pause
