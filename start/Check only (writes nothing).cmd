@echo off
rem Shows what an update would change in PnL Record.xlsx, without writing anything.
cd /d "%~dp0.."
python tools\update_record.py --dry
echo.
pause
