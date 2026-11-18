@echo off
rem Shows what an update would change in PnL Record.xlsx, without writing anything.
cd /d "%~dp0.."
rem the py launcher when installed: on some PCs "python" is only the Microsoft Store stub
set "PY=python"
where py >nul 2>&1 && set "PY=py -3"
%PY% tools\update_record.py --dry
echo.
pause
