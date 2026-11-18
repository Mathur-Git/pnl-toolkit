@echo off
rem Fetches the newest toolkit (start\, tools\, README.md) from GitHub. Your statements, ledger
rem and PnL Record.xlsx are never touched. The record takes the new look at its next update.
cd /d "%~dp0.."
git pull --ff-only
echo.
pause
