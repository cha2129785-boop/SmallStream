@echo off
cd /d "%~dp0"
py -3.12 -m pip install -r requirements.txt
echo.
echo Done. Run SmallStream.pyw or run.bat
pause
