@echo off
rem SmallStream - make installer (registers in Windows "Installed apps").
rem All steps and Korean messages are in packaging\build_installer.py
rem (cmd misreads UTF-8 Korean lines in .bat files - keep this file ASCII only)
rem Default: dist_installer\SmallStream_Setup_<ver>.exe   Option: build.bat --portable (single EXE, no install)
cd /d "%~dp0"
set "PY="
py -3.12 -c "import sys" >nul 2>nul && set "PY=py -3.12"
if not defined PY (python -c "import sys" >nul 2>nul && set "PY=python")
if not defined PY (echo [ERROR] Python not found. Install it from python.org and run again. & pause & exit /b 1)
%PY% packaging\build_installer.py %*
echo.
pause
