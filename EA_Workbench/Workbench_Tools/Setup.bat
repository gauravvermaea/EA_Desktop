@echo off
rem EA Workbench Setup - opens the setup window (or the menu with --menu)
cd /d "%~dp0"
set PY=
where py >nul 2>nul && set PY=py -3
if not defined PY where python >nul 2>nul && set PY=python
if not defined PY (
  echo Python 3 is not installed on this computer.
  echo Install it from https://www.python.org/downloads/ and tick "Add python.exe to PATH", then run Setup again.
  pause
  exit /b 1
)
%PY% workbench_setup.py %*
if errorlevel 1 pause
