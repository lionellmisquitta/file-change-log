@echo off
setlocal
cd /d "%~dp0"
if exist "runtime\python.exe" (
  "runtime\python.exe" "run.py"
) else (
  echo This source checkout needs Python 3.11+ and requirements.txt installed.
  echo For the no-install version, use the Windows x64 download package.
  py -3 "run.py"
)
if errorlevel 1 (
  echo.
  echo Recorder could not start. See %%LOCALAPPDATA%%\FileChangeLog\startup-error.txt
  pause
)
