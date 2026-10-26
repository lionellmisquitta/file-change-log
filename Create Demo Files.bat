@echo off
cd /d "%~dp0"
if exist "runtime\python.exe" (
  "runtime\python.exe" "demo\create_demo.py"
) else (
  py -3 "demo\create_demo.py"
)
pause
