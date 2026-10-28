@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo First launch: installing required components...
  python -m venv .venv
  if errorlevel 1 goto fail
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 goto fail
)
start "" ".venv\Scripts\pythonw.exe" -m editable_pptx %*
exit /b 0
:fail
echo Setup failed. Install Python 3 and try again.
pause
exit /b 1
