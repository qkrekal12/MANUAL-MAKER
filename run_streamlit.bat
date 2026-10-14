@echo off
setlocal
cd /d "%~dp0"
title AI Manual Studio - Streamlit

where py >nul 2>nul
if errorlevel 1 (
  echo Python Launcher (py) was not found. Install Python 3 and enable PATH, then retry.
  pause
  exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
  echo Creating the isolated Python environment...
  py -3 -m venv .venv
  if errorlevel 1 goto failed
)

if not exist ".venv\Scripts\streamlit.exe" (
  echo Installing Streamlit app dependencies. This may take a few minutes on first launch...
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
  if errorlevel 1 goto failed
)

start "" http://127.0.0.1:8767
".venv\Scripts\python.exe" -m streamlit run run_streamlit.py --server.address 127.0.0.1 --server.port 8767 --browser.gatherUsageStats false
if errorlevel 1 goto failed
exit /b 0

:failed
echo.
echo Streamlit could not start. Review the error above and confirm Python and internet access.
pause
exit /b 1
