@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Manual Studio

where node >nul 2>nul
if errorlevel 1 goto :missing_node
where npm >nul 2>nul
if errorlevel 1 goto :missing_npm

node -e "const v=process.versions.node.split('.')[0]; if(Number(v)<20) process.exit(1)"
if errorlevel 1 goto :old_node

if not exist "node_modules\docx\dist\index.iife.js" (
  echo Installing the required document library. Internet access is needed once.
  call npm install --omit=dev --no-audit --no-fund
  if errorlevel 1 goto :failed
)

if not exist ".env" (
  copy /y ".env.example" ".env" >nul
  echo First run setup: enter your OpenAI API key in the .env file that will open.
  start "" /wait notepad.exe ".env"
)

node "local_server.mjs"
if errorlevel 1 goto :failed
exit /b 0

:missing_node
echo Node.js 20 or later was not found. Install Node.js LTS, then run this file again.
goto :failed
:missing_npm
echo npm was not found. Reinstall Node.js LTS with npm enabled, then retry.
goto :failed
:old_node
echo This app requires Node.js 20 or later. Install Node.js LTS, then retry.
goto :failed
:failed
echo.
echo Manual Studio did not start. See README.md for setup requirements.
pause
exit /b 1
