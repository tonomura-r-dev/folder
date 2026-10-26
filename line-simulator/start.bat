@echo off
rem LINE simulator editor launcher.
rem Keep this file ASCII only: cmd can break UTF-8 Japanese text in .bat files.
rem Japanese messages are printed by tools\say.ts instead.
chcp 65001 > nul
cd /d "%~dp0"

where node > nul 2> nul
if errorlevel 1 (
  echo [line-simulator] Node.js was not found. Please install Node.js 24 or later, then run this file again.
  pause
  exit /b 1
)

if not exist node_modules (
  node tools\say.ts install
  call npm install --no-audit --no-fund
  if errorlevel 1 (
    node tools\say.ts installFailed
    pause
    exit /b 1
  )
)

node tools\say.ts starting
call npm run dev
pause
