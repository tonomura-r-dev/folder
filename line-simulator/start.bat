@echo off
chcp 65001 > nul
cd /d "%~dp0"

rem 初回だけ依存を入れる（数分かかります）
if not exist node_modules (
  echo 初回の準備をしています...
  call npm install --no-audit --no-fund
  if errorlevel 1 (
    echo 準備に失敗しました。画面をそのまま担当者に見せてください。
    pause
    exit /b 1
  )
)

rem Phase 2a: 編集画面はまだ無いので、検証案件を Day 30 まで流した結果を出す
call npm run sim -- sample-staff --until-day 30
pause
