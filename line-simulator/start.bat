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

echo 編集ツールを起動しています。ブラウザが開いたら、この黒い画面は閉じずにそのままにしてください。
echo 終わるときは、編集画面で保存してから、この画面を閉じてください。
call npm run dev
pause
