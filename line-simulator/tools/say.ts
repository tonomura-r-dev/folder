// start.bat から呼ぶ、日本語の案内。
// cmd は UTF-8 の .bat の日本語を行の途中で壊すことがあるので、.bat 自体は英数字だけにして、表示はここで行う。

const MESSAGES: Record<string, string> = {
  install: '初回の準備をしています（数分かかります）…',
  installFailed: '準備に失敗しました。この画面をそのまま担当者に見せてください。',
  starting: [
    '編集ツールを起動しています。',
    'ブラウザが開いたら、この黒い画面は閉じずにそのままにしてください。',
    '終わるときは、編集画面で保存してから、この画面を閉じてください。',
  ].join('\n'),
};

console.log(MESSAGES[process.argv[2] ?? ''] ?? '');
