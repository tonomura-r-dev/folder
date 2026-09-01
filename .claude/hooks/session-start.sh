#!/bin/bash
# セッション開始時に「案件インデックス」を読み込ませるフック。
# これにより、どの端末・どのセッションから始めても
# 過去の案件の場所と進捗を最初から把握した状態で会話が始まる。
set -euo pipefail

PROJECT_DIR="${CLAUDE_PROJECT_DIR:-$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)}"
INDEX="$PROJECT_DIR/_memory/INDEX.md"

if [ ! -f "$INDEX" ]; then
  exit 0
fi

cat <<'HEADER'
=== 📦 案件インデックス（_memory/INDEX.md の自動読み込み）===
これは過去のセッションから引き継いだ「記憶」です。
案件の所在・進捗を聞かれたら、まずこの索引を根拠に答えること。
作業が一区切りしたら、この索引を更新して commit & push すること。
HEADER

cat "$INDEX"

echo "=== 案件インデックス ここまで ==="
