# LINE施策統合シミュレーター

LINE公式アカウントの施策を、実装前に組み立てて「ユーザーに何が見えるか」と「裏で何が起きているか」を見せるツール。
設計は [`docs/`](docs/GOAL.md) を参照。

## いまできること（Phase 2a）

画面はまだ無い。検証案件を架空ユーザー4人で Day 30 まで流し、配信・分岐・停止の結果を表で出す。

- `start.bat` をダブルクリック（初回は準備に数分かかる）
- またはターミナルで：

```bash
npm run sim -- sample-staff --until-day 30
```

| オプション | 中身 |
|---|---|
| `--until-day N` | Day N まで流す（既定 30） |
| `--user p_osaka` | 1人だけ表示 |
| `--json` | 比較用のJSONで出す |

## 確認コマンド

```bash
npm run check
```

型チェック → エンジンの汎用性検査（日本語・禁止API・外部 import が0件か）→ テスト の順に走る。

## フォルダ

| 場所 | 中身 |
|---|---|
| `packages/engine/` | シナリオエンジン。依存ゼロ・DOMなし・日本語なし。編集ツールと営業用HTMLで同じものを使う |
| `packages/player/` | 画面側。いまは判定結果を日本語にする部分だけ |
| `projects/<案件>/project.json` | 案件（シナリオ設定） |
| `tools/` | CLI・検査 |
| `tests/` | テスト。`acceptance.test.ts` が Phase 2a の検収条件 |
