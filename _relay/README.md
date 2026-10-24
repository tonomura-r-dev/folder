# _relay — ChatGPT との GitHub 経由リレー

Claude（claude.ai/code）と ChatGPT を、このリポジトリを共通の置き場にしてつなぐ。
**ChatGPT に読ませてよいのはこのフォルダ（`_relay/`）だけ。** 他のフォルダには社内数字が入っている。

```
_relay/
  to_chatgpt/     ← Claude が置く。ChatGPT への依頼（調査・画像生成）
  from_chatgpt/   ← ChatGPT（Codex）が置く。調査結果。to_chatgpt と同じファイル名
```

## ファイル名
`YYYYMMDD_<案件名>_<調査|画像>.md`（例：`20261024_チェングロウス_調査.md`）

## 殿村さんが ChatGPT に送る文（コピペ用）

調査：
```
GitHub の tonomura-r-dev/folder にある _relay/to_chatgpt/ の最新ファイルを読んで、そこに書いてある依頼を実行してください。
読むのは _relay/ の中だけにしてください。
結果は Codex で _relay/from_chatgpt/ に同じファイル名で commit してください。
```

画像：
```
GitHub の tonomura-r-dev/folder にある _relay/to_chatgpt/ の最新ファイルを読んで、そこに書いてあるプロンプトで画像を生成してください。
読むのは _relay/ の中だけにしてください。
```
→ 出た画像はダウンロードして Claude のチャットに添付する（Codex は画像を作れないので、ここだけ手で運ぶ）。

## Claude 側の約束
- 依頼ファイルは **ChatGPT がそれだけ読めば動ける完成形** にする（背景・やること・出力形式・出典ルール）
- 調査の出力形式は `lineoa-proposal-flow` の型（表のみ・出典URLつき・確認できないものは「確認できず」）
- `from_chatgpt/` の結果は、出典URLのない数字を資料に載せない（CLAUDE.md 2026-09-28）
- 処理済みの依頼は消さず残す（履歴になる）
