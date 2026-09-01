# 📦 案件インデックス（おもちゃ箱の目次）

**セッションを始めたら、Claudeはまずこのファイルを読むこと。**
「あの案件どうなってたっけ」はここを見れば全部わかる、という状態を維持する。

最終更新：2026-09-01

---

## ⚠️ 最重要：成果物はブランチに散らばっている

このリポジトリは**masterに統合されていない**。完成したPPTX/XLSXの多くは
各作業ブランチにしか存在しない。**「masterに無い＝存在しない」ではない。**
探すときは必ず下の対応表を見ること。

### ブランチ → 中に何があるか

| ブランチ | 入っている成果物 |
|---|---|
| `master` | 注文住宅業界の一式／ビイサイドプランニング／vivical／LP分析 |
| `claude/pptx-creation-revision-5a4v0d` | **最も成果物が多い。** EC業界PPTX＋チェックリスト／賃貸VerA・VerB PPTX＋チェックリスト／小林住宅（20260901）／注文住宅 |
| `claude/kochira-douzo-mzomvp` | 賃貸業界 VerA・VerB のPPTX＋チェックリスト／注文住宅 |
| `claude/industry-standard-materials-ndzvb9` | `DYM_LINEOA_業界汎用_ご提案.pptx`（業界を絞らない汎用デッキ） |
| `claude/line-oa-greeting-message-znn29r` | `プリントアース_LINEOA_配信企画案_2026-09.xlsx` |
| `claude/special-plan-change-slide-d3rpsl` | LINE料金改定・特別プラン関連 2本（20260826） |
| `claude/proposal-management-sheet-f87u96` | `20260831_案件一覧_部署集計_11列.xlsx`／`_build/gyokai.py`（業界分類）／`teian_import_gs.gs` |
| `claude/task-management-sheet-nxzci8` | `task_management_sheet.gs`（Googleスプレッドシートのタスク管理。GASを貼って使う） |
| `claude/operations-task-sync-1v4z8h` | `_drafts/` の調査資料が最も揃っている（EC・単品通販・注文住宅・賃貸の全4案件） |
| `claude/multi-ai-workflow-integration-mpzlwz` | マルチAIリレーのスキル（`.claude/skills/multi-ai-relay/`）＋この索引 |

**ブランチ横断で探すコマンド**（ファイル名で探す場合）：
```bash
for ref in $(git for-each-ref --format='%(refname)' refs/remotes/origin/); do
  git -c core.quotepath=false ls-tree -r --name-only "$ref" | grep -i "探したい語" | sed "s|^|$ref: |"
done
```

---

## 業界汎用資料（業界名だけで作るシリーズ）

スキル `lineoa-industry` で生成。構成は36〜39枚。**業界汎用なのでSIMは入れない。**

| 案件 | 進捗 | 成果物の場所 | 残タスク |
|---|---|---|---|
| **注文住宅業界** | ✅ PPTX 36枚 完成 | `master`：`注文住宅業界_LINEOA施策提案.pptx`／`_drafts/注文住宅業界_要件定義（引き継ぎ用）.md` を最初に読む | ⚠️**対外提出不可**。チェックリスト赤4枚（S18/S24/S25/S29＝出典未記載の数値）。競合の友だち数が未取得（S9/S10が差込枠のまま） |
| **賃貸業界 Ver.A（入居者集客）** | ✅ PPTX 完成 | `claude/pptx-creation-revision-5a4v0d` または `claude/kochira-douzo-mzomvp` | 公表判断チェックリストの確認 |
| **賃貸業界 Ver.B（オーナー開拓）** | ✅ PPTX 完成 | 同上 | 同上 |
| **EC業界（総合EC）** | ✅ PPTX＋チェックリスト 完成 | `claude/pptx-creation-revision-5a4v0d`：`EC業界_LINEOA施策提案.pptx` | ⚠️CLAUDE.mdに記録が無かった案件。内容の再確認が必要 |
| **単品通販（D2C）業界** | 🔸 プロンプトまで完成、**PPTX未確認** | `_drafts/単品通販業界_{アジェンダ,要旨,調査結果,スライド作成プロンプト}.md` | PPTX化がまだの可能性が高い |
| **業界汎用（業界を絞らない）** | ✅ 完成 | `claude/industry-standard-materials-ndzvb9`：`DYM_LINEOA_業界汎用_ご提案.pptx` | — |

### 各業界の設計上の判断（引き継ぎで効く要点）

- **賃貸**：Ver.A＝入居者集客（B2C・主KPI＝CPO）／Ver.B＝オーナー開拓（B2B・主KPI＝面談CPA）。
  効いた発見＝**検索は年中フラット（ピーク7月）なのに契約は1-3月に集中**＝業界は1-3月しか刈っていない。
  広告の話は Ver.A の S06 だけに留める。
- **EC**：**賃貸のルール（広告を削る）を持ち込まない。** CPC推移＋広告審査は章2に必要。
  理由＝「獲得単価が上がるから回収側に回す」が資料の背骨。薬機法がLINE配信文面の制約になる。
  主KPI＝CPO と F2転換率の2本立て。
- **単品通販**：与件が**法律**。特商法改正（2022/6/1施行）で**解約妨害に罰則**。
  だから「解約を止める」ではなく**「解約の手前で受け止める」**を軸にする。主KPI＝F2転換率＋月次解約率、締めはLTV÷CPO。
- **注文住宅**：請負会社のLINE公式アカウント**公式事例は存在しない**（＝先行者になれる）。
  近い事例＝auka／オープンハウスG／ミサワホーム北越／LIFULL HOME'S。

---

## 個別クライアント案件

| 案件 | 成果物 | ブランチ |
|---|---|---|
| ビイサイドプランニング | LINE公式アカウント活用のご提案（2本） | `master` |
| 求人vivical『しが就職・転職フェア』 | ご提案／ver2 | `master` |
| 小林住宅株式会社 | `20260901_小林住宅株式会社 御中_LINE構築提案.pptx` | `claude/pptx-creation-revision-5a4v0d` |
| プリントアース | `プリントアース_LINEOA_配信企画案_2026-09.xlsx` | `claude/line-oa-greeting-message-znn29r` |
| LINE料金改定・特別プラン | 提供方針＋案内資料（20260826） | `claude/special-plan-change-slide-d3rpsl` |

---

## 社内の仕組み・ツール

| 名前 | 中身 | 場所 |
|---|---|---|
| タスク管理シート | GASをGoogleスプレッドシートに貼って使う。期日で色が変わる／毎朝カレンダー通知 | `claude/task-management-sheet-nxzci8`：`task_management_sheet.gs` |
| 案件一覧・部署集計 | 提案管理シート＋業界自動分類 | `claude/proposal-management-sheet-f87u96` |
| 資料ビルドスクリプト | `_build/build_*.py`（デッキ生成）／`qa_render.py`（検品）／`make_images.py`（図版生成） | 全ブランチ共通 |

---

## ❓ 行方不明のもの（見つかったらここを消す）

| 探しているもの | 状況 |
|---|---|
| **カフェのSIM** | 2026-09-01時点で**このリポジトリのどのブランチにも存在しない**。心当たりのセッション＝「LINEOA SIM作成」（`session_014asRKWSr3nzjoCoKUH5yMQ` 2026-08-19／`session_012S8Wmodct1x7MRJTHyyeGp` 2026-08-07）。どちらもコミット形跡なし＝**そのセッションのコンテナ内で作って保存せずに終わった可能性が高い**。殿村さんがそのセッションを開いて回収する必要がある |

> **教訓：セッション内で作ったファイルは、コミット・プッシュしないと消える。**
> 作業の最後に必ず push すること。これを怠ると今回のカフェSIMと同じことが起きる。

---

## 📥 この箱の使い方（ルール）

**読むとき**（セッション開始時）
1. まずこのINDEXを読む
2. 該当案件の「成果物の場所」のブランチを見に行く
3. 案件の詳細が要るなら `_drafts/<案件名>_*.md` を読む

**書くとき**（作業が一区切りしたら・セッションを閉じる前に）
1. 作った成果物を**必ず commit & push**（これをやらないと消える）
2. このINDEXの該当行を更新（進捗・場所・残タスク）
3. 新しい案件なら行を1つ足す
4. 判断の理由や引き継ぎ事項が多いときは `_drafts/<案件名>_要件定義（引き継ぎ用）.md` に書き、ここからリンクする

**書く粒度**：1案件1行＋残タスク。長い経緯はここに書かず `_drafts/` に逃がす。
INDEXは「どこに何があるか」の地図であって、資料そのものではない。
