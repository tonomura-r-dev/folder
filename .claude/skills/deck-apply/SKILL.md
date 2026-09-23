---
name: deck-apply
description: 確定した修正指示をPowerPoint資料（.pptx）に反映するときの実装スキル。「この修正を資料に入れて」「P17をこう直して」「表の数値を差し替えて」「フォントを揃えて」など、内容が決まっている変更をpptxへ適用する作業で使う。python-pptxとPowerPoint COMの既知の落とし穴（一括置換の巻き込み・書式が落ちる条件・PDF書き出しの型エラー）を回避して、PNG目視まで行う。資料の中身の基準（構成・配色・文字サイズ・出典）は lineoa-industry §11 と dym-format が持っているので、そちらを必ず併読する。何をどう直すかを設計する作業は対象外。
---

# 資料への修正反映（実装編）

**確定した指示をpptxに入れて、崩れていないことをPNGで確かめるまで**が担当。

## 併読するもの（中身の基準はこちらが正）

このスキルは**手の動かし方だけ**を持つ。何を良しとするかは既存スキルにある。

| | 参照先 |
|---|---|
| 文字サイズ・出典の扱い・完成度基準・ファイル運用 | **`lineoa-industry` §11「資料作成ノウハウ」** |
| PPTXのQA項目（EMUサイズ・脚注のy座標・見出しband） | **`lineoa-industry` §8** |
| 配色・スライド種別・コンポーネント | **`dym-format`** |

ページを新規に足すときは**同じ資料内の似たページをコピーして中身を差し替える**。ゼロから描かない。

## 進め方

```
① 元ファイル → scratchpad\demo.pptx にコピー
② demo.pptx を編集するスクリプトを書く
③ PNG書き出し → 目視
④ 崩れがあれば ①からやり直す（毎回コピーし直して冪等に）
⑤ PNGを見せてOKをもらう
⑥ 指示された出力先へ反映（直前にバックアップを scratchpad へ）
```

**反映のタイミングと出力先（上書き／別名／PDFも要るか）は、その都度の指示に従う。**決め打ちしない。指示がなければ聞く。

**②〜④を何度も回すので、スクリプトは毎回「元ファイルをコピーし直してから編集」する形にする。**編集済みに追加編集を重ねると失敗時に戻せない。

---

## python-pptx の落とし穴

### グループを再帰で拾う

DYMのFMTはグループが多い。`slide.shapes` を直接回すと中身が取れない。

```python
def walk(shs, out):
    for s in shs:
        if s.shape_type == 6:      # GROUP
            walk(s.shapes, out)
        else:
            out.append(s)
```

### ★一括置換をしない

「¥50,000を全部¥100,000に」は他の行を巻き込む。
**実例：サンクスの初期を上げたら、CPF広告の月額まで10万になった。**

行・列を特定して書き換える。

```python
for ri, r in enumerate(tbl.rows):
    if 'サンクス' in r.cells[0].text:
        setcell(tbl.cell(ri, 3), '¥100,000')      # 行と列を両方指定
```

### テキスト図形は「座標＋現在値」で特定する

同じ文言が複数あるので、テキスト一致だけでは足りない。

```python
cand = [s for s in shapes
        if s.has_text_frame
        and abs(s.top / 914400 - top) < 0.25
        and abs(s.left / 914400 - left) < 0.25
        and s.text_frame.text.strip().startswith(old)]
if len(cand) != 1:
    LOG.append('★見つからず')
```

**「見つからず」は必ず報告する。**黙ってスキップすると直ったつもりで直っていない。

### ★段落を潰すと色が消える

段落数と行数が一致するなら、**段落ごとに書き換える**。
**実例：通算コスト／通算CVの2段落を1つにまとめたら、赤字が黒になった。**

```python
lines = new.split('\n')
paras = tf.paragraphs
if len(lines) == len(paras) and all(p.runs for p in paras):
    for p, line in zip(paras, lines):
        r = p.runs[0]
        for extra in p.runs[1:]:
            extra._r.getparent().remove(extra._r)
        r.text = line          # 各段落の書式はそのまま
else:
    ...  # 段落数が合わないときだけ集約
```

### ★空セルは隣から書式を複製する

run がないセルに `add_run()` すると既定フォント（18pt相当）になり、**行が膨らんで表が枠外にはみ出す。**

```python
import copy
from pptx.oxml.ns import qn

def setcell(cell, text, like=None):
    tf = cell.text_frame
    p = tf.paragraphs[0]
    for extra_p in tf.paragraphs[1:]:
        extra_p._p.getparent().remove(extra_p._p)
    src = like.text_frame.paragraphs[0].runs[0] if like is not None else None
    if p.runs:
        r = p.runs[0]
        for extra in p.runs[1:]:
            extra._r.getparent().remove(extra._r)
    else:
        r = p.add_run()
    if src is not None:
        old = r._r.find(qn('a:rPr'))
        newp = src._r.find(qn('a:rPr'))
        if newp is not None:
            if old is not None:
                r._r.remove(old)
            r._r.insert(0, copy.deepcopy(newp))
    r.text = text
    if like is not None:
        p.alignment = like.text_frame.paragraphs[0].alignment
```

**`like` は同じ列の別セルを渡す。**列ごとに色が違うことがある。
**実例：対応項目列は `srgbClr 000000`（黒）、初期列は `schemeClr tx1 + lumMod 65%`（グレー）。**別列を渡して全部グレーになった。

### テーブルの行を入れ替える

```python
def swap_rows(tbl, a, b):        # a < b
    trs = tbl._tbl.tr_lst
    tr_a, tr_b = trs[a], trs[b]
    tbl._tbl.remove(tr_a)
    tr_b.addnext(tr_a)
```

### ★SIM表の列オフセット

DYMのSIM表（10列）は **列3＝現状、列4〜9＝1〜6か月目**。
`cell(r, 4)` から7個入れると列外で `IndexError`。**7個なら `col_start=3`。**

### フォントは latin / ea / cs の3つを設定する

```python
def set_font(run):
    rPr = run._r.get_or_add_rPr()
    for tag, face in (('a:latin', 'Meiryo'), ('a:ea', 'メイリオ'), ('a:cs', 'Meiryo')):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {})
            rPr.append(el)
        el.set('typeface', face)
```

英数字 `Meiryo` ／日本語 `メイリオ` の**2本立てが正常**。混在は異常ではない。
`Times New Roman` や「指定なし」が混ざっていたら直す。

**指摘されたページ以外も全走査する。実例：指摘はP17・P36だったが、原因はP3・P4・P7・P8にあった。**
文字幅が変わるので、直したページは必ずPNGで崩れを確認する。

---

## PowerShell COM の落とし穴

### ps1 は Python から UTF-8 BOM で書き出す

日本語パスを含むps1をWriteツールで直接書くと、Shift-JIS誤読でパースエラーになる。

```python
io.open(path, 'w', encoding='utf-8-sig').write(ps)
```

### PNG書き出し

```powershell
$pr = $pp.Presentations.Open($src, $true, $false, $false)   # 読み取り専用
$pr.Slides.Item($n).Export($path, "PNG", 1600, 900)
```

### ★PDFは SaveCopyAs を使う

`ExportAsFixedFormat` は PowerShell から**引数の型変換で落ちる**
（`Cannot convert the "2" value of type "int" to type "Object"`）。

```powershell
$pr.SaveCopyAs($out, 32)      # 32 = ppSaveAsPDF
```

### ★非表示スライドはPDFに出ない

総スライド数とPDFのページ数が食い違ったらこれ。

```powershell
foreach ($s in $pr.Slides) { if ($s.SlideShowTransition.Hidden -eq -1) { $hidden += $s.SlideIndex } }
```

### PermissionError / LOCKED

**まず開いていないか確認する。**

```powershell
Get-Process POWERPNT, EXCEL -ErrorAction SilentlyContinue | Select-Object Id, MainWindowTitle
```

`MainWindowTitle` が空なら COM の孤児プロセス。落として安全。

```powershell
Get-Process EXCEL | Where-Object { [string]::IsNullOrWhiteSpace($_.MainWindowTitle) } | Stop-Process -Force
```

タイトルが出ていればユーザーが開いている。**閉じてもらう。**

### 後始末

```powershell
$pr.Close(); $pp.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
```

省くと孤児プロセスが残る。

### PDFのテキスト検証はできない

日本語フォントを埋め込んでいるためテキストがグリフIDで格納され、**バイト検索では文字を拾えない**。
「古い数値が残っていないか」の検証は**PDFではなくpptxのPNGで行う**。
PDFで確認できるのはページ数（`/Count`）とタイムスタンプまで。

---

## 反映が終わったら

**元データとの突き合わせを表で出す。**

```
        SIM本体    資料
初期    345,000    P16・P32・P34  ✅
半年CV  63件       P17・P32・P34  ✅
```

報告は**変更したセルだけ**に絞る（スマホで見られる）。バックアップの場所を伝える。

---

## 指示が曖昧なら着手しない

「よしなに」「見やすく」「最新にして」は、先に表の形で確定させる。

| ページ | 対象 | 現在 | 変更後 |
|---|---|---|---|

**資料内で数値が食い違っている場合は勝手に決めない。**
**実例：P31=44件／P17=65件／SIM本体=63件の三つ巴。**どれが正データかを明示してもらう。
