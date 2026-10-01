#!/usr/bin/env python3
"""企画案の行データ（JSON）から、配信管理シートに貼るタブ区切りの行を作る。

使い方:
  python3 make_tsv.py <rows.json>          # シートに貼るタブ区切り（1本ずつ）
  python3 make_tsv.py <rows.json> --view   # 確認用（列名：値）

rows.json は行のリスト。キーはシートの列名（COLUMNS）。無いキーは空欄。
COLUMNS 以外のキー（アイテム名・配信時間の参考・テキストのリンク など）は
タブ区切りには入れず、--view にだけ出す（管理画面で入れる物）。
GA_URL が空なら URL と ga_source／ga_medium／ga_campaign から作る。
気になる点（日付のずれ・字数・1行の長さ・空の必須欄）は標準エラーに出す。
"""
import json
import re
import sys

COLUMNS = [
    "No.", "納品日", "配信日", "Status", "project", "企画案", "狙い概要", "ターゲット",
    "サイズ", "完成バナー", "テキスト", "備考", "バナー参考", "先方FB", "URL",
    "ga_source", "ga_medium", "ga_campaign", "aa", "messageID ※投稿後", "GA_URL",
]
REQUIRED = ["配信日", "企画案", "狙い概要", "ターゲット", "テキスト", "URL"]
TEXT_LIMIT = 500   # LINEのテキストは1吹き出し500字まで（改行も1字）
LINE_LIMIT = 15    # プレビューの幅は全角14〜15字


def cell(value):
    s = "" if value is None else str(value)
    s = s.replace("\r\n", "\n").replace("\t", " ")
    if "\n" in s or '"' in s:
        s = '"' + s.replace('"', '""') + '"'
    return s


def build_ga_url(row):
    url = (row.get("URL") or "").strip()
    params = [(k, (row.get(c) or "").strip()) for k, c in
              (("utm_source", "ga_source"), ("utm_medium", "ga_medium"), ("utm_campaign", "ga_campaign"))]
    query = "&".join(f"{k}={v}" for k, v in params if v)
    if not url or not query:
        return ""
    return url + ("&" if "?" in url else "?") + query


def warnings(row):
    name = f"No.{row.get('No.', '?')} {row.get('配信日', '')}"
    found = []
    for col in REQUIRED:
        if not (row.get(col) or "").strip():
            found.append(f"{col}が空")
    m = re.match(r"(\d{4})/(\d{1,2})/(\d{1,2})", row.get("配信日") or "")
    camp = row.get("ga_campaign") or ""
    if m and camp:
        ymd = f"{m.group(1)}{int(m.group(2)):02d}{int(m.group(3)):02d}"
        if not camp.startswith(ymd):
            found.append(f"ga_campaign の日付（{camp[:8]}）が配信日（{ymd}）と違う")
    text = row.get("テキスト") or ""
    if len(text) > TEXT_LIMIT:
        found.append(f"テキストが{len(text)}字（{TEXT_LIMIT}字まで）")
    for line in text.split("\n"):
        if line.startswith("http"):
            continue
        if len(line) > LINE_LIMIT:
            found.append(f"1行が{len(line)}字：「{line}」")
    for w in found:
        print(f"⚠️ {name}：{w}", file=sys.stderr)


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    with open(sys.argv[1], encoding="utf-8") as f:
        rows = json.load(f)
    view = "--view" in sys.argv[2:]
    for row in rows:
        if not (row.get("GA_URL") or "").strip():
            row["GA_URL"] = build_ga_url(row)
        warnings(row)
        title = f"No.{row.get('No.', '')} {row.get('配信日', '')} {row.get('企画案', '')}"
        if view:
            print(f"## {title}\n")
            for col in COLUMNS + [k for k in row if k not in COLUMNS]:
                value = row.get(col, "")
                mark = "（管理画面）" if col not in COLUMNS else ""
                if "\n" in str(value):
                    print(f"【{col}】{mark}\n```\n{value}\n```")
                else:
                    print(f"【{col}】{mark}{value}")
            print()
        else:
            print(f"### {title}\n```")
            print("\t".join(cell(row.get(c, "")) for c in COLUMNS))
            print("```\n")


if __name__ == "__main__":
    main()
