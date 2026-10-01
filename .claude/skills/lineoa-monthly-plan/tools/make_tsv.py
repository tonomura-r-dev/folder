#!/usr/bin/env python3
"""企画案の行データ（JSON）から、配信管理シートに貼るタブ区切りの行を作る。

使い方:
  python3 make_tsv.py <rows.json>               # 標準の21列（インクアートのシート）
  python3 make_tsv.py <rows.json> --fmt norst   # ノーストのシートの並び
  python3 make_tsv.py <rows.json> --view        # 確認用（列名：値）。--fmt と併用できる

rows.json は行のリスト。キーはシートの列名。無いキーは空欄。
シートの列に無いキー（配信日・計測URL・画像生成プロンプト など）は
タブ区切りには入れず、--view にだけ出す（管理画面で入れる物）。
シートの見出しが空の列（結合セルなど）は、いつも空欄で出す。
ただし見出しが空でも行に値が入っている列は名前を付けて持つ（ノーストの「AL列（見出しなし）」）。
標準の形では、GA_URL が空なら URL と ga_source／ga_medium／ga_campaign から作る。
気になる点（日付のずれ・字数・1行の長さ・空の必須欄）は標準エラーに出す。
日付のずれは ga_campaign の列と、テキストやリンクの中の utm_campaign を全部、配信日（無ければ納品日）と比べる。
"""
import json
import re
import sys

FORMATS = {
    "standard": {
        "columns": [
            "No.", "納品日", "配信日", "Status", "project", "企画案", "狙い概要", "ターゲット",
            "サイズ", "完成バナー", "テキスト", "備考", "バナー参考", "先方FB", "URL",
            "ga_source", "ga_medium", "ga_campaign", "aa", "messageID ※投稿後", "GA_URL",
        ],
        "required": ["配信日", "企画案", "狙い概要", "ターゲット", "テキスト", "URL"],
    },
    "norst": {
        # 2026-10-01 殿村さん受領の見出し（50列）。"" は見出しが空の列。
        # 見出しが空でも、シートの行に値が入っている列は名前を付けて持つ（AL列＝No.5で「1540×1000px」）
        "columns": [
            "No.", "納品日", "ディレクション担当", "制作担当", "進行度合", "用途", "型", "サイズ", "",
            "アイテム名（表示名）", "企画案", "狙い概要", "ターゲット", "備考", "テキスト",
            "参考イメージ", "CR概要", "", "", "", "", "", "", "", "", "", "完成バナー",
        ] + [""] * 10 + ["AL列（見出しなし）"] + [""] * 12,
        "required": ["納品日", "アイテム名（表示名）", "企画案", "狙い概要", "ターゲット", "テキスト"],
    },
}
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


def delivery_ymd(row):
    """配信日（無ければ納品日。ノーストのシートは納品日＝配信日）を YYYYMMDD で返す"""
    for key in ("配信日", "納品日"):
        m = re.match(r"(\d{4})/(\d{1,2})/(\d{1,2})", row.get(key) or "")
        if m:
            return f"{m.group(1)}{int(m.group(2)):02d}{int(m.group(3)):02d}"
    return ""


def warnings(row, required):
    name = f"No.{row.get('No.', '?')} {row.get('配信日') or row.get('納品日', '')}"
    found = []
    for col in required:
        if not (row.get(col) or "").strip():
            found.append(f"{col}が空")
    ymd = delivery_ymd(row)
    # ga_campaign の列と、テキストや画像のリンクの中の utm_campaign を全部見る
    camps = {}
    if row.get("ga_campaign"):
        camps.setdefault(row["ga_campaign"].strip(), "ga_campaign")
    for key, value in row.items():
        for camp in re.findall(r"utm_campaign=([^&\s]+)", str(value)):
            camps.setdefault(camp, key)
    for camp, key in camps.items():
        if ymd and not camp.startswith(ymd):
            found.append(f"{key}の utm_campaign（{camp}）の日付が配信日（{ymd}）と違う")
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
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        sys.exit(1)
    fmt = "standard"
    if "--fmt" in args:
        fmt = args[args.index("--fmt") + 1]
    columns = FORMATS[fmt]["columns"]
    named = [c for c in columns if c]
    with open(args[0], encoding="utf-8") as f:
        rows = json.load(f)
    view = "--view" in args
    for row in rows:
        if "GA_URL" in columns and not (row.get("GA_URL") or "").strip():
            row["GA_URL"] = build_ga_url(row)
        warnings(row, FORMATS[fmt]["required"])
        title = f"No.{row.get('No.', '')} {row.get('配信日', '')} {row.get('企画案', '')}"
        if view:
            print(f"## {title}\n")
            for col in named + [k for k in row if k not in named]:
                value = row.get(col, "")
                mark = "（シートの列に無い）" if col not in named else ""
                if "\n" in str(value):
                    print(f"【{col}】{mark}\n```\n{value}\n```")
                else:
                    print(f"【{col}】{mark}{value}")
            print()
        else:
            print(f"### {title}\n```")
            print("\t".join(cell(row.get(c, "")) if c else "" for c in columns))
            print("```\n")


if __name__ == "__main__":
    main()
