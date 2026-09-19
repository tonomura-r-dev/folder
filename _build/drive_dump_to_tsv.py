#!/usr/bin/env python3
"""Googleドライブコネクタのスプシ読み取り結果（JSON/markdown表）をタブ区切りTSVに変換する。

使い方:
  python3 _build/drive_dump_to_tsv.py <dump.txt> <出力ディレクトリ>

<dump.txt> は mcp__Google_Drive__read_file_content の保存結果
（{"fileContent": "..."} のJSON、または markdown表そのもの）。
markdownの表ブロックごとに 1つのTSV（table_01.tsv, table_02.tsv, ...）を出力する。

注意:
- 読み取り専用の変換。元のスプシには一切書き込まない（禁止事項）。
- コネクタの出力は巨大シートだと途中で切れることがある。行数を必ず目視確認すること。
- セル内のタブ・改行は空白に潰す（列ズレ防止。11列TSVで実際に事故った）。
"""
import json
import re
import sys
from pathlib import Path


def load_content(path: str) -> str:
    text = Path(path).read_text(encoding="utf-8")
    try:
        return json.loads(text)["fileContent"]
    except (json.JSONDecodeError, KeyError, TypeError):
        return text


def split_tables(content: str):
    """連続する markdown 表の行ブロックを切り出す。"""
    tables, cur = [], []
    for line in content.splitlines():
        if line.lstrip().startswith("|"):
            cur.append(line)
        else:
            if cur:
                tables.append(cur)
                cur = []
    if cur:
        tables.append(cur)
    return tables


def row_cells(line: str):
    cells = [c.strip() for c in line.strip().strip("|").split("|")]
    # markdownエスケープを戻し、タブ・改行を潰す
    return [re.sub(r"[\t\n\r]+", " ", c.replace("\\_", "_").replace("\\&", "&").replace("\\#", "#")) for c in cells]


def is_separator(line: str) -> bool:
    return bool(re.fullmatch(r"[|\s:\-]+", line.strip()))


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    content = load_content(sys.argv[1])
    outdir = Path(sys.argv[2])
    outdir.mkdir(parents=True, exist_ok=True)

    tables = split_tables(content)
    if not tables:
        sys.exit("markdown表が見つからない。dumpの中身を確認すること。")

    for i, tbl in enumerate(tables, 1):
        rows = [row_cells(l) for l in tbl if not is_separator(l)]
        rows = [r for r in rows if any(r)]  # 全セル空の行は捨てる
        out = outdir / f"table_{i:02d}.tsv"
        out.write_text("\n".join("\t".join(r) for r in rows), encoding="utf-8")
        print(f"{out}: {len(rows)}行 x {len(rows[0]) if rows else 0}列")


if __name__ == "__main__":
    main()
