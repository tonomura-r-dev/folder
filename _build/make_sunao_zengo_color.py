# -*- coding: utf-8 -*-
"""すなお 前後検索「投資セミナー」を3分類で色分け（2026-10-20）。タナカホーム版（make_tanaka_zengo_color.py）の部品をそのまま使う。
  A=ほかのセミナー・イベントと比べる（青）／B=不安を確かめる（紫）／C=自分で始める（ティール）
  オレンジは使わない（資料に黄色・オレンジを使わないルール）。分類に入らないラベルは色なし（薄いまま）。
  ※検索しているのは個人の投資家（保険営業の方ではない）。本命は前後検索「保険営業」「セミナー」で、これは来るまでの控え。
分類データ：_data/zengo/すなお_投資セミナー_クエリ分類.csv
出力：_images/sunao_zengo_toushi_seminar_color.png（凡例なし＝資料用）／ --legend <出力> で凡例つき（確認用）
"""
import sys

from _build import make_tanaka_zengo_color as base

base.SRC = base.ROOT / "_data/zengo/すなお_前後検索_投資セミナー.webp"
base.CSV = base.ROOT / "_data/zengo/すなお_投資セミナー_クエリ分類.csv"
base.OUT = base.ROOT / "_images/sunao_zengo_toushi_seminar_color.png"
base.COL = {"A": (52, 103, 178), "B": (142, 78, 198), "C": (0, 137, 123)}
base.LEGEND = [("A", "ほかのセミナー・イベントと比べる"), ("B", "不安を確かめる"), ("C", "自分で始める")]

if __name__ == "__main__":
    im = base.render()
    if "--legend" in sys.argv:
        base.with_legend(im, sys.argv[sys.argv.index("--legend") + 1])
