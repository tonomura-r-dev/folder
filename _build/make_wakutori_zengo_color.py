# -*- coding: utf-8 -*-
"""ワクトリ 前後検索「派遣 登録」を3分類で色分け（2026-10-28）。作りは横浜シミズ版（make_yokohama_shimizu_zengo_color.py）をそのまま使う。
  A=働き方・求人サービスを比べる（青）／B=登録・応募の準備（紫）／C=お金・保険を調べる（ティール）
分類データ：_data/zengo/ワクトリ_派遣登録_クエリ分類.csv　元画像：_data/wakutori/zengo/前後検索_派遣_登録.webp（殿村さんから受領）
出力：_images/wakutori_zengo_haken_touroku_color.png（--legend <出力> で凡例付きの確認用も出す）
"""
import sys

from _build import make_yokohama_shimizu_zengo_color as base

base.SRC = base.ROOT / "_data/wakutori/zengo/前後検索_派遣_登録.webp"
base.CSV = base.ROOT / "_data/zengo/ワクトリ_派遣登録_クエリ分類.csv"
base.OUT = base.ROOT / "_images/wakutori_zengo_haken_touroku_color.png"
base.LEGEND = [("A", "働き方・求人サービスを比べる"), ("B", "登録・応募の準備"), ("C", "お金・保険を調べる")]

if __name__ == "__main__":
    im = base.render()
    if "--legend" in sys.argv:
        base.with_legend(im, sys.argv[sys.argv.index("--legend") + 1])
