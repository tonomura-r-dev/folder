# -*- coding: utf-8 -*-
"""すなお ver7.2：ver7.1と同じ作り方で、想定動線（①〜④／⑤〜⑧）から「離脱課題」を外し、「LINEの役割」を「心理」のすぐ下に上げる（2026-10-22 殿村さん）。

    python -m _build.patch_sunao_v72_no_risk <PC保存版ver6.6.pptx> <ver7.2.pptx>

  ベースは殿村さんのPC保存版ver6.6（`_data/sunao/PC保存版_20261022b_すなおコンサルティング_ver6.6.pptx`）。ほかの直しは ver7.1 と同じ。
"""
import sys

from _build.patch_sunao_v71_pc66 import main

if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], risk_row=False)
