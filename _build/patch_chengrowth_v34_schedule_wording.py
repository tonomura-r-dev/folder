# -*- coding: utf-8 -*-
"""チェングロウス ver3.3 → ver3.4：23枚目、サンクスLINE誘導・Profile+は先方の意向とサイトの状況に応じて決める書き方に（2026-10-02）
  python3 _build/patch_chengrowth_v34_schedule_wording.py <ver3.3.pptx>
"""
import sys
from pathlib import Path
from pptx import Presentation

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "20260929_チェングロウス_LINE公式アカウント運用のご提案ver3.4.pptx"
prs = Presentation(sys.argv[1])
sh = next(x for x in prs.slides[22].shapes if x.name == "Rounded Rectangle 7")
p = sh.text_frame.paragraphs[1]
p.runs[0].text = "あいさつ・リッチメニュー・配信の設計、離脱防止の設置（サンクスLINE誘導・Profile+はご要望に応じて）"
for r in p.runs[1:]:
    r._r.getparent().remove(r._r)
from pptx.util import Pt
p.runs[0].font.size = Pt(10)      # 1行に収める
prs.save(str(OUT))
print("saved:", OUT.name, len(prs.slides), "枚")
