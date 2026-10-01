# -*- coding: utf-8 -*-
"""サンクスLINE：殿村さんがChatGPTで生成した画像を差し込む（2026-10-01）。上書き保存。
- 「LINEは、申込みの「後」につながる」のページ：四角と矢印の図を _images/thanks_line_journey.png に差し替え（下の一文は残す）
- 「実績と費用」のページ：6つの実績カードに業種アイコン（_images/thanks_icon_*.png）を付ける
  python3 _build/patch_thanks_line_images.py <pptx> [<pptx> ...]
"""
import sys
from pathlib import Path
from pptx import Presentation
from pptx.util import Cm

ROOT = Path(__file__).resolve().parent.parent
IMG = ROOT / "_images"
ICONS = {  # 実績カードの1行目（何をしたか）→ アイコン
    "完了画面からLINEへ": "beauty", "予約日時をLINEで": "hoken", "予約の通知をLINEで": "bus",
    "友だち追加後のステップ": "hifuka", "LINEからの問い合わせ": "shushoku", "メッセージを開封した": "cafe",
}

for path in sys.argv[1:]:
    prs = Presentation(path)
    SW = prs.slide_width / 360000
    for s in prs.slides:
        title = s.shapes[0].text_frame.text if s.shapes[0].has_text_frame else ""
        if title.startswith("LINEは、申込みの「後」"):
            for sh in list(s.shapes):
                if sh.name in ("Rounded Rectangle 4", "Right Arrow 5", "Rounded Rectangle 6", "Right Arrow 7",
                               "Right Arrow 9", "Rounded Rectangle 14", "Rounded Rectangle 15",
                               "Rounded Rectangle 16", "Rounded Rectangle 17"):
                    sh._element.getparent().remove(sh._element)
            W = 21.0
            H = W * 941 / 1672
            pic = s.shapes.add_picture(str(IMG / "thanks_line_journey.png"), Cm((SW - W) / 2), Cm(4.15), Cm(W), Cm(H))
            pic.name = "広告とLINEの役割"
            band = next(x for x in s.shapes if x.name == "Rounded Rectangle 13")
            band.left, band.top, band.width = Cm(1.46), Cm(4.15 + H + 0.3), Cm(SW - 2.92)
        if title.startswith("実績と費用"):
            for sh in list(s.shapes):
                if not sh.has_text_frame:
                    continue
                first = sh.text_frame.paragraphs[0].text
                key = next((k for k in ICONS if first.startswith(k)), None)
                if not key:
                    continue
                d = 2.3
                s.shapes.add_picture(str(IMG / f"thanks_icon_{ICONS[key]}.png"),
                                     sh.left + Cm(0.25), sh.top + (sh.height - Cm(d)) // 2, Cm(d), Cm(d))
                sh.text_frame.margin_left = Cm(d + 0.35)
    prs.save(path)
    print("saved:", path)
