"""
10月号の右上の逆三角形（認知/検討/獲得/伸長）に、トピックごとの強調段の色を付ける（2026-10-26 殿村さん指示）。
まずシリアルコード（P.7）まで。P.8以降は各トピックを直すときに足す。
使い方: python3 _build/patch_trend_202610_funnel.py <pptx>
"""
import sys
from pptx import Presentation
import trend_report_engine as engine

F = sys.argv[1]
prs = Presentation(F)
slides = list(prs.slides)
STAGES = {  # ページ番号: 強調する段
    3: ["獲得"], 4: ["獲得"],  # ネイティブアプリ遷移
    5: [],                     # UX追加申請の請求書（手続きの話）
    6: ["獲得"], 7: ["獲得"],  # シリアルコードスタンプ
}
for page, stages in STAGES.items():
    engine.set_funnel_highlight(slides[page - 1], stages)
prs.save(F)
print("saved", F)
