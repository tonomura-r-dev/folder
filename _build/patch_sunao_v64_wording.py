# -*- coding: utf-8 -*-
"""すなお ver6.3 → ver6.4：本編の日本語の直し（2026-10-21 殿村さん「直してほしい」）。

    python -m _build.patch_sunao_v64_wording <ver6.3.pptx> <ver6.4.pptx>

  言い切りでない文末（〜可能です／〜ございます）、社内の略語・言い方、言葉の抜け、保証に読める言い方を直す。
  段落の文字がそのまま一致したところだけを置き換える（書式はその段落の最初の文字の書式）。Appendix は触らない。
"""
import sys

from pptx import Presentation
from pptx.oxml.ns import qn

EDITS = {  # ページ番号（ver6.3）: [(元の段落, 新しい段落)]
    13: [("友だち獲得を増やし、ブロック率を抑えながら、配信可能な人数を積み上げる",
          "友だちを増やし、ブロック率を抑えながら、配信可能な人数を積み上げる")],
    17: [("⑦  (成約)", "⑦ 入会"),
         ("※個別チャット対応は効率化のみで基本貴社対応", "※個別チャットは貴社で対応（自動応答などの効率化は弊社で設定）")],
    18: [("トラック分岐の設計思想", "出し分けの考え方")],
    24: [(" ☑ 複数のLINE流入経路からステップ配信を構築し、CV数を増加。 ", " ☑ 流入経路ごとにステップ配信を構築し、CV数の増加を図る。 "),
         (" に基づいて、契約における条件の調整が可能です。", " に基づいて、契約における条件の調整が可能。")],
    27: [("LINE以外の注力動線への誘導も同バナーから実施が可能です。", "LINE以外の注力動線への誘導も同バナーから実施可能。")],
    28: [("LINE以外の注力動線への誘導も同バナーから実施が可能です。", "LINE以外の注力動線への誘導も同バナーから実施可能。"),
         ("簡易SIMへ", "セミナーLPへ")],
    30: [("LINE友だち追加～15日を境に大きくアクティブ率が下がる（ブロック率が上がる）",
          "友だち追加から15日ほどで反応が下がり、ブロックも増えやすい。"),
         ("為、段階に合わせて配信を変え、イベントの無い時期は回数を絞る。",
          "そのため、段階に合わせて配信を変え、イベントのない時期は回数を絞る。")],
    32: [("熱が高いうちに、サロンをご案内", "関心が高いうちに、サロンをご案内"),
         ("谷岡さんの経歴：契約件数全国1位で表彰・週2日の営業で毎年TOT", "谷岡様の経歴：契約件数全国1位で表彰・週2日の営業で毎年TOT")],
    36: [("リス・SNS運用", "広告・SNS運用")],
    38: [("※契約期間は6か月。貴社作業・確認にかかる時間により前後する可能性がございます。",
          "※契約期間は6か月。貴社作業・確認にかかる時間により前後する場合がある。")],
}


def ptext(p):
    return "".join(t.text or "" for t in p.iter(qn("a:t")))


def set_para(p, text):
    runs = p.findall(qn("a:r"))
    for r in runs[1:]:
        p.remove(r)
    runs[0].find(qn("a:t")).text = text


def main(src, out):
    prs = Presentation(src)
    slides = list(prs.slides)
    assert len(slides) == 59
    for no, pairs in EDITS.items():
        paras = list(slides[no - 1]._element.iter(qn("a:p")))
        for old, new in pairs:
            hits = [p for p in paras if ptext(p) == old]
            assert hits, (no, old)
            for p in hits:
                set_para(p, new)
    prs.save(out)
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
