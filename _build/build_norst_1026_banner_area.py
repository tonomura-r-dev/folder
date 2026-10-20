"""ノースト 10月3本目（10/26）のエリア別バナー（1040×1040）。v9の実績の3枚目の札だけをエリアの院数に変える。

2026-10-20 殿村さん「同じ配信内容で、セグメント（地域分け）で切って、同時にめっちゃ配信してもいい」
→ 全国を6つのエリアに分け、同じ中身で「近くの院」だけ変えた6本を同時に送る。
配信先はOAMの属性（地域）。院名は公式の「全国一覧から探す」（2026-10-20 殿村さんのスクショ）に合わせた。開院準備中（奈良・松山・熊本）と移転準備中（豊田）は載せない。
三重県は公式どおり四日市院の東海に、新潟県は甲信越に入れた。エリアの定義・人数の根拠は `_data/norst/clinics_list_20261020.md` と
`_data/norst/friend_demographic_area_20261020.csv`（`_drafts/東京ノーストクリニック_数値分析_2026-10.md`）。

使い方:
  python3 _build/build_norst_1026_banner_area.py            # 6本まとめて
  python3 _build/build_norst_1026_banner_area.py kanto      # 1本だけ
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import build_norst_1023_banner as base  # noqa: E402
from build_norst_1019_rough import font, tw, th, text, text_c, rrect, ORANGE, GREEN_DARK, WHITE  # noqa: E402

ROOT = Path(__file__).resolve().parent.parent

# key: (札の上の言葉, 院数, 都道府県, 院名の行（テキスト用・1行14字以内）, utm)
AREAS = {
    "kanto": ("東京・神奈川・埼玉・千葉", 8, "東京都・神奈川県・埼玉県・千葉県",
              ["上野／新宿／渋谷", "町田／八王子", "横浜／大宮／千葉"], "20261026_Alone_Kanto"),
    "kitakanto": ("茨城・栃木・群馬", 3, "茨城県・栃木県・群馬県",
                  ["水戸／宇都宮／高崎"], "20261026_Alone_KitaKanto"),
    "kinki": ("近畿", 4, "大阪府・兵庫県・京都府・奈良県・滋賀県・和歌山県",
              ["大阪／梅田", "京都／姫路"], "20261026_Alone_Kinki"),
    "tokai": ("東海・北陸・甲信越", 8, "愛知県・静岡県・岐阜県・三重県・石川県・新潟県・富山県・福井県・山梨県・長野県",
              ["名古屋／静岡／岐阜", "四日市／金沢／新潟", "甲府／長野"], "20261026_Alone_Tokai"),
    "hokkaido_tohoku": ("北海道・東北", 5, "北海道・宮城県・福島県・岩手県・秋田県・山形県・青森県",
                        ["札幌／函館／盛岡", "仙台／郡山"], "20261026_Alone_Hokkaido"),
    "nishi": ("中国・四国・九州", 7, "広島県・岡山県・山口県・島根県・鳥取県・香川県・愛媛県・徳島県・高知県・福岡県・長崎県・熊本県・鹿児島県・宮崎県・大分県・佐賀県・沖縄県",
              ["広島／岡山／松江", "高松／福岡／長崎", "鹿児島"], "20261026_Alone_Nishi"),
}


def stat_card_fit(d, x1, y1, x2, y2, label, num, unit):
    """base.stat_card と同じ札。上の言葉が長いときだけ文字を小さくして札の幅に収める（28→22px）。"""
    rrect(d, (x1 + 5, y1 + 5, x2 + 5, y2 + 5), 18, fill=GREEN_DARK)  # 影
    rrect(d, (x1, y1, x2, y2), 18, fill=WHITE)
    cx = (x1 + x2) / 2
    size = 28
    while tw(font(size), label) > (x2 - x1) - 20 and size > 20:
        size -= 1
    f_l, f_n, f_u = font(size), font(104), font(44)
    text_c(d, label, f_l, cx, y1 + 16 + (28 - size) / 2, GREEN_DARK)
    w = tw(f_n, num) + tw(f_u, unit) + 4
    x = cx - w / 2
    top = y1 + 58
    text(d, num, f_n, x, top, ORANGE)
    text(d, unit, f_u, x + tw(f_n, num) + 4, top + th(f_n, num) - th(f_u, unit), GREEN_DARK)


def build(key, out=None):
    label, n, _, _, _ = AREAS[key]
    out = out or ROOT / "_images" / f"norst_1026_banner_{key}_1040.jpg"
    orig = base.stat_card

    def stat_card(d, x1, y1, x2, y2, lab, num, unit):
        if lab == "全国に":
            lab, num = label, str(n)
        stat_card_fit(d, x1, y1, x2, y2, lab, num, unit)

    base.stat_card = stat_card
    try:
        base.build(str(out))
    finally:
        base.stat_card = orig
    return out


if __name__ == "__main__":
    keys = sys.argv[1:] or list(AREAS)
    for k in keys:
        print(build(k))
