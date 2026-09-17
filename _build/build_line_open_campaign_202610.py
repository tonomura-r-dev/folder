"""
LINEオープンキャンペーン（26年10-12月・期間限定特別プラン）専用の単発デッキ。

トンマナはトレンドレポートと同じ trend_report_engine.py（DYM_LINEOA_TREND_FMT.pptx
＝2026年7月号の実ファイルを複製する方式）を流用。トレンドレポート本体とは別の
単独ファイルとして、このキャンペーンだけを説明する3ページ構成にする。
中身（特典内容／実施スケジュール／注意点）は build_trend_report_202609.py の
⑥LINEオープンキャンペーンと同じデータを使う。

使い方: python3 _build/build_line_open_campaign_202610.py
"""

from pathlib import Path
import trend_report_engine as engine

ROOT = Path(__file__).resolve().parent.parent
OUT = str(ROOT / "LINEオープンキャンペーン_26年10-12月_ご案内.pptx")
COVER_TITLE = "LINEオープンキャンペーン　26年10-12月　ご案内"

TOPICS = [
    dict(
        bracket="LINEオープンキャンペーン",
        headline="特典内容",
        industry="友だち獲得を強化したい全業種",
        category="オプション\n商材",
        description_lines=[
            "2026年10月20日開始〜12月25日までに終了する案件限定で、",
            "基本費用・掲載期間・配信通数の3点が特別条件になるキャンペーン。",
            "友だち獲得を一気に伸ばしたいクライアントへの提案材料になる。",
        ],
        render="before_after",
        as_is="通常プラン｜基本費用1,800万円・掲載期間7日間・配信900万通",
        to_be_lines=["特別プラン｜基本費用1,500万円", "掲載期間15日間・配信1,050万通に拡大"],
        merits=[
            ("特別価格で提案しやすい",
             "基本費用が1,800万円→1,500万円に。期間限定の特別プランとして提案の後押しになる。"),
            ("景品はLINEポイント2pt総付",
             "応募者全員にLINEポイント2ptを進呈するアンケート型キャンペーン。対象はLINE公式アカウント・Messaging API対応かつ認証プロバイダー必須。"),
        ],
        source="https://workers-hub.box.com/s/odilqpe44tp398zlfko148yggcgj68ny",
    ),
    dict(
        bracket="LINEオープンキャンペーン",
        headline="実施スケジュール",
        industry="友だち獲得を強化したい全業種",
        category="オプション\n商材",
        description_lines=[
            "申込みから掲載開始まで、審査・枠おさえ・入稿の3段階を経る。",
            "配信枠は早い者勝ちのため、各締切の厳守が前提条件になる。",
            "土日祝日の掲載開始は不可なので、平日での日程調整が必要。",
        ],
        render="flow",
        steps=[
            ("30営業日前", "企業・商材審査／企画審査"),
            ("25営業日前 17時", "配信枠おさえ・申込み"),
            ("20営業日前 17時", "入稿物の提出"),
            ("開始日", "掲載開始（土日祝日不可）"),
        ],
        note_headline="早めの申込みが必須",
        note_desc="期間・条件限定のキャンペーンのため、早めの提案・申込み誘導がカギ。配信枠は先着で埋まるため、審査期間も見込んだ逆算スケジュールを組む。",
        source="https://workers-hub.box.com/s/odilqpe44tp398zlfko148yggcgj68ny",
    ),
    dict(
        bracket="LINEオープンキャンペーン",
        headline="注意点・ご利用条件",
        industry="友だち獲得を強化したい全業種",
        category="オプション\n商材",
        description_lines=[
            "キャンセル規定・割引の併用可否など、提案前に必ず確認したい4点。",
            "特にキャンセル規定は発注後の費用に直結するため要注意。",
            "景品内容は変更となる場合があり、事前の合意形成が欠かせない。",
        ],
        render="grid_notes",
        items=[
            ("キャンセル規定", "発注期日超過後のキャンセルは、最低発注金額1,500万円を請求。"),
            ("他割引との併用不可", "既存の各種割引施策とは併用不可。特別プラン単体での提案が前提。"),
            ("景品変更の可能性あり", "景品はLINEポイント2pt固定だが変更となる場合があり、変更不可のお申し出は受付不可。"),
            ("対象は認証プロバイダー必須", "LINE公式アカウント・Messaging API対応かつ認証プロバイダーのアカウントが対象。"),
        ],
        source="https://workers-hub.box.com/s/odilqpe44tp398zlfko148yggcgj68ny",
    ),
]

for t in TOPICS:
    t["title"] = f"【{t['bracket']}】{t['headline']}"


if __name__ == "__main__":
    engine.build_standalone(OUT, COVER_TITLE, TOPICS)
