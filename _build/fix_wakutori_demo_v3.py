# -*- coding: utf-8 -*-
"""ワクトリのLINE施策デモ v2 → v3：提案FMTの具体施策8枚のうち、台本が無かった3つを足す（2026-10-30 殿村さん）。

    python3 _build/fix_wakutori_demo_v3.py <v2 の project.json> <出力 project.json>

  ・休眠顧客の有効化：追加から25日、ボタン・メニューを1度も押していない方に「次のシーズンの開始時期」を聞き直す
    → 答えた開始時期で合う求人カルーセル（今の診断と同じ出し分け）。台本の人：佐藤 美咲（海・3ヶ月以上先→1〜2ヶ月先）
  ・リピート施策：応募（ENTRY）から135日後に「お仕事は終わりましたか？」→「お仕事を終えました」で次のシーズンの求人
    （実際は担当が勤務実績を確認してタグを付ける。デモではご本人の回答で代わりにしている＝タグ名にも書く）
  ・コミュニティ運営：次のお仕事にも応募した方（くり返し応募）に、体験談の募集と友だちへの紹介の案内
    台本の人：高橋 蓮（冬のスキー場→夏の高原→紹介）
  ・説明（ナレーション）・道筋の「資料p.◯」を、資料v8の番号に振り直す
  書き出し：line-simulator で  node tools/export.ts client-c-resort --confirm-fictional
"""
import copy
import json
import re
import sys

DAY = 1440
PAGE_MAP = {23: 32, 24: 33, 25: 34, 37: 38, 40: 41, 27: 24, 28: 24, 29: 26, 30: 26, 31: 29, 32: 25, 33: 27}   # v6 → v8
START_TAGS = ["start_asap", "start_1_2", "start_3plus", "start_unknown"]


def node(i, typ, x, y, **kw):
    return {"id": i, "type": typ, "position": {"x": x, "y": y}, **kw}


def edge(fid, k, a, port, b):
    return {"id": f"{fid}_e{k}", "from": {"nodeId": a, "port": port}, "to": b}


def chain(fid, nodes, links):
    return [edge(fid, k + 1, a, port, b) for k, (a, port, b) in enumerate(links)]


def card(cid, title, img, place, start, work, cond, page, bid):
    return {"id": cid, "title": title, "assetId": img, "badge": "デモ用サンプル求人",
            "facts": [{"label": "勤務地", "value": place}, {"label": "開始時期", "value": start},
                      {"label": "仕事内容", "value": work}, {"label": "条件", "value": cond}],
            "buttons": [{"id": bid, "label": "求人の詳細を見る", "action": {"kind": "openPage", "pageId": page}}]}


def remap(text):
    """「資料p.23〜24・30」の番号を、v8の番号に"""
    def rep(m):
        return re.sub(r"\d+", lambda n: str(PAGE_MAP.get(int(n.group(0)), int(n.group(0)))), m.group(0))
    return re.sub(r"資料p\.[0-9・、〜]+", rep, text).replace("資料p.24〜24", "資料p.24")


def main(src, out):
    p = json.loads(remap(open(src, encoding="utf-8").read()))     # 先に、既存の説明の「資料p.◯」を振り直す
    tags = {t["id"] for t in p["tags"]}
    assert "dormant" not in tags
    p["tags"] += [
        {"id": "dormant", "label": "反応なし（追加から25日）"},
        {"id": "worked", "label": "お仕事を終えた（デモではご本人の回答）"},
        {"id": "repeater", "label": "くり返し応募"},
    ]
    M = {m["id"]: m for m in p["messages"]}
    car = {c["id"]: c for m in p["messages"] for b in m["bubbles"] if b["kind"] == "carousel" for c in b["cards"]}

    # ---------- 休眠顧客の有効化 ----------
    p["messages"] += [
        {"id": "msg_dormant", "bubbles": [
            {"kind": "buttons", "text": "お久しぶりです、ワクトリです。\n次のシーズンのお仕事、開始時期は決まりましたか？\n今の希望に合う求人だけをお届けします。",
             "buttons": [{"id": "btn_dm_1_2", "label": "1〜2ヶ月先に始めたい", "action": {"kind": "startFlow", "flowId": "flow_dm_1_2"}},
                         {"id": "btn_dm_3", "label": "3ヶ月以上先", "action": {"kind": "startFlow", "flowId": "flow_dm_3"}},
                         {"id": "btn_dm_unknown", "label": "まだわからない", "action": {"kind": "startFlow", "flowId": "flow_dm_unknown"}}]}]},
    ]
    fid = "flow_dormant"
    nodes = [node("n_dm_entry", "entry", 0, 0),
             node("n_dm_wait", "wait", 0, 150, label="追加から25日目の12:00", wait={"kind": "dayAt", "day": 25, "time": "12:00"}),
             node("n_dm_b", "branch", 0, 300, label="追加から25日、ボタン・メニューを押したか",
                  cases=[{"id": "active", "label": "押した（反応あり）", "when": {"op": "or", "children": [
                      {"op": "did", "event": e, "cmp": "exists", "withinMinutes": 25 * DAY}
                      for e in ("tapButton", "tapMenu")]}}]),
             node("n_dm_tag", "tag", 200, 450, label="反応なし", add=["dormant"]),
             node("n_dm_msg", "message", 200, 600, label="休眠：次のシーズンの開始時期を聞き直す", messageId="msg_dormant"),
             node("n_dm_end", "end", 0, 750)]
    p["flows"].append({"id": fid, "label": "休眠の掘り起こし（追加25日・反応なしの方に開始時期を聞き直す）",
                       "trigger": {"kind": "friendAdd"}, "priority": 25, "concurrency": "parallel", "reentry": "never",
                       "stopWhen": {"op": "or", "children": [{"op": "friend", "state": "blocked"},
                                                             {"op": "cv", "cvId": "cv_entry", "done": True}]},
                       "nodes": nodes,
                       "edges": chain(fid, nodes, [("n_dm_entry", "out", "n_dm_wait"), ("n_dm_wait", "out", "n_dm_b"),
                                                   ("n_dm_b", "case:active", "n_dm_end"), ("n_dm_b", "else", "n_dm_tag"),
                                                   ("n_dm_tag", "out", "n_dm_msg"), ("n_dm_msg", "out", "n_dm_end")])})
    for code, tag, label in (("1_2", "start_1_2", "1〜2ヶ月先"), ("3", "start_3plus", "3ヶ月以上先"), ("unknown", "start_unknown", "まだわからない")):
        f = f"flow_dm_{code}"
        nodes = [node(f"n_dm{code}_entry", "entry", 0, 0),
                 node(f"n_dm{code}_tag", "tag", 0, 150, label=f"開始時期を「{label}」に答え直す",
                      add=[tag], remove=[t for t in START_TAGS if t != tag] + ["dormant"]),
                 node(f"n_dm{code}_jump", "jump", 0, 300, label="答えに合う求人へ", flowId="flow_match")]
        p["flows"].append({"id": f, "label": f"休眠：開始時期「{label}」→ 合う求人へ", "trigger": {"kind": "manual"},
                           "priority": 10, "concurrency": "exclusive", "reentry": "always", "nodes": nodes,
                           "edges": chain(f, nodes, [(f"n_dm{code}_entry", "out", f"n_dm{code}_tag"),
                                                     (f"n_dm{code}_tag", "out", f"n_dm{code}_jump")])})

    # ---------- リピート施策・コミュニティ運営 ----------
    p["messages"] += [
        {"id": "msg_work_check", "bubbles": [
            {"kind": "buttons", "text": "お仕事、おつかれさまでした。\nワクトリでのお仕事は、終わりましたか？\n次のシーズンのお仕事のご案内に使います。",
             "buttons": [{"id": "btn_wc_done", "label": "お仕事を終えました", "action": {"kind": "startFlow", "flowId": "flow_worked"}},
                         {"id": "btn_wc_yet", "label": "まだ勤務中です", "action": {"kind": "sendText", "text": "まだ勤務中です"}}]}]},
        {"id": "msg_repeat_mt", "bubbles": [
            {"kind": "text", "text": "冬の山・高原でのお仕事、ありがとうございました。\n春・夏に始められる次のお仕事を、先にご案内します。"},
            {"kind": "carousel", "cards": [
                card("card_rp_m4", "高原リゾートの夏のスタッフ", "img_job_summer_highland", "長野県（高原）", "7月〜", "レストラン・売店", "寮・食事つき（例）", "page_job_m4", "b_rp_m4"),
                card("card_rp_h1", "高原リゾートホテルの客室係", "img_job_highland", "群馬県（高原）", "2月以降（春休み）", "客室の清掃・準備", "寮・食事つき（例）", "page_job_h1", "b_rp_h1"),
                card("card_rp_m1", "夏のビーチリゾートホテル", "img_job_summer_beach", "沖縄県（恩納村）", "6月〜", "プールの監視・レストラン", "寮つき（例）", "page_job_m1", "b_rp_m1")]}]},
        {"id": "msg_repeat_other", "bubbles": [
            {"kind": "text", "text": "お仕事、ありがとうございました。\n次のシーズンに始められるお仕事を、先にご案内します。"},
            {"kind": "carousel", "cards": [
                card("card_rpo_w1", "スキー場ホテルのフロント・接客", "img_job_ski_hotel", "長野県（スキー場エリア）", "12月中旬〜", "フロント受付・レストランの接客", "寮・食事つき（例）", "page_job_w1", "b_rpo_w1"),
                card("card_rpo_m4", "高原リゾートの夏のスタッフ", "img_job_summer_highland", "長野県（高原）", "7月〜", "レストラン・売店", "寮・食事つき（例）", "page_job_m4", "b_rpo_m4")]}]},
        {"id": "msg_community", "bubbles": [
            {"kind": "text", "text": "いつもワクトリをご利用いただき、ありがとうございます。"},
            {"kind": "buttons", "text": "これから働く方のために、経験者の声を集めています。\nまた、お友だちにもワクトリを紹介できます。",
             "buttons": [{"id": "btn_cm_voice", "label": "体験談を送る", "action": {"kind": "openPage", "pageId": "page_voice_post"}},
                         {"id": "btn_cm_ref", "label": "友だちに紹介する", "action": {"kind": "openPage", "pageId": "page_referral"}}]}]},
    ]
    entry_page = next(w for w in p["webPages"] if w["id"] == "page_entry")
    post = copy.deepcopy(entry_page)
    post.update({"id": "page_voice_post", "title": "体験談を送る", "url": "https://example.com/resortbaito/voice/post"})
    post["blocks"] = [{"kind": "heading", "text": "体験談を送る"},
                      {"kind": "text", "text": "働いた場所・期間と、はじめての不安が実際どうだったかを教えてください。\n（デモ用の枠です。いただいた声は、貴社の確認のうえ体験談のページに掲載します）"}]
    post["form"] = {"fields": [{"id": "v_place", "label": "働いた場所", "type": "text"},
                               {"id": "v_term", "label": "期間", "type": "text"},
                               {"id": "v_text", "label": "ひとこと", "type": "text"}],
                    "onSubmit": {}}
    for k in ("exitPopup", "lineCta"):
        post.pop(k, None)
    ref = {"id": "page_referral", "kind": "lp", "title": "お友だちに紹介する", "url": "https://example.com/resortbaito/referral",
           "blocks": [{"kind": "heading", "text": "お友だちにワクトリを紹介する"},
                      {"kind": "text", "text": "このLINEアカウントを、LINEでお友だちに送れます（LINEの転送）。"},
                      {"kind": "text", "text": "紹介の特典（例）\n内容は貴社で決めます（デモでは表示しません）"},
                      {"kind": "button", "label": "LINEで友だちに送る", "action": {"kind": "none"}}]}
    p["webPages"] += [post, ref]

    fid = "flow_after_work"
    nodes = [node("n_aw_entry", "entry", 0, 0),
             node("n_aw_wait", "wait", 0, 150, label="応募から135日目の19:00（冬のお仕事が終わるころ）",
                  wait={"kind": "dayAt", "day": 135, "time": "19:00", "from": "flowStart"}),
             node("n_aw_msg", "message", 0, 300, label="お仕事は終わりましたか？", messageId="msg_work_check"),
             node("n_aw_end", "end", 0, 450)]
    p["flows"].append({"id": fid, "label": "お仕事の後（応募135日後）：お仕事は終わりましたか？",
                       "trigger": {"kind": "cv", "cvId": "cv_entry"}, "priority": 6, "concurrency": "parallel", "reentry": "never",
                       "stopWhen": {"op": "friend", "state": "blocked"}, "nodes": nodes,
                       "edges": chain(fid, nodes, [("n_aw_entry", "out", "n_aw_wait"), ("n_aw_wait", "out", "n_aw_msg"),
                                                   ("n_aw_msg", "out", "n_aw_end")])})
    fid = "flow_worked"
    nodes = [node("n_wk_entry", "entry", 0, 0),
             node("n_wk_tag", "tag", 0, 150, label="お仕事を終えた（実際は担当が勤務実績を確認して付ける）", add=["worked"]),
             node("n_wk_b", "branch", 0, 300, label="前回のエリア",
                  cases=[{"id": "mt", "label": "山・高原", "when": {"op": "hasTag", "tagId": "area_mountain"}}]),
             node("n_wk_rep_mt", "message", -200, 450, label="次のシーズン：春・夏の高原・海", messageId="msg_repeat_mt"),
             node("n_wk_rep_other", "message", 200, 450, label="次のシーズンのお仕事", messageId="msg_repeat_other"),
             node("n_wk_wait", "wait", 0, 600, label="28日後の19:00", wait={"kind": "dayAt", "day": 28, "time": "19:00", "from": "flowStart"}),
             node("n_wk_b2", "branch", 0, 750, label="次のお仕事にも応募したか",
                  cases=[{"id": "again", "label": "くり返し応募", "when": {"op": "did", "event": "submitForm", "where": {"pageId": "page_entry"},
                                                                     "cmp": "countGte", "value": 2}}]),
             node("n_wk_tag2", "tag", -200, 900, label="くり返し応募", add=["repeater"]),
             node("n_cm_msg", "message", -200, 1050, label="コミュニティ：体験談の募集・友だちへの紹介", messageId="msg_community"),
             node("n_wk_end", "end", 0, 1200)]
    p["flows"].append({"id": fid, "label": "お仕事を終えた方：次のシーズンのお仕事 → くり返し応募の方に体験談・紹介",
                       "trigger": {"kind": "manual"}, "priority": 6, "concurrency": "parallel", "reentry": "never",
                       "stopWhen": {"op": "friend", "state": "blocked"}, "nodes": nodes,
                       "edges": chain(fid, nodes, [("n_wk_entry", "out", "n_wk_tag"), ("n_wk_tag", "out", "n_wk_b"),
                                                   ("n_wk_b", "case:mt", "n_wk_rep_mt"), ("n_wk_b", "else", "n_wk_rep_other"),
                                                   ("n_wk_rep_mt", "out", "n_wk_wait"), ("n_wk_rep_other", "out", "n_wk_wait"),
                                                   ("n_wk_wait", "out", "n_wk_b2"), ("n_wk_b2", "case:again", "n_wk_tag2"),
                                                   ("n_wk_b2", "else", "n_wk_end"), ("n_wk_tag2", "out", "n_cm_msg"),
                                                   ("n_cm_msg", "out", "n_wk_end")])})

    # ---------- 台本の人 ----------
    def step(m, ev):
        return {"atMinutes": m, "event": ev}
    ans = lambda m, q, c: step(m, {"type": "answer", "surveyId": "sv_diag", "questionId": q, "choiceIds": [c]})
    d25 = 25 * DAY + 12 * 60
    p["personas"].append({"id": "p_dormant", "name": "反応がなく25日（休眠）→ 開始時期を答え直して応募", "personName": "佐藤 美咲",
                          "attrs": {"年代": "20代", "希望": "3ヶ月以上先・海"},
                          "script": [step(720, {"type": "visitPage", "pageId": "page_top"}),
                                     step(722, {"type": "tapPageButton", "pageId": "page_top", "blockIndex": 2}),
                                     ans(724, "q_entry", "c_ent_yet"), ans(725, "q_start", "c_st_3"), ans(726, "q_area", "c_ar_sea"),
                                     step(d25 + 3, {"type": "tapButton", "buttonId": "btn_dm_1_2"}),
                                     step(d25 + 5, {"type": "viewCard", "messageId": "msg_car_sea_1_2", "cardId": next(
                                         c["id"] for c in car.values() if c["id"].startswith("card_car_sea_1_2"))}),
                                     step(d25 + 6, {"type": "tapButton", "buttonId": "b_car_sea_1_2_s1"}),
                                     step(d25 + 9, {"type": "tapPageButton", "pageId": "page_job_s1", "blockIndex": 4}),
                                     step(d25 + 12, {"type": "submitForm", "pageId": "page_entry",
                                                     "values": {"f_name": "デモ 美咲", "f_tel": "000-0000-0000", "f_start": "1〜2ヶ月先", "f_point": "立地"}})]})
    w = 135 * DAY + 19 * 60 + 10          # 応募（Day0 10:10ごろ）から135日目の19:00
    c = w + 28 * DAY
    p["personas"].append({"id": "p_repeat", "name": "お仕事の後、次のシーズンにも応募 → 紹介（リピート・コミュニティ）", "personName": "高橋 蓮",
                          "attrs": {"年代": "20代", "希望": "1〜2ヶ月先・山・高原"},
                          "script": [step(600, {"type": "visitPage", "pageId": "page_top"}),
                                     step(602, {"type": "tapPageButton", "pageId": "page_top", "blockIndex": 2}),
                                     ans(604, "q_entry", "c_ent_yet"), ans(605, "q_start", "c_st_1_2"), ans(606, "q_area", "c_ar_mt"),
                                     step(608, {"type": "tapButton", "buttonId": "b_car_mt_1_2_w1"}),
                                     step(611, {"type": "tapPageButton", "pageId": "page_job_w1", "blockIndex": 4}),
                                     step(615, {"type": "submitForm", "pageId": "page_entry",
                                                "values": {"f_name": "デモ 蓮", "f_tel": "000-0000-0000", "f_start": "1〜2ヶ月先", "f_point": "稼げる"}}),
                                     step(1500, {"type": "closePage", "pageId": "page_entry_thanks"}),
                                     step(w, {"type": "tapButton", "buttonId": "btn_wc_done"}),
                                     step(w + 2, {"type": "viewCard", "messageId": "msg_repeat_mt", "cardId": "card_rp_m4"}),
                                     step(w + 3, {"type": "tapButton", "buttonId": "b_rp_m4"}),
                                     step(w + 6, {"type": "tapPageButton", "pageId": "page_job_m4", "blockIndex": 4}),
                                     step(w + 9, {"type": "submitForm", "pageId": "page_entry",
                                                  "values": {"f_name": "デモ 蓮", "f_tel": "000-0000-0000", "f_start": "3ヶ月以上先", "f_point": "出会い多数"}}),
                                     step(w + 600, {"type": "closePage", "pageId": "page_entry_thanks"}),
                                     step(c + 30, {"type": "tapButton", "buttonId": "btn_cm_ref"})]})

    # ---------- 説明・道筋 ----------
    p["narration"] += [
        {"id": "nar_dormant", "on": {"kind": "node", "nodeId": "n_dm_msg"}, "eyebrow": "休眠顧客の有効化",
         "title": "反応が止まった方に、次のシーズンの開始時期を聞き直す",
         "body": "追加から25日、ボタン・メニューを押していない方だけに送ります。答え直した開始時期で、合う求人を出し直します（資料p.28）。"},
        {"id": "nar_worked", "on": {"kind": "node", "nodeId": "n_wk_tag"}, "eyebrow": "リピート施策",
         "title": "お仕事を終えた方に、次のシーズンのお仕事を",
         "body": "前回のエリアから、次に始められる求人を先にご案内します。実際は担当が勤務実績を確認してタグを付けます（デモではご本人の回答で代わりにしています。資料p.30）。"},
        {"id": "nar_community", "on": {"kind": "node", "nodeId": "n_cm_msg"}, "eyebrow": "コミュニティ運営",
         "title": "くり返し働く方に、体験談の募集と紹介のご案内",
         "body": "次のお仕事にも応募した方だけに送ります。体験談は、これから働く方の不安を減らす配信に使えます（資料p.31）。"},
    ]
    J = p["journey"]
    pos = next(i for i, j in enumerate(J) if j["id"] == "j_long") + 1
    J.insert(pos, {"id": "j_dormant", "label": "反応のない方の掘り起こし（追加25日）", "nodeIds": ["n_dm_msg"],
                   "purpose": "ボタン・メニューを押していない方に、次のシーズンの開始時期を聞き直し、合う求人を出し直す（資料p.28）",
                   "feasibility": "lstep", "feasibilityNote": "反応（クリック）の有無はタグ・行動の絞り込みで判定"})
    J += [{"id": "j_repeat", "label": "お仕事の後：次のシーズンのお仕事", "nodeIds": ["n_wk_rep_mt", "n_wk_rep_other"],
           "purpose": "お仕事を終えた方に、前回のエリアから次に始められる求人を先にご案内（資料p.30）",
           "feasibility": "conditional", "feasibilityNote": "勤務実績は担当が確認してタグを付ける（システム連携は前提にしない）"},
          {"id": "j_community", "label": "くり返し働く方：体験談・紹介", "nodeIds": ["n_cm_msg"],
           "purpose": "くり返し応募した方に、体験談の募集と友だちへの紹介のご案内（資料p.31）",
           "feasibility": "lstep", "feasibilityNote": "紹介はLINEの転送。特典の中身は貴社で決める"}]

    json.dump(p, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print("saved:", out)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
