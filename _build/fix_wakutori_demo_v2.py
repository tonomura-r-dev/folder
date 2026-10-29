# -*- coding: utf-8 -*-
"""ワクトリのLINE施策デモ（line-simulator の案件 client-c-resort）を直す（2026-10-29 殿村さん）。

    python3 _build/fix_wakutori_demo_v2.py <元の project.json> <出力 project.json>

  1. 成果を1つにする：ENTRY（会員登録）＝求人への応募（資料・SIMと同じ。同じ登録フォーム）。
     「応募（個別の求人）」の成果・応募フォーム17枚・応募の完了画面・「ENTRYと応募の違い」を外し、
     求人ページの「この求人に応募する」は ENTRY のフォームへ。成果の名前は「応募（ENTRY）」。
     応募（ENTRY）したら、ステップ配信・特集は止まる（担当エージェントが引き継ぐ）。
  2. 診断のエリアの選択肢を、今のアンケートから「離島」「市街地」を外した4つにする：海／山・高原／温泉地／湖。
     「湖」の求人はデモ用サンプルを2件足す（画像は既存の代替画像を使う）。「離島」の付いたサンプル求人は「海辺」に言い換える。
  書き出し：line-simulator で  node tools/export.ts client-c-resort --confirm-fictional
"""
import copy
import json
import re
import sys


def main(src, out):
    p = json.load(open(src, encoding="utf-8"))

    # ---------- 2. エリア：文字の言い換え（IDは変えない。area_any だけ area_lake に） ----------
    raw = json.dumps(p, ensure_ascii=False)
    for a, b in [("area_any", "area_lake"), ("c_ar_any", "c_ar_lake"),
                 ("山・高原（スキー場など）", "山・高原"), ("海・離島", "海"), ("温泉・観光地", "温泉地"),
                 ("沖縄・離島", "沖縄"), ("離島リゾートホテル", "海辺のリゾートホテル"), ("離島の旅館", "海辺の旅館"),
                 ("離島の民宿", "海辺の民宿"), ("沖縄県（離島）", "沖縄県（恩納村）"), ("鹿児島県（離島）", "鹿児島県（海沿い）"),
                 ("エリアこだわらず", "エリア未回答"), ("まだわからない・こだわらない", "まだわからない・湖")]:
        raw = raw.replace(a, b)
    p = json.loads(raw)

    q = p["surveys"][0]["questions"][2]
    ch = {c["id"]: c for c in q["choices"]}
    ch["c_ar_lake"]["label"] = "湖"
    q["choices"] = [ch["c_ar_sea"], ch["c_ar_mt"], ch["c_ar_onsen"], ch["c_ar_lake"]]   # 今のアンケートの並び
    tags = {t["id"]: t for t in p["tags"]}
    tags["area_lake"]["label"] = "エリア：湖"

    # 湖のサンプル求人（ページ・応募の流れ・カード）
    pages = {w["id"]: w for w in p["webPages"]}
    lake = [("l1", "湖畔のリゾートホテルの客室係", "山梨県（河口湖）", "11月〜", "客室の清掃・準備", "寮つき（例）", "img_job_highland"),
            ("l2", "湖畔の旅館の仲居・配膳", "北海道（洞爺湖）", "12月〜", "お部屋への配膳と接客", "寮つき・食事つき（例）", "img_job_onsen")]
    flows = {f["id"]: f for f in p["flows"]}
    for code, title, place, start, work, cond, img in lake:
        pg = copy.deepcopy(pages["page_job_o1"])
        pg["id"], pg["title"], pg["url"] = f"page_job_{code}", f"求人：{title}", f"https://example.com/resortbaito/job/{code}"
        pg["blocks"][0]["assetId"] = img
        pg["blocks"][1]["text"] = title
        pg["blocks"][3]["text"] = (f"勤務地：{place}\n開始時期：{start}\n期間：3か月〜\n仕事内容：{work}\n条件：{cond}\n"
                                   "給与：実際の求人で表示（デモでは表示しません）")
        for b in pg["blocks"]:
            if b.get("action", {}).get("flowId") == "flow_pick_o1":
                b["action"]["flowId"] = f"flow_pick_{code}"
        p["webPages"].insert(p["webPages"].index(pages["page_job_m4"]) + 1, pg)
        f = json.loads(json.dumps(flows["flow_pick_o1"], ensure_ascii=False).replace("_o1", f"_{code}"))
        f["label"] = f["label"].replace("温泉旅館の仲居・配膳", title)
        p["flows"].insert(p["flows"].index(flows["flow_pick_m4"]) + 1, f)
        pages["page_search"]["blocks"].append({"kind": "button", "label": f"{title}｜{place}・{start}",
                                               "action": {"kind": "openPage", "pageId": f"page_job_{code}"}})
    p["tags"].insert(p["tags"].index(tags["want_m4"]) + 1, {"id": "want_l1", "label": "応募したい求人：湖畔のリゾートホテルの客室係", "group": "応募したい求人"})
    p["tags"].insert(p["tags"].index(tags["want_m4"]) + 2, {"id": "want_l2", "label": "応募したい求人：湖畔の旅館の仲居・配膳", "group": "応募したい求人"})

    def lake_cards(prefix, with_apply):
        cards = []
        for code, title, place, start, work, cond, img in lake:
            c = {"id": f"card_{prefix}_{code}", "title": title, "assetId": img, "badge": "デモ用サンプル求人",
                 "facts": [{"label": "勤務地", "value": place}, {"label": "開始時期", "value": start},
                           {"label": "仕事内容", "value": work}, {"label": "条件", "value": cond}],
                 "buttons": [{"id": f"b_{prefix}_{code}", "label": "求人の詳細を見る", "action": {"kind": "openPage", "pageId": f"page_job_{code}"}}]}
            cards.append(c)
        return cards

    msgs = {m["id"]: m for m in p["messages"]}

    def lake_msg(new_id, base_id, prefix, text=None):
        m = copy.deepcopy(msgs[base_id])
        m["id"] = new_id
        for b in m["bubbles"]:
            if b["kind"] == "carousel":
                b["cards"] = lake_cards(prefix, False)
            elif b["kind"] == "text" and text and "{{" not in b["text"]:
                pass
        if text:
            m["bubbles"][0]["text"] = text
        for bb in m["bubbles"]:                                       # ボタンのIDが元のメッセージと重ならないように
            for x in bb.get("buttons", []):
                x["id"] = f"b_{prefix}_{x['id']}"
        p["messages"].insert(p["messages"].index(msgs[base_id]) + 1, m)
        msgs[new_id] = m
        return m

    lake_msg("msg_car_lake", "msg_car_onsen_unknown", "car_lake",
             "回答ありがとうございます！\nご希望（湖）に合う求人です。\n左右に送って、気になる求人をチェックしてみてください。")
    lake_msg("msg_agent_lake", "msg_agent_onsen_unknown", "agent_lake")
    lake_msg("msg_s4_lake", flows["flow_step"]["nodes"][[n["id"] for n in flows["flow_step"]["nodes"]].index("n_s4_onsen_unknown")]["messageId"], "s4_lake")
    lake_msg("msg_k_lake", "msg_k_onsen_1_2", "k_lake", "＼ LINE友だちへの特集 ／\n10〜12月から働ける、湖のお仕事を集めました。")
    for mid in ("msg_agent_lake", "msg_s4_lake"):                     # 「ご希望（温泉地…）」の書き出しを湖に
        for b in msgs[mid]["bubbles"]:
            if b["kind"] in ("text", "buttons") and "温泉地" in b.get("text", ""):
                b["text"] = b["text"].replace("温泉地", "湖")

    def add_case(fid, branch_id, node, cond, edge_to=None, first=False):
        f = flows[fid]
        br = next(n for n in f["nodes"] if n["id"] == branch_id)
        case = {"id": "lake", "label": "湖", "when": cond}
        if first:
            br["cases"].insert(0, case)
        else:
            br["cases"].append(case)
        if node is not None:
            f["nodes"].append(node)
            f["edges"].append({"id": f"e_{node['id']}_in", "from": {"nodeId": branch_id, "port": "case:lake"}, "to": node["id"]})
            f["edges"].append({"id": f"e_{node['id']}_out", "from": {"nodeId": node["id"], "port": "out"}, "to": edge_to})
        else:
            f["edges"].append({"id": f"e_{branch_id}_lake", "from": {"nodeId": branch_id, "port": "case:lake"}, "to": edge_to})

    lake_tag = {"op": "hasTag", "tagId": "area_lake"}
    add_case("flow_match", "n_m_a", {"id": "n_m_lake", "type": "message", "label": "湖：湖畔のリゾートホテルの客室係／湖畔の旅館の仲居・配膳",
                                     "messageId": "msg_car_lake", "position": {"x": 1600, "y": 300}}, lake_tag, "n_m_end")
    add_case("flow_entry", "n_e_a", {"id": "n_e_lake", "type": "message", "label": "湖：湖畔のリゾートホテルの客室係／湖畔の旅館の仲居・配膳",
                                     "messageId": "msg_agent_lake", "position": {"x": 1600, "y": 900}}, lake_tag, "n_e_w2")
    add_case("flow_step", "n_s4_a", {"id": "n_s4_lake", "type": "message", "label": "湖：ENTRYの案内＋湖畔の求人",
                                     "messageId": "msg_s4_lake", "position": {"x": 2300, "y": 1500}}, lake_tag, "n_s_end")
    near = {"op": "and", "children": [lake_tag, {"op": "or", "children": [{"op": "hasTag", "tagId": "start_asap"}, {"op": "hasTag", "tagId": "start_1_2"}]}]}
    add_case("flow_seg", "n_k_a", {"id": "n_k_lake", "type": "message", "label": "湖（開始が近い方）：湖畔のリゾートホテルの客室係／湖畔の旅館の仲居・配膳",
                                   "messageId": "msg_k_lake", "position": {"x": 2000, "y": 450}}, near, "n_k_end")
    br = next(n for n in flows["flow_seg"]["nodes"] if n["id"] == "n_k_a")
    br["cases"].append({"id": "lake_far", "label": "湖（開始が先・未定）", "when": lake_tag})
    flows["flow_seg"]["edges"].append({"id": "e_n_k_a_lake_far", "from": {"nodeId": "n_k_a", "port": "case:lake_far"}, "to": "n_k_skip"})

    # ---------- 1. 成果を1つにする（ENTRY＝応募） ----------
    p["cvs"] = [c for c in p["cvs"] if c["id"] != "cv_apply"]
    p["cvs"][0].update({"label": "応募（ENTRY）", "demoAction": "応募後へ"})
    p["settings"]["labels"] = {"cv_entry": "応募"}
    drop_pages = {w["id"] for w in p["webPages"] if w["id"].startswith("page_apply_")} | {"page_about_entry"}
    p["webPages"] = [w for w in p["webPages"] if w["id"] not in drop_pages]
    p["flows"] = [f for f in p["flows"] if f["id"] not in ("flow_apply", "flow_back_apply")]
    p["messages"] = [m for m in p["messages"] if m["id"] not in ("msg_apply_thanks", "msg_s4_apply")]
    p["tags"] = [t for t in p["tags"] if t["id"] != "applied"]
    pages = {w["id"]: w for w in p["webPages"]}
    flows = {f["id"]: f for f in p["flows"]}
    msgs = {m["id"]: m for m in p["messages"]}

    for w in p["webPages"]:                                            # 求人ページ：ENTRYと応募の2つのボタンを1つに
        if w["id"].startswith("page_job_"):
            bl = w["blocks"]
            assert bl[4]["label"].startswith("かんたん1分 ENTRY") and bl[5]["label"] == "この求人に応募する", w["id"]
            bl[4]["label"] = "この求人に応募する（かんたん1分 ENTRY）"
            del bl[5]
    ent = pages["page_entry"]
    ent["title"] = "かんたん1分 ENTRY（応募）"
    ent["blocks"][0]["text"] = "かんたん1分 ENTRY（応募）"
    ent["blocks"][2]["text"] = "このENTRY（会員登録）が、求人への応募になります。選んだ求人は、専属エージェントにお伝えします。"
    th = pages["page_entry_thanks"]
    th["title"] = "応募（ENTRY）の完了"
    th["blocks"] = [b for b in th["blocks"] if b.get("action", {}).get("flowId") != "flow_back_apply"]
    th["blocks"][0]["text"] = "応募（ENTRY）を受け付けました"
    th["blocks"][1]["text"] = "専属エージェントから、面談とお仕事開始までの流れをご連絡します。（デモ）"

    for f in p["flows"]:                                               # 求人を選ぶ → ENTRY（応募）のフォームへ
        if f["id"].startswith("flow_pick_"):
            name = f["label"].split("：", 1)[1].split(" → ")[0]
            f["label"] = f"求人を選ぶ：{name} → 応募（ENTRY）のフォームへ"
            tag = next(n for n in f["nodes"] if n["type"] == "tag")
            form = next(n for n in f["nodes"] if n["type"] == "page" and n["pageId"] == "page_entry")
            form["label"] = "応募（ENTRY）のフォーム"
            end = next(n for n in f["nodes"] if n["type"] == "end")
            f["nodes"] = [n for n in f["nodes"] if n["type"] in ("entry", "tag", "end") or n is form]
            entry = next(n for n in f["nodes"] if n["type"] == "entry")
            f["edges"] = [{"id": f"{f['id']}_e1", "from": {"nodeId": entry["id"], "port": "out"}, "to": tag["id"]},
                          {"id": f"{f['id']}_e2", "from": {"nodeId": tag["id"], "port": "out"}, "to": form["id"]},
                          {"id": f"{f['id']}_e3", "from": {"nodeId": form["id"], "port": "out"}, "to": end["id"]}]
            tag["remove"] = [t for t in tag["remove"] if t != tag["add"][0]] + [t for t in ("want_l1", "want_l2") if t not in tag["add"] and t not in tag["remove"]]
    for f in p["flows"]:
        if f["id"].startswith("flow_pick_"):
            tag = next(n for n in f["nodes"] if n["type"] == "tag")
            allw = [t["id"] for t in p["tags"] if t["id"].startswith("want_")]
            tag["remove"] = [t for t in allw if t not in tag["add"]]

    # 止まる条件：応募（ENTRY）で止まる
    for fid in ("flow_step", "flow_seg"):
        flows[fid]["stopWhen"] = {"op": "cv", "cvId": "cv_entry", "done": True}
    flows["flow_entry"]["stopWhen"] = {"op": "friend", "state": "blocked"}   # 応募で止める条件は外す（応募で始まる流れなので）
    flows["flow_entry"]["label"] = "応募（ENTRY）→ 受付・担当のおすすめ求人 → フォロー"
    # ステップ配信 Day10：ENTRY済みの分岐を外す（応募で止まるので届かない）
    st = flows["flow_step"]
    b = next(n for n in st["nodes"] if n["id"] == "n_s4_b")
    into = [e for e in st["edges"] if e["to"] == "n_s4_b"]
    st["nodes"] = [n for n in st["nodes"] if n["id"] not in ("n_s4_b", "n_s4_apply")]
    st["edges"] = [e for e in st["edges"] if e["from"]["nodeId"] not in ("n_s4_b", "n_s4_apply") and e["to"] not in ("n_s4_b", "n_s4_apply")]
    for e in into:
        e["to"] = "n_s4_a"
        st["edges"].append(e)
    st["label"] = st["label"].replace("Day1〜10", "Day1〜10・応募で止まる")
    # 中長期：応募済み＝応募（ENTRY）済み
    raw = json.dumps(p, ensure_ascii=False).replace('"cvId": "cv_apply"', '"cvId": "cv_entry"')
    p = json.loads(raw)
    flows = {f["id"]: f for f in p["flows"]}
    msgs = {m["id"]: m for m in p["messages"]}

    # ENTRY後のおすすめ求人：カードの「応募する」ボタンを外す（すでに応募済みのため。詳細を見るは残す）
    for m in p["messages"]:
        for bb in m["bubbles"]:
            for c in bb.get("cards", []):
                c["buttons"] = [x for x in c.get("buttons", []) if not x["id"].endswith("_ap")]
            if "buttons" in bb:
                bb["buttons"] = [x for x in bb["buttons"] if x.get("action", {}).get("pageId") != "page_about_entry"]

    # 言葉：ENTRY（会員登録）→ 応募（ENTRY）（成果・流れの名前）
    for c in p["cvs"]:
        pass
    for x in p["personas"]:
        x["name"] = x["name"].replace("カルーセルからENTRY", "カルーセルから応募（ENTRY）").replace("ENTRY済みで友だちに", "応募（ENTRY）してから友だちに")
        sc = []
        for e in x["script"]:
            ev = e["event"]
            if ev.get("pageId", "").startswith("page_apply_") or ev.get("buttonId", "").endswith("_ap") or ev.get("buttonId") == "btn_s4ap_search":
                continue
            if ev.get("type") == "tapPageButton" and ev.get("pageId", "").startswith("page_job_") and ev["blockIndex"] > 5:
                ev["blockIndex"] -= 1
            sc.append(e)
        x["script"] = sc
    pe = next(x for x in p["personas"] if x["id"] == "p_entered")      # 応募（ENTRY）してから友だち：その後は担当のおすすめを見る
    pe["script"] = [e for e in pe["script"] if not (e["event"].get("pageId") in ("page_search", "page_job_o1") and e["event"]["type"] == "tapPageButton")]
    pe["script"].append({"atMinutes": 2700, "event": {"type": "viewCard", "messageId": "msg_agent_onsen_asap",
                                                     "cardId": msgs["msg_agent_onsen_asap"]["bubbles"][-1]["cards"][0]["id"]
                                                     if msgs["msg_agent_onsen_asap"]["bubbles"][-1]["kind"] == "carousel"
                                                     else next(b for b in msgs["msg_agent_onsen_asap"]["bubbles"] if b["kind"] == "carousel")["cards"][0]["id"]}})
    pe["script"].sort(key=lambda e: e["atMinutes"])

    for d in p["settings"]["demoEntries"]:
        d["label"] = d["label"].replace("ENTRY（会員登録）から開始", "応募（ENTRY）から開始")
    # 道筋（③）と説明
    J = {j["id"]: j for j in p["journey"]}
    p["journey"] = [j for j in p["journey"] if j["id"] != "j_apply"]
    J["j_entry"].update({"label": "応募（ENTRY）",
                         "purpose": "かんたん1分のENTRY（会員登録）が、そのまま求人への応募（資料の概算お見積りの「応募」と同じ）。専属エージェントが付く（資料p.24・32）。応募したら、求人を案内する配信（ステップ・特集）は止まり、担当が引き継ぐ"})
    J["j_agent"]["label"] = "応募（ENTRY）後：担当のおすすめ求人"
    J["j_agent"]["purpose"] = "応募（ENTRY）の受付と、エリア・開始時期に合う求人をカルーセルで。合う求人が無いときは担当が探す案内。2日後に「気になる求人はありましたか？」（デモ用仮設定）"
    J["j_detail"]["purpose"] = J["j_detail"]["purpose"].replace("→ ENTRY」", "→ 応募（ENTRY）」")
    J["j_s4"]["purpose"] = "応募（ENTRY）がまだの方に「かんたん1分 ENTRY」と、開始時期・エリアに合う求人を届ける（資料p.24）。応募した方には届かない（ステップ配信は応募で止まる）"
    J["j_diag"]["purpose"] = "今のアンケート3問（エントリー状況・開始時期・エリア）を生かし、答えでタグを付ける（資料p.29）。エリアは今の選択肢から「離島」「市街地」を外した4つ（海・山・高原・温泉地・湖）"
    for jid, node in (("j_match", "n_m_lake"), ("j_seg", "n_k_lake")):
        for v in J[jid]["variants"]:
            if "area_lake" in json.dumps(v["when"]):
                v["nodeIds"] = [node]
                v["label"] = "湖" if jid == "j_match" else "近い×湖"
                v["title"] = "湖の求人カルーセル" if jid == "j_match" else "湖の特集"
                v["purpose"] = "湖畔のリゾートホテルの客室係／湖畔の旅館の仲居・配膳（デモ用サンプル求人）"
                if jid == "j_seg":
                    v["when"] = near
    for a_, b_ in (("ENTRYありがとうございます", "ご応募（ENTRY）ありがとうございます"),
                   ("気になる求人は、詳細を見てそのまま応募できます。", "気になる求人は、詳細を見て担当にご相談ください。"),
                   ("ご希望（湖・開始時期は未定）", "ご希望（湖）")):
        p = json.loads(json.dumps(p, ensure_ascii=False).replace(a_, b_))
    J = {j["id"]: j for j in p["journey"]}
    next(t_ for t_ in p["tags"] if t_["id"] == "entry_done")["label"] = "応募（ENTRY）済み"
    J["j_s4"]["variants"] = [v for v in J["j_s4"].get("variants", []) if "n_s4_apply" not in v.get("nodeIds", [])]
    raw = json.dumps(p, ensure_ascii=False).replace("エリア・開始時期にこだわらない方へ", "エリア・開始時期がまだ決まっていない方へ")
    p = json.loads(raw)
    for nar in p["narration"]:
        if nar["id"] == "nar_s4":
            nar.update({"on": {"kind": "node", "nodeId": "n_s4_a"}, "title": "応募（ENTRY）がまだの方に、合う求人とENTRYを",
                        "body": "応募（ENTRY）した方には届きません。ステップ配信は応募で止まり、担当エージェントが引き継ぎます（資料p.24）。"})

        if nar["id"] == "nar_apply":
            nar.update({"on": {"kind": "node", "nodeId": "n_e_tag"}, "eyebrow": "応募（ENTRY）＝GOAL",
                        "title": "応募したら、担当エージェントが引き継ぐ",
                        "body": "ENTRY（会員登録）が、そのまま求人への応募です（資料の「応募」と同じ）。応募した方には、ステップ配信・特集を止め、担当エージェントが面談とお仕事開始までを伴走します。"})

    json.dump(p, open(out, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    left = re.findall(r".{0,20}(?:離島|こだわら|cv_apply|page_apply|ENTRYと応募).{0,20}", json.dumps(p, ensure_ascii=False))
    print("saved:", out, "残り:", len(left))
    for s in left[:40]:
        print("  ", s)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
