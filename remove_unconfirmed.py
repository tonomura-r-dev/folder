import io

path = "使用済み企業リスト.txt"
with io.open(path, "r", encoding="utf-8") as f:
    lines = f.readlines()

remove_set = {
"株式会社水野（まねきや）",
"株式会社銀蔵",
"グローバルトレード株式会社（東京ぶらんど）",
"株式会社新岐阜商会（新岐阜切手古銭商会）",
"いわの美術株式会社",
"株式会社楽器堂（楽器堂OPUS）",
"株式会社フジヤカメラ店",
"株式会社GKファクトリー（ゴルフキング）",
"株式会社じゃんぱら",
"OMO株式会社（工具UP）",
"株式会社タックルベリー",
"株式会社VOLCA（マウンテンシティ）",
"株式会社交通趣味ギャラリー",
"株式会社ナガツマ（バイクボーイ）",
"株式会社ネクストビート（保育士バンク！）",
"クックビズ株式会社",
"株式会社ニッソーネット（ほいく畑）",
"ロジHR株式会社（ロジリク）",
"株式会社ビューティースリー（C3 シースリー）",
"医療法人社団あおばクリニック",
"FAVORIX BEAUTY株式会社（GINZA BLV）",
"株式会社ソシエ・ワールド（ソシエ）",
"医療法人社団光芒会（ジェニークリニック）",
"株式会社PMKメディカルラボ（PMK）",
"弁護士法人グレイス",
"弁護士法人東京新宿法律事務所",
"円満相続税理士法人",
"弁護士法人エース",
"弁護士法人春田法律事務所",
"司法書士法人NCP",
"杠司法書士法人",
"税理士法人朝日中央綜合事務所",
"弁護士法人小杉法律事務所",
"株式会社アネストワン",
"株式会社シアーズホーム（ジャストホーム）",
"株式会社フレッシュハウス",
"Areti株式会社",
"DINETTE株式会社（PHOEBE BEAUTY UP）",
"ドクターリセラ株式会社",
"株式会社長寿乃里（然-しかり-よかせっけん）",
"合同会社Play Works",
"有限会社アーバンワールド（アーバンマリッジ）",
"株式会社ブライダルドリーム（DUO BRIDAL）",
"株式会社ウィルコミュニケーションズ（ホワイトキー）",
"アイマップス株式会社（婚活NANA）",
"株式会社Realing（リアリングエージェント）",
"株式会社Real Standard（コスラボ）",
"株式会社STAYGOLD",
"株式会社山徳（ビーレコーズ）",
"株式会社美術刀剣みらい",
"株式会社トリアイナ（はれや）",
"株式会社リサイクルマイスター",
"株式会社トレードランド",
"株式会社フェアーメディカル",
"株式会社アイエイト",
"株式会社スモビー（ラミパス）",
"株式会社ノースフィールド",
"株式会社山田書店",
"株式会社我楽洞",
"株式会社コードファクトリー",
"株式会社まるげん",
"株式会社Blanc",
"株式会社シンメトリー（Symmetry）",
"株式会社DRN（ドクターネイル爪革命）",
"株式会社横浜ヘルシー",
"株式会社クォーク（美男 bidan）",
"株式会社T・S・S（mellow-wax）",
}

kept_lines = []
removed_names = []
for line in lines:
    stripped = line.rstrip("\n")
    if stripped in remove_set:
        removed_names.append(stripped)
        continue
    kept_lines.append(line)

with io.open(path, "w", encoding="utf-8") as f:
    f.writelines(kept_lines)

print(f"removed {len(removed_names)} / expected {len(remove_set)}")
missing = remove_set - set(removed_names)
if missing:
    print("NOT FOUND IN FILE (check exact text):")
    for m in missing:
        print(" -", m)
