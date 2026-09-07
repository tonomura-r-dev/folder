// 八十吉：ご予約の流れ（パフォーマー指名）1枚スライド
//   node _build/yasokichi_build_flow_slide.js
//   → 20260907_八十吉_ご予約の流れ_予約状況ビュー版.pptx（ルート直下）
// 既存デッキには組み込まない前提の単独1枚。文言を直すときはここを触る。
const pptxgen = require("pptxgenjs");

const path = require("path");

const OUT = path.join(__dirname, "..", "20260907_八十吉_ご予約の流れ_予約状況ビュー版.pptx");
const FONT = "メイリオ";
const NAVY = "1F285A", INK = "333333", MUTED = "6B7280", LINE = "D9D9D9";
const GREEN = "06C755", GREEN_SOFT = "E6F8EE", GREEN_INK = "0B7A3A";
const AMBER = "D9661F", AMBER_SOFT = "FBEEE4";
const BLUE_SOFT = "F4F7FF";

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625 in
const s = pres.addSlide();
s.background = { color: "FFFFFF" };

// タイトル
s.addText("ご予約の流れ（パフォーマー指名）", { x: 0.45, y: 0.22, w: 6.5, h: 0.42, fontFace: FONT, fontSize: 18, bold: true, color: NAVY, isTextBox: true, margin: 0 });
s.addText("お席の管理はレストランボードのまま。LINEは連絡と案内、「空き状況」画面は見るだけ。", { x: 0.45, y: 0.66, w: 7.2, h: 0.3, fontFace: FONT, fontSize: 10.5, color: MUTED, isTextBox: true, margin: 0 });

// 凡例
const legend = [["LINE", GREEN_SOFT, GREEN_INK], ["空き状況（新設）", BLUE_SOFT, NAVY], ["食べログ", AMBER_SOFT, AMBER]];
let lx = 6.2;
legend.forEach(([t, fill, col]) => {
  const w = t.length * 0.14 + 0.3;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: lx, y: 0.26, w, h: 0.26, fill: { color: fill }, line: { color: fill }, rectRadius: 0.13 });
  s.addText(t, { x: lx, y: 0.26, w, h: 0.26, fontFace: FONT, fontSize: 8.5, bold: true, color: col, align: "center", valign: "middle", isTextBox: true, margin: 0 });
  lx += w + 0.08;
});

// 6ステップ
const steps = [
  { n: "1", who: "お客様 → パフォーマー", place: "LINE", what: "日時・人数・ご指名をトークで送る（記入フォーマットは自動応答で提示）" },
  { n: "2", who: "パフォーマー → お客様", place: "LINE", what: "受け取りの返事。「この時点では確定していません」を明記" },
  { n: "2.5", who: "パフォーマー", place: "空き状況", what: "その日の予約一覧と残席を見る。レストランボードは開かない。売上は出ない" },
  { n: "2.5", who: "パフォーマー → お客様", place: "LINE", what: "ご案内できる時間帯を提示。先着順" },
  { n: "3", who: "パフォーマー → お客様", place: "LINE", what: "時間・人数を確定し、予約ページのリンクと確保期限を送る。要望欄に「指名：〇〇」と書いてもらう" },
  { n: "4", who: "お客様", place: "食べログ", what: "本予約。お席の在庫はここで判定＝二重予約なし。予約できたら日時をトークに送る" },
  { n: "5", who: "パフォーマー（お店も同じ画面）", place: "空き状況", what: "予約の行が増えたのを見て、LINEで確定の連絡「〇〇で承りました。当日お待ちしています」" },
];
const placeStyle = { "LINE": [GREEN_SOFT, GREEN_INK], "空き状況": [BLUE_SOFT, NAVY], "食べログ": [AMBER_SOFT, AMBER] };

const X0 = 0.45, Y0 = 1.08, ROW = 0.56, W_LEFT = 6.35;
steps.forEach((st, i) => {
  const y = Y0 + i * ROW;
  const [fill, col] = placeStyle[st.place];
  // 番号
  s.addShape(pres.shapes.OVAL, { x: X0, y: y + 0.09, w: 0.34, h: 0.34, fill: { color: NAVY }, line: { color: NAVY } });
  s.addText(st.n, { x: X0, y: y + 0.09, w: 0.34, h: 0.34, fontFace: FONT, fontSize: st.n.length > 1 ? 8.5 : 10, bold: true, color: "FFFFFF", align: "center", valign: "middle", isTextBox: true, margin: 0 });
  // 場所チップ
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: X0 + 0.44, y: y + 0.13, w: 0.86, h: 0.26, fill: { color: fill }, line: { color: fill }, rectRadius: 0.13 });
  s.addText(st.place, { x: X0 + 0.44, y: y + 0.13, w: 0.86, h: 0.26, fontFace: FONT, fontSize: 8.5, bold: true, color: col, align: "center", valign: "middle", isTextBox: true, margin: 0 });
  // 誰が・何を
  s.addText([
    { text: st.who, options: { bold: true, color: NAVY, fontSize: 10, breakLine: true } },
    { text: st.what, options: { color: INK, fontSize: 9 } },
  ], { x: X0 + 1.4, y: y + 0.02, w: W_LEFT - 1.4, h: ROW - 0.04, fontFace: FONT, valign: "middle", isTextBox: true, margin: 0, paraSpaceAfter: 1 });
  // 区切り線
  if (i < steps.length - 1) s.addShape(pres.shapes.LINE, { x: X0 + 0.44, y: y + ROW - 0.02, w: W_LEFT - 0.44, h: 0, line: { color: LINE, width: 0.5 } });
});

// 右：新しく作るもの
const RX = 7.05, RY = 1.08, RW = 2.5, RH = 3.95;
s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: RX, y: RY, w: RW, h: RH, fill: { color: BLUE_SOFT }, line: { color: LINE, width: 0.75 }, rectRadius: 0.1 });
s.addText("新しく作るもの", { x: RX + 0.18, y: RY + 0.12, w: RW - 0.36, h: 0.28, fontFace: FONT, fontSize: 11, bold: true, color: NAVY, isTextBox: true, margin: 0 });
s.addText("「空き状況」画面（LINE内）", { x: RX + 0.18, y: RY + 0.4, w: RW - 0.36, h: 0.26, fontFace: FONT, fontSize: 9.5, color: INK, isTextBox: true, margin: 0 });
const pipe = ["レストランボード", "予約通知メール", "Gmail", "GAS（5分おき）", "スプレッドシート", "空き状況（LINE内）"];
pipe.forEach((p, i) => {
  const y = RY + 0.78 + i * 0.4;
  const last = i === pipe.length - 1;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: RX + 0.18, y, w: RW - 0.36, h: 0.28, fill: { color: last ? NAVY : "FFFFFF" }, line: { color: last ? NAVY : LINE, width: 0.75 }, rectRadius: 0.06 });
  s.addText(p, { x: RX + 0.18, y, w: RW - 0.36, h: 0.28, fontFace: FONT, fontSize: 9, bold: true, color: last ? "FFFFFF" : NAVY, align: "center", valign: "middle", isTextBox: true, margin: 0 });
  if (!last) s.addText("▼", { x: RX + 0.18, y: y + 0.27, w: RW - 0.36, h: 0.14, fontFace: FONT, fontSize: 7, color: MUTED, align: "center", valign: "middle", isTextBox: true, margin: 0 });
});
s.addText([
  { text: "ログインなし／手入力なし", options: { breakLine: true } },
  { text: "月額0円／配信通数0", options: { breakLine: true } },
  { text: "売上はこの経路を通らない" },
], { x: RX + 0.18, y: RY + 3.2, w: RW - 0.36, h: 0.66, fontFace: FONT, fontSize: 8.5, bold: true, color: GREEN_INK, isTextBox: true, margin: 0, paraSpaceAfter: 2 });

// 下段の一言
s.addShape(pres.shapes.RECTANGLE, { x: 0.45, y: 5.12, w: 9.1, h: 0.34, fill: { color: NAVY }, line: { color: NAVY } });
s.addText("パフォーマーが触るのは、LINEのトークと「空き状況」画面の2つだけ。レストランボードには誰もログインしない。", { x: 0.45, y: 5.12, w: 9.1, h: 0.34, fontFace: FONT, fontSize: 10, bold: true, color: "FFFFFF", align: "center", valign: "middle", isTextBox: true, margin: 0 });

// pptxgenjs は fontFace を a:latin / a:ea / a:cs の3つに書くので、後付けの加工はしない
// （2026-09-07：後付けで a:ea/a:cs を重ねたら要素が二重になり PowerPoint で開けなくなった）
pres.writeFile({ fileName: OUT }).then(() => console.log("wrote", OUT));
