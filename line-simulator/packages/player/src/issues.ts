// 検証結果（code）を日本語の説明にする。エンジンは code と path しか返さない。

import type { Issue } from '../../engine/src/index.ts';

const KIND: Record<string, string> = {
  tag: 'タグ',
  cv: 'CV',
  route: '友だち追加経路',
  message: '文面',
  richMenu: 'リッチメニュー',
  cell: 'メニューの区画',
  autoReply: '自動応答',
  survey: 'アンケート',
  question: 'アンケートの質問',
  choice: '選択肢',
  page: 'Webページ',
  flow: 'シナリオ',
  node: 'ノード',
  button: 'ボタン',
  asset: '画像',
};

const LIMIT: Record<string, string> = {
  bubblesPerMessage: '1回の配信の吹き出しは5個まで',
  buttonsPerBubble: 'ボタンは4個まで',
  textLength: '本文は5,000字まで',
  stepDays: 'ステップ配信は365日まで',
  menuCells: 'リッチメニューは20区画まで',
};

const str = (v: unknown) => (v === undefined || v === null || v === '' ? '（未設定）' : String(v));

export function describeIssue(i: Issue): string {
  const p = i.params ?? {};
  switch (i.code) {
    case 'id.duplicate':
      return `ID「${str(p['id'])}」が重複しています`;
    case 'ref.missing': {
      const kind = KIND[String(p['kind'])] ?? String(p['kind']);
      // 新しく置いたノードなど、まだ選んでいないだけのときは「選んでください」と言う
      return p['id'] === '' || p['id'] === undefined ? `${kind}が選ばれていません` : `${kind}「${str(p['id'])}」が見つかりません`;
    }
    case 'edge.danglingSource':
      return `線の出発点のノード「${str(p['nodeId'])}」がありません`;
    case 'edge.danglingTarget':
      return p['reason'] === 'entry' ? '入口ノードに線が入っています' : `線の行き先のノード「${str(p['nodeId'])}」がありません`;
    case 'edge.invalidPort':
      return p['reason'] === 'duplicate' ? '同じ出口から線が2本出ています' : `このノードには「${str(p['port'])}」という出口がありません`;
    case 'expr.invalid':
      return '条件式の書き方が正しくありません';
    case 'flow.noEntry':
      return `入口ノードは1つだけ必要です（いま${str(p['count'])}個）`;
    case 'flow.zeroTimeCycle':
      return '時間が進まないまま同じところを回るループがあります（無限ループ）';
    case 'branch.noExit':
      return '分岐に「それ以外」の行き先がありません';
    case 'node.requiredMissing':
      return p['reason'] === 'noButtons'
        ? 'ボタン選択には、ボタン付きの文面を選んでください'
        : p['reason'] === 'outOfBounds'
          ? 'メニューの区画が枠からはみ出しています'
          : '必須の設定が空です';
    case 'asset.fileMissing':
      return `画像ファイルがありません：${str(p['path'])}`;
    case 'prov.unapproved':
      return 'AIの推定が承認されていません（承認するまで書き出せません）';
    case 'patch.approvedByAi':
      return 'AIの差分に「承認済み」が付いています（承認は人だけが付けます）';
    case 'tag.unused':
      return `タグ「${str(p['id'])}」がどこでも使われていません`;
    case 'node.unreachable':
      return '入口からたどり着けないノードです';
    case 'node.deadEnd':
      return '出口がどこにも繋がっていません（ここでシナリオが終わります）';
    case 'optional.empty':
      return '任意の項目が空です';
    case 'delivery.possibleDuplicate':
      return `同じ文面に、別々のシナリオから届く経路があります（重複配信の可能性）`;
    case 'flow.noStop':
      return '停止条件がありません（CV済みの人にも送り続ける可能性）';
    case 'flow.friendRequired':
      return '友だち追加前に始まるシナリオが、待機を挟まずに配信しようとしています（届きません）';
    case 'persona.noneMatch':
      return `ペルソナを台本どおり${str(p['days'])}日流しても、誰もこのシナリオに入りません（デモで見せられません）`;
    case 'flow.neverStarted':
      return 'どこからも呼ばれないシナリオです';
    case 'text.unknownVar':
      return `差し込み「{{${str(p['key'])}}}」は使えません`;
    case 'line.limit':
      return `LINEの上限を超えています（${LIMIT[String(p['limit'])] ?? String(p['limit'])}）`;
    default:
      return i.code;
  }
}
