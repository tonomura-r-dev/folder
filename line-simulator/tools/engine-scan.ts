// エンジンのソース検査の中身（NFR-07）。check-engine.ts とテストから使う。
//   1. 文字列リテラルに日本語が無い（コメントは対象外）
//   2. DOM・Node.js・現在時刻・乱数を使っていない（決定論と、ブラウザ／サーバー両対応のため）
//   3. エンジンの外のファイルを import していない

const JAPANESE = /[　-ヿ㐀-䶿一-鿿＀-￯]/;
const BANNED = ['Date', 'document', 'window', 'process', 'require', 'localStorage', 'sessionStorage', 'fetch', 'XMLHttpRequest', 'performance'];

type Token = { kind: 'string' | 'code'; text: string; line: number };

/** コメントを飛ばし、文字列リテラルとそれ以外に分ける最小限の字句解析 */
function tokenize(src: string): Token[] {
  const out: Token[] = [];
  let i = 0;
  let line = 1;
  let code = '';
  let codeLine = 1;
  const flush = () => {
    if (code) out.push({ kind: 'code', text: code, line: codeLine });
    code = '';
  };
  while (i < src.length) {
    const c = src[i]!;
    const n = src[i + 1];
    if (c === '/' && n === '/') {
      while (i < src.length && src[i] !== '\n') i++;
      continue;
    }
    if (c === '/' && n === '*') {
      i += 2;
      while (i < src.length && !(src[i] === '*' && src[i + 1] === '/')) (src[i] === '\n' && line++, i++);
      i += 2;
      continue;
    }
    if (c === '"' || c === "'" || c === '`') {
      flush();
      const start = line;
      let s = '';
      i++;
      while (i < src.length && src[i] !== c) {
        if (src[i] === '\\') (s += src[i]! + (src[i + 1] ?? ''), (i += 2));
        else (src[i] === '\n' && line++, (s += src[i]!), i++);
      }
      i++;
      out.push({ kind: 'string', text: s, line: start });
      codeLine = line;
      continue;
    }
    if (!code) codeLine = line;
    if (c === '\n') line++;
    code += c;
    i++;
  }
  flush();
  return out;
}

export function scanEngineSource(src: string, name: string): string[] {
  const problems: string[] = [];
  const tokens = tokenize(src);
  tokens.forEach((t, idx) => {
    if (t.kind === 'string') {
      if (JAPANESE.test(t.text)) problems.push(`${name}:${t.line}  文字列に日本語: ${JSON.stringify(t.text.slice(0, 40))}`);
      const prev = tokens[idx - 1];
      if (prev?.kind === 'code' && /\b(from|import)\s*\(?\s*$/.test(prev.text) && !t.text.startsWith('./'))
        problems.push(`${name}:${t.line}  エンジン外の import: ${t.text}`);
      return;
    }
    for (const word of BANNED) {
      if (new RegExp(`(?<![\\w.$])${word}\\b`).test(t.text)) problems.push(`${name}:${t.line}  使用禁止: ${word}`);
    }
    if (/Math\.random\b/.test(t.text)) problems.push(`${name}:${t.line}  使用禁止: Math.random`);
  });
  return problems;
}
