// 営業用HTMLに入れる JS を1本に束ねる。中身は packages/player/src/standalone.ts と、そこから読む player・エンジンだけ。
// Vite（中身は rolldown）の build を使うので、追加のツールは要らない。

import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { build, type Rollup } from 'vite';
import { ROOT } from './project-io.ts';

export const PLAYER_CSS = join(ROOT, 'packages', 'player', 'src', 'player.css');

let cached: string | null = null;

/** IIFE 形式の JS（グローバル LineSim に simulate・mount が入る） */
export async function bundlePlayer(): Promise<string> {
  if (cached) return cached;
  const result = await build({
    configFile: false,
    root: ROOT,
    logLevel: 'silent',
    publicDir: false,
    build: {
      write: false,
      minify: true,
      target: 'es2022',
      lib: { entry: join(ROOT, 'packages', 'player', 'src', 'standalone.ts'), formats: ['iife'], name: 'LineSim', fileName: () => 'player.js' },
    },
  });
  const outputs = (Array.isArray(result) ? result : [result]) as Rollup.RollupOutput[];
  const chunk = outputs.flatMap((o) => o.output).find((o): o is Rollup.OutputChunk => o.type === 'chunk');
  if (!chunk) throw new Error('JSを束ねられませんでした');
  cached = chunk.code;
  return cached;
}

export const playerCss = () => readFileSync(PLAYER_CSS, 'utf8');
