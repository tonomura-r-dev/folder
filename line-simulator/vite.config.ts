// 編集ツールの開発サーバー。保存API（tools/dev-api.ts）を同じサーバーに差し込む。
// 社内PCのこのフォルダでだけ動かす前提なので、外からの接続は受けない（localhost のみ）。

import { realpathSync } from 'node:fs';
import react from '@vitejs/plugin-react';
import { defineConfig, type Plugin } from 'vite';
import { createApiHandler } from './tools/dev-api.ts';
import { ProjectStore } from './tools/project-store.ts';

// Windows の短い名前（TONOMU~1 など）で開かれると、ファイル監視が落ちることがある（libuv の fs-event）。
// 長い名前に直してから動かす。
const ROOT = realpathSync.native(import.meta.dirname);
process.chdir(ROOT);

function projectApi(): Plugin {
  return {
    name: 'line-simulator-api',
    configureServer(server) {
      const handler = createApiHandler(new ProjectStore(ROOT));
      server.middlewares.use((req, res, next) => void handler(req, res, next));
    },
  };
}

export default defineConfig({
  root: 'packages/editor',
  plugins: [react(), projectApi()],
  server: {
    host: '127.0.0.1',
    port: 5317,
    open: '/',
    // projects/ の保存でページが再読み込みされないように、監視から外す
    watch: { ignored: ['**/projects/**', '**/.trash/**'] },
  },
  build: { outDir: '../../dist/editor', emptyOutDir: true },
});
