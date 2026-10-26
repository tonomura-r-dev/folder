// テストの設定。vite.config.ts は編集画面用（root が packages/editor）なので、テストは別にする。

import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    root: import.meta.dirname,
    include: ['tests/**/*.test.ts'],
  },
});
