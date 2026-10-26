// 編集ツールのデモ画面。営業用HTMLと同じ player をそのまま埋め込む。
// 編集中（未保存）の内容で動くので、直したらすぐ動きを確かめられる。

import { useEffect, useRef } from 'react';
import type { Demo } from '../../player/src/controller.ts';
import { createPlayer, type PlayerHandle } from '../../player/src/player.ts';
import '../../player/src/player.css';
import { assetUrl } from './api.ts';
import { useEditor } from './ctx.ts';

/** 画面を切り替えて戻ってきても続きから見られるよう、案件ごとにデモの状態を覚えておく */
const kept = new Map<string, Demo>();

export function DemoView() {
  const { project, slug, issues } = useEditor();
  const box = useRef<HTMLDivElement>(null);
  const player = useRef<PlayerHandle | null>(null);

  useEffect(() => {
    const saved = kept.get(slug);
    const p = createPlayer(box.current!, {
      project,
      ...(saved ? { demo: saved } : {}),
      assetUrl: (a) => assetUrl(slug, a.path),
      onChange: (d) => kept.set(slug, d),
    });
    player.current = p;
    return () => {
      kept.set(slug, p.getDemo());
      p.destroy();
      player.current = null;
    };
    // 案件を開き直したときだけ作り直す。編集の反映は下の setProject
  }, [slug]);

  useEffect(() => player.current?.setProject(project), [project]);

  const errors = issues.filter((i) => i.level === 'error').length;
  return (
    <div className="detail demo-view" style={{ width: '100%', padding: 0 }}>
      <div className="demo-head">
        <h2>デモ（動作確認）</h2>
        <span className="muted small">スマホの中を操作すると、編集中の内容（未保存を含む）で動きます。右の「時間」で日を進めると、ステップ配信が届きます。</span>
        {errors > 0 && <span className="badge err">重大エラー {errors}件：動きが正しくない可能性があります</span>}
      </div>
      <div ref={box} />
    </div>
  );
}
