import { useCallback, useState } from 'react';
import { ProjectEditor } from './ProjectEditor.tsx';
import { ProjectList } from './ProjectList.tsx';
import { useRoute } from './route.ts';

export function App() {
  const route = useRoute();
  const [toast, setToast] = useState<{ msg: string; kind: 'ok' | 'err' } | null>(null);
  const show = useCallback((msg: string, kind: 'ok' | 'err' = 'ok') => {
    setToast({ msg, kind });
    window.setTimeout(() => setToast((t) => (t?.msg === msg ? null : t)), kind === 'err' ? 6000 : 2500);
  }, []);
  return (
    <>
      {route.page === 'list' ? <ProjectList toast={show} /> : <ProjectEditor key={route.slug} route={route} toast={show} />}
      {toast && (
        <div className={`toast${toast.kind === 'err' ? ' err' : ''}`} role="status">
          {toast.msg}
        </div>
      )}
    </>
  );
}
