import type { ExceptionGroup } from '../lib/types';

interface ExceptionsPanelProps {
  groups: ExceptionGroup[];
  onNavigate: (rowId: string) => void;
}

export function ExceptionsPanel({ groups, onNavigate }: ExceptionsPanelProps): JSX.Element {
  return (
    <div className="mx-auto flex max-w-4xl flex-col gap-4 p-4">
      {groups.map((g) => (
        <section key={g.key} className="rounded-md border border-neutral-300 bg-white shadow-sm">
          <header className="flex items-center gap-2 border-b border-neutral-200 px-3 py-2">
            <h2 className="text-sm font-bold">{g.title}</h2>
            <span
              className={`rounded-full px-2 py-px text-[11px] font-semibold ${
                g.items.length > 0 ? 'bg-amber-100 text-amber-800' : 'bg-green-100 text-green-800'
              }`}
            >
              {g.items.length}
            </span>
          </header>
          {g.note && <p className="border-b border-neutral-100 px-3 py-1.5 text-[11px] italic text-neutral-500">{g.note}</p>}
          {g.items.length === 0 ? (
            <p className="px-3 py-2 text-xs text-neutral-400">Nothing to review.</p>
          ) : (
            <ul className="divide-y divide-neutral-100">
              {g.items.map((item, i) => (
                <li key={i} className="flex items-start gap-2 px-3 py-1.5 text-xs">
                  <span className="mt-0.5 text-amber-500">⚠</span>
                  <span className="min-w-0 flex-1 break-words">{item.message}</span>
                  {item.rowId && (
                    <button
                      type="button"
                      className="shrink-0 rounded border border-blue-300 bg-blue-50 px-1.5 py-0.5 text-[11px] font-medium text-blue-700 hover:bg-blue-100"
                      onClick={() => onNavigate(item.rowId as string)}
                    >
                      Show row →
                    </button>
                  )}
                </li>
              ))}
            </ul>
          )}
        </section>
      ))}
    </div>
  );
}
