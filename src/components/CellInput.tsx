/**
 * Round 8 §12.2 — the one editable table cell used by every grid.
 *
 * Commits on blur / Enter, reverts on Escape, and keeps the scroll viewport
 * where it was through the re-render the commit triggers. Enter also steps down
 * one cell in the same column, Excel-style.
 */
import { useRef, useState, type CSSProperties } from 'react';
import { commitCell } from './cellNav';

export interface CellInputProps {
  /** Formatted current value (also the placeholder-free display). */
  value: string;
  onCommit: (text: string) => void;
  className?: string;
  title?: string;
  placeholder?: string;
  style?: CSSProperties;
  /** Set false for inputs outside a table body (Enter just commits). */
  moveOnEnter?: boolean;
}

export function CellInput({
  value,
  onCommit,
  className,
  title,
  placeholder,
  style,
  moveOnEnter = true,
}: CellInputProps): JSX.Element {
  const [draft, setDraft] = useState<string | null>(null);
  const handled = useRef(false);

  const commit = (text: string | null): void => {
    if (text !== null && text !== value) onCommit(text);
  };

  return (
    <input
      className={`cell-input ${className ?? ''}`}
      value={draft ?? value}
      title={title}
      placeholder={placeholder}
      style={style}
      onChange={(e) => setDraft(e.target.value)}
      onFocus={() => setDraft(value)}
      onBlur={(e) => {
        if (handled.current) {
          handled.current = false;
          setDraft(null);
          return;
        }
        const text = draft;
        setDraft(null);
        commitCell(e.currentTarget, () => commit(text), 'none');
      }}
      onKeyDown={(e) => {
        const el = e.currentTarget;
        if (e.key === 'Enter') {
          e.preventDefault();
          const text = draft;
          handled.current = true;
          setDraft(null);
          commitCell(el, () => commit(text), moveOnEnter ? 'down' : 'none');
        } else if (e.key === 'Escape') {
          handled.current = true;
          setDraft(null);
          el.blur();
        }
      }}
    />
  );
}
