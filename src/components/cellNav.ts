/**
 * Round 8 §12.2 — shared editable-grid behavior.
 *
 * Committing a cell edit re-renders the whole tab from App state, which reset
 * the scroll container (most visibly to the far left on the wide TimeNavi
 * grid). Every commit path therefore snapshots the scroll offsets of the
 * scrollable ancestors first and re-applies them across the frames React needs
 * to finish its re-render. A real user scroll (wheel / touch) cancels the
 * restore so we never fight the mouse.
 */

interface ScrollSnapshot {
  el: Element;
  left: number;
  top: number;
}

function scrollSnapshots(el: HTMLElement): ScrollSnapshot[] {
  const out: ScrollSnapshot[] = [];
  for (let node: HTMLElement | null = el; node; node = node.parentElement) {
    const scrolls = node.scrollWidth > node.clientWidth || node.scrollHeight > node.clientHeight;
    if (scrolls) out.push({ el: node, left: node.scrollLeft, top: node.scrollTop });
  }
  const doc = document.scrollingElement;
  if (doc && !out.some((s) => s.el === doc)) {
    out.push({ el: doc, left: doc.scrollLeft, top: doc.scrollTop });
  }
  return out;
}

/** Frames to keep re-applying the snapshot (React commit + layout + paint). */
const RESTORE_FRAMES = 4;

/**
 * Capture the viewport around `el` and hold it steady while the caller's state
 * change re-renders. `after` runs once the restore window closes.
 */
export function keepScrollPosition(el: HTMLElement, after?: () => void): void {
  const snapshots = scrollSnapshots(el);
  let cancelled = false;
  const cancel = (): void => {
    cancelled = true;
  };
  window.addEventListener('wheel', cancel, { passive: true });
  window.addEventListener('touchmove', cancel, { passive: true });

  const finish = (): void => {
    window.removeEventListener('wheel', cancel);
    window.removeEventListener('touchmove', cancel);
    after?.();
  };

  const apply = (): void => {
    for (const s of snapshots) {
      if (s.el.scrollLeft !== s.left) s.el.scrollLeft = s.left;
      if (s.el.scrollTop !== s.top) s.el.scrollTop = s.top;
    }
  };

  apply();
  let frames = 0;
  const tick = (): void => {
    if (cancelled) {
      finish();
      return;
    }
    apply();
    if (++frames < RESTORE_FRAMES) requestAnimationFrame(tick);
    else finish();
  };
  requestAnimationFrame(tick);
}

/** Run `mutate` (a state change) without letting the viewport jump. */
export function withScrollKept(el: HTMLElement | null, mutate: () => void): void {
  if (!el) {
    mutate();
    return;
  }
  keepScrollPosition(el);
  mutate();
}

function nearestScrollParent(el: HTMLElement): HTMLElement | null {
  for (let node: HTMLElement | null = el.parentElement; node; node = node.parentElement) {
    if (node.scrollHeight > node.clientHeight) return node;
  }
  return null;
}

/**
 * Vertical-only "scroll into view": nudges the container just enough to show
 * `el`, never touching scrollLeft (an Excel-like Enter must not slide the grid
 * sideways).
 */
export function ensureVisibleVertically(el: HTMLElement): void {
  const container = nearestScrollParent(el);
  if (!container) return;
  const cell = el.getBoundingClientRect();
  const box = container.getBoundingClientRect();
  // Sticky headers overlap the top of the box; keep a small margin.
  const margin = 28;
  if (cell.top < box.top + margin) {
    container.scrollTop -= box.top + margin - cell.top;
  } else if (cell.bottom > box.bottom) {
    container.scrollTop += cell.bottom - box.bottom;
  }
}

const EDITABLE_SELECTOR = 'input:not([type="checkbox"]):not([disabled]), textarea:not([disabled])';

/**
 * Excel-like vertical move: the editable cell in the same column, `dir` rows
 * away. Rows without an editable cell in that column are skipped.
 */
export function findCellInDirection(from: HTMLElement, dir: 1 | -1): HTMLElement | null {
  const cell = from.closest('td, th');
  const row = from.closest('tr');
  if (!cell || !row) return null;
  const colIndex = (cell as HTMLTableCellElement).cellIndex;

  const table = row.closest('table');
  if (!table) return null;
  const rows = Array.from(table.querySelectorAll('tr'));
  let i = rows.indexOf(row);
  if (i < 0) return null;

  for (i += dir; i >= 0 && i < rows.length; i += dir) {
    const target = rows[i].cells[colIndex];
    if (!target) continue;
    const input = target.querySelector<HTMLElement>(EDITABLE_SELECTOR);
    if (input) return input;
  }
  return null;
}

/**
 * Commit a cell edit: hold the viewport, apply the change, and (on Enter) move
 * focus one row down, scrolling only as far as that cell needs vertically.
 */
export function commitCell(
  el: HTMLElement,
  commit: () => void,
  move: 'down' | 'up' | 'none',
): void {
  const next = move === 'none' ? null : findCellInDirection(el, move === 'down' ? 1 : -1);
  keepScrollPosition(el, next ? () => ensureVisibleVertically(next) : undefined);
  if (next) {
    next.focus({ preventScroll: true });
    if (next instanceof HTMLInputElement) next.select();
  } else if (move !== 'none') {
    el.blur();
  }
  commit();
}
