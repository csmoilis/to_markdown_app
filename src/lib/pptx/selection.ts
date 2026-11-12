/**
 * Which slides/pages the user removed from the export, with an undo history.
 * Pure and immutable so React state updates stay simple and it can be unit tested.
 */
export interface Selection {
  /** `slide.index` values left out of the preview, copy and download. */
  removed: ReadonlySet<number>;
  /** Earlier `removed` sets, newest last, for Undo. */
  history: ReadonlyArray<ReadonlySet<number>>;
}

const HISTORY_LIMIT = 100;

export const emptySelection: Selection = { removed: new Set(), history: [] };

function commit(selection: Selection, removed: Set<number>): Selection {
  if (removed.size === selection.removed.size && [...removed].every((index) => selection.removed.has(index))) {
    return selection;
  }
  const history = [...selection.history, selection.removed].slice(-HISTORY_LIMIT);
  return { removed, history };
}

export function removeSection(selection: Selection, index: number): Selection {
  const removed = new Set(selection.removed);
  removed.add(index);
  return commit(selection, removed);
}

export function restoreSection(selection: Selection, index: number): Selection {
  const removed = new Set(selection.removed);
  removed.delete(index);
  return commit(selection, removed);
}

export function toggleSection(selection: Selection, index: number): Selection {
  return selection.removed.has(index) ? restoreSection(selection, index) : removeSection(selection, index);
}

export function restoreAll(selection: Selection): Selection {
  return commit(selection, new Set());
}

export function undo(selection: Selection): Selection {
  const previous = selection.history[selection.history.length - 1];
  if (!previous) return selection;
  return { removed: previous, history: selection.history.slice(0, -1) };
}
