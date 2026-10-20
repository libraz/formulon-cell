import type { WorkbookHandle } from '../../engine/workbook-handle.js';
import { mutators, type SpreadsheetStore } from '../../store/store.js';
import type { History } from '../history.js';
import {
  type ChartClip,
  clearSessionChart,
  copySessionChart,
  pasteSessionChart,
  sessionChartById,
} from '../session-chart.js';
import {
  clearSessionIllustration,
  copySessionIllustration,
  type IllustrationClip,
  pasteSessionIllustration,
  sessionIllustrationById,
} from '../session-illustration.js';

type DrawingClip =
  | { kind: 'chart'; clip: ChartClip }
  | { kind: 'illustration'; clip: IllustrationClip };

/**
 * The chart, picture or shape last copied or cut in a workbook. A copied
 * object has no text form, so its name stands in for it on the system
 * clipboard: while the clipboard still holds that name, a paste is the object.
 */
export interface DrawingClipboard {
  /** Copy, or cut, the selected object. Returns the text for the system
   *  clipboard, or null when no object is selected or it cannot be copied. */
  take(mode: 'copy' | 'cut'): string | null;
  /** Whether pasting the system clipboard's `text` pastes the copied object;
   *  `text` is null when the system clipboard cannot be read. */
  owns(text: string | null): boolean;
  /** Paste the copied object at the active cell and select it; false when refused. */
  paste(): boolean;
  /** Forget the copied object, as a later copy of cells does. */
  clear(): void;
}

export const createDrawingClipboard = (deps: {
  store: SpreadsheetStore;
  wb: WorkbookHandle;
  history: History | null;
}): DrawingClipboard => {
  const { store, wb, history } = deps;
  let held: { clip: DrawingClip; text: string } | null = null;

  const copySelected = (id: string): DrawingClip | null => {
    const state = store.getState();
    if (sessionChartById(state, id)) {
      const clip = copySessionChart(store, id, wb);
      return clip ? { kind: 'chart', clip } : null;
    }
    if (sessionIllustrationById(state, id)) {
      const clip = copySessionIllustration(store, id, wb);
      return clip ? { kind: 'illustration', clip } : null;
    }
    return null;
  };

  const remove = (clip: DrawingClip, id: string): boolean =>
    clip.kind === 'chart'
      ? clearSessionChart(store, id, history, wb, { origin: 'clipboard', commandId: 'chartCut' })
      : clearSessionIllustration(store, id, history, wb, {
          origin: 'clipboard',
          commandId: 'pictureCut',
        });

  return {
    take(mode) {
      const selected = store.getState().ui.selectedObject;
      if (!selected) return null;
      const clip = copySelected(selected.id);
      // Like the spreadsheet, a cut object leaves the sheet at once.
      if (!clip || (mode === 'cut' && !remove(clip, selected.id))) return null;
      held = { clip, text: selected.name };
      return selected.name;
    },
    owns(text) {
      return held !== null && (text === null || text === held.text);
    },
    paste() {
      if (!held) return false;
      const { active } = store.getState().selection;
      const authorization = {
        origin: 'clipboard',
        commandId: held.clip.kind === 'chart' ? 'chartPaste' : 'picturePaste',
      } as const;
      const pasted =
        held.clip.kind === 'chart'
          ? pasteSessionChart(store, held.clip.clip, active, history, wb, authorization)
          : pasteSessionIllustration(store, held.clip.clip, active, history, wb, authorization);
      if (!pasted) return false;
      // The overlay showing the object selects it and names it.
      mutators.setSelectedObject(store, { id: pasted.id, name: '' });
      return true;
    },
    clear() {
      held = null;
    },
  };
};
