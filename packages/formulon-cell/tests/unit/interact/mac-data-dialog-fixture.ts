import type { CellValue } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { createI18nController, type I18nController } from '../../../src/i18n/controller.js';
import type { SpreadsheetInstance } from '../../../src/mount/types.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

export interface MacDataDialogFixture {
  readonly instance: SpreadsheetInstance;
  readonly workbook: WorkbookHandle;
  readonly store: SpreadsheetStore;
  readonly host: HTMLElement;
  readonly i18n: I18nController;
  dispose(): void;
}

/** A minimal instance backed by the real engine and a pass-through command sink. */
export const createMacDataDialogFixture = async (): Promise<MacDataDialogFixture> => {
  const workbook = await WorkbookHandle.createDefault();
  const store = createSpreadsheetStore();
  const i18n = createI18nController({ locale: 'en' });
  const host = document.createElement('div');
  host.tabIndex = -1;
  document.body.appendChild(host);
  const commands = {
    canExecute: () => ({ allowed: true }),
    execute(command: {
      changes: readonly {
        addr: { sheet: number; row: number; col: number };
        value: CellValue;
        formula?: string | null;
      }[];
    }) {
      const atomic = workbook.applyCellPatchAtomic(command.changes);
      return {
        status: atomic.changed.length > 0 ? 'applied' : 'noop',
        applied: atomic.changed,
        rejected: [],
        revision: 1,
      } as const;
    },
  };
  const instance = { workbook, store, commands, host, i18n } as unknown as SpreadsheetInstance;
  return {
    instance,
    workbook,
    store,
    host,
    i18n,
    dispose(): void {
      workbook.dispose();
      i18n.dispose();
      document.body.innerHTML = '';
    },
  };
};

/** Dispatch a bubbling Enter keydown from `target`. */
export const pressEnter = (target: Element, init: KeyboardEventInit = {}): KeyboardEvent => {
  const event = new KeyboardEvent('keydown', {
    key: 'Enter',
    bubbles: true,
    cancelable: true,
    ...init,
  });
  target.dispatchEvent(event);
  return event;
};

export const byId = <T extends HTMLElement>(id: string): T => {
  const el = document.getElementById(id);
  if (!el) throw new Error(`missing #${id}`);
  return el as T;
};

export const byAction = (action: string): HTMLButtonElement => {
  const el = document.querySelector<HTMLButtonElement>(`[data-fc-mac-action="${action}"]`);
  if (!el) throw new Error(`missing action ${action}`);
  return el;
};
