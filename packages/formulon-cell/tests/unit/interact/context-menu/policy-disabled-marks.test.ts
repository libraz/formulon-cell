import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type {
  CellBatchCommand,
  OperationIntent,
} from '../../../../src/commands/interaction-policy.js';
import type { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { ja } from '../../../../src/i18n/strings.js';
import {
  attachContextMenu,
  type ContextMenuHandle,
} from '../../../../src/interact/context-menu.js';
import type { ContextMenuInteractionController } from '../../../../src/interact/context-menu-options.js';
import { disposeOverlayPortal } from '../../../../src/interact/overlay-portal.js';
import { createSpreadsheetStore } from '../../../../src/store/store.js';
import { fireContextMenu, item, newWb } from './fixtures.js';

const command: CellBatchCommand = {
  type: 'cellBatch',
  operation: 'clear',
  origin: 'instanceApi',
  changes: [],
};

const parentOf = (id: string): HTMLButtonElement | null =>
  document.querySelector<HTMLButtonElement>(`.fc-ctxmenu [data-fc-submenu="${id}"]`);

describe('policy-disabled marks on explicitly disabled items', () => {
  let host: HTMLElement;
  let wb: WorkbookHandle;
  let detach: ContextMenuHandle | null;

  beforeEach(async () => {
    host = document.createElement('div');
    host.tabIndex = -1;
    document.body.appendChild(host);
    wb = await newWb();
    detach = null;
  });

  afterEach(() => {
    detach?.();
    disposeOverlayPortal(host);
    document.body.innerHTML = '';
  });

  it('marks a plain item and a submenu parent alike, and clears both on refresh', () => {
    let allowed = false;
    let refresh = (): void => {};
    const controller: ContextMenuInteractionController = {
      canExecute: (_intent: OperationIntent) =>
        allowed
          ? { allowed: true }
          : { allowed: false, code: 'operationDenied', reason: 'Blocked by host policy.' },
      execute: vi.fn(),
      subscribe: (listener) => {
        refresh = listener;
        return () => {};
      },
    };
    detach = attachContextMenu({
      getLocale: () => 'en-US',
      strings: ja,
      host,
      store: createSpreadsheetStore(),
      wb,
      interactionController: controller,
      options: {
        mode: 'builtIn',
        items: [],
        transform: () => [
          { id: 'app.plain', label: 'Plain', command, disabled: true, disabledReason: 'Mine' },
          {
            id: 'app.more',
            label: 'More',
            command,
            disabled: true,
            disabledReason: 'Mine',
            children: [{ id: 'app.nested', label: 'Nested', command }],
          },
        ],
      },
    });
    fireContextMenu(host, 200, 70);
    const plain = item('app.plain');
    const more = parentOf('app.more');
    expect(plain?.dataset.policyDisabled).toBe('true');
    expect(more?.dataset.policyDisabled).toBe('true');
    expect(more?.dataset.disabledReason).toBe(plain?.dataset.disabledReason);

    allowed = true;
    refresh();
    for (const button of [plain, more]) {
      expect(button?.dataset.policyDisabled).toBeUndefined();
      expect(button?.disabled).toBe(true);
      expect(button?.dataset.disabledReason).toBe('Mine');
    }
  });
});
