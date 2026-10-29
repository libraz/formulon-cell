import { expect, it, vi } from 'vitest';
import { en } from '../../../src/i18n/strings.js';
import {
  disposeCellShiftDialog,
  openCellShiftDialog,
} from '../../../src/interact/cell-shift-dialog.js';
import { mountStubSheet } from '../../test-utils/mount.js';

const capture = (options?: boolean | AddEventListenerOptions): boolean =>
  typeof options === 'boolean' ? options : (options?.capture ?? false);

it('disposes an open cell-shift dialog and its document listeners with the host', async () => {
  const sheet = await mountStubSheet();
  const add = vi.spyOn(document, 'addEventListener');
  const remove = vi.spyOn(document, 'removeEventListener');
  const outside = document.createElement('button');
  document.body.appendChild(outside);
  try {
    openCellShiftDialog({ host: sheet.host, strings: en, kind: 'insert', onSubmit: vi.fn() });
    const listeners = add.mock.calls.filter(([type]) => type === 'keydown');
    expect(listeners.length).toBeGreaterThan(0);
    expect(document.querySelector('.fc-cellshift')).not.toBeNull();
    outside.focus();
    sheet.instance.dispose();
    expect(document.querySelector('.fc-cellshift')).toBeNull();
    expect(document.activeElement).toBe(outside);
    for (const [type, listener, options] of listeners) {
      expect(
        remove.mock.calls.some(
          ([removedType, removedListener, removedOptions]) =>
            removedType === type &&
            removedListener === listener &&
            capture(removedOptions) === capture(options),
        ),
      ).toBe(true);
    }
  } finally {
    disposeCellShiftDialog(sheet.host);
    sheet.dispose();
    outside.remove();
    vi.restoreAllMocks();
  }
});
