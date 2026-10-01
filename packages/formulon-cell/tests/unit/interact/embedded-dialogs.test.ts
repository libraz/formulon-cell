import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { en } from '../../../src/i18n/strings/en.js';
import { openCellShiftDialog } from '../../../src/interact/cell-shift-dialog.js';
import { openInsertCopiedCellsDialog } from '../../../src/interact/insert-copied-cells-dialog.js';
import { disposeOverlayPortal, ensureOverlayPortal } from '../../../src/interact/overlay-portal.js';

describe('host-bound direction dialogs', () => {
  beforeEach(() => {
    document.body.replaceChildren();
  });

  afterEach(() => {
    document.body.replaceChildren();
  });

  it('mounts each dialog in its host portal instead of the first document host', () => {
    const firstHost = document.createElement('div');
    firstHost.className = 'fc-host';
    const secondHost = document.createElement('div');
    secondHost.className = 'fc-host';
    document.body.append(firstHost, secondHost);
    const firstRoot = document.createElement('section');
    const secondRoot = document.createElement('section');
    document.body.append(firstRoot, secondRoot);
    ensureOverlayPortal(firstHost, { root: firstRoot });
    ensureOverlayPortal(secondHost, { root: secondRoot });

    openInsertCopiedCellsDialog({
      host: secondHost,
      strings: en,
      onSubmit: () => undefined,
    });
    expect(secondRoot.querySelector('.fc-insertcopied')).not.toBeNull();
    expect(firstRoot.querySelector('.fc-insertcopied')).toBeNull();

    openCellShiftDialog({
      host: firstHost,
      strings: en,
      kind: 'insert',
      onSubmit: () => undefined,
    });
    expect(firstRoot.querySelector('.fc-cellshift')).not.toBeNull();
    expect(secondRoot.querySelector('.fc-cellshift')).toBeNull();

    disposeOverlayPortal(firstHost);
    disposeOverlayPortal(secondHost);
  });

  it('creates dialogs in a non-global owner document', () => {
    const ownerDocument = document.implementation.createHTMLDocument('embedded');
    const host = ownerDocument.createElement('div');
    host.className = 'fc-host';
    ownerDocument.body.appendChild(host);

    openCellShiftDialog({
      host,
      strings: en,
      kind: 'delete',
      onSubmit: () => undefined,
    });

    const dialog = ownerDocument.querySelector('.fc-cellshift');
    expect(dialog).not.toBeNull();
    expect(dialog?.ownerDocument).toBe(ownerDocument);
  });
});
