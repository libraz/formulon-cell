import { afterEach, describe, expect, it } from 'vitest';
import { disposeOverlayPortal, ensureOverlayPortal } from '../../../src/interact/overlay-portal.js';
import { createDialogShell } from '../../../src/toolbar/dialogs/shell.js';

describe('toolbar dialog shell overlay boundary', () => {
  afterEach(() => {
    document.body.replaceChildren();
  });

  it('uses the explicit host portal and keeps the host-owned root intact', () => {
    const host = document.createElement('div');
    host.className = 'fc-host';
    const root = document.createElement('section');
    document.body.append(host, root);
    ensureOverlayPortal(host, { root });

    const shell = createDialogShell({ title: 'Embedded dialog', host });

    expect(root.contains(shell.overlay)).toBe(true);
    shell.overlay.remove();
    disposeOverlayPortal(host);
    expect(document.body.contains(root)).toBe(true);
  });
});
