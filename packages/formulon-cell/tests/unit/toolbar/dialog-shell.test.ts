import { afterEach, describe, expect, it } from 'vitest';
import { disposeOverlayPortal, ensureOverlayPortal } from '../../../src/interact/overlay-portal.js';
import { createDialogShell, mountDialog } from '../../../src/toolbar/dialogs/shell.js';

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

  it('can focus a dialog control synchronously for immediate keyboard dismissal', () => {
    const opener = document.createElement('button');
    const host = document.createElement('div');
    document.body.append(opener, host);
    opener.focus();

    const shell = createDialogShell({ title: 'Immediate dialog', host });
    const control = document.createElement('button');
    shell.body.appendChild(control);

    mountDialog(shell, control, { immediate: true });

    expect(document.activeElement).toBe(control);
    shell.overlay.remove();
    opener.remove();
    host.remove();
  });

  it('does not focus a detached dialog from a deferred callback', async () => {
    const opener = document.createElement('button');
    const host = document.createElement('div');
    document.body.append(opener, host);
    opener.focus();

    const shell = createDialogShell({ title: 'Deferred dialog', host });
    const control = document.createElement('button');
    shell.body.appendChild(control);
    mountDialog(shell, control);
    shell.overlay.remove();

    await new Promise<void>((resolve) => {
      const frame = document.defaultView?.requestAnimationFrame ?? requestAnimationFrame;
      frame(() => resolve());
    });

    expect(document.activeElement).toBe(opener);
    opener.remove();
    host.remove();
  });
});
