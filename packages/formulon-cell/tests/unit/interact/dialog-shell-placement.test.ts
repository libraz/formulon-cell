import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { withDialogAnchor } from '../../../src/interact/dialog-placement.js';
import { createDialogShell } from '../../../src/interact/dialog-shell.js';
import { disposeOverlayPortal } from '../../../src/interact/overlay-portal.js';
import {
  createDialogShell as createToolbarDialogShell,
  mountDialog,
} from '../../../src/toolbar/dialogs/shell.js';

const sized = (panel: HTMLElement): void => {
  panel.getBoundingClientRect = () => ({ width: 400, height: 300, left: 0, top: 0 }) as DOMRect;
};

describe('shared dialog shells on the Mac profile', () => {
  let host: HTMLElement;

  beforeEach(() => {
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(1600);
    vi.spyOn(window, 'innerHeight', 'get').mockReturnValue(1000);
    host = document.createElement('div');
    host.classList.add('fc-host');
    host.dataset.fcPlatform = 'mac';
    document.body.appendChild(host);
  });
  afterEach(() => {
    vi.restoreAllMocks();
    disposeOverlayPortal(host);
    document.body.replaceChildren();
  });

  it('pins an interact dialog on open and frees it on close', () => {
    const shell = createDialogShell({
      host,
      className: 'fc-test',
      ariaLabel: 'test',
      onDismiss: () => {},
    });
    sized(shell.panel);
    shell.open();
    expect([shell.panel.style.left, shell.panel.style.top]).toEqual(['600px', '350px']);
    shell.close();
    expect(shell.panel.style.left).toBe('');
    shell.dispose();
  });

  it('centres an interact dialog on the anchor of the dispatch that opens it', () => {
    const shell = createDialogShell({
      host,
      className: 'fc-test',
      ariaLabel: 'test',
      onDismiss: () => {},
    });
    sized(shell.panel);
    withDialogAnchor({ x: 1039, y: 615 }, () => shell.open());
    expect([shell.panel.style.left, shell.panel.style.top]).toEqual(['839px', '465px']);
    shell.dispose();
  });

  it('places a toolbar dialog from the anchor captured when its shell was built', () => {
    const shell = withDialogAnchor({ x: 1039, y: 615 }, () =>
      createToolbarDialogShell({ host, title: 'test' }),
    );
    sized(shell.panel);
    // The dispatch is over by the time the dialog is mounted; the captured anchor still applies.
    mountDialog(shell, null);
    expect([shell.panel.style.left, shell.panel.style.top]).toEqual(['839px', '465px']);
  });

  it('leaves the desktop profile centred by the stylesheet', () => {
    host.dataset.fcPlatform = 'default';
    const shell = createToolbarDialogShell({ host, title: 'test' });
    sized(shell.panel);
    mountDialog(shell, null);
    expect(shell.panel.style.left).toBe('');
  });
});
