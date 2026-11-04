import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  clearDialogPlacement,
  currentDialogAnchor,
  placeDialogPanel,
  withDialogAnchor,
} from '../../../src/interact/dialog-placement.js';

const build = (platform: string, size = { width: 400, height: 300 }) => {
  const portal = document.createElement('div');
  portal.dataset.fcPlatform = platform;
  const overlay = document.createElement('div');
  const panel = document.createElement('div');
  overlay.appendChild(panel);
  portal.appendChild(overlay);
  document.body.appendChild(portal);
  panel.getBoundingClientRect = () => ({ ...size, left: 0, top: 0 }) as DOMRect;
  return { overlay, panel };
};

describe('dialog placement', () => {
  beforeEach(() => {
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(1600);
    vi.spyOn(window, 'innerHeight', 'get').mockReturnValue(1000);
  });
  afterEach(() => {
    vi.restoreAllMocks();
    document.body.replaceChildren();
  });

  it('centres a Mac dialog on the window, then keeps its top-left corner', () => {
    const { overlay, panel } = build('mac');
    placeDialogPanel(overlay, panel);
    expect([panel.style.left, panel.style.top]).toEqual(['600px', '350px']);
    // The panel grows while open; the pinned corner does not move.
    panel.getBoundingClientRect = () =>
      ({ width: 440, height: 300, left: 600, top: 350 }) as DOMRect;
    expect([panel.style.left, panel.style.top]).toEqual(['600px', '350px']);
    expect(overlay.classList.contains('fc-fmtdlg--placed')).toBe(true);
  });

  it('centres a Mac dialog on the context-menu row it came from', () => {
    const { overlay, panel } = build('mac');
    withDialogAnchor({ x: 1039, y: 615 }, () =>
      placeDialogPanel(overlay, panel, currentDialogAnchor()),
    );
    expect([panel.style.left, panel.style.top]).toEqual(['839px', '465px']);
  });

  it('keeps the dialog inside the window', () => {
    const { overlay, panel } = build('mac');
    placeDialogPanel(overlay, panel, { x: 1590, y: 20 });
    expect([panel.style.left, panel.style.top]).toEqual(['1200px', '0px']);
  });

  it('leaves other profiles to the stylesheet', () => {
    const { overlay, panel } = build('default');
    placeDialogPanel(overlay, panel, { x: 100, y: 100 });
    expect(panel.style.position).toBe('');
    expect(overlay.classList.contains('fc-fmtdlg--placed')).toBe(false);
  });

  it('forgets the anchor when the dispatch returns, also when it throws', () => {
    expect(currentDialogAnchor()).toBeNull();
    withDialogAnchor({ x: 1, y: 2 }, () => {
      expect(currentDialogAnchor()).toEqual({ x: 1, y: 2 });
    });
    expect(currentDialogAnchor()).toBeNull();
    expect(() =>
      withDialogAnchor({ x: 1, y: 2 }, () => {
        throw new Error('boom');
      }),
    ).toThrow('boom');
    expect(currentDialogAnchor()).toBeNull();
  });

  it('clears the pinned position for the next opening', () => {
    const { overlay, panel } = build('mac');
    placeDialogPanel(overlay, panel, { x: 300, y: 300 });
    clearDialogPlacement(overlay, panel);
    expect(panel.style.left).toBe('');
    expect(overlay.classList.contains('fc-fmtdlg--placed')).toBe(false);
  });
});
