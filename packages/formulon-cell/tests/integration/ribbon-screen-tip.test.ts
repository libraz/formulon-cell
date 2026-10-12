import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  attachRibbonScreenTips,
  formatRibbonShortcut,
  ribbonScreenTipText,
} from '../../src/interact/ribbon-screen-tip.js';

const makeButton = (title: string): HTMLButtonElement => {
  const button = document.createElement('button');
  button.type = 'button';
  button.dataset.ribbonCommand = 'copy';
  button.title = title;
  button.setAttribute('aria-label', title);
  return button;
};

const pointer = (
  type: 'pointerover' | 'pointerout',
  target: Element,
  relatedTarget?: EventTarget,
) =>
  target.dispatchEvent(
    new PointerEvent(type, { bubbles: true, relatedTarget: relatedTarget ?? null }),
  );

describe('ribbon screen tips', () => {
  afterEach(() => {
    vi.useRealTimers();
    document.body.replaceChildren();
  });

  it('formats only declared platform shortcuts', () => {
    expect(formatRibbonShortcut('Control+C Meta+C', 'default')).toBe('Ctrl+C');
    expect(formatRibbonShortcut('Control+C Meta+C', 'mac')).toBe('⌘C');
    expect(formatRibbonShortcut('F5 Control+G Meta+G', 'mac')).toBe('⌘G');
    expect(formatRibbonShortcut('F9', 'default')).toBe('F9');
    expect(formatRibbonShortcut(null, 'default')).toBeNull();
  });

  it('does not duplicate shortcut text already present in model titles', () => {
    const button = makeButton('太字 (Cmd+B)');
    button.setAttribute('aria-keyshortcuts', 'Control+B Meta+B');
    expect(ribbonScreenTipText(button, 'mac')).toBe('太字 (Cmd+B)');
    button.title = 'Bold (Ctrl+B)';
    expect(ribbonScreenTipText(button, 'default')).toBe('Bold (Ctrl+B)');
    button.title = 'Fett (Ctrl/Cmd+B)';
    expect(ribbonScreenTipText(button, 'default')).toBe('Fett (Ctrl+B)');
    expect(ribbonScreenTipText(button, 'mac')).toBe('Fett (⌘B)');
  });

  it('can use a localized accessible name when a quick-access button has no title', () => {
    const button = makeButton('保存');
    button.removeAttribute('title');
    const host = document.createElement('div');
    host.appendChild(button);
    document.body.appendChild(host);
    const handle = attachRibbonScreenTips({ host, delayMs: 0 });
    pointer('pointerover', button);
    expect(document.querySelector<HTMLElement>('[role="tooltip"]')?.textContent).toBe('保存');
    handle.detach();
  });

  it('uses current title text and restores native metadata after hover', () => {
    vi.useFakeTimers();
    const host = document.createElement('div');
    const button = makeButton('Copy');
    button.setAttribute('aria-keyshortcuts', 'Control+C Meta+C');
    host.appendChild(button);
    document.body.appendChild(host);
    const handle = attachRibbonScreenTips({ host, delayMs: 100 });

    pointer('pointerover', button);
    expect(host.querySelector('[role="tooltip"]')).toBeNull();
    vi.advanceTimersByTime(99);
    expect(host.ownerDocument.querySelector('[role="tooltip"]')).toBeNull();
    vi.advanceTimersByTime(1);

    const tip = host.ownerDocument.querySelector<HTMLElement>('[role="tooltip"]');
    expect(tip?.textContent).toBe('Copy\nCtrl+C');
    expect(tip?.hidden).toBe(false);
    expect(button.hasAttribute('title')).toBe(false);
    expect(button.getAttribute('aria-describedby')).toBe(tip?.id);

    handle.hide();
    expect(tip?.hidden).toBe(true);
    expect(button.getAttribute('title')).toBe('Copy');
    expect(button.hasAttribute('aria-describedby')).toBe(false);
    handle.detach();
    expect(tip?.isConnected).toBe(false);
  });

  it('shows immediately for focus, preserves described-by, and hides on escape', () => {
    const host = document.createElement('div');
    const button = makeButton('コピー');
    button.setAttribute('aria-keyshortcuts', 'Control+C Meta+C');
    button.setAttribute('aria-describedby', 'validation-message');
    host.appendChild(button);
    document.body.appendChild(host);
    const handle = attachRibbonScreenTips({ host, getPlatform: () => 'mac' });

    button.dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
    const tip = document.querySelector<HTMLElement>('[role="tooltip"]');
    expect(tip?.textContent).toBe('コピー\n⌘C');
    expect(button.getAttribute('aria-describedby')).toBe(`validation-message ${tip?.id}`);

    button.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(tip?.hidden).toBe(true);
    expect(button.getAttribute('title')).toBe('コピー');
    expect(button.getAttribute('aria-describedby')).toBe('validation-message');
    handle.detach();
  });

  it('keeps a title/description update made while the custom tip is open', () => {
    const host = document.createElement('div');
    host.lang = 'de';
    const button = makeButton('Kopieren');
    host.appendChild(button);
    document.body.appendChild(host);
    const handle = attachRibbonScreenTips({ host, delayMs: 0 });
    button.dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
    const tip = document.querySelector<HTMLElement>('[role="tooltip"]');
    expect(tip?.lang).toBe('de');

    button.title = 'Kopieren (aktualisiert)';
    button.setAttribute('aria-describedby', `${tip?.id} live-update`);
    handle.hide();
    expect(button.title).toBe('Kopieren (aktualisiert)');
    expect(button.getAttribute('aria-describedby')).toBe('live-update');
    handle.detach();
  });

  it('keeps a scrollable tip open while the pointer crosses from a control', () => {
    vi.useFakeTimers();
    const host = document.createElement('div');
    const button = makeButton('長い説明を読むためのツールチップ');
    host.appendChild(button);
    document.body.appendChild(host);
    const handle = attachRibbonScreenTips({ host, delayMs: 0 });
    pointer('pointerover', button);
    const tip = document.querySelector<HTMLElement>('[role="tooltip"]');
    expect(tip?.hidden).toBe(false);

    pointer('pointerout', button, document.body);
    host.dispatchEvent(new PointerEvent('pointerleave', { relatedTarget: document.body }));
    vi.advanceTimersByTime(100);
    expect(tip?.hidden).toBe(false);
    tip?.dispatchEvent(new PointerEvent('pointerenter', { bubbles: true }));
    vi.advanceTimersByTime(180);
    expect(tip?.hidden).toBe(false);
    tip?.dispatchEvent(new Event('scroll'));
    expect(tip?.hidden).toBe(false);
    tip?.dispatchEvent(
      new PointerEvent('pointerleave', { bubbles: true, relatedTarget: document.body }),
    );
    vi.advanceTimersByTime(179);
    expect(tip?.hidden).toBe(false);
    vi.advanceTimersByTime(1);
    expect(tip?.hidden).toBe(true);
    handle.detach();
  });

  it('updates a visible tip when the command projection changes its explanation', async () => {
    const host = document.createElement('div');
    const button = makeButton('Paste');
    host.appendChild(button);
    document.body.appendChild(host);
    const handle = attachRibbonScreenTips({ host, delayMs: 0 });
    pointer('pointerover', button);
    const tip = document.querySelector<HTMLElement>('[role="tooltip"]');
    button.title = 'Paste — clipboard unavailable';
    await Promise.resolve();
    expect(tip?.textContent).toBe('Paste — clipboard unavailable');
    expect(button.hasAttribute('title')).toBe(false);
    handle.hide();
    expect(button.title).toBe('Paste — clipboard unavailable');
    handle.detach();
  });

  it('dismisses before a ribbon pointer action starts', () => {
    const host = document.createElement('div');
    const button = makeButton('Save');
    host.appendChild(button);
    document.body.appendChild(host);
    const handle = attachRibbonScreenTips({ host, delayMs: 0 });
    button.dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
    const tip = document.querySelector<HTMLElement>('[role="tooltip"]');
    expect(tip?.hidden).toBe(false);
    button.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true }));
    expect(tip?.hidden).toBe(true);
    handle.detach();
  });

  it('dismisses hover-only tips on Escape when focus remains outside the ribbon', () => {
    const host = document.createElement('div');
    const button = makeButton('Bold');
    host.appendChild(button);
    document.body.appendChild(host);
    const handle = attachRibbonScreenTips({ host, delayMs: 0 });
    pointer('pointerover', button);
    const tip = document.querySelector<HTMLElement>('[role="tooltip"]');
    expect(tip?.hidden).toBe(false);
    document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(tip?.hidden).toBe(true);
    expect(button.title).toBe('Bold');
    expect(button.hasAttribute('aria-describedby')).toBe(false);
    handle.detach();
  });

  it('cancels a pending tip when the pointer leaves for another host element', () => {
    vi.useFakeTimers();
    const host = document.createElement('div');
    const button = makeButton('保存');
    const label = document.createElement('span');
    label.textContent = 'Ribbon group';
    host.append(button, label);
    document.body.appendChild(host);
    const handle = attachRibbonScreenTips({ host, delayMs: 100 });
    pointer('pointerover', button);
    pointer('pointerout', button, label);
    vi.advanceTimersByTime(100);
    expect(document.querySelector('[role="tooltip"]')).toBeNull();
    handle.detach();
  });

  it('hides on pointer exit, scroll, and rerender without leaving detached owners', async () => {
    vi.useFakeTimers();
    const host = document.createElement('div');
    const button = makeButton('Format cells');
    host.appendChild(button);
    document.body.appendChild(host);
    const handle = attachRibbonScreenTips({ host, delayMs: 0 });
    pointer('pointerover', button);
    const tip = document.querySelector<HTMLElement>('[role="tooltip"]');
    expect(tip?.hidden).toBe(false);

    pointer('pointerout', button, document.body);
    vi.advanceTimersByTime(180);
    expect(tip?.hidden).toBe(true);
    button.dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
    expect(tip?.hidden).toBe(false);
    document.dispatchEvent(new Event('scroll'));
    expect(tip?.hidden).toBe(true);

    button.dispatchEvent(new FocusEvent('focusin', { bubbles: true }));
    const replacement = makeButton('Format cells');
    host.replaceChildren(replacement);
    await Promise.resolve();
    expect(tip?.hidden).toBe(true);
    handle.detach();
    expect(tip?.isConnected).toBe(false);
  });

  it('keeps translated labels and line wrapping text content-safe', () => {
    const button = makeButton('セルの書式設定 / รูปแบบเซลล์');
    button.setAttribute('aria-keyshortcuts', 'Control+1 Meta+1');
    expect(ribbonScreenTipText(button, 'default')).toBe('セルの書式設定 / รูปแบบเซลล์\nCtrl+1');
    expect(button.innerHTML).toBe('');
  });
});
