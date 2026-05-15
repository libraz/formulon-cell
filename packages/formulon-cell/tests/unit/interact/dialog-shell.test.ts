import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { createDialogShell } from '../../../src/interact/dialog-shell.js';

describe('interact/dialog-shell', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    host.remove();
    // Sanity check: the body should be free of any test-leaked overlays.
    for (const child of Array.from(document.body.querySelectorAll('[role="dialog"]'))) {
      child.remove();
    }
  });

  it('mounts an aria-correct overlay portaled to body and starts hidden', () => {
    const shell = createDialogShell({ host, className: 'fc-test', ariaLabel: 'Test' });
    expect(shell.overlay.parentElement).toBe(document.body);
    expect(shell.overlay.getAttribute('role')).toBe('dialog');
    expect(shell.overlay.getAttribute('aria-modal')).toBe('true');
    expect(shell.overlay.getAttribute('aria-label')).toBe('Test');
    expect(shell.overlay.className).toBe('fc-test');
    expect(shell.panel.className).toBe('fc-test__panel');
    expect(shell.overlay.hidden).toBe(true);
    expect(shell.isOpen()).toBe(false);
    shell.dispose();
  });

  it('open()/close() flip visibility and isOpen()', () => {
    const shell = createDialogShell({ host, className: 'fc-x', ariaLabel: 'X' });
    shell.open();
    expect(shell.overlay.hidden).toBe(false);
    expect(shell.isOpen()).toBe(true);
    shell.close();
    expect(shell.overlay.hidden).toBe(true);
    expect(shell.isOpen()).toBe(false);
    shell.dispose();
  });

  it('restores focus to the opener when close() hides an active dialog', () => {
    host.tabIndex = -1;
    const shell = createDialogShell({ host, className: 'fc-x', ariaLabel: 'X' });
    const ok = document.createElement('button');
    shell.panel.appendChild(ok);

    host.focus();
    shell.open();
    expect(document.activeElement).toBe(ok);
    shell.close();

    expect(shell.isOpen()).toBe(false);
    expect(document.activeElement).toBe(host);
    shell.dispose();
  });

  it('keeps externally moved focus when close() runs after focus already left the dialog', () => {
    const outside = document.createElement('button');
    document.body.appendChild(outside);
    const shell = createDialogShell({ host, className: 'fc-x', ariaLabel: 'X' });
    const ok = document.createElement('button');
    shell.panel.appendChild(ok);

    host.focus();
    shell.open();
    outside.focus();
    shell.close();

    expect(document.activeElement).toBe(outside);
    shell.dispose();
    outside.remove();
  });

  it('focuses the first focusable control on open when focus is outside', () => {
    const shell = createDialogShell({ host, className: 'fc-x', ariaLabel: 'X' });
    const first = document.createElement('button');
    const second = document.createElement('button');
    shell.panel.append(first, second);

    host.focus();
    shell.open();
    expect(document.activeElement).toBe(first);
    shell.dispose();
  });

  it('traps Tab and Shift+Tab focus within the dialog while open', () => {
    const shell = createDialogShell({ host, className: 'fc-x', ariaLabel: 'X' });
    const first = document.createElement('button');
    const middle = document.createElement('input');
    const last = document.createElement('button');
    shell.panel.append(first, middle, last);
    shell.open();

    last.focus();
    last.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true }));
    expect(document.activeElement).toBe(first);

    first.focus();
    first.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Tab', shiftKey: true, bubbles: true }),
    );
    expect(document.activeElement).toBe(last);
    shell.dispose();
  });

  it('does not trap Tab while hidden', () => {
    const outside = document.createElement('button');
    document.body.appendChild(outside);
    const shell = createDialogShell({ host, className: 'fc-x', ariaLabel: 'X' });
    const first = document.createElement('button');
    shell.panel.appendChild(first);
    shell.close();

    outside.focus();
    outside.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true }));
    expect(document.activeElement).toBe(outside);
    shell.dispose();
    outside.remove();
  });

  it('invokes onDismiss on Escape only while open', () => {
    const onDismiss = vi.fn();
    const shell = createDialogShell({ host, className: 'fc-x', ariaLabel: 'X', onDismiss });

    // Closed: ignored.
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    expect(onDismiss).not.toHaveBeenCalled();

    shell.open();
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    expect(onDismiss).toHaveBeenCalledTimes(1);

    // Non-Escape keys ignored.
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }));
    expect(onDismiss).toHaveBeenCalledTimes(1);

    shell.dispose();
  });

  it('invokes onDismiss on backdrop click but not on panel-internal click', () => {
    const onDismiss = vi.fn();
    const shell = createDialogShell({ host, className: 'fc-x', ariaLabel: 'X', onDismiss });
    shell.open();

    // Click on the panel (not the overlay) → no dismiss.
    shell.panel.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(onDismiss).not.toHaveBeenCalled();

    // Click whose target IS the overlay → dismiss.
    shell.overlay.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(onDismiss).toHaveBeenCalledTimes(1);

    shell.dispose();
  });

  it('does NOT install Escape/backdrop listeners when onDismiss is omitted', () => {
    const shell = createDialogShell({ host, className: 'fc-x', ariaLabel: 'X' });
    shell.open();
    // Should not throw and should not log.
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    shell.overlay.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(shell.isOpen()).toBe(true); // Still open — nothing dismissed it.
    shell.dispose();
  });

  it('on() tracks listeners and dispose() removes every one of them', () => {
    const shell = createDialogShell({ host, className: 'fc-x', ariaLabel: 'X' });
    const handlerA = vi.fn();
    const handlerB = vi.fn();

    const button = document.createElement('button');
    shell.panel.appendChild(button);
    shell.on(button, 'click', handlerA);
    shell.on(document, 'keydown', handlerB);

    button.click();
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'a' }));
    expect(handlerA).toHaveBeenCalledTimes(1);
    expect(handlerB).toHaveBeenCalledTimes(1);

    shell.dispose();

    button.click();
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'a' }));
    expect(handlerA).toHaveBeenCalledTimes(1);
    expect(handlerB).toHaveBeenCalledTimes(1);
  });

  it('dispose is idempotent and removes the overlay from the DOM', () => {
    const shell = createDialogShell({ host, className: 'fc-x', ariaLabel: 'X' });
    expect(document.body.contains(shell.overlay)).toBe(true);
    shell.dispose();
    expect(document.body.contains(shell.overlay)).toBe(false);
    // Calling again is a no-op.
    expect(() => shell.dispose()).not.toThrow();
  });

  it('after dispose, on()/open() are inert and do not re-attach', () => {
    const shell = createDialogShell({ host, className: 'fc-x', ariaLabel: 'X' });
    shell.dispose();
    const handler = vi.fn();
    const btn = document.createElement('button');
    shell.on(btn, 'click', handler);
    btn.click();
    expect(handler).not.toHaveBeenCalled();

    shell.open();
    expect(shell.isOpen()).toBe(false); // overlay was already removed; hidden flag flipped but element gone
  });

  it('setAriaLabel mutates the live aria-label without remounting', () => {
    const shell = createDialogShell({ host, className: 'fc-x', ariaLabel: 'Old' });
    expect(shell.overlay.getAttribute('aria-label')).toBe('Old');
    shell.setAriaLabel('New');
    expect(shell.overlay.getAttribute('aria-label')).toBe('New');
    shell.dispose();
  });

  it('does not leak document keydown listeners across many dispose cycles', () => {
    // Functional regression check: 50 mount/dispose cycles must not result in
    // the document accumulating active listeners that still fire onDismiss.
    const onDismiss = vi.fn();
    for (let i = 0; i < 50; i++) {
      const shell = createDialogShell({ host, className: 'fc-x', ariaLabel: 'X', onDismiss });
      shell.open();
      shell.dispose();
    }
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    expect(onDismiss).not.toHaveBeenCalled();
  });
});
