import { afterEach, describe, expect, it, vi } from 'vitest';
import { disposeOverlayPortal, ensureOverlayPortal } from '../../../src/interact/overlay-portal.js';
import {
  createDialogShell,
  installDialogLifecycle,
  mountDialog,
} from '../../../src/toolbar/dialogs/shell.js';

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

  it('uses a checked second named radio as the first tab stop for Shift+Tab', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const shell = createDialogShell({ title: 'Radio shift', host });
    const firstRadio = document.createElement('input');
    firstRadio.type = 'radio';
    firstRadio.name = 'choice';
    const checkedRadio = document.createElement('input');
    checkedRadio.type = 'radio';
    checkedRadio.name = 'choice';
    checkedRadio.checked = true;
    const last = document.createElement('button');
    shell.body.append(firstRadio, checkedRadio, last);
    const lifecycle = installDialogLifecycle({
      shell,
      resolve: () => undefined,
      onCancel: () => undefined,
    });

    try {
      checkedRadio.focus();
      const event = new KeyboardEvent('keydown', {
        key: 'Tab',
        shiftKey: true,
        bubbles: true,
        cancelable: true,
      });
      checkedRadio.dispatchEvent(event);

      expect(event.defaultPrevented).toBe(true);
      expect(document.activeElement).toBe(last);
    } finally {
      lifecycle.finish(undefined);
      host.remove();
    }
  });

  it('wraps from the last control to the checked second named radio', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const shell = createDialogShell({ title: 'Radio forward', host });
    const firstRadio = document.createElement('input');
    firstRadio.type = 'radio';
    firstRadio.name = 'choice';
    const checkedRadio = document.createElement('input');
    checkedRadio.type = 'radio';
    checkedRadio.name = 'choice';
    checkedRadio.checked = true;
    const last = document.createElement('button');
    shell.body.append(firstRadio, checkedRadio, last);
    const lifecycle = installDialogLifecycle({
      shell,
      resolve: () => undefined,
      onCancel: () => undefined,
    });

    try {
      last.focus();
      const event = new KeyboardEvent('keydown', {
        key: 'Tab',
        bubbles: true,
        cancelable: true,
      });
      last.dispatchEvent(event);

      expect(event.defaultPrevented).toBe(true);
      expect(document.activeElement).toBe(checkedRadio);
    } finally {
      lifecycle.finish(undefined);
      host.remove();
    }
  });

  it('advances between a checked radio and Cancel in both Tab directions', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const shell = createDialogShell({ title: 'Radio interior', host });
    const firstRadio = document.createElement('input');
    firstRadio.type = 'radio';
    firstRadio.name = 'choice';
    const checkedRadio = document.createElement('input');
    checkedRadio.type = 'radio';
    checkedRadio.name = 'choice';
    checkedRadio.checked = true;
    const cancel = document.createElement('button');
    const ok = document.createElement('button');
    shell.body.append(firstRadio, checkedRadio, cancel, ok);
    const lifecycle = installDialogLifecycle({
      shell,
      resolve: () => undefined,
      onCancel: () => undefined,
    });

    try {
      checkedRadio.focus();
      const forward = new KeyboardEvent('keydown', {
        key: 'Tab',
        bubbles: true,
        cancelable: true,
      });
      checkedRadio.dispatchEvent(forward);
      expect(forward.defaultPrevented).toBe(true);
      expect(document.activeElement).toBe(cancel);
      expect(firstRadio.checked).toBe(false);
      expect(checkedRadio.checked).toBe(true);

      const backward = new KeyboardEvent('keydown', {
        key: 'Tab',
        shiftKey: true,
        bubbles: true,
        cancelable: true,
      });
      cancel.dispatchEvent(backward);
      expect(backward.defaultPrevented).toBe(true);
      expect(document.activeElement).toBe(checkedRadio);
      expect(checkedRadio.checked).toBe(true);
    } finally {
      lifecycle.finish(undefined);
      host.remove();
    }
  });

  it('submits Enter only for an unhandled non-composing input key', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const shell = createDialogShell({ title: 'Submit guard', host });
    const cancel = document.createElement('button');
    cancel.textContent = 'Cancel';
    const input = document.createElement('input');
    shell.body.append(cancel, input);
    const submit = vi.fn();
    const cancelClick = vi.fn();
    cancel.addEventListener('click', cancelClick);
    const lifecycle = installDialogLifecycle({
      shell,
      resolve: () => undefined,
      onSubmit: submit,
      onCancel: () => undefined,
    });

    try {
      cancel.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }),
      );
      expect(submit).not.toHaveBeenCalled();
      cancel.click();
      expect(cancelClick).toHaveBeenCalledTimes(1);

      input.dispatchEvent(
        new KeyboardEvent('keydown', {
          key: 'Enter',
          bubbles: true,
          cancelable: true,
          isComposing: true,
        }),
      );
      expect(submit).not.toHaveBeenCalled();

      input.addEventListener('keydown', (event) => event.preventDefault(), { once: true });
      const preventedEnter = new KeyboardEvent('keydown', {
        key: 'Enter',
        bubbles: true,
        cancelable: true,
      });
      input.dispatchEvent(preventedEnter);
      expect(preventedEnter.defaultPrevented).toBe(true);
      expect(submit).not.toHaveBeenCalled();

      input.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true }),
      );
      expect(submit).toHaveBeenCalledTimes(1);
    } finally {
      lifecycle.finish(undefined);
      host.remove();
    }
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
