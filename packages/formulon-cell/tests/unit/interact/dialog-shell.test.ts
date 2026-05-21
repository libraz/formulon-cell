import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  appendDialogActions,
  appendDialogButton,
  appendDialogFrame,
  appendDialogIconButton,
  appendDialogOptionButton,
  appendDialogTabPair,
  clearDialogError,
  createDialogButton,
  createDialogShell,
  createDialogToggleButton,
  focusAndSelectInput,
  showDialogError,
} from '../../../src/interact/dialog-shell.js';

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

  it('enhances native selects into keyboardable custom comboboxes without breaking change events', () => {
    const shell = createDialogShell({ host, className: 'fc-x', ariaLabel: 'X' });
    const select = document.createElement('select');
    select.setAttribute('aria-label', 'Number format');
    for (const { value, label } of [
      { value: 'general', label: 'General' },
      { value: 'currency', label: 'Currency' },
      { value: 'date', label: 'Date' },
    ]) {
      const option = document.createElement('option');
      option.value = value;
      option.textContent = label;
      select.appendChild(option);
    }
    const onChange = vi.fn();
    select.addEventListener('change', onChange);
    shell.panel.appendChild(select);

    shell.open();

    const combo = shell.panel.querySelector<HTMLButtonElement>('.fc-select__button');
    expect(combo).not.toBeNull();
    expect(select.classList.contains('fc-select__native')).toBe(true);
    expect(combo?.textContent).toContain('General');

    combo?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    combo?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));

    expect(select.value).toBe('currency');
    expect(onChange).toHaveBeenCalledTimes(1);
    expect(combo?.textContent).toContain('Currency');
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

  it('builds shared dialog frame and action button contracts', () => {
    const shell = createDialogShell({ host, className: 'fc-x', ariaLabel: 'X' });
    const frame = appendDialogFrame(shell, {
      title: 'Title',
      panelClasses: ['fc-fmtdlg__panel', 'fc-x__panel'],
      bodyClass: 'fc-fmtdlg__body fc-x__body',
    });

    expect(shell.panel.classList.contains('fc-fmtdlg__panel')).toBe(true);
    expect(shell.panel.classList.contains('fc-x__panel')).toBe(true);
    expect(frame.header.className).toBe('fc-fmtdlg__header');
    expect(frame.header.textContent).toBe('Title');
    expect(frame.body.className).toBe('fc-fmtdlg__body fc-x__body');
    expect(frame.footer.className).toBe('fc-fmtdlg__footer');

    const { cancelBtn, okBtn } = appendDialogActions(frame.footer, {
      cancelLabel: 'Cancel',
      okLabel: 'OK',
    });
    expect(cancelBtn.type).toBe('button');
    expect(cancelBtn.className).toBe('fc-fmtdlg__btn');
    expect(cancelBtn.textContent).toBe('Cancel');
    expect(okBtn.type).toBe('button');
    expect(okBtn.className).toBe('fc-fmtdlg__btn fc-fmtdlg__btn--primary');
    expect(okBtn.textContent).toBe('OK');

    const standalone = createDialogButton({ label: 'Apply', variant: 'primary' });
    expect(standalone.type).toBe('button');
    expect(standalone.className).toBe('fc-fmtdlg__btn fc-fmtdlg__btn--primary');
    expect(standalone.textContent).toBe('Apply');

    const custom = appendDialogButton(frame.footer, {
      label: 'Close',
      variant: 'primary',
      baseClass: 'fc-custom__btn',
    });
    expect(custom.className).toBe('fc-custom__btn fc-custom__btn--primary');
    expect(custom.textContent).toBe('Close');

    const customSecondary = appendDialogButton(frame.footer, {
      label: 'Cancel',
      variant: 'secondary',
      baseClass: 'fc-custom__btn',
      secondaryClass: 'fc-custom__btn--secondary',
    });
    expect(customSecondary.className).toBe('fc-custom__btn fc-custom__btn--secondary');

    const icon = appendDialogIconButton(frame.header, {
      label: '×',
      ariaLabel: 'Close',
      title: 'Close',
      baseClass: 'fc-x__close',
    });
    expect(icon.type).toBe('button');
    expect(icon.className).toBe('fc-x__close');
    expect(icon.textContent).toBe('×');
    expect(icon.getAttribute('aria-label')).toBe('Close');
    expect(icon.title).toBe('Close');

    const tabs = document.createElement('div');
    const { button: tab, panel: tabPanel } = appendDialogTabPair(tabs, frame.body, {
      id: 'page',
      label: 'Page',
      tabId: 'tab-page',
      panelId: 'panel-page',
      tabDatasetKey: 'testTab',
      panelDatasetKey: 'testTab',
    });
    expect(tab.type).toBe('button');
    expect(tab.className).toBe('fc-fmtdlg__tab');
    expect(tab.getAttribute('role')).toBe('tab');
    expect(tab.getAttribute('aria-selected')).toBe('false');
    expect(tab.getAttribute('aria-controls')).toBe('panel-page');
    expect(tab.dataset.testTab).toBe('page');
    expect(tabPanel.className).toBe('fc-fmtdlg__panel-tab');
    expect(tabPanel.getAttribute('role')).toBe('tabpanel');
    expect(tabPanel.getAttribute('aria-labelledby')).toBe('tab-page');
    expect(tabPanel.dataset.testTab).toBe('page');
    expect(tabPanel.hidden).toBe(true);

    const option = appendDialogOptionButton(frame.body, {
      label: 'Fixed',
      baseClass: 'fc-test__option',
      datasetKey: 'testOption',
      value: 'fixed',
      selected: true,
      extraClass: 'fc-test__option--accent',
    });
    expect(option.type).toBe('button');
    expect(option.className).toBe('fc-test__option fc-test__option--accent');
    expect(option.textContent).toBe('Fixed');
    expect(option.getAttribute('role')).toBe('option');
    expect(option.getAttribute('aria-selected')).toBe('true');
    expect(option.tabIndex).toBe(-1);
    expect(option.dataset.testOption).toBe('fixed');

    const toggle = createDialogToggleButton({
      label: 'Top border',
      baseClass: 'fc-test__toggle',
      extraClass: 'fc-test__toggle--top',
      datasetKey: 'borderSide',
      value: 'top',
      title: 'Top border',
    });
    expect(toggle.type).toBe('button');
    expect(toggle.className).toBe('fc-test__toggle fc-test__toggle--top');
    expect(toggle.getAttribute('aria-label')).toBe('Top border');
    expect(toggle.getAttribute('aria-pressed')).toBe('false');
    expect(toggle.dataset.borderSide).toBe('top');
    expect(toggle.title).toBe('Top border');
    shell.dispose();
  });

  it('can build a shared dialog frame with a form body', () => {
    const shell = createDialogShell({ host, className: 'fc-x', ariaLabel: 'X' });
    const frame = appendDialogFrame(shell, { title: 'Form', bodyTag: 'form' });

    expect(frame.body).toBeInstanceOf(HTMLFormElement);
    expect(frame.body.className).toBe('fc-fmtdlg__body');
    shell.dispose();
  });

  it('centralizes input focus/select and inline error row updates', () => {
    const input = document.createElement('input');
    input.value = 'Sheet1!A1';
    document.body.appendChild(input);
    focusAndSelectInput(input);
    expect(document.activeElement).toBe(input);
    expect(input.selectionStart).toBe(0);
    expect(input.selectionEnd).toBe(input.value.length);

    const errorRow = document.createElement('div');
    errorRow.hidden = true;
    showDialogError(errorRow, 'Required');
    expect(errorRow.hidden).toBe(false);
    expect(errorRow.textContent).toBe('Required');
    clearDialogError(errorRow);
    expect(errorRow.hidden).toBe(true);
    expect(errorRow.textContent).toBe('');

    input.remove();
  });

  it('keeps migrated dialog error/focus handling centralized in dialog-shell', () => {
    for (const file of ['src/interact/hyperlink-dialog.ts', 'src/interact/named-range-dialog.ts']) {
      const source = readFileSync(join(process.cwd(), file), 'utf8');
      expect(source).toContain('showDialogError(');
      expect(source).toContain('clearDialogError(');
      expect(source).toContain('focusAndSelectInput(');
      expect(source).not.toContain('errorRow.textContent');
      expect(source).not.toContain('errorRow.hidden = false');
      expect(source).not.toContain('.select();');
    }
  });

  it('keeps migrated dialog frame/action DOM centralized in dialog-shell', () => {
    for (const file of [
      'src/interact/paste-special.ts',
      'src/interact/goto-dialog.ts',
      'src/interact/evaluate-formula-dialog.ts',
      'src/interact/pivot-table-dialog.ts',
      'src/interact/iterative-dialog.ts',
      'src/interact/hyperlink-dialog.ts',
      'src/interact/fx-dialog.ts',
      'src/interact/pivot-field-settings.ts',
    ]) {
      const source = readFileSync(join(process.cwd(), file), 'utf8');
      expect(source).toContain('appendDialogFrame(');
      expect(source).not.toContain("const footer = document.createElement('div')");
      expect(source).not.toContain("const cancelBtn = document.createElement('button')");
      expect(source).not.toContain("const okBtn = document.createElement('button')");
      expect(source).not.toContain("const insertBtn = document.createElement('button')");
    }
  });

  it('keeps migrated action buttons centralized even when dialogs own custom footer placement', () => {
    for (const file of [
      'src/interact/page-setup-dialog.ts',
      'src/interact/comment-dialog.ts',
      'src/interact/insert-copied-cells-dialog.ts',
    ]) {
      const source = readFileSync(join(process.cwd(), file), 'utf8');
      expect(source).toContain('appendDialogActions(footer');
      expect(source).not.toContain("const cancelBtn = document.createElement('button')");
      expect(source).not.toContain("const okBtn = document.createElement('button')");
    }
  });

  it('keeps migrated single action buttons centralized in dialog-shell', () => {
    for (const file of ['src/interact/external-links-dialog.ts', 'src/interact/validation.ts']) {
      const source = readFileSync(join(process.cwd(), file), 'utf8');
      expect(source).toContain('appendDialogButton(footer');
      expect(source).not.toContain("const closeBtn = document.createElement('button')");
      expect(source).not.toContain("const okBtn = document.createElement('button')");
    }
  });

  it('keeps migrated mixed footer/action buttons centralized in dialog-shell', () => {
    for (const file of [
      'src/interact/conditional-dialog.ts',
      'src/interact/cf-rules-dialog.ts',
      'src/interact/named-range-dialog.ts',
    ]) {
      const source = readFileSync(join(process.cwd(), file), 'utf8');
      expect(source).toContain('appendDialogButton(');
      expect(source).not.toContain("const closeBtn = document.createElement('button')");
      expect(source).not.toContain("const clearAllBtn = document.createElement('button')");
      expect(source).not.toContain("const editorOkBtn = document.createElement('button')");
      expect(source).not.toContain("const editorCancelBtn = document.createElement('button')");
      expect(source).not.toContain("const deleteOkBtn = document.createElement('button')");
      expect(source).not.toContain("const deleteCancelBtn = document.createElement('button')");
      expect(source).not.toContain("const addBtn = document.createElement('button')");
      expect(source).not.toContain("const newBtn = document.createElement('button')");
      expect(source).not.toContain("const editBtn = document.createElement('button')");
      expect(source).not.toContain("const deleteBtn = document.createElement('button')");
      expect(source).not.toContain("const filterBtn = document.createElement('button')");
    }
    const conditionalSource = readFileSync(
      join(process.cwd(), 'src/interact/conditional-dialog.ts'),
      'utf8',
    );
    expect(conditionalSource).not.toContain("const removeBtn = document.createElement('button')");
  });

  it('keeps migrated panel action buttons centralized in dialog-shell', () => {
    for (const file of ['src/interact/slicer.ts', 'src/interact/watch-panel.ts']) {
      const source = readFileSync(join(process.cwd(), file), 'utf8');
      expect(source).toContain('appendDialogButton(');
      expect(source).not.toContain("const addBtn = document.createElement('button')");
      expect(source).not.toContain("const clearBtn = document.createElement('button')");
      expect(source).not.toContain("const closeBtn = document.createElement('button')");
      expect(source).not.toContain("const removeBtn = document.createElement('button')");
    }
  });

  it('keeps migrated icon buttons centralized in dialog-shell', () => {
    for (const file of [
      'src/interact/comment-dialog.ts',
      'src/interact/format-dialog-view.ts',
      'src/interact/named-range-dialog.ts',
      'src/interact/page-setup-dialog.ts',
      'src/interact/workbook-objects.ts',
    ]) {
      const source = readFileSync(join(process.cwd(), file), 'utf8');
      expect(source).toContain('appendDialogIconButton(');
      expect(source).not.toContain("const closeBtn = document.createElement('button')");
      expect(source).not.toContain("const headerCloseBtn = document.createElement('button')");
      expect(source).not.toContain("const removeBtn = document.createElement('button')");
      expect(source).not.toContain("const quickCommitBtn = document.createElement('button')");
      expect(source).not.toContain("const quickCancelBtn = document.createElement('button')");
    }
  });

  it('keeps migrated shared dialog tabs centralized in dialog-shell', () => {
    for (const file of [
      'src/interact/format-dialog-view.ts',
      'src/interact/page-setup-dialog.ts',
    ]) {
      const source = readFileSync(join(process.cwd(), file), 'utf8');
      expect(source).toContain('appendDialogTabPair(');
      expect(source).not.toContain("className = 'fc-fmtdlg__tab'");
      expect(source).not.toContain("setAttribute('role', 'tab')");
      expect(source).not.toContain("setAttribute('role', 'tabpanel')");
    }
  });

  it('keeps migrated option buttons centralized in dialog-shell', () => {
    const source = readFileSync(join(process.cwd(), 'src/interact/format-dialog-view.ts'), 'utf8');
    expect(source).toContain('appendDialogOptionButton(');
    expect(source).not.toContain("className = 'fc-fmtdlg__cat-item'");
    expect(source).not.toContain("className = 'fc-fmtdlg__negative-item'");
    expect(source).not.toContain("setAttribute('role', 'option')");
  });

  it('keeps migrated format dialog action button helpers centralized in dialog-shell', () => {
    const source = readFileSync(join(process.cwd(), 'src/interact/format-dialog-dom.ts'), 'utf8');
    expect(source).toContain('createDialogButton(');
    expect(source).not.toContain("className = primary ? 'fc-fmtdlg__btn fc-fmtdlg__btn--primary'");
  });

  it('keeps migrated format dialog toggle buttons centralized in dialog-shell', () => {
    for (const file of [
      'src/interact/format-dialog-dom.ts',
      'src/interact/format-dialog-tabs/border.ts',
    ]) {
      const source = readFileSync(join(process.cwd(), file), 'utf8');
      expect(source).toContain('createDialogToggleButton(');
      expect(source).not.toContain("setAttribute('aria-pressed', 'false')");
    }
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
