import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { dictionaries } from '../../../src/i18n/strings.js';
import { createFormatDialogView } from '../../../src/interact/format-dialog-view.js';

const en = dictionaries.en;

describe('interact/format-dialog-view', () => {
  let host: HTMLElement;

  beforeEach(() => {
    host = document.createElement('div');
    document.body.appendChild(host);
  });

  afterEach(() => {
    host.remove();
  });

  it('renders the overlay hidden + with role="dialog" aria-modal="true"', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    expect(view.overlay.hidden).toBe(true);
    expect(view.overlay.getAttribute('role')).toBe('dialog');
    expect(view.overlay.getAttribute('aria-modal')).toBe('true');
    expect(view.overlay.getAttribute('aria-label')).toBe(en.formatDialog.title);
  });

  it('creates one button + one panel per tab id', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    const expectedTabs = ['number', 'align', 'font', 'border', 'fill', 'protection', 'more'];
    for (const id of expectedTabs) {
      expect(view.tabButtons.has(id as never), `button missing for tab ${id}`).toBe(true);
      expect(view.tabPanels.has(id as never), `panel missing for tab ${id}`).toBe(true);
    }
  });

  it('starts every tab panel hidden — the controller toggles based on state', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    for (const panel of view.tabPanels.values()) {
      expect(panel.hidden).toBe(true);
    }
  });

  it('renders the 11 number-category buttons with role="option"', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    expect(view.catButtons.size).toBe(11);
    for (const btn of view.catButtons.values()) {
      expect(btn.getAttribute('role')).toBe('option');
    }
  });

  it('renders both alignment fieldsets (horizontal + vertical) with 4 radios each', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    expect(view.hAlignRadios.size).toBe(4); // default + left + center + right
    expect(view.vAlignRadios.size).toBe(4); // default + top + middle + bottom
  });

  it('wires data-fc-check on the font + alignment checkboxes', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    expect(view.boldCk.input.dataset.fcCheck).toBe('bold');
    expect(view.italicCk.input.dataset.fcCheck).toBe('italic');
    expect(view.underlineCk.input.dataset.fcCheck).toBe('underline');
    expect(view.strikeCk.input.dataset.fcCheck).toBe('strike');
    expect(view.wrapCk.input.dataset.fcCheck).toBe('wrap');
    expect(view.lockedCk.input.dataset.fcCheck).toBe('locked');
  });

  it('appends the overlay to the supplied host (not document.body) so callers control mount point', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    expect(host.contains(view.overlay)).toBe(true);
  });

  it('renders 6 border-style buttons with the right data attribute', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    expect(view.borderStyleButtons.size).toBe(6);
    const ids = Array.from(view.borderStyleButtons.keys()).sort();
    expect(ids).toEqual(['dashed', 'dotted', 'double', 'medium', 'thick', 'thin']);
    for (const [id, btn] of view.borderStyleButtons.entries()) {
      expect(btn.dataset.borderStyle).toBe(id);
      expect(btn.getAttribute('aria-pressed')).toBe('false');
    }
  });

  it('renders the validation kind/op selectors with the expected option counts', () => {
    const view = createFormatDialogView({ host, strings: en, t: en.formatDialog });
    expect(view.validationKindSelect.options.length).toBe(8); // none + 7 kinds
    expect(view.validationOpSelect.options.length).toBe(8); // between/notBetween/=/<>/</<=/>/>=
  });
});
