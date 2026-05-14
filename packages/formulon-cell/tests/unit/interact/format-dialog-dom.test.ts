import { describe, expect, it } from 'vitest';

import {
  makeButton,
  makeCheckbox,
  makeListSourceRadio,
  makeSection,
  makeSwatches,
  makeVisualSideButton,
} from '../../../src/interact/format-dialog-dom.js';
import { type SideKey, THEME_SWATCHES } from '../../../src/interact/format-dialog-model.js';

describe('interact/format-dialog-dom', () => {
  describe('makeCheckbox', () => {
    it('wraps an input + label span and sets the dialog class', () => {
      const { wrap, input } = makeCheckbox('Wrap text');
      expect(wrap.tagName).toBe('LABEL');
      expect(wrap.className).toBe('fc-fmtdlg__check');
      expect(input.type).toBe('checkbox');
      expect(wrap.querySelector('span')?.textContent).toBe('Wrap text');
    });
  });

  describe('makeButton', () => {
    it('produces a non-primary button by default', () => {
      const b = makeButton('Cancel');
      expect(b.type).toBe('button');
      expect(b.className).toBe('fc-fmtdlg__btn');
      expect(b.textContent).toBe('Cancel');
    });

    it('adds the primary modifier when requested', () => {
      const b = makeButton('OK', true);
      expect(b.className).toBe('fc-fmtdlg__btn fc-fmtdlg__btn--primary');
    });
  });

  describe('makeSwatches', () => {
    it('renders 12 theme swatches with color metadata', () => {
      const group = makeSwatches('font');
      expect(group.dataset.swatches).toBe('font');
      const swatches = group.querySelectorAll('button');
      expect(swatches.length).toBe(THEME_SWATCHES.length);
      for (let i = 0; i < swatches.length; i += 1) {
        const btn = swatches[i] as HTMLButtonElement;
        expect(btn.dataset.color).toBe(THEME_SWATCHES[i]);
        expect(btn.getAttribute('aria-label')).toBe(THEME_SWATCHES[i]);
      }
    });
  });

  describe('makeVisualSideButton', () => {
    it('registers the button in the shared map and tags the side', () => {
      const map = new Map<SideKey, HTMLButtonElement[]>();
      const btn = makeVisualSideButton(map, 'top', 'Top border');
      expect(btn.dataset.borderSide).toBe('top');
      expect(btn.getAttribute('aria-label')).toBe('Top border');
      expect(btn.getAttribute('aria-pressed')).toBe('false');
      expect(map.get('top')).toEqual([btn]);
    });

    it('accumulates multiple buttons for the same side', () => {
      const map = new Map<SideKey, HTMLButtonElement[]>();
      makeVisualSideButton(map, 'left', 'Left');
      makeVisualSideButton(map, 'left', 'Left');
      expect(map.get('left')?.length).toBe(2);
    });
  });

  describe('makeSection', () => {
    it('builds a section with a titled heading', () => {
      const sec = makeSection('Alignment');
      expect(sec.className).toBe('fc-fmtdlg__section');
      expect(sec.querySelector('.fc-fmtdlg__section-title')?.textContent).toBe('Alignment');
    });
  });

  describe('makeListSourceRadio', () => {
    it('builds a radio in the shared validation-list-source group', () => {
      const { input } = makeListSourceRadio('range', 'Range');
      expect(input.type).toBe('radio');
      expect(input.name).toBe('fc-validation-list-source');
      expect(input.value).toBe('range');
    });
  });
});
