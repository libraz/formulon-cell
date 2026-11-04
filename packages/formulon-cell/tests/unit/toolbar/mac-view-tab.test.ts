import { afterEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { mutators } from '../../../src/store/store.js';
import { macRibbonLabelForCommand } from '../../../src/toolbar/ribbon/mac/labels.js';
import { projectMacRibbonState } from '../../../src/toolbar/ribbon/mac/state-projection.js';
import { type MountedStubSheet, mountStubSheet } from '../../test-utils/mount.js';

describe('Mac View tab', () => {
  it('names the zoom commands as the desktop does', () => {
    expect(macRibbonLabelForCommand('mac.view.zoom100', 'ja')).toBe('ズーム (100%)');
    expect(macRibbonLabelForCommand('mac.view.zoom100', 'en')).toBe('Zoom (100%)');
    expect(macRibbonLabelForCommand('mac.view.zoomSelection', 'ja')).toBe(
      '選択範囲に合わせて拡大/縮小',
    );
  });

  describe('toggles', () => {
    let sheet: MountedStubSheet | undefined;
    afterEach(() => {
      sheet?.dispose();
      sheet = undefined;
    });

    it('shows Formula Bar pressed while the bar is shown, Ruler on, Zoom to Selection off', async () => {
      sheet = await mountStubSheet({
        workbook: await WorkbookHandle.createDefault({ preferStub: true }),
        locale: 'ja',
      });
      const { instance } = sheet;
      const host = document.createElement('div');
      const button = (id: string) => {
        const el = document.createElement('button');
        el.dataset.ribbonCommand = id;
        host.appendChild(el);
        return el;
      };
      const formulaBar = button('mac.view.formulaBar');
      const ruler = button('mac.view.ruler');
      const zoomSelection = button('mac.view.zoomSelection');
      projectMacRibbonState(host, instance);
      expect(formulaBar.getAttribute('aria-pressed')).toBe('true');
      expect(ruler.getAttribute('aria-pressed')).toBe('true');
      expect(zoomSelection.getAttribute('aria-pressed')).toBe('false');
      mutators.setFormulaBarVisible(instance.store, false);
      projectMacRibbonState(host, instance);
      expect(formulaBar.getAttribute('aria-pressed')).toBe('false');
    });
  });
});
