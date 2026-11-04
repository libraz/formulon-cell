import { describe, expect, it } from 'vitest';
import { macRibbonLabelForCommand } from '../../../src/toolbar/ribbon/mac/labels.js';

describe('Mac View tab', () => {
  it('names the zoom commands as the desktop does', () => {
    expect(macRibbonLabelForCommand('mac.view.zoom100', 'ja')).toBe('ズーム (100%)');
    expect(macRibbonLabelForCommand('mac.view.zoom100', 'en')).toBe('Zoom (100%)');
    expect(macRibbonLabelForCommand('mac.view.zoomSelection', 'ja')).toBe(
      '選択範囲に合わせて拡大/縮小',
    );
  });
});
