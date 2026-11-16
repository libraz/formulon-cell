import { afterEach, describe, expect, it, vi } from 'vitest';
import { dictionaries } from '../../../../src/i18n/strings.js';
import { createSelectColorRibbon } from '../../../../src/toolbar/ribbon/select-color.js';
import { buildRibbonModel } from '../../../../src/toolbar/ribbon-model.js';

afterEach(() => {
  vi.useRealTimers();
  vi.restoreAllMocks();
  document.body.replaceChildren();
});

const build = (id: 'fontSize' | 'fontColor') => {
  const command = buildRibbonModel('en')
    .find((tab) => tab.id === 'home')
    ?.groups.flatMap((group) => group.commands)
    .find((candidate) => candidate.id === id);
  if (!command) throw new Error('command unavailable');
  const ribbon = createSelectColorRibbon({
    ribbonLang: 'en',
    ribbonText: dictionaries.en.ribbon,
    pageScaleText: { automatic: 'Automatic', page: 'Page', pages: 'Pages' },
    getInst: () => null,
    currentRibbonControlValue: () => '11',
    applyRibbonControl: () => {},
    shouldShowFontOption: () => true,
    createRibbonIcon: () => null,
  });
  const wrap =
    id === 'fontSize' ? ribbon.createRibbonSelect(command) : ribbon.createRibbonColor(command);
  document.body.append(wrap);
  const trigger = wrap.querySelector<HTMLButtonElement>(
    '.fc-tb__rb-dd__btn, .fc-tb__rb-color__btn',
  );
  if (!trigger) throw new Error('trigger unavailable');
  return { ribbon, wrap, trigger };
};

describe('ribbon dropdown outside-press timer', () => {
  for (const id of ['fontSize', 'fontColor'] as const) {
    it(`${id}: closing before the deferred registration installs no listener`, () => {
      vi.useFakeTimers();
      const { trigger } = build(id);
      trigger.click(); // open
      trigger.click(); // close before the timer fires
      const spy = vi.spyOn(document, 'addEventListener');
      vi.runAllTimers();
      expect(spy.mock.calls.filter(([type]) => type === 'mousedown')).toHaveLength(0);
    });

    it(`${id}: closing through closeOpenRibbonDropdowns installs no listener`, () => {
      vi.useFakeTimers();
      const { ribbon, trigger } = build(id);
      trigger.click();
      ribbon.closeOpenRibbonDropdowns();
      const spy = vi.spyOn(document, 'addEventListener');
      vi.runAllTimers();
      expect(spy.mock.calls.filter(([type]) => type === 'mousedown')).toHaveLength(0);
    });

    it(`${id}: reopening before the deferred registration installs one listener`, () => {
      vi.useFakeTimers();
      const { ribbon, trigger } = build(id);
      trigger.click();
      ribbon.closeOpenRibbonDropdowns();
      trigger.click();
      const spy = vi.spyOn(document, 'addEventListener');
      vi.runAllTimers();
      expect(spy.mock.calls.filter(([type]) => type === 'mousedown')).toHaveLength(1);
    });
  }
});
