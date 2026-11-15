import { afterEach, describe, expect, it } from 'vitest';
import { dictionaries } from '../../../../src/i18n/strings.js';
import { buildMacRibbonModel } from '../../../../src/toolbar/ribbon/mac/model.js';
import { createSelectColorRibbon } from '../../../../src/toolbar/ribbon/select-color.js';
import { buildRibbonModel, type RibbonTabModel } from '../../../../src/toolbar/ribbon-model.js';

const cleanups: Array<() => void> = [];
afterEach(() => {
  for (const cleanup of cleanups.splice(0)) cleanup();
});

const profiles: [string, () => RibbonTabModel[]][] = [
  ['default', () => buildRibbonModel('en')],
  [
    'mac',
    () => {
      const model = buildRibbonModel('en');
      const home = model.find((tab) => tab.id === 'home');
      if (!home) throw new Error('home tab unavailable');
      return buildMacRibbonModel('en', home);
    },
  ],
];

function select(model: RibbonTabModel[], id: string, values: { current: string }) {
  const command = model
    .find((tab) => tab.id === 'home')
    ?.groups.flatMap((group) => group.commands)
    .find((candidate) => candidate.id === id);
  if (!command) throw new Error(`${id} command unavailable`);
  const ribbon = createSelectColorRibbon({
    ribbonLang: 'en',
    ribbonText: dictionaries.en.ribbon,
    pageScaleText: { automatic: 'Automatic', page: 'Page', pages: 'Pages' },
    getInst: () => null,
    currentRibbonControlValue: () => values.current,
    applyRibbonControl: () => {},
    shouldShowFontOption: () => true,
    createRibbonIcon: () => null,
  });
  const wrap = ribbon.createRibbonSelect(command);
  document.body.append(wrap);
  cleanups.push(() => wrap.remove());
  const button = wrap.querySelector<HTMLButtonElement>('.fc-tb__rb-dd__btn');
  const value = wrap.querySelector<HTMLElement>('.fc-tb__rb-dd__value');
  if (!button || !value) throw new Error('dropdown parts unavailable');
  return { wrap, button, value, title: command.title };
}

describe.each(profiles)('ribbon dropdown accessible name (%s profile)', (_name, build) => {
  it.each([
    ['fontSize', '11', '14'],
    ['numberFormat', 'general', 'percent'],
    ['fontFamily', 'Arial', 'Georgia'],
  ])('%s: the name contains the visible value before and after a change', (id, first, second) => {
    const values = { current: first };
    const { wrap, button, value, title } = select(build(), id, values);
    expect(value.textContent).not.toBe('');
    expect(button.getAttribute('aria-label')).toContain(value.textContent?.trim());
    expect(button.getAttribute('aria-label')?.startsWith(title ?? '')).toBe(true);

    values.current = second;
    wrap.dispatchEvent(new Event('fc-ribbon-control-refresh'));
    expect(button.getAttribute('aria-label')).toContain(value.textContent?.trim());
    expect(button.getAttribute('aria-label')).not.toContain(first === '11' ? '11:' : 'General');
  });
});
