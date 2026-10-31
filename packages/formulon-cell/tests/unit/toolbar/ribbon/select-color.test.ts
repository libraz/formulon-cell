import { afterEach, describe, expect, it } from 'vitest';
import { dictionaries } from '../../../../src/i18n/strings.js';
import { createSelectColorRibbon } from '../../../../src/toolbar/ribbon/select-color.js';
import { buildRibbonModel } from '../../../../src/toolbar/ribbon-model.js';

const cleanups: Array<() => void> = [];
afterEach(async () => {
  // Let the deferred outside-click listener install before closing the picker.
  await new Promise<void>((resolve) => setTimeout(resolve, 0));
  for (const cleanup of cleanups.splice(0)) cleanup();
});

function picker(target?: HTMLElement, id: 'fontSize' | 'fontColor' | 'fillColor' = 'fontSize') {
  const command = buildRibbonModel('en')
    .find((tab) => tab.id === 'home')
    ?.groups.flatMap((group) => group.commands)
    .find((command) => command.id === id);
  if (!command) throw new Error('Font size command unavailable');
  let current = id === 'fontSize' ? '11' : '#000000';
  const applied: string[] = [];
  let openDuringDispatch = false;
  const ribbon = createSelectColorRibbon({
    ribbonLang: 'en',
    ribbonText: dictionaries.en.ribbon,
    pageScaleText: { automatic: 'Automatic', page: 'Page', pages: 'Pages' },
    getInst: () => null,
    currentRibbonControlValue: () => current,
    applyRibbonControl: (_id, value) => {
      openDuringDispatch =
        select.classList.contains('fc-tb__rb-dd--open') ||
        select.classList.contains('fc-tb__rb-color--open');
      current = value;
      applied.push(value);
      target?.focus();
    },
    shouldShowFontOption: () => true,
    createRibbonIcon: () => null,
  });
  const select =
    id === 'fontSize' ? ribbon.createRibbonSelect(command) : ribbon.createRibbonColor(command);
  document.body.append(select);
  if (target) document.body.append(target);
  cleanups.push(() => {
    ribbon.closeOpenRibbonDropdowns();
    select.remove();
    target?.remove();
  });
  const trigger = select.querySelector<HTMLButtonElement>(
    '.fc-tb__rb-dd__btn, .fc-tb__rb-color__btn',
  );
  if (!trigger) throw new Error('Font picker trigger unavailable');
  trigger.click();
  const option = select.querySelector<HTMLButtonElement>(
    id === 'fontSize' ? '[role="option"][data-value="24"]' : 'button[data-color]',
  );
  if (!option) throw new Error('24 point option unavailable');
  return { select, trigger, option, applied, openDuringDispatch: () => openDuringDispatch };
}

describe('ribbon select focus after dispatch', () => {
  it('keeps sheet focus chosen by a clicked formatting command', () => {
    const sheet = document.createElement('div');
    sheet.tabIndex = 0;
    const control = picker(sheet);
    control.option.click();
    expect(document.activeElement).toBe(sheet);
    expect(control.openDuringDispatch()).toBe(false);
    expect(control.applied).toEqual(['24']);
  });

  it('keeps dialog input focus chosen by a keyboard selection', () => {
    const input = document.createElement('input');
    const control = picker(input);
    control.option.focus();
    control.option.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    expect(document.activeElement).toBe(input);
    expect(control.openDuringDispatch()).toBe(false);
    expect(control.applied).toEqual(['24']);
  });

  it('returns to the trigger when the command does not choose another focus target', () => {
    const control = picker();
    control.option.click();
    expect(document.activeElement).toBe(control.trigger);
    expect(control.trigger.getAttribute('aria-expanded')).toBe('false');
    expect(control.select.querySelector('[role="listbox"]')).toBeNull();
  });

  it('returns to the trigger without applying an option on Escape', () => {
    const control = picker();
    control.option.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(document.activeElement).toBe(control.trigger);
    expect(control.applied).toEqual([]);
  });

  it.each(['fontColor', 'fillColor'] as const)(
    'keeps sheet focus after choosing a %s swatch',
    (id) => {
      const sheet = document.createElement('div');
      sheet.tabIndex = 0;
      const control = picker(sheet, id);
      control.option.click();
      expect(document.activeElement).toBe(sheet);
      expect(control.openDuringDispatch()).toBe(false);
      expect(control.applied).toEqual([control.option.dataset.color]);
    },
  );
});
