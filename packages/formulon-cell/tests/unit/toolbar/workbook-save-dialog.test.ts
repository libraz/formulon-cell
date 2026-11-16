import { afterEach, describe, expect, it } from 'vitest';
import { defaultStrings } from '../../../src/i18n/strings.js';
import { confirmWorkbookSaveLosses } from '../../../src/toolbar/dialogs/workbook-save.js';

const losses = [{ id: 'notes', certain: true }] as never;

const open = (): Promise<boolean> => confirmWorkbookSaveLosses(defaultStrings, 'xlsx', losses);

const pressEnter = (target: Element, init: KeyboardEventInit = {}): KeyboardEvent => {
  const event = new KeyboardEvent('keydown', {
    key: 'Enter',
    bubbles: true,
    cancelable: true,
    ...init,
  });
  target.dispatchEvent(event);
  return event;
};

describe('workbook save confirmation dialogs', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('confirms with Enter pressed on the dialog body', async () => {
    const result = open();
    const body = document.body.querySelector('.fc-tb__dlg__body, .fc-tb__dlg') as Element;
    pressEnter(body);
    await expect(result).resolves.toBe(true);
  });

  it('ignores an Enter that confirms IME composition', async () => {
    const result = open();
    const body = document.body.querySelector('.fc-tb__dlg__body, .fc-tb__dlg') as Element;
    pressEnter(body, { isComposing: true });
    expect(document.body.querySelector('.fc-tb__dlg')).toBeTruthy();
    document.body.querySelector('button')?.click();
    await expect(result).resolves.toBe(false);
  });
});
