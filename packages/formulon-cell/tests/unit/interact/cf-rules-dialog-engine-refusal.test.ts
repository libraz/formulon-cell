import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { en } from '../../../src/i18n/strings.js';
import { attachCfRulesDialog } from '../../../src/interact/cf-rules-dialog.js';
import { createSpreadsheetStore } from '../../../src/store/store.js';

const sqref = [{ firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 0 }];

describe('cf rules dialog with an engine that refuses a rewrite', () => {
  let host: HTMLElement;
  let wb: WorkbookHandle;

  beforeEach(async () => {
    host = document.createElement('div');
    document.body.appendChild(host);
    wb = await WorkbookHandle.createDefault();
  });

  afterEach(() => {
    wb.dispose();
    while (document.body.firstChild) document.body.removeChild(document.body.firstChild);
  });

  /** Refuses any add that carries `field === value`, as a refused formula would. */
  const refuseAdds = (field: 'priority' | 'stopIfTrue', value: number | boolean): void => {
    const add = wb.addConditionalFormat.bind(wb);
    wb.addConditionalFormat = (sheet, input) =>
      (input as unknown as Record<string, unknown>)[field] === value ? -1 : add(sheet, input);
  };
  const alertText = (): string | null => {
    const row = document.querySelector<HTMLElement>('.fc-cfrulesdlg [role="alert"]');
    return row && !row.hidden ? row.textContent : null;
  };
  const q = <T extends HTMLElement>(sel: string): T => {
    const el = document.querySelector<T>(sel);
    if (!el) throw new Error(`missing ${sel}`);
    return el;
  };
  const attach = (platform?: 'mac') => {
    if (platform) host.dataset.fcPlatform = platform;
    const handle = attachCfRulesDialog({
      host,
      getWb: () => wb,
      getActiveSheet: () => 0,
      store: createSpreadsheetStore(),
      strings: en,
    });
    handle.open();
    return handle;
  };

  it('shows the refusal when a loaded rule cannot take Stop If True', () => {
    wb.addConditionalFormat(0, { ranges: sqref, type: 17, priority: 4 });
    refuseAdds('stopIfTrue', true);
    const handle = attach();
    const box = q<HTMLInputElement>('.fc-cfrulesdlg__cell-stop input');
    box.checked = true;
    box.dispatchEvent(new Event('change'));

    expect(wb.getConditionalFormats(0)).toMatchObject([{ priority: 4, stopIfTrue: false }]);
    expect(alertText()).toBe(en.commandRefusal.engineFailed);
    handle.detach();
  });

  it('shows the refusal when a loaded rule cannot change priority', () => {
    wb.addConditionalFormat(0, { ranges: sqref, type: 17, priority: 1 });
    wb.addConditionalFormat(0, { ranges: sqref, type: 16, priority: 2 });
    refuseAdds('priority', 2);
    const handle = attach();
    q<HTMLButtonElement>('.fc-cfrulesdlg__move-down').click();

    expect(
      [...wb.getConditionalFormats(0)]
        .sort((a, b) => a.priority - b.priority)
        .map((e) => [e.type, e.priority]),
    ).toEqual([
      [17, 1],
      [16, 2],
    ]);
    expect(alertText()).toBe(en.commandRefusal.engineFailed);
    handle.detach();
  });

  it('keeps the Mac manager open and shows the refusal when OK cannot apply a staged edit', () => {
    wb.addConditionalFormat(0, { ranges: sqref, type: 17, priority: 4 });
    refuseAdds('stopIfTrue', true);
    const handle = attach('mac');
    const box = q<HTMLInputElement>('.fc-cfrulesdlg__cell-stop input');
    box.checked = true;
    box.dispatchEvent(new Event('change'));
    q<HTMLButtonElement>('.fc-cfrulesdlg__footer .fc-fmtdlg__btn--primary').click();

    expect(wb.getConditionalFormats(0)).toMatchObject([{ priority: 4, stopIfTrue: false }]);
    expect(document.querySelector<HTMLElement>('.fc-cfrulesdlg')?.hidden).toBe(false);
    expect(alertText()).toBe(en.commandRefusal.engineFailed);
    handle.detach();
  });
});
