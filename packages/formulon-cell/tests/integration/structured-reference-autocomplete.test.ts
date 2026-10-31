import { afterEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { type AutocompleteHandle, attachAutocomplete } from '../../src/interact/autocomplete.js';

describe('integration: structured-reference autocomplete', () => {
  let workbook: WorkbookHandle | null = null;
  let handle: AutocompleteHandle | null = null;
  let input: HTMLTextAreaElement | null = null;

  afterEach(() => {
    handle?.detach();
    input?.remove();
    if (workbook && !workbook.isStub) workbook.dispose();
    handle = null;
    input = null;
    workbook = null;
  });

  it('encodes an escaped special header into a formula that calculates correctly', async () => {
    workbook = await WorkbookHandle.createDefault();
    expect(workbook.isStub).toBe(false);

    workbook.setText({ sheet: 0, row: 0, col: 0 }, 'A]B');
    workbook.setNumber({ sheet: 0, row: 1, col: 0 }, 7);
    workbook.setNumber({ sheet: 0, row: 2, col: 0 }, 8);
    expect(
      workbook.createTable({
        sheetIndex: 0,
        ref: 'A1:A3',
        name: 'Sales',
        columns: ['A]B'],
      }),
    ).toBe(0);
    expect(workbook.getTables()[0]?.columns).toEqual(['A]B']);

    const resultCell = { sheet: 0, row: 0, col: 1 };
    workbook.setFormula(resultCell, '=SUM(A2:A3)');
    const plainColumnResult = workbook.getValue(resultCell);
    expect(plainColumnResult).toEqual({ kind: 'number', value: 15 });

    input = document.createElement('textarea');
    input.value = "=SUM(Sales[A')";
    input.setSelectionRange(input.value.indexOf(')'), input.value.indexOf(')'));
    document.body.append(input);
    handle = attachAutocomplete({ input, getTables: () => workbook?.getTables() ?? [] });
    handle.refresh();

    expect(document.querySelector('.fc-autocomplete__name')?.textContent).toBe('A]B');
    expect(handle.acceptHighlighted()).toBe(true);
    expect(input.value).toBe("=SUM(Sales[[A']B]])");

    workbook.setFormula(resultCell, input.value);
    expect(workbook.getValue(resultCell)).toEqual(plainColumnResult);
  });
});
