import { afterEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachAutocomplete } from '../../../src/interact/autocomplete.js';
import {
  workbookAutocompleteCatalog,
  workbookCustomFunctions,
  workbookFunctionHelp,
} from '../../../src/interact/workbook-function-catalog.js';

const handles: Array<{ detach(): void }> = [];
const books: WorkbookHandle[] = [];
afterEach(() => {
  for (const h of handles.splice(0)) h.detach();
  for (const wb of books.splice(0)) wb.dispose();
  document.body.replaceChildren();
});

/** The function names the formula editor lists for `text` typed on `sheet`, with no hints. */
const editorList = (wb: WorkbookHandle, sheet: number, text: string): string[] => {
  const input = document.createElement('textarea');
  document.body.appendChild(input);
  input.value = text;
  input.setSelectionRange(text.length, text.length);
  const custom = workbookCustomFunctions(
    () => wb,
    () => sheet,
    { list: () => [], args: () => undefined },
  );
  const handle = attachAutocomplete({
    input,
    ...custom,
    ...workbookAutocompleteCatalog(() => wb, custom.getCustomFunctionArgs),
  });
  handles.push(handle);
  handle.refresh();
  return Array.from(
    document.querySelectorAll('.fc-autocomplete__item .fc-autocomplete__name'),
    (el) => el.textContent ?? '',
  );
};

describe('the workbook function help', () => {
  it('lists exactly what the formula editor lists, for built-ins and LAMBDA names in each scope', async () => {
    const wb = await WorkbookHandle.createDefault({ locale: 'en' });
    books.push(wb);
    wb.addSheet('Second');
    wb.setDefinedNameEntry('Twice', 'LAMBDA(x,x*2)');
    wb.setDefinedNameEntry('Thrice', 'LAMBDA(x,x*3)', 1);
    for (const [text, sheet] of [
      ['=su', 0],
      ['=t', 0],
      ['=t', 1],
      ['=1+co', 1],
    ] as const) {
      const editor = editorList(wb, sheet, text);
      expect(editor.length, text).toBeGreaterThan(0);
      expect(workbookFunctionHelp(wb, sheet).suggest(text, text.length)?.names, text).toEqual(
        editor,
      );
      document.body.replaceChildren();
    }
  });
});
