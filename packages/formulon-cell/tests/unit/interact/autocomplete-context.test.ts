import { afterEach, describe, expect, it } from 'vitest';
import {
  type AutocompleteHandle,
  attachAutocomplete,
  suggestStructuredRef,
} from '../../../src/interact/autocomplete.js';

const table = { name: 'Sales', columns: ['Region', 'Revenue'] };

describe('formula autocomplete respects caret context', () => {
  const attached: { input: HTMLTextAreaElement; handle: AutocompleteHandle }[] = [];

  afterEach(() => {
    for (const { input, handle } of attached.splice(0)) {
      handle.detach();
      input.remove();
    }
  });

  const setup = (value: string, caret = value.length, functionNames: string[] = []) => {
    const input = document.createElement('textarea');
    input.value = value;
    input.setSelectionRange(caret, caret);
    document.body.append(input);
    const handle = attachAutocomplete({
      input,
      getTables: () => [table],
      getNames: () => [{ name: 'SumRate', detail: '=Sheet1!$A$1' }],
      getCustomFunctions: () => ['SUMCUSTOM'],
      // Isolate runtime custom-function / name suggestions from the built-in
      // catalog, whose caret-context guard is covered by commands/refs tests.
      getFunctionNames: () => functionNames,
    });
    attached.push({ input, handle });
    handle.refresh();
    return { input, handle };
  };

  it('does not suggest custom functions or defined names inside string literals', () => {
    for (const [value, caret] of [
      ['="SUM"', 4],
      ["='SUM'!A1", 5],
    ] as const) {
      const { handle } = setup(value, caret);
      expect(handle.isOpen()).toBe(false);
    }
  });

  it('does not accept a function or name suggestion into a quoted string', () => {
    const { input, handle } = setup('="SU');

    expect(handle.isOpen()).toBe(false);
    expect(handle.acceptHighlighted()).toBe(false);
    expect(input.value).toBe('="SU');
  });

  it('does not suggest functions, custom functions, or names inside an unquoted sheet name', () => {
    const { handle } = setup('=SUMSheet!A1', 3);

    expect(handle.isOpen()).toBe(false);
  });

  it('hides suggestions when the editor has a nonempty selection', () => {
    const { input, handle } = setup('=SUFFIX+1');
    input.setSelectionRange(3, 6);

    handle.refresh();

    expect(handle.isOpen()).toBe(false);
  });

  it('does not accept a stale suggestion after the selection moves over its suffix', () => {
    const value = '=SU+FFIX+1';
    const { input, handle } = setup(value, 3);
    expect(handle.isOpen()).toBe(true);
    input.setSelectionRange(1, 6);

    expect(handle.acceptHighlighted()).toBe(false);
    expect(input.value).toBe(value);
  });

  it('does not suggest a completion when the caret splits a function or name token', () => {
    const { handle: custom } = setup('=SUFFIX+1', 3);
    expect(custom.isOpen()).toBe(false);
    const { handle: catalog } = setup('=SUFFIX+1', 3, ['SUM', 'SUMIF']);
    expect(catalog.isOpen()).toBe(false);
    const { handle: name } = setup('=SumRate+1', 4);
    expect(name.isOpen()).toBe(false);
  });

  it('does not complete a structured reference into a retained suffix or closing bracket', () => {
    expect(suggestStructuredRef('=Sales[Region]', 9, [table])).toBeNull();
    expect(suggestStructuredRef('=Sales[Re]', 9, [table])).toBeNull();
  });

  it('does not suggest table columns from text inside a string', () => {
    expect(suggestStructuredRef('="Sales[Re', '="Sales[Re'.length, [table])).toBeNull();
    const { handle } = setup('="Sales[Re');
    expect(handle.isOpen()).toBe(false);
  });

  it('keeps structured-reference suggestions available inside a real table bracket', () => {
    const { input, handle } = setup('=Sales[Re');

    expect(handle.isOpen()).toBe(true);
    expect(document.querySelector('.fc-autocomplete__name')?.textContent).toBe('Region');
    expect(handle.acceptHighlighted()).toBe(true);
    expect(input.value).toBe('=Sales[Region]');
  });

  it('reports the accepted table-column candidate kind', () => {
    const input = document.createElement('textarea');
    input.value = '=Sales[Re';
    input.setSelectionRange(input.value.length, input.value.length);
    document.body.append(input);
    let outcome: { kind: string } | undefined;
    const handle = attachAutocomplete({
      input,
      getTables: () => [table],
      onAfterInsert: (result) => {
        outcome = result;
      },
    });
    attached.push({ input, handle });
    handle.refresh();

    expect(handle.acceptHighlighted()).toBe(true);
    expect(input.value).toBe('=Sales[Region]');
    expect(outcome).toEqual({ kind: 'tableColumn' });
  });

  it('keeps function and name suggestions active in formula expression text', () => {
    const { handle } = setup('=SUMC');

    expect(handle.isOpen()).toBe(true);
    expect(
      [...document.querySelectorAll('.fc-autocomplete__name')].map((node) => node.textContent),
    ).toContain('SUMCUSTOM');
  });
});
