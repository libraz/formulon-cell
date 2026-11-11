import { afterEach, describe, expect, it } from 'vitest';
import type { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachArgHelper } from '../../../src/interact/arg-helper.js';
import {
  attachAutocomplete,
  collectAutocompleteNames,
  collectLambdaFunctions,
  lambdaParameters,
} from '../../../src/interact/autocomplete.js';
import { workbookCustomFunctions } from '../../../src/interact/workbook-function-catalog.js';

const DEFINED = [
  { name: 'MYADD', formula: 'LAMBDA(a, b, a+b)', localSheetId: -1 },
  { name: 'RATE1', formula: 'Sheet1!$A$1', localSheetId: -1 },
  { name: 'OtherSheetFn', formula: 'LAMBDA(x, x)', localSheetId: 1 },
];
const workbook = { definedNames: () => DEFINED } as unknown as WorkbookHandle;

const listRows = (): string[] =>
  [...document.querySelectorAll('.fc-autocomplete__list > *')].map((row) =>
    row.classList.contains('fc-autocomplete__group')
      ? `# ${row.textContent}`
      : (row.querySelector('.fc-autocomplete__name')?.textContent ?? ''),
  );

afterEach(() => {
  document.body.replaceChildren();
});

describe('LAMBDA names', () => {
  it('reads the parameter names of a LAMBDA formula', () => {
    expect(lambdaParameters('=LAMBDA(a, b, a+b)')).toEqual(['a', 'b']);
    expect(lambdaParameters('lambda(x, IF(x>0, "a,""b", SUM({1,2})))')).toEqual(['x']);
    expect(lambdaParameters('LAMBDA(42)')).toEqual([]);
    expect(lambdaParameters('Sheet1!$A$1')).toBeNull();
    expect(lambdaParameters('LAMBDA(a, ')).toBeNull();
  });

  it('completes as functions, not names, from the sheets that see them', () => {
    expect(collectAutocompleteNames(DEFINED, [], 0).map((n) => n.name)).toEqual(['RATE1']);
    expect([...collectLambdaFunctions(DEFINED, 0).values()]).toEqual([
      { name: 'MYADD', params: ['a', 'b'] },
    ]);
    const custom = workbookCustomFunctions(
      () => workbook,
      () => 0,
      { list: () => ['HOSTFN'], args: (name) => (name === 'HOSTFN' ? ['v'] : undefined) },
    );
    expect(custom.getCustomFunctions()).toEqual(['HOSTFN', 'MYADD']);
    expect(custom.getCustomFunctionArgs('MYADD')).toEqual(['a', 'b']);
    expect(custom.getCustomFunctionArgs('OTHERSHEETFN')).toBeUndefined();
  });
});

describe('formula autocomplete list layout', () => {
  const open = (value: string, grouped: boolean) => {
    const input = document.createElement('textarea');
    input.value = value;
    document.body.appendChild(input);
    input.setSelectionRange(value.length, value.length);
    const custom = workbookCustomFunctions(
      () => workbook,
      () => 0,
      { list: () => [], args: () => undefined },
    );
    const handle = attachAutocomplete({
      input,
      ...custom,
      getNames: () => collectAutocompleteNames(DEFINED, [], 0),
      getFunctionNames: () => ['RATE', 'INTRATE', 'SUM'],
      groupNames: () => grouped,
      labels: { functions: '関数', namedRanges: '名前付き一覧' },
    });
    handle.refresh();
    return { input, handle, custom };
  };

  it('on the Mac, heads a LAMBDA name as a function and inserts it with its argument tooltip', () => {
    const { input, handle, custom } = open('=MYA', true);
    expect(listRows()).toEqual(['# 関数', 'MYADD']);
    expect(document.querySelector('.fc-autocomplete__item')?.getAttribute('data-fc-kind')).toBe(
      'function',
    );
    expect(handle.acceptHighlighted()).toBe(true);
    expect(input.value).toBe('=MYADD(');
    const helper = attachArgHelper({
      input,
      getCustomFunctionArgs: custom.getCustomFunctionArgs,
    });
    helper.refresh();
    expect(document.querySelector('.fc-arghelper')?.textContent).toBe('MYADD(a, b)');
    helper.detach();
    handle.detach();
  });

  it('on the Mac, lists other names in their own group above the functions and inserts them bare', () => {
    const { input, handle } = open('=RAT', true);
    expect(listRows()).toEqual(['# 名前付き一覧', 'RATE1', '# 関数', 'RATE']);
    expect(handle.acceptHighlighted()).toBe(true);
    expect(input.value).toBe('=RATE1');
    handle.detach();
  });

  // Windows is not measured: it keeps the documented mixed alphabetical list without headers.
  it('on Windows, mixes names alphabetically among the functions without headers', () => {
    const { handle } = open('=RAT', false);
    expect(listRows()).toEqual(['RATE', 'RATE1']);
    handle.detach();
  });
});
