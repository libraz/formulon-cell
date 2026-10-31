import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';
import {
  attachAutocomplete,
  collectAutocompleteNames,
  pickListValues,
  suggestColumnHistory,
  suggestStructuredRef,
} from '../../../src/interact/autocomplete.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const liveHandles = new Set<ReturnType<typeof attachAutocomplete>>();
const attach = (deps: Parameters<typeof attachAutocomplete>[0]) => {
  const handle = attachAutocomplete(deps);
  liveHandles.add(handle);
  return handle;
};

afterEach(() => {
  for (const handle of liveHandles) handle.detach();
  liveHandles.clear();
  document.body.replaceChildren();
});

const sales = {
  name: 'Sales',
  columns: ['Region', 'Revenue', 'Quarter'],
};

describe('pickListValues', () => {
  const cell = (
    row: number,
    col: number,
    value: string | number,
    formula: string | null = null,
  ) => ({
    addr: { sheet: 0, row, col },
    value:
      typeof value === 'string'
        ? { kind: 'text' as const, value }
        : { kind: 'number' as const, value },
    formula,
  });

  it('collects deduped text above the row in the column, nearest first', () => {
    const cells = [
      cell(0, 0, 'East'),
      cell(1, 0, 'West'),
      cell(2, 0, 'East'),
      cell(3, 0, 42),
      cell(4, 0, 'Total', '="Total"'),
      cell(5, 0, ''),
      cell(1, 1, 'Other column'),
      cell(9, 0, 'Below'),
    ];
    expect(pickListValues(cells, 0, 6)).toEqual(['East', 'West']);
  });

  it('caps the list at ten values', () => {
    const cells = Array.from({ length: 12 }, (_, row) => cell(row, 0, `v${row}`));
    expect(pickListValues(cells, 0, 12)).toEqual(
      Array.from({ length: 10 }, (_, i) => `v${11 - i}`),
    );
  });
});

describe('suggestStructuredRef', () => {
  it('suggests all columns just inside an empty bracket', () => {
    const text = '=Sales[';
    const ctx = suggestStructuredRef(text, text.length, [sales]);
    expect(ctx?.matches).toEqual(['Quarter', 'Region', 'Revenue']);
    expect(ctx?.tokenStart).toBe('=Sales['.length);
    expect(ctx?.tokenEnd).toBe('=Sales['.length);
    expect(ctx?.insertSuffix).toBe(']');
  });

  it('filters by partial prefix, case-insensitive', () => {
    const text = '=Sales[re';
    const ctx = suggestStructuredRef(text, text.length, [sales]);
    expect(ctx?.matches).toEqual(['Region', 'Revenue']);
  });

  it('returns null when caret is outside any open bracket', () => {
    expect(suggestStructuredRef('=SUM(A1)', 8, [sales])).toBeNull();
  });

  it('returns null when the table name does not match', () => {
    expect(suggestStructuredRef('=Other[r', 8, [sales])).toBeNull();
  });

  it('returns null when text is not a formula', () => {
    expect(suggestStructuredRef('Sales[r', 7, [sales])).toBeNull();
  });

  it('returns null when bracket already closed before caret', () => {
    expect(suggestStructuredRef('=Sales[Region]', 14, [sales])).toBeNull();
  });

  it('returns null when no tables are passed', () => {
    expect(suggestStructuredRef('=Sales[r', 8, [])).toBeNull();
  });

  it('matches the full Unicode and backslash table identifier', () => {
    const tables = [
      { name: '売上', columns: ['金額'] },
      { name: '\\売上', columns: ['金額'] },
    ];
    expect(suggestStructuredRef('=売上[金', 5, tables)?.matches).toEqual(['金額']);
    expect(suggestStructuredRef('=\\売上[金', 6, tables)?.matches).toEqual(['金額']);
  });

  it('does not fall back to a table-name suffix after a dot', () => {
    const tables = [
      { name: 'Sales.Data', columns: ['Revenue +Data Region'] },
      { name: 'Data', columns: ['Region'] },
    ];
    const text = '=Sales.Data[Re';
    expect(suggestStructuredRef(text, text.length, tables)?.matches).toEqual([
      'Revenue +Data Region',
    ]);
  });

  it('deduplicates header labels case-insensitively and keeps the list sorted', () => {
    const ctx = suggestStructuredRef('=Sales[R', 8, [
      { name: 'Sales', columns: ['Revenue', 'Region', 'region', 'Rate'] },
    ]);
    expect(ctx?.matches).toEqual(['Rate', 'Region', 'Revenue']);
  });
});

describe('structured table autocomplete insertion', () => {
  const setup = (value: string, tables: readonly { name: string; columns: string[] }[]) => {
    const input = document.createElement('textarea');
    input.value = value;
    input.setSelectionRange(value.length, value.length);
    document.body.append(input);
    const handle = attach({ input, getTables: () => tables });
    handle.refresh();
    return { input, handle };
  };

  it('matches and inserts a full dotted table name without a suffix-table collision', () => {
    const { input, handle } = setup('=Sales.Data[Re', [
      { name: 'Sales.Data', columns: ['Revenue +Data Region'] },
      { name: 'Data', columns: ['Region'] },
    ]);
    const candidates = [...document.querySelectorAll('.fc-autocomplete__name')].map(
      (node) => node.textContent,
    );

    expect.soft(candidates).toEqual(['Revenue +Data Region']);
    expect(handle.acceptHighlighted()).toBe(true);
    expect(input.value).toBe('=Sales.Data[[Revenue +Data Region]]');
    expect(input.selectionStart).toBe(input.value.length);
  });

  it.each([
    ['=売上[金', [{ name: '売上', columns: ['金額'] }], '金額', '=売上[金額]'],
    ['=\\売上[金', [{ name: '\\売上', columns: ['金額'] }], '金額', '=\\売上[金額]'],
  ])('matches and inserts the full table name for %s', (value, tables, label, expected) => {
    const { input, handle } = setup(value, tables);
    expect(document.querySelector('.fc-autocomplete__name')?.textContent).toBe(label);
    expect(handle.acceptHighlighted()).toBe(true);
    expect(input.value).toBe(expected);
    expect(input.selectionStart).toBe(expected.length);
  });

  it.each([
    ['A]B', "=Sales[A'", "=Sales[[A']B]]"],
    ['A[B', "=Sales[A'[", "=Sales[[A'[B]]"],
    ['Cost#', "=Sales[Cost'", "=Sales[[Cost'#]]"],
    ['Qty@', "=Sales[Qty'", "=Sales[[Qty'@]]"],
    ["O'Brien", "=Sales[O'", "=Sales[[O''Brien]]"],
  ])(
    'shows raw header %s and inserts its escaped structured reference',
    (header, value, expected) => {
      const { input, handle } = setup(value, [{ name: 'Sales', columns: [header] }]);
      expect(document.querySelector('.fc-autocomplete__name')?.textContent).toBe(header);
      expect(handle.acceptHighlighted()).toBe(true);
      expect(input.value).toBe(expected);
      expect(input.selectionStart).toBe(expected.length);
    },
  );

  it.each([
    ['Price, Net', '=Sales[P', '=Sales[[Price, Net]]'],
    ['Tax:Total', '=Sales[T', '=Sales[[Tax:Total]]'],
    ['A.B', '=Sales[A', '=Sales[[A.B]]'],
  ])(
    'adds extra brackets for special header %s without changing its label',
    (header, value, expected) => {
      const { input, handle } = setup(value, [{ name: 'Sales', columns: [header] }]);
      expect(document.querySelector('.fc-autocomplete__name')?.textContent).toBe(header);
      expect(handle.acceptHighlighted()).toBe(true);
      expect(input.value).toBe(expected);
      expect(input.selectionStart).toBe(expected.length);
    },
  );

  it.each([
    ['=Sales[[Re', '=Sales[[Region]]'],
    ['=Sales[[#Headers],[Re', '=Sales[[#Headers],[Region]]'],
    ["=Sales[A']", "=Sales[[A']B]]"],
    ['=SUM(Sales[[Re)+1', '=SUM(Sales[[Region]])+1'],
  ])(
    'completes already-nested and escaped partial %s while preserving the suffix',
    (value, expected) => {
      const { input, handle } = setup(value, [{ name: 'Sales', columns: ['Region', 'A]B'] }]);
      const caret = value.includes(')+') ? value.indexOf(')') : value.length;
      input.setSelectionRange(caret, caret);
      handle.refresh();
      expect(document.querySelector('.fc-autocomplete__name')?.textContent).toBe(
        value.includes("A']") ? 'A]B' : 'Region',
      );
      expect(handle.acceptHighlighted()).toBe(true);
      expect(input.value).toBe(expected);
      expect(input.selectionStart).toBe(
        value.includes(')+') ? expected.indexOf(')+') : expected.length,
      );
    },
  );
});

describe('suggestColumnHistory', () => {
  it('matches by case-insensitive prefix and replaces the whole token', () => {
    const text = 'ap';
    const ctx = suggestColumnHistory(text, text.length, ['Apple', 'Banana', 'Apricot']);
    expect(ctx?.matches).toEqual(['Apple', 'Apricot']);
    expect(ctx?.tokenStart).toBe(0);
    expect(ctx?.tokenEnd).toBe(text.length);
    expect(ctx?.insertSuffix).toBe('');
    expect(ctx?.kind).toBe('column');
  });

  it('preserves caller-supplied order (nearest-first) and dedupes', () => {
    // The values array is the editor's responsibility — it must come in
    // nearest-first / deduped. Verify the suggester keeps that order intact
    // and never re-emits a duplicate it sees later in the list.
    const ctx = suggestColumnHistory('a', 1, ['Apricot', 'Apple', 'Avocado']);
    expect(ctx?.matches).toEqual(['Apricot', 'Apple', 'Avocado']);
  });

  it('returns null when nothing prefix-matches', () => {
    expect(suggestColumnHistory('zz', 2, ['Apple', 'Banana'])).toBeNull();
  });

  it('returns null on empty input or mid-edit caret', () => {
    expect(suggestColumnHistory('', 0, ['Apple'])).toBeNull();
    // Caret at offset 1 of "apple" — user is correcting the middle, not
    // appending; spreadsheets don't pop the list there either.
    expect(suggestColumnHistory('apple', 1, ['Apricot'])).toBeNull();
  });

  it('skips exact-length matches (would be a no-op insert)', () => {
    // "App" already typed — no point offering "App" itself, but longer
    // prefix-matches should still surface.
    const ctx = suggestColumnHistory('App', 3, ['App', 'Apple', 'Application']);
    expect(ctx?.matches).toEqual(['Apple', 'Application']);
  });
});

describe('autocomplete insertion outcomes', () => {
  it('brackets a successful insertion with before and after callbacks', () => {
    const input = document.createElement('textarea');
    input.value = '=SU';
    input.setSelectionRange(input.value.length, input.value.length);
    document.body.append(input);
    const calls: string[] = [];
    const handle = attach({
      input,
      getFunctionNames: () => ['SUM'],
      onBeforeInsert: () => calls.push(`before:${input.value}`),
      onAfterInsert: (result) => calls.push(`after:${input.value}:${result?.kind}`),
    });
    handle.refresh();

    expect(handle.acceptHighlighted()).toBe(true);
    expect(calls).toEqual(['before:=SU', 'after:=SUM(:function']);
    handle.detach();
    input.remove();
  });

  it('reports the accepted column-history candidate kind', () => {
    const input = document.createElement('textarea');
    input.value = 'App';
    input.setSelectionRange(input.value.length, input.value.length);
    document.body.append(input);
    let outcome: { kind: string } | undefined;
    const handle = attach({
      input,
      editingAddr: { sheet: 0, row: 4, col: 2 },
      getColumnValues: () => ['Apple'],
      onAfterInsert: (result) => {
        outcome = result;
      },
    });
    handle.refresh();

    expect(handle.acceptHighlighted()).toBe(true);
    expect(input.value).toBe('Apple');
    expect(outcome).toEqual({ kind: 'column' });
  });
});

describe('attachAutocomplete labels', () => {
  it('renders localized labels for custom functions', () => {
    const input = document.createElement('textarea');
    input.value = '=CUS';
    document.body.appendChild(input);
    input.setSelectionRange(input.value.length, input.value.length);
    const handle = attach({
      input,
      getCustomFunctions: () => ['CUSTOM_RATE'],
      labels: {
        customFunction: 'カスタム関数',
        structuredTableColumn: '構造化テーブル列',
        pickFromList: 'リストから選択',
      },
    });

    handle.refresh();
    expect(document.body.textContent).toContain('カスタム関数');

    handle.detach();
    input.remove();
  });

  it('resets dismissal when a canceled edit is programmatically closed', () => {
    const input = document.createElement('textarea');
    input.value = '=SU';
    document.body.appendChild(input);
    input.setSelectionRange(input.value.length, input.value.length);
    const handle = attach({ input, getFunctionNames: () => ['SUM', 'SUMIF'] });
    handle.refresh();
    expect(handle.isOpen()).toBe(true);

    handle.dismiss?.();
    handle.refresh();
    expect(handle.isOpen()).toBe(false);
    input.blur();
    handle.close();

    // Cancel and begin another edit with the same draft text and caret.
    input.value = '=SU';
    input.setSelectionRange(input.value.length, input.value.length);
    input.focus();
    handle.refresh();
    expect(handle.isOpen()).toBe(true);

    handle.detach();
    input.remove();
  });

  it('shows a custom function signature from its args and falls back to the label', () => {
    const input = document.createElement('textarea');
    input.value = '=CUS';
    document.body.appendChild(input);
    input.setSelectionRange(input.value.length, input.value.length);
    const handle = attach({
      input,
      getCustomFunctions: () => ['CUSTOM_RATE', 'CUSTOM_PLAIN'],
      getCustomFunctionArgs: (name) =>
        name === 'CUSTOM_RATE' ? ['base', 'rate', '[step]'] : undefined,
      getListSeparator: () => ';',
      labels: { customFunction: 'Custom function' },
    });

    handle.refresh();
    handle.move(1);
    const text = document.body.textContent ?? '';
    expect(text).toContain('CUSTOM_RATE(base; rate; [step])');
    expect(text).not.toContain('CUSTOM_PLAIN(');
    handle.move(-1);
    expect(document.body.textContent).toContain('Custom function');

    handle.detach();
    input.remove();
  });

  it('updates visible labels in place', () => {
    const input = document.createElement('textarea');
    input.value = '=Sales[';
    document.body.appendChild(input);
    input.setSelectionRange(input.value.length, input.value.length);
    const handle = attach({
      input,
      getTables: () => [sales],
      labels: {
        customFunction: 'カスタム関数',
        structuredTableColumn: '構造化テーブル列',
        pickFromList: 'リストから選択',
      },
    });

    handle.refresh();
    expect(document.body.textContent).toContain('構造化テーブル列');
    handle.setLabels({ structuredTableColumn: 'テーブル列' });
    expect(document.body.textContent).toContain('テーブル列');

    handle.detach();
    input.remove();
  });

  it('wires listbox state to the editor input and selected option', () => {
    const input = document.createElement('textarea');
    input.value = '=S';
    document.body.appendChild(input);
    input.setSelectionRange(input.value.length, input.value.length);
    const handle = attach({
      input,
      getFunctionNames: () => ['SUM', 'SUBTOTAL'],
    });

    handle.refresh();
    const root = document.querySelector<HTMLElement>('.fc-autocomplete');
    const listbox = root?.querySelector<HTMLElement>('.fc-autocomplete__list');
    expect(root?.getAttribute('role')).toBeNull();
    expect(listbox?.getAttribute('role')).toBe('listbox');
    expect(input.getAttribute('aria-expanded')).toBe('true');
    expect(input.getAttribute('aria-controls')).toBe(listbox?.id);

    const options = Array.from(listbox?.querySelectorAll<HTMLElement>('[role="option"]') ?? []);
    expect(options[0]?.getAttribute('aria-selected')).toBe('true');
    expect(input.getAttribute('aria-activedescendant')).toBe(options[0]?.id);

    handle.move(1);
    const nextOptions = Array.from(
      document.querySelectorAll<HTMLElement>('.fc-autocomplete__list [role="option"]'),
    );
    expect(nextOptions[0]?.getAttribute('aria-selected')).toBe('false');
    expect(nextOptions[1]?.getAttribute('aria-selected')).toBe('true');
    expect(input.getAttribute('aria-activedescendant')).toBe(nextOptions[1]?.id);

    handle.close();
    expect(input.getAttribute('aria-expanded')).toBe('false');
    expect(input.hasAttribute('aria-controls')).toBe(false);
    expect(input.hasAttribute('aria-activedescendant')).toBe(false);
    input.remove();
  });

  it('suggests localized Unicode aliases and resolves their canonical signature', () => {
    const input = document.createElement('textarea');
    input.value = '=合';
    document.body.appendChild(input);
    input.setSelectionRange(input.value.length, input.value.length);
    const handle = attach({
      input,
      getFunctionNames: () => ['SUM'],
      getFunctionDisplayNames: () => ['合計'],
      resolveFunctionName: (rawName) => (rawName === '合計' ? 'SUM' : rawName.toUpperCase()),
      getListSeparator: () => ';',
    });
    handle.refresh();
    expect(document.body.textContent).toContain('合計(number1; [number2]; ...)');
    expect(handle.acceptHighlighted()).toBe(true);
    expect(input.value).toBe('=合計(');
    handle.detach();
    input.remove();
  });

  it('keeps option buttons on the autocomplete item helper', () => {
    const source = readFileSync(join(root, 'src/interact/autocomplete.ts'), 'utf8');
    expect(source).toContain("import { createInteractionButton } from './chip-button.js'");
    expect(source).toContain('function createAutocompleteItemButton(');
    expect(source).toContain('createAutocompleteItemButton({');
    expect(source).not.toContain("document.createElement('button')");
  });

  it('keeps autocomplete list and ScreenTip compact separate surfaces', () => {
    const css = readFileSync(join(root, 'src/styles/core/app/popups/formula-helpers.css'), 'utf8');

    expect(css).toMatch(/\.fc-autocomplete\s*\{[\s\S]*?pointer-events: none;/);
    expect(css).toMatch(
      /\.fc-autocomplete__list\s*\{[\s\S]*?max-height: min\(200px, calc\(100vh - 20px\)\);[\s\S]*?background: var\(--fc-bg-elev, Canvas\);[\s\S]*?border-radius: 2px;[\s\S]*?0 5px 16px rgba\(0, 0, 0, 0\.16\)/,
    );
    expect(css).toMatch(
      /\.fc-autocomplete__item\s*\{[\s\S]*?grid-template-columns: 18px minmax\(0, 1fr\);[\s\S]*?min-height: 24px;[\s\S]*?padding: 2px 7px;/,
    );
    expect(css).toMatch(
      /\.fc-autocomplete__help\s*\{[\s\S]*?width: 224px;[\s\S]*?border-radius: 2px;[\s\S]*?0 4px 14px rgba\(0, 0, 0, 0\.14\)/,
    );
    expect(css).not.toMatch(/\.fc-autocomplete__item\s*\{[^}]*border-bottom:/);
    expect(css).toMatch(/\.fc-autocomplete__badge\s*\{[\s\S]*?width: 18px;[\s\S]*?height: 16px;/);
    expect(css).not.toContain('text-transform: uppercase');
    expect(css).not.toContain('box-shadow: var(--fc-shadow-8)');
    expect(css).not.toContain('background: #f3f2f1');
    expect(css).toContain(
      'outline: 1px solid var(--fc-accent-soft, var(--fc-accent, currentColor))',
    );
  });
});

describe('attachAutocomplete defined names', () => {
  const setup = (value: string, names: ReturnType<typeof collectAutocompleteNames>) => {
    const input = document.createElement('textarea');
    input.value = value;
    document.body.appendChild(input);
    input.setSelectionRange(value.length, value.length);
    const handle = attach({
      input,
      getNames: () => names,
      getFunctionNames: () => ['SUMIF', 'SUM'],
    });
    handle.refresh();
    return { input, handle };
  };

  it('scopes names to the sheet and includes table names', () => {
    const names = collectAutocompleteNames(
      [
        { name: 'SumRate', formula: 'Sheet1!$A$1', localSheetId: -1 },
        { name: 'SumLocal', formula: 'Sheet1!$A$2', localSheetId: 0 },
        { name: 'SumOther', formula: 'Sheet2!$A$1', localSheetId: 1 },
      ],
      [{ name: 'SumTable', columns: ['a'] }],
      0,
    );
    expect(names.map((n) => n.name)).toEqual(['SumRate', 'SumLocal', 'SumTable']);
  });

  it('lists names with functions alphabetically and inserts the bare name', () => {
    const { input, handle } = setup('=su', [{ name: 'SumRate', detail: '=Sheet1!$A$1' }]);
    const items = [...document.querySelectorAll<HTMLElement>('.fc-autocomplete__item')];
    expect(items.map((i) => i.querySelector('.fc-autocomplete__name')?.textContent)).toEqual([
      'SUM',
      'SUMIF',
      'SumRate',
    ]);
    expect(items[2]?.classList.contains('fc-autocomplete__item--name')).toBe(true);
    handle.move(-1);
    expect(handle.acceptHighlighted()).toBe(true);
    expect(input.value).toBe('=SumRate');
    handle.detach();
    input.remove();
  });

  it('matches Japanese names by prefix', () => {
    const { input, handle } = setup('=売', [{ name: '売上', detail: '=Sheet1!$A$1' }]);
    expect(document.querySelector('.fc-autocomplete__name')?.textContent).toBe('売上');
    handle.detach();
    input.remove();
  });

  it('uses the full dotted identifier for custom-function and defined-name matches', () => {
    const input = document.createElement('textarea');
    input.value = '=Rate.F';
    input.setSelectionRange(input.value.length, input.value.length);
    document.body.append(input);
    let outcome: { kind: string } | undefined;
    const handle = attach({
      input,
      getNames: () => [{ name: 'Rate.Factor', detail: '=Sheet1!$A$1' }],
      getCustomFunctions: () => ['FOO', 'Rate.FOO'],
      getFunctionNames: () => [],
      onAfterInsert: (result) => {
        outcome = result;
      },
    });
    handle.refresh();

    const candidates = [...document.querySelectorAll('.fc-autocomplete__name')].map(
      (node) => node.textContent,
    );
    expect.soft(candidates).toEqual(['Rate.Factor', 'Rate.FOO']);
    const nameIndex = candidates.indexOf('Rate.Factor');
    for (let i = 0; i < nameIndex; i += 1) handle.move(1);
    expect(handle.acceptHighlighted()).toBe(true);
    expect(input.value).toBe('=Rate.Factor');
    expect(outcome).toEqual({ kind: 'name' });
  });

  it('keeps backslash-prefixed names whole and does not suggest a built-in suffix', () => {
    const input = document.createElement('textarea');
    input.value = '=\\SU';
    input.setSelectionRange(input.value.length, input.value.length);
    document.body.append(input);
    const handle = attach({
      input,
      getNames: () => [{ name: '\\SUMRate', detail: '=Sheet1!$A$1' }],
      getFunctionNames: () => ['SUM'],
    });
    handle.refresh();

    expect
      .soft(
        [...document.querySelectorAll('.fc-autocomplete__name')].map((node) => node.textContent),
      )
      .toEqual(['\\SUMRate']);
    expect(handle.acceptHighlighted()).toBe(true);
    expect(input.value).toBe('=\\SUMRate');
  });

  it('matches dotted Unicode custom functions against the entire identifier', () => {
    const input = document.createElement('textarea');
    input.value = '=売上_2.電';
    input.setSelectionRange(input.value.length, input.value.length);
    document.body.append(input);
    const handle = attach({
      input,
      getCustomFunctions: () => ['電算', '売上_2.電算'],
      getFunctionNames: () => [],
    });
    handle.refresh();

    expect
      .soft(
        [...document.querySelectorAll('.fc-autocomplete__name')].map((node) => node.textContent),
      )
      .toEqual(['売上_2.電算']);
    expect(handle.acceptHighlighted()).toBe(true);
    expect(input.value).toBe('=売上_2.電算(');
  });

  it.each([
    ['=SU(A1)', 3, '=SUM(A1)', 5],
    ['=IF(1,SU(A1),0)', 8, '=IF(1,SUM(A1),0)', 10],
    ['=SU  (A1)', 3, '=SUM  (A1)', 7],
  ])(
    'completes a function name before its existing call: %s',
    (value, caret, expected, expectedCaret) => {
      const input = document.createElement('textarea');
      input.value = value;
      input.setSelectionRange(caret, caret);
      document.body.append(input);
      let outcome: { kind: string } | undefined;
      const handle = attach({
        input,
        getFunctionNames: () => ['SUM'],
        getCustomFunctions: () => ['SUMCUSTOM'],
        getNames: () => [{ name: 'SumRate', detail: '=Sheet1!$A$1' }],
        onAfterInsert: (result) => {
          outcome = result;
        },
      });
      handle.refresh();

      expect(handle.isOpen()).toBe(true);
      expect(
        [...document.querySelectorAll('.fc-autocomplete__name')].map((node) => node.textContent),
      ).toEqual(['SUM', 'SUMCUSTOM']);
      expect(handle.acceptHighlighted()).toBe(true);
      expect(input.value).toBe(expected);
      expect(input.selectionStart).toBe(expectedCaret);
      expect(outcome).toEqual({ kind: 'function' });
    },
  );

  it('does not offer names when a call-name prefix has no function matches', () => {
    const input = document.createElement('textarea');
    input.value = '=SU(';
    input.setSelectionRange(input.value.length - 1, input.value.length - 1);
    document.body.append(input);
    const handle = attach({
      input,
      getFunctionNames: () => [],
      getNames: () => [{ name: 'SumRate', detail: '=Sheet1!$A$1' }],
    });

    handle.refresh();

    expect(handle.isOpen()).toBe(false);
  });

  it('keeps a remaining function-name suffix intact before an existing parenthesis', () => {
    const input = document.createElement('textarea');
    input.value = '=SUM(A1)';
    input.setSelectionRange(3, 3);
    document.body.append(input);
    const handle = attach({ input, getFunctionNames: () => ['SUM', 'SUMIF'] });

    handle.refresh();

    expect(handle.isOpen()).toBe(false);
    expect(handle.acceptHighlighted()).toBe(false);
    expect(input.value).toBe('=SUM(A1)');
  });
});
