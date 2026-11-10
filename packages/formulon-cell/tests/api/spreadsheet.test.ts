/**
 * Consumer tests for the UI object model returned by `mount()`.
 *
 * `update(patch)` combinations are pairwise (strength 2, 13 rows, 93/93 pairs) over:
 *   profile  ∈ { minimal, standard, full }
 *   policy   ∈ { none, viewer, fixedForm }
 *   viewport ∈ { auto, fixed }
 *   toolbar  ∈ { none, ribbon }
 *   theme    ∈ { paper, ink, contrast }
 *   locale   ∈ { en, ja }
 */
import { ValueKind } from '@libraz/formulon';
import {
  fixedFormPolicy,
  mount,
  type SpreadsheetInstance,
  type SpreadsheetPatch,
  viewerPolicy,
  Workbook,
} from '@libraz/formulon-cell';
import { ribbon } from '@libraz/formulon-cell/toolbar';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { spreadsheetInternals } from '../../src/mount/internals.js';
import { makeHost } from './helpers.js';

const live: Array<{ dispose(): void }> = [];

const mountSheet = async (
  options: Parameters<typeof mount>[1] = {},
): Promise<SpreadsheetInstance> => {
  const instance = await mount(makeHost(), { locale: 'en', ...options });
  live.push(instance);
  return instance;
};

afterEach(() => {
  for (const d of live.splice(0).reverse()) d.dispose();
});

describe('spreadsheet: sub-objects', () => {
  it('selection: set, extend and read; a sheet-qualified ref switches the active sheet', async () => {
    const sheet = await mountSheet();
    sheet.workbook.addSheet('Data');
    const changes = vi.fn();
    sheet.selection.on('change', changes);
    sheet.selection.set('B2');
    sheet.selection.extend('C4');
    expect(sheet.selection.get()).toMatchObject({
      sheet: 0,
      active: { row: 1, col: 1 },
      range: { firstRow: 1, firstCol: 1, lastRow: 3, lastCol: 2 },
    });
    sheet.selection.set('Data!A1');
    expect(sheet.selection.get().sheet).toBe(1);
    expect(changes).toHaveBeenCalled();
    expect(changes.mock.calls.at(-1)?.[0]).toHaveProperty('range');
  });

  it('history: a facade write on a mounted workbook is one undo step', async () => {
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    expect(sheet.history.canUndo).toBe(false);
    sheet.workbook.batch(() => {
      s.setNumber('A1', 1);
      s.setNumber('A2', 2);
    });
    expect(sheet.history.canUndo).toBe(true);
    sheet.history.undo();
    expect(s.getTypedValue('A1').kind).toBe(ValueKind.Blank);
    expect(s.getTypedValue('A2').kind).toBe(ValueKind.Blank);
    expect(sheet.history.canRedo).toBe(true);
    sheet.history.clear();
    expect(sheet.history.canRedo).toBe(false);
  });

  it('commands: a ribbon command acts on the current selection', async () => {
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    s.setNumber('A1', 7);
    sheet.selection.set('A1');
    expect(sheet.commands.can('insertRowsAbove')).toBe(true);
    sheet.commands.execute('insertRowsAbove');
    expect(s.getTypedValue('A2')).toEqual({ kind: ValueKind.Number, number: 7 });
  });

  it('dialogs: open and close by id', async () => {
    const sheet = await mountSheet();
    expect(sheet.dialogs.canOpen('formatCells')).toBe(true);
    sheet.dialogs.open('formatCells');
    sheet.dialogs.close('formatCells');
  });

  it('audit: trace precedents and clear them', async () => {
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    s.setNumber('A1', 1);
    s.setFormula('B1', '=A1');
    sheet.audit.tracePrecedents('B1');
    expect(sheet.audit.list().length).toBeGreaterThan(0);
    sheet.audit.clear();
    expect(sheet.audit.list()).toEqual([]);
  });

  it('registries: cells and formulaHints share one shape', async () => {
    const sheet = await mountSheet();
    const offCell = sheet.cells.add({ id: 'stars', match: () => false, format: () => '' } as never);
    expect(sheet.cells.get('stars')).not.toBeNull();
    expect(sheet.cells.get('missing')).toBeNull();
    offCell();
    expect(sheet.cells.list()).toEqual([]);

    const offHint = sheet.formulaHints.add({ name: 'MYFN', args: ['x'] } as never);
    expect(sheet.formulaHints.has('MYFN')).toBe(true);
    offHint();
    expect(sheet.formulaHints.get('MYFN')).toBeNull();
  });

  it('typed sub-handles are null, never undefined, when their feature is off', async () => {
    const sheet = await mountSheet({
      ui: { features: { clipboard: false, formatPainter: false, borderDraw: false } },
    });
    expect(sheet.clipboard).toBeNull();
    expect(sheet.formatPainter).toBeNull();
    expect(sheet.borderDraw).toBeNull();
    expect(sheet.toolbar).toBeNull();
  });

  it('extensions: get returns null for an absent id and remove throws', async () => {
    const sheet = await mountSheet();
    expect(sheet.extensions.get('no-such-extension' as never)).toBeNull();
    expect(() => sheet.extensions.remove('no-such-extension')).toThrow();
  });
});

describe('spreadsheet: update', () => {
  it('update({}) is a no-op without an update event', async () => {
    const sheet = await mountSheet();
    const updates = vi.fn();
    sheet.on('update', updates);
    sheet.update({});
    expect(updates).not.toHaveBeenCalled();
  });

  it('merges a patch, reports the changed keys and resets a nullable key with null', async () => {
    const sheet = await mountSheet({ theme: 'ink' });
    const updates = vi.fn();
    sheet.on('update', updates);
    sheet.update({ locale: 'ja' });
    expect(sheet.options.theme).toBe('ink');
    expect(sheet.options.locale).toBe('ja');
    expect(updates).toHaveBeenCalledWith(expect.objectContaining({ keys: ['locale'] }));
    sheet.update({ policy: viewerPolicy() });
    sheet.update({ policy: null });
    expect(sheet.commands.can('insertRowsAbove')).toBe(true);
  });

  it('a policy does not disable dialogs whose features are on', async () => {
    const sheet = await mountSheet({ policy: viewerPolicy() });
    expect(sheet.dialogs.canOpen('formatCells')).toBe(true);
  });

  const PAIRWISE: ReadonlyArray<
    readonly [
      'minimal' | 'standard' | 'full',
      'none' | 'viewer' | 'fixedForm',
      'auto' | 'fixed',
      'none' | 'ribbon',
      'paper' | 'ink' | 'contrast',
      'en' | 'ja',
    ]
  > = [
    ['standard', 'fixedForm', 'auto', 'ribbon', 'paper', 'ja'],
    ['full', 'viewer', 'fixed', 'ribbon', 'contrast', 'en'],
    ['full', 'none', 'auto', 'none', 'ink', 'ja'],
    ['minimal', 'viewer', 'auto', 'none', 'contrast', 'ja'],
    ['minimal', 'fixedForm', 'fixed', 'none', 'ink', 'en'],
    ['standard', 'none', 'fixed', 'none', 'paper', 'en'],
    ['minimal', 'none', 'fixed', 'ribbon', 'contrast', 'ja'],
    ['standard', 'viewer', 'fixed', 'ribbon', 'ink', 'en'],
    ['full', 'none', 'auto', 'none', 'paper', 'en'],
    ['standard', 'fixedForm', 'fixed', 'ribbon', 'contrast', 'ja'],
    ['minimal', 'fixedForm', 'fixed', 'ribbon', 'paper', 'ja'],
    ['full', 'fixedForm', 'fixed', 'none', 'contrast', 'ja'],
    ['minimal', 'viewer', 'fixed', 'ribbon', 'paper', 'en'],
  ];

  it.each(PAIRWISE)(
    'applies profile=%s policy=%s viewport=%s toolbar=%s theme=%s locale=%s',
    async (profile, policy, viewport, toolbar, theme, locale) => {
      const sheet = await mountSheet();
      const patch: SpreadsheetPatch = {
        ui: { profile },
        policy:
          policy === 'viewer'
            ? viewerPolicy()
            : policy === 'fixedForm'
              ? fixedFormPolicy({ ranges: ['A1:B2'] })
              : null,
        viewport: viewport === 'fixed' ? { mode: 'fixed', range: 'A1:D10' } : { mode: 'auto' },
        toolbar: toolbar === 'ribbon' ? ribbon() : null,
        theme,
        locale,
      };
      sheet.update(patch);
      expect(sheet.options.theme).toBe(theme);
      expect(sheet.options.locale).toBe(locale);
      expect(sheet.i18n.locale).toBe(locale);
      expect(sheet.viewport.options.mode).toBe(viewport);
      expect(sheet.toolbar === null).toBe(toolbar === 'none');
      expect(sheet.commands.can('insertRowsAbove')).toBe(policy === 'none');
    },
  );
});

describe('spreadsheet: events', () => {
  it('carries the origin on cell and change events', async () => {
    const sheet = await mountSheet();
    const cells = vi.fn();
    const changes = vi.fn();
    sheet.on('cell', cells);
    sheet.on('change', changes);
    sheet.applyChanges([{ addr: 'A1', input: '5' }], { history: 'record', origin: 'host:import' });
    expect(cells).toHaveBeenCalledWith(
      expect.objectContaining({
        sheet: 0,
        at: { row: 0, col: 0 },
        value: { kind: ValueKind.Number, number: 5 },
        origin: 'host:import',
      }),
    );
    expect(changes).toHaveBeenCalledTimes(1);
  });

  it('i18n announces a locale change on its own object', async () => {
    const sheet = await mountSheet();
    const changes = vi.fn();
    sheet.i18n.on('change', changes);
    sheet.update({ locale: 'ja' });
    expect(changes).toHaveBeenCalledTimes(1);
    expect(sheet.i18n.locale).toBe('ja');
  });
});

describe('spreadsheet: ownership', () => {
  it('update({ workbook }) swaps in a caller-owned workbook and disposes one the instance created', async () => {
    const sheet = await mountSheet();
    const created = sheet.workbook;
    const next = await Workbook.createDefault();
    live.push(next);
    sheet.update({ workbook: next });
    expect(sheet.workbook).toBe(next);
    expect(created.disposed).toBe(true);
    sheet.dispose();
    expect(next.disposed).toBe(false);
  });
});

describe('spreadsheet: facade write while mounted', () => {
  it('a table and a cell format written through the facade reach the UI state', async () => {
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    s.setText('A1', 'Item');
    s.setText('B1', 'Qty');
    s.addTable({ range: 'A1:B3', name: 'Stock' });
    s.updateCellFormat('B2:B3', { numberFormat: '0.00' });

    // The UI renders from its own state; read that state through the internal seam tests use.
    const { store } = spreadsheetInternals(sheet);
    const state = store.getState();
    expect(state.tables.tables.map((t) => t.name)).toEqual(['Stock']);
    expect(state.format.formats.get('0:1:1')?.numFmt).toBeDefined();

    s.setNumber('B2', 3);
    sheet.history.undo();
    expect(s.getTypedValue('B2').kind).toBe(ValueKind.Blank);
  });
});
