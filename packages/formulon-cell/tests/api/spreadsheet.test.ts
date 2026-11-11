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
import { PageOrientation, PaperSize, ValueKind } from '@libraz/formulon';
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
import { workbookHandleOf } from '../../src/api/workbook.js';
import { setComment } from '../../src/commands/comment.js';
import { spreadsheetInternals } from '../../src/mount/internals.js';
import { expectFailure, makeHost } from './helpers.js';

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

  it('history: repeat runs the latest repeatable action again and canRepeat tracks it', async () => {
    const sheet = await mountSheet();
    const changes = vi.fn();
    sheet.history.on('change', changes);
    expect(sheet.history.canRepeat).toBe(false);
    expect(sheet.history.repeat()).toBe(false);

    sheet.selection.set('A1');
    sheet.commands.execute('bold');
    expect(sheet.history.canRepeat).toBe(true);
    expect(changes).toHaveBeenCalled();

    sheet.selection.set('B2');
    const before = changes.mock.calls.length;
    expect(sheet.history.repeat()).toBe(true);
    expect(changes.mock.calls.length).toBeGreaterThan(before);
    expect(sheet.history.canUndo).toBe(true);

    sheet.history.clear();
    expect(sheet.history.canRepeat).toBe(false);
    expect(sheet.history.repeat()).toBe(false);
  });

  it('printDocument: builds the document from the sheet page setup the engine holds', async () => {
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    s.setText('A1', 'Quarterly');
    s.updatePageSetup({ orientation: PageOrientation.Landscape, paperSize: PaperSize.Letter });
    const doc = sheet.printDocument();
    expect(doc.html).toContain('<!DOCTYPE html>');
    expect(doc.html).toContain('Quarterly');
    expect(doc.html).toContain('@page { size: 792pt 612pt;');
    expect(doc.cssVars['--fc-print-orient']).toBe('landscape');
    expect(doc.cssVars['--fc-print-paper']).toBe('letter');

    s.updatePageSetup({ orientation: PageOrientation.Portrait });
    expect(sheet.printDocument().cssVars['--fc-print-orient']).toBe('portrait');
  });

  it('printDocument: print() renders the same document', async () => {
    const sheet = await mountSheet();
    sheet.workbook.sheet(0).setText('A1', 'Rendered');
    const host = sheet.host;
    // happy-dom's window has no print(); the render path falls back to it.
    const win = window as unknown as { print?: () => void };
    const original = win.print;
    win.print = vi.fn();
    sheet.print('pdf');
    const frame = host.querySelector('iframe');
    expect(frame?.srcdoc).toContain('Rendered');
    expect(frame?.srcdoc).toContain('@page');
    await new Promise((resolve) => setTimeout(resolve, 0));
    frame?.remove();
    win.print = original;
  });

  it('commands: a ribbon command acts on the current selection', async () => {
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    s.setNumber('A1', 7);
    sheet.selection.set('A1');
    expect(sheet.commands.can('insertRows')).toBe(true);
    sheet.commands.execute('insertRows');
    expect(s.getTypedValue('A2')).toEqual({ kind: ValueKind.Number, number: 7 });
  });

  it('dialogs: open and close by id', async () => {
    const sheet = await mountSheet();
    expect(sheet.dialogs.canOpen('format')).toBe(true);
    sheet.dialogs.open('format');
    sheet.dialogs.close('format');
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
    expect(updates).toHaveBeenCalledWith({ keys: ['locale'], source: 'api' });
    sheet.update({ policy: viewerPolicy() });
    sheet.update({ policy: null });
    expect(sheet.commands.can('insertRows')).toBe(true);
  });

  it('Show All Notes from the UI reports a user update, writes no note and leaves the file as it was', async () => {
    const sheet = await mountSheet();
    const { store } = spreadsheetInternals(sheet);
    setComment(store, { sheet: 0, row: 1, col: 1 }, 'note', workbookHandleOf(sheet.workbook));
    const formats = store.getState().format.formats;
    const saved = sheet.workbook.save();
    const updates = vi.fn();
    sheet.on('update', updates);

    sheet.commands.execute('showAllNotesReview');
    expect(updates).toHaveBeenCalledTimes(1);
    expect(updates).toHaveBeenCalledWith({ keys: ['showAllNotes'], source: 'user' });
    expect(sheet.options.showAllNotes).toBe(true);
    expect(store.getState().format.formats).toBe(formats);
    expect(sheet.workbook.save()).toEqual(saved);

    sheet.update({ showAllNotes: false });
    expect(updates).toHaveBeenLastCalledWith({ keys: ['showAllNotes'], source: 'api' });
    expect(updates).toHaveBeenCalledTimes(2);
    expect(store.getState().ui.showAllNotes).toBe(false);
  });

  it('showAllNotes set at mount draws every note from the start', async () => {
    const sheet = await mountSheet({ showAllNotes: true });
    expect(sheet.options.showAllNotes).toBe(true);
    expect(spreadsheetInternals(sheet).store.getState().ui.showAllNotes).toBe(true);
  });

  it('update({ host }) replaces the function providers and null removes them', async () => {
    const firstRefresh = vi.fn(() => []);
    const secondRefresh = vi.fn(() => []);
    const firstCapture = vi.fn(() => 'data:image/png;base64,first');
    const secondCapture = vi.fn(() => 'data:image/png;base64,second');
    const sheet = await mountSheet({
      host: { refreshPrinterProfiles: firstRefresh, captureScreenClip: firstCapture },
    });

    sheet.update({
      host: { refreshPrinterProfiles: secondRefresh, captureScreenClip: secondCapture },
    });
    await sheet.refreshPrinterProfiles();
    expect(await sheet.captureScreenClip()).toMatchObject({ src: 'data:image/png;base64,second' });
    expect(firstRefresh).not.toHaveBeenCalled();
    expect(firstCapture).not.toHaveBeenCalled();
    expect(secondRefresh).toHaveBeenCalledTimes(1);

    sheet.update({ host: null });
    await sheet.refreshPrinterProfiles();
    expect(await sheet.captureScreenClip()).toBeNull();
    expect(secondRefresh).toHaveBeenCalledTimes(1);
    expect(secondCapture).toHaveBeenCalledTimes(1);
  });

  it('update({ ui }) leaves the theme and the toolbar the host set', async () => {
    const sheet = await mountSheet({ theme: 'contrast' });
    sheet.update({ ui: { ...sheet.options.ui, features: { slicer: true } } });
    expect(sheet.host.dataset.fcTheme).toBe('contrast');

    sheet.update({ theme: 'ink', toolbar: ribbon() });
    sheet.update({ ui: { ...sheet.options.ui, features: { slicer: false } } });
    expect(sheet.host.dataset.fcTheme).toBe('ink');
    expect(sheet.toolbar).not.toBeNull();
  });

  it('decides a dialog as the ribbon decides the command that opens it', async () => {
    const sheet = await mountSheet({ policy: viewerPolicy() });
    expect(sheet.dialogs.canOpen('format')).toBe(false);
    expect(sheet.dialogs.canOpen('goTo')).toBe(true);
    expect(sheet.dialogs.canOpen('findReplace')).toBe(true);
    expectFailure(() => sheet.dialogs.open('format'), 'policy-denied');
  });

  it('refuses dialogs the protected sheet refuses, as the ribbon and shortcuts do', async () => {
    const sheet = await mountSheet();
    expect(sheet.dialogs.canOpen('pivotTable')).toBe(true);
    expect(sheet.dialogs.canOpen('hyperlink')).toBe(true);
    await sheet.workbook.sheet(0).protect();
    expect(sheet.dialogs.canOpen('pivotTable')).toBe(false);
    expect(sheet.dialogs.canOpen('hyperlink')).toBe(false);
    expectFailure(() => sheet.dialogs.open('pivotTable'), 'protected');
    expectFailure(() => sheet.dialogs.open('hyperlink'), 'protected');
    await sheet.workbook.sheet(0).unprotect();
    expect(sheet.dialogs.canOpen('pivotTable')).toBe(true);
    expect(sheet.dialogs.canOpen('hyperlink')).toBe(true);
  });

  it('asks the policy about the dialogs that lead straight to an edit', async () => {
    const denied = await mountSheet({ policy: viewerPolicy() });
    expect(denied.dialogs.canOpen('cfRules')).toBe(false);
    expect(denied.dialogs.canOpen('filterDropdown')).toBe(false);
    expect(denied.dialogs.canOpen('pivotTable')).toBe(false);
    expect(denied.dialogs.canOpen('watchWindow')).toBe(true);
    expectFailure(() => denied.dialogs.open('cfRules'), 'policy-denied');
    const allowed = await mountSheet({
      policy: {
        defaultOperation: 'deny',
        operations: { format: true, filter: true, pivot: true },
      },
    });
    expect(allowed.dialogs.canOpen('cfRules')).toBe(true);
    expect(allowed.dialogs.canOpen('filterDropdown')).toBe(true);
    expect(allowed.dialogs.canOpen('pivotTable')).toBe(true);
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
      expect(sheet.commands.can('insertRows')).toBe(policy === 'none');
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
    sheet.applyChanges([{ at: 'A1', input: '5' }], { history: 'record', origin: 'host:import' });
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
    const swapped = new Promise((resolve) => sheet.on('workbook', resolve));
    sheet.update({ workbook: next });
    expect(sheet.workbook).toBe(created);
    expect(await swapped).toEqual({ workbook: next });
    expect(sheet.workbook).toBe(next);
    expect(created.disposed).toBe(true);
    sheet.dispose();
    expect(next.disposed).toBe(false);
  });

  it('update({ workbook }) keeps the current workbook when the mount cannot switch', async () => {
    const created = await Workbook.createDefault();
    created.addSheet('Input');
    live.push(created);
    const sheet = await mountSheet({
      workbook: created,
      viewport: { mode: 'fixed', range: 'Input!A1:B2' },
    });
    const next = await Workbook.createDefault();
    live.push(next);
    const swaps = vi.fn();
    sheet.on('workbook', swaps);
    const failed = new Promise((resolve) => sheet.on('error', resolve));

    sheet.update({ workbook: next });

    expect(await failed).toMatchObject({ source: 'workbook', error: expect.any(Error) });
    expect(swaps).not.toHaveBeenCalled();
    expect(sheet.workbook).toBe(created);
    expect(sheet.options.workbook).toBe(created);
    expect(next.disposed).toBe(false);
    // The refused workbook is free again for another spreadsheet.
    const other = await mountSheet({ workbook: next });
    expect(other.workbook).toBe(next);
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
