import { ValueKind } from '@libraz/formulon';
import { mount, type SpreadsheetInstance, viewerPolicy, Workbook } from '@libraz/formulon-cell';
import { afterEach, describe, expect, it } from 'vitest';
import {
  type JournalCodec,
  type JournalScope,
  registerJournalOverride,
} from '../../src/api/journal/index.js';
import { workbookHandleOf } from '../../src/api/workbook.js';
import { makeHost } from './helpers.js';

const live: Array<{ dispose(): void }> = [];
const cleanup: Array<() => void> = [];

const mountSheet = async (
  options: Parameters<typeof mount>[1] = {},
): Promise<SpreadsheetInstance> => {
  const instance = await mount(makeHost(), { locale: 'en', ...options });
  live.push(instance);
  return instance;
};

afterEach(() => {
  for (const off of cleanup.splice(0)) off();
  for (const d of live.splice(0).reverse()) d.dispose();
});

const RED = '#ff0000';
const BLUE = '#0000ff';

/** A codec over the sheet tab colour, which the engine reports and restores faithfully. */
const tabColorCodec = (): JournalCodec & { captures: number } => {
  const codec = {
    id: 'test-tab-color',
    captures: 0,
    scopeKey: (scope: JournalScope) => `tab:${scope.sheet}`,
    capture: (scope: JournalScope) => {
      codec.captures += 1;
      const sheet = scope.engine.sheet(scope.sheet as number);
      const color = sheet.getTabColor();
      return () => sheet.setTabColor(color);
    },
  };
  return codec;
};

const journalTabColor = (): ReturnType<typeof tabColorCodec> => {
  const codec = tabColorCodec();
  cleanup.push(registerJournalOverride('sheet', 'setTabColor', codec));
  return codec;
};

describe('facade journal core', () => {
  it('keeps a cell batch as one undo step', async () => {
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    sheet.workbook.batch(() => {
      s.setNumber('A1', 1);
      s.setNumber('A2', 2);
    });
    sheet.history.undo();
    expect(s.getTypedValue('A1').kind).toBe(ValueKind.Blank);
    expect(s.getTypedValue('A2').kind).toBe(ValueKind.Blank);
    expect(sheet.history.canUndo).toBe(false);
    sheet.history.redo();
    expect(s.getTypedValue('A2').kind).toBe(ValueKind.Number);
  });

  it('records a batch of cells and a facade write as one step', async () => {
    journalTabColor();
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    expect(s.getTabColor()).toBeNull();
    sheet.workbook.batch(() => {
      s.setNumber('A1', 1);
      s.setTabColor(RED);
      s.setNumber('A2', 2);
    });
    expect(s.getTabColor()).not.toBeNull();
    sheet.history.undo();
    expect(s.getTabColor()).toBeNull();
    expect(s.getTypedValue('A1').kind).toBe(ValueKind.Blank);
    expect(s.getTypedValue('A2').kind).toBe(ValueKind.Blank);
    expect(sheet.history.canUndo).toBe(false);
    sheet.history.redo();
    expect(s.getTabColor()).not.toBeNull();
    expect(s.getTypedValue('A1').kind).toBe(ValueKind.Number);
  });

  it('captures a slice written twice in a batch once and undoes to the first before', async () => {
    const codec = journalTabColor();
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    sheet.workbook.batch(() => {
      s.setTabColor(RED);
      s.setTabColor(BLUE);
    });
    // One capture on first touch; the seal adds the after-state.
    expect(codec.captures).toBe(2);
    sheet.history.undo();
    expect(s.getTabColor()).toBeNull();
    sheet.history.redo();
    expect(s.getTabColor()?.argb).toBe(0xff0000ff);
  });

  it('records no entry for a refused batch and keeps the cell rollback', async () => {
    journalTabColor();
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    expect(() =>
      sheet.workbook.batch(() => {
        s.setNumber('A1', 1);
        s.setTabColor(RED);
        s.setNumber('not a cell', 2);
      }),
    ).toThrow();
    expect(s.getTypedValue('A1').kind).toBe(ValueKind.Blank);
    expect(s.getTabColor()).not.toBeNull();
    expect(sheet.history.canUndo).toBe(false);
  });

  it('captures nothing on a headless workbook', async () => {
    const codec = journalTabColor();
    const wb = await Workbook.createDefault({ locale: 'en' });
    try {
      const s = wb.sheet(0);
      wb.batch(() => {
        s.setNumber('A1', 1);
        s.setTabColor(RED);
      });
      expect(codec.captures).toBe(0);
      expect(s.getTypedValue('A1').kind).toBe(ValueKind.Number);
      expect(s.getTabColor()).not.toBeNull();
    } finally {
      wb.dispose();
    }
  });

  it('does not journal a write a listener makes while an undo replays', async () => {
    const codec = journalTabColor();
    const sheet = await mountSheet();
    const wb = sheet.workbook;
    const s = wb.sheet(0);
    wb.batch(() => {
      s.setNumber('A1', 1);
      s.setTabColor(RED);
    });
    const afterWrite = codec.captures;
    let wrote = false;
    const off = workbookHandleOf(wb).subscribeFacadeChanges(() => {
      if (wrote) return;
      wrote = true;
      s.setTabColor(BLUE);
    });
    sheet.history.undo();
    off();
    expect(wrote).toBe(true);
    expect(codec.captures).toBe(afterWrite);
    expect(sheet.history.canUndo).toBe(false);
    expect(sheet.history.canRedo).toBe(true);
  });

  it('clears History after the frame closes when a write resets it', async () => {
    cleanup.push(registerJournalOverride('sheet', 'setTabColor', 'reset'));
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    s.setNumber('B1', 5);
    expect(sheet.history.canUndo).toBe(true);
    sheet.workbook.batch(() => {
      s.setNumber('A1', 1);
      s.setTabColor(RED);
    });
    expect(sheet.history.canUndo).toBe(false);
    expect(s.getTypedValue('A1').kind).toBe(ValueKind.Number);
  });

  it('leaves History alone when the reset runs inside an open frame', async () => {
    cleanup.push(registerJournalOverride('sheet', 'setTabColor', 'reset'));
    const sheet = await mountSheet();
    const wb = sheet.workbook;
    const s = wb.sheet(0);
    s.setNumber('B1', 5);
    const history = workbookHandleOf(wb).attachedHistory();
    expect(history).not.toBeNull();
    history?.transact(() => {
      wb.batch(() => {
        s.setNumber('A1', 1);
        s.setTabColor(RED);
      });
    });
    expect(sheet.history.canUndo).toBe(true);
  });

  it('treats a codec that cannot capture as a reset', async () => {
    cleanup.push(
      registerJournalOverride('sheet', 'setTabColor', {
        id: 'test-null',
        scopeKey: () => 'null',
        capture: () => null,
      }),
    );
    const sheet = await mountSheet();
    const s = sheet.workbook.sheet(0);
    s.setNumber('B1', 5);
    sheet.workbook.batch(() => s.setTabColor(RED));
    expect(sheet.history.canUndo).toBe(false);
  });

  it('refuses undo of a facade entry under a policy, cell batch included', async () => {
    journalTabColor();
    const sheet = await mountSheet({ policy: viewerPolicy() });
    const s = sheet.workbook.sheet(0);
    const history = workbookHandleOf(sheet.workbook).attachedHistory();
    sheet.workbook.batch(() => s.setNumber('A1', 1));
    expect(history?.undo()).toBe(false);
    expect(s.getTypedValue('A1').kind).toBe(ValueKind.Number);
    sheet.workbook.batch(() => {
      s.setNumber('A2', 2);
      s.setTabColor(RED);
    });
    expect(history?.undo()).toBe(false);
    expect(s.getTabColor()).not.toBeNull();
    expect(sheet.history.canUndo).toBe(true);
  });
});
