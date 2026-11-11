import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import {
  type AutoCorrectOptions,
  AutoCorrectStore,
  registerAutoCorrect,
} from '../../../src/commands/autocorrect.js';
import { commitCellInput } from '../../../src/commands/cell-input-commit.js';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { InlineEditor } from '../../../src/interact/editor.js';
import { createSpreadsheetStore, type SpreadsheetStore } from '../../../src/store/store.js';

const A1 = { sheet: 0, row: 0, col: 0 };

describe('AutoCorrect on cell entry', () => {
  let store: SpreadsheetStore;
  let wb: WorkbookHandle;
  let autoCorrect: AutoCorrectStore;
  let options: AutoCorrectOptions;
  let unregister: () => void;

  beforeEach(async () => {
    store = createSpreadsheetStore();
    wb = await WorkbookHandle.createDefault({ preferStub: true });
    options = {};
    autoCorrect = new AutoCorrectStore(
      () => options,
      () => 'en',
      (next) => {
        options = next;
      },
    );
    registerAutoCorrect(store, autoCorrect);
    const controller = new InteractionController({
      store,
      getWb: () => wb,
      history: new History(),
    });
    unregister = registerInteractionController(store, controller);
  });

  afterEach(() => {
    unregister();
    wb.dispose();
    document.body.replaceChildren();
  });

  const typed = (raw: string, isTyped = true): unknown => {
    commitCellInput({ store, wb, addr: A1, raw, origin: 'editor', typed: isTyped });
    return wb.getValue(A1);
  };

  it('corrects a typed constant and leaves formulas as typed', () => {
    expect(typed('(c) 2026')).toEqual({ kind: 'text', value: '© 2026' });
    typed('="(c) x"');
    expect(wb.cellFormula(A1)).toBe('="(c) x"');
  });

  it('applies the capitalisation rules to a typed constant', () => {
    expect(typed('THe end. see you monday')).toEqual({
      kind: 'text',
      value: 'The end. See you Monday',
    });
  });

  it('leaves an entry that was not typed, such as a list pick, as it is', () => {
    expect(typed('(c) 2026', false)).toEqual({ kind: 'text', value: '(c) 2026' });
  });

  it('stops correcting when Replace text as you type is off', () => {
    autoCorrect.update({ replaceAsYouType: false });
    expect(typed('(c) 2026')).toEqual({ kind: 'text', value: '(c) 2026' });
  });

  it('corrects every cell Ctrl+Enter fills', () => {
    const host = document.createElement('div');
    const grid = document.createElement('div');
    host.appendChild(grid);
    document.body.appendChild(host);
    store.setState((s) => ({
      ...s,
      selection: {
        active: A1,
        anchor: A1,
        range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 1 },
        extraRanges: [],
      },
    }));
    const editor = new InlineEditor({ host, grid, store, wb, onAfterCommit: () => {} });
    editor.begin('(tm)');
    editor.commitMulti();
    expect(wb.getValue(A1)).toEqual({ kind: 'text', value: '™' });
    expect(wb.getValue({ sheet: 0, row: 0, col: 1 })).toEqual({ kind: 'text', value: '™' });
    if (editor.isActive()) editor.cancel();
  });
});
