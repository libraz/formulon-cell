import { describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  deleteDefinedName,
  listDefinedNames,
  recordDefinedNamesChange,
  upsertDefinedName,
} from '../../../src/commands/named-ranges.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { collectAutocompleteNames } from '../../../src/interact/autocomplete.js';
import { rangeDefinedNames, resolveAreaText } from '../../../src/mount/ref-utils.js';

const canLoadWasm = (): boolean => typeof WebAssembly !== 'undefined';

describe.skipIf(!canLoadWasm())('defined name comment and hidden', () => {
  it('round-trips a comment through save and load', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      const result = upsertDefinedName(wb, 'Rate', 'Sheet1!$A$1', -1, { comment: 'Tax rate' });
      expect(result).toMatchObject({ ok: true, entry: { name: 'Rate', comment: 'Tax rate' } });
      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        expect(listDefinedNames(restored)).toMatchObject([{ name: 'Rate', comment: 'Tax rate' }]);
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('keeps the comment when only the formula changes and clears it on request', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      upsertDefinedName(wb, 'Rate', 'Sheet1!$A$1', -1, { comment: 'note' });
      upsertDefinedName(wb, 'Rate', 'Sheet1!$B$2');
      expect(listDefinedNames(wb)).toMatchObject([{ formula: 'Sheet1!$B$2', comment: 'note' }]);
      upsertDefinedName(wb, 'Rate', 'Sheet1!$B$2', -1, { comment: '' });
      expect(listDefinedNames(wb)[0]?.comment).toBeUndefined();
    } finally {
      wb.dispose();
    }
  });

  it('undoes and redoes a comment edit', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      upsertDefinedName(wb, 'Rate', 'Sheet1!$A$1', -1, { comment: 'one' });
      const history = new History();
      recordDefinedNamesChange(history, wb, () =>
        upsertDefinedName(wb, 'Rate', 'Sheet1!$A$1', -1, { comment: 'two' }),
      );
      expect(listDefinedNames(wb)[0]?.comment).toBe('two');
      history.undo();
      expect(listDefinedNames(wb)[0]?.comment).toBe('one');
      history.redo();
      expect(listDefinedNames(wb)[0]?.comment).toBe('two');
    } finally {
      wb.dispose();
    }
  });

  it('keeps hidden names in the engine, out of lists, and resolvable', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      upsertDefinedName(wb, 'Shown', 'Sheet1!$A$1');
      upsertDefinedName(wb, 'Secret', 'Sheet1!$C$3');
      expect(wb.setDefinedNameAttrs('Secret', -1, { hidden: true, comment: 'internal' })).toBe(
        true,
      );

      const restored = await WorkbookHandle.loadBytes(wb.save());
      try {
        const names = [...restored.definedNames()];
        expect(names.map((n) => [n.name, !!n.hidden])).toEqual(
          expect.arrayContaining([
            ['Shown', false],
            ['Secret', true],
          ]),
        );

        // Lists skip it; resolution does not.
        expect(collectAutocompleteNames(restored.definedNames(), [], 0).map((n) => n.name)).toEqual(
          ['Shown'],
        );
        expect(rangeDefinedNames(restored, 0).find((n) => n.name === 'Secret')?.hidden).toBe(true);
        expect(resolveAreaText(restored, 'Secret', 0)).toMatchObject({ r0: 2, c0: 2 });

        restored.setNumber({ sheet: 0, row: 2, col: 2 }, 42);
        restored.setFormula({ sheet: 0, row: 0, col: 5 }, '=Secret*2');
        restored.recalcAuto();
        expect(restored.getValue({ sheet: 0, row: 0, col: 5 })).toEqual({
          kind: 'number',
          value: 84,
        });
      } finally {
        restored.dispose();
      }
    } finally {
      wb.dispose();
    }
  });

  it('preserves a hidden name across edits and deletes of other names', async () => {
    const wb = await WorkbookHandle.createDefault();
    try {
      upsertDefinedName(wb, 'Secret', 'Sheet1!$C$3', -1, { comment: 'keep', hidden: true });
      upsertDefinedName(wb, 'Other', 'Sheet1!$A$1');
      const history = new History();
      recordDefinedNamesChange(history, wb, () =>
        upsertDefinedName(wb, 'Other', 'Sheet1!$A$2', -1, { comment: 'edited' }),
      );
      recordDefinedNamesChange(history, wb, () => deleteDefinedName(wb, 'Other'));
      const secret = () => listDefinedNames(wb).find((n) => n.name === 'Secret');
      expect(secret()).toMatchObject({ hidden: true, comment: 'keep', formula: 'Sheet1!$C$3' });

      history.undo();
      history.undo();
      expect(secret()).toMatchObject({ hidden: true, comment: 'keep' });
      history.redo();
      expect(secret()).toMatchObject({ hidden: true, comment: 'keep' });
    } finally {
      wb.dispose();
    }
  });
});
