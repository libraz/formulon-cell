import { describe, expect, it, vi } from 'vitest';
import { commitCellInput } from '../../src/commands/cell-input-commit.js';
import { applySelectionFormatAction, setFont } from '../../src/commands/format.js';
import { hyperlinkAt } from '../../src/commands/hyperlinks.js';
import { interactionControllerFor } from '../../src/commands/interaction-controller.js';
import type { InteractionPolicy } from '../../src/commands/interaction-policy.js';
import { setProtectedSheet } from '../../src/commands/protection.js';
import { flushFormatToEngine } from '../../src/engine/cell-format-sync.js';
import { fontPointsToCssPx } from '../../src/engine/font-units.js';
import { cssColorFromArgb } from '../../src/engine/theme-color.js';
import type { Addr } from '../../src/engine/types.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { createSpreadsheetStore, mutators } from '../../src/store/store.js';
import { mountStubSheet } from '../test-utils/mount.js';

const A1 = { sheet: 0, row: 0, col: 0 } as const;
const B1 = { sheet: 0, row: 0, col: 1 } as const;
const C1 = { sheet: 0, row: 0, col: 2 } as const;
const D1 = { sheet: 0, row: 0, col: 3 } as const;
const FAR = { sheet: 0, row: 100, col: 20 } as const;
const URL = 'https://example.com/formulon-native-check?source=typed#cell';

function rawNamedStyleMetadata(workbook: WorkbookHandle) {
  return Array.from({ length: workbook.cellStyleCount() }, (_, index) =>
    workbook.getCellStyle(index),
  ).filter((style) => style !== null);
}

function addLocalizedHyperlinkStyle(workbook: WorkbookHandle, hidden: boolean) {
  const normalXf = workbook.getCellStyleXf(0);
  const normalFont = workbook.getFontRecord(0);
  const theme = workbook.getTheme();
  const hyperlinkColor = theme?.colors[10];
  if (!normalXf || !normalFont || hyperlinkColor === undefined)
    throw new Error('default font, Normal XF, or hyperlink theme colour is unavailable');

  const fontIndex = workbook.addFontRecord({
    ...normalFont,
    size: normalFont.size,
    bold: false,
    italic: false,
    strike: false,
    underline: 1,
    vertAlign: 0,
    colorArgb: hyperlinkColor,
    color: { kind: 2, rgb: 0, theme: 10, tint: 0, indexed: 0 },
  });
  const xfId = workbook.addCellStyleXfRecord({ ...normalXf, fontIndex });
  if (fontIndex < 0 || xfId < 0) throw new Error('failed to author localized Hyperlink style');
  if (
    !workbook.setNamedCellStyleRecord({
      name: 'LienWeb',
      xfId,
      builtinId: 8,
      iLevel: 0,
      hidden,
      customBuiltin: true,
    })
  ) {
    throw new Error('failed to register localized Hyperlink style');
  }
  const font = workbook.getResolvedFontRecord(fontIndex);
  if (!font) throw new Error('localized Hyperlink font could not be resolved');
  return font;
}

function fontAt(workbook: WorkbookHandle, addr: Addr) {
  const effective = workbook.getEffectiveStyle(addr.sheet, addr.row, addr.col);
  const xf = effective ? workbook.getCellXf(effective.xfIndex) : null;
  const font = xf ? workbook.getResolvedFontRecord(xf.fontIndex) : null;
  if (!font) throw new Error(`effective font missing at ${addr.sheet}:${addr.row}:${addr.col}`);
  return font;
}

describe('typed URL autoformat', () => {
  it.each([true, false])(
    'uses the raw localized built-in style and keeps its catalog record unchanged (hidden=%s)',
    async (hidden) => {
      const workbook = await WorkbookHandle.createDefault({ locale: 'en' });
      const hyperlinkFont = addLocalizedHyperlinkStyle(workbook, hidden);
      const beforeMountCatalog = rawNamedStyleMetadata(workbook);
      const mounted = await mountStubSheet({ workbook, locale: 'en' });
      const controller = interactionControllerFor(mounted.instance.store);
      const notifiedHyperlinks: Array<string | null> = [];
      const unsubscribe = controller?.subscribe(() => {
        notifiedHyperlinks.push(hyperlinkAt(mounted.instance.store.getState(), A1));
      });
      try {
        expect(workbook.isStub).toBe(false);
        expect(rawNamedStyleMetadata(workbook)).toEqual(beforeMountCatalog);

        expect(
          commitCellInput({
            store: mounted.instance.store,
            wb: workbook,
            addr: A1,
            raw: URL,
            origin: 'editor',
          }),
        ).toMatchObject({ status: 'applied' });
        expect(workbook.getValue(A1)).toEqual({ kind: 'text', value: URL });
        expect(workbook.cellFormula(A1)).toBeNull();
        expect(hyperlinkAt(mounted.instance.store.getState(), A1)).toBe(URL);
        expect(workbook.getHyperlinksFull(0)).toMatchObject([{ row: 0, col: 0, target: URL }]);
        expect(mounted.instance.store.getState().format.formats.get('0:0:0')).toMatchObject({
          fontFamily: hyperlinkFont.name,
          fontSize: fontPointsToCssPx(hyperlinkFont.size),
          bold: hyperlinkFont.bold,
          italic: hyperlinkFont.italic,
          strike: hyperlinkFont.strike,
          underline: 'single',
          color: cssColorFromArgb(hyperlinkFont.colorArgb),
        });
        expect(notifiedHyperlinks).toEqual([URL]);
        expect(rawNamedStyleMetadata(workbook)).toEqual(beforeMountCatalog);

        expect(mounted.instance.history.undo()).toBe(true);
        expect(workbook.getValue(A1)).toEqual({ kind: 'text', value: URL });
        expect(hyperlinkAt(mounted.instance.store.getState(), A1)).toBeNull();
        expect(workbook.getHyperlinksFull(0)).toEqual([]);

        expect(mounted.instance.history.undo()).toBe(true);
        expect(workbook.getValue(A1)).toEqual({ kind: 'blank' });
        expect(hyperlinkAt(mounted.instance.store.getState(), A1)).toBeNull();

        expect(mounted.instance.history.redo()).toBe(true);
        expect(workbook.getValue(A1)).toEqual({ kind: 'text', value: URL });
        expect(hyperlinkAt(mounted.instance.store.getState(), A1)).toBeNull();
        expect(mounted.instance.history.redo()).toBe(true);
        expect(hyperlinkAt(mounted.instance.store.getState(), A1)).toBe(URL);
        expect(workbook.getHyperlinksFull(0)).toMatchObject([{ row: 0, col: 0, target: URL }]);
        expect(rawNamedStyleMetadata(workbook)).toEqual(beforeMountCatalog);
      } finally {
        unsubscribe?.();
        mounted.dispose();
        workbook.dispose();
      }
    },
  );

  it('skips automatic format inside an existing outer history transaction', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en' });
    addLocalizedHyperlinkStyle(workbook, true);
    const mounted = await mountStubSheet({ workbook, locale: 'en' });
    try {
      const transaction = mounted.instance.history.begin();
      const result = commitCellInput({
        store: mounted.instance.store,
        wb: workbook,
        addr: A1,
        raw: URL,
        origin: 'editor',
      });

      expect(result.status).toBe('applied');
      expect(workbook.getValue(A1)).toEqual({ kind: 'text', value: URL });
      expect(hyperlinkAt(mounted.instance.store.getState(), A1)).toBeNull();
      expect(() => mounted.instance.history.end(transaction)).not.toThrow();
      expect(mounted.instance.history.undo()).toBe(true);
      expect(workbook.getValue(A1)).toEqual({ kind: 'blank' });
      expect(hyperlinkAt(mounted.instance.store.getState(), A1)).toBeNull();
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it('keeps non-editor origins, apostrophe input, formulas, no-ops, and spanning links unconverted', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en' });
    addLocalizedHyperlinkStyle(workbook, false);
    const mounted = await mountStubSheet({ workbook, locale: 'en' });
    const controller = interactionControllerFor(mounted.instance.store);
    try {
      const commit = (addr: Addr, raw: string, origin: 'editor' | 'formulaBar') =>
        commitCellInput({ store: mounted.instance.store, wb: workbook, addr, raw, origin });

      expect(commit(B1, URL, 'formulaBar').status).toBe('applied');
      expect(hyperlinkAt(mounted.instance.store.getState(), B1)).toBeNull();
      expect(commit(B1, URL, 'editor').status).toBe('applied');
      expect(hyperlinkAt(mounted.instance.store.getState(), B1)).toBeNull();

      expect(commit(C1, `'${URL}`, 'editor').status).toBe('applied');
      expect(workbook.getValue(C1)).toEqual({ kind: 'text', value: URL });
      expect(hyperlinkAt(mounted.instance.store.getState(), C1)).toBeNull();

      expect(commit(D1, '=1+1', 'editor').status).toBe('applied');
      expect(workbook.cellFormula(D1)).toBe('=1+1');
      expect(hyperlinkAt(mounted.instance.store.getState(), D1)).toBeNull();

      controller?.execute({
        type: 'cellBatch',
        operation: 'paste',
        origin: 'clipboard',
        changes: [{ addr: { sheet: 0, row: 1, col: 0 }, input: URL }],
      });
      expect(
        hyperlinkAt(mounted.instance.store.getState(), { sheet: 0, row: 1, col: 0 }),
      ).toBeNull();
      controller?.applyChanges([{ addr: { sheet: 0, row: 1, col: 1 }, input: URL }], {
        history: 'record',
        origin: 'editor',
      });
      expect(
        hyperlinkAt(mounted.instance.store.getState(), { sheet: 0, row: 1, col: 1 }),
      ).toBeNull();

      workbook.setText(B1, 'before span');
      mutators.setCellFormat(mounted.instance.store, A1, {
        hyperlink: 'https://anchor.example/',
        hyperlinkSpan: { lastRow: 1, lastCol: 2 },
      });
      expect(commit(B1, URL, 'editor').status).toBe('applied');
      expect(hyperlinkAt(mounted.instance.store.getState(), A1)).toBe('https://anchor.example/');
      expect(hyperlinkAt(mounted.instance.store.getState(), B1)).toBe('https://anchor.example/');

      mutators.setCellFormat(mounted.instance.store, A1, null);
      workbook.setText(C1, 'before own link');
      mutators.setCellFormat(mounted.instance.store, C1, { hyperlink: 'https://old.example/' });
      expect(commit(C1, URL, 'editor').status).toBe('applied');
      expect(hyperlinkAt(mounted.instance.store.getState(), C1)).toBe(URL);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it.each([
    { format: false, hyperlink: true },
    { format: true, hyperlink: false },
  ])(
    'keeps value input when policy denies automatic $format/$hyperlink formatting',
    async (policy) => {
      const workbook = await WorkbookHandle.createDefault({ locale: 'en' });
      addLocalizedHyperlinkStyle(workbook, false);
      const mounted = await mountStubSheet({ workbook, locale: 'en' });
      const controller = interactionControllerFor(mounted.instance.store);
      try {
        controller?.setPolicy({
          operations: { valueEdit: true, format: policy.format, hyperlink: policy.hyperlink },
        });
        const result = commitCellInput({
          store: mounted.instance.store,
          wb: workbook,
          addr: A1,
          raw: URL,
          origin: 'editor',
        });
        expect(result.status).toBe('applied');
        expect(workbook.getValue(A1)).toEqual({ kind: 'text', value: URL });
        expect(hyperlinkAt(mounted.instance.store.getState(), A1)).toBeNull();

        controller?.setPolicy({
          operations: { valueEdit: true, format: false, hyperlink: true },
        });
        mutators.setPendingFormat(mounted.instance.store, { addr: B1, format: { bold: true } });
        const explicitDenied = commitCellInput({
          store: mounted.instance.store,
          wb: workbook,
          addr: B1,
          raw: URL,
          origin: 'editor',
        });
        expect(explicitDenied.status).toBe('rejected');
        expect(workbook.getValue(B1)).toEqual({ kind: 'blank' });
        expect(hyperlinkAt(mounted.instance.store.getState(), B1)).toBeNull();
      } finally {
        mounted.dispose();
        workbook.dispose();
      }
    },
  );

  it.each([
    { formatCells: true, insertHyperlinks: true, linked: true },
    { formatCells: false, insertHyperlinks: true, linked: false },
    { formatCells: true, insertHyperlinks: false, linked: false },
  ])('requires both protected-sheet permissions for an unlocked cell', async (permissions) => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en' });
    addLocalizedHyperlinkStyle(workbook, false);
    const mounted = await mountStubSheet({ workbook, locale: 'en' });
    try {
      mutators.setCellFormat(mounted.instance.store, A1, { locked: false });
      setProtectedSheet(mounted.instance.store, 0, true, {
        permissions: {
          formatCells: permissions.formatCells,
          insertHyperlinks: permissions.insertHyperlinks,
        },
      });
      const result = commitCellInput({
        store: mounted.instance.store,
        wb: workbook,
        addr: A1,
        raw: URL,
        origin: 'editor',
      });
      expect(result.status).toBe('applied');
      expect(workbook.getValue(A1)).toEqual({ kind: 'text', value: URL });
      expect(hyperlinkAt(mounted.instance.store.getState(), A1) === URL).toBe(permissions.linked);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it.each(['store', 'policy', 'workbook'] as const)(
    'skips stale autoformat when a reentrant host restriction changes %s state',
    async (reentrantChange) => {
      const workbook = await WorkbookHandle.createDefault({ locale: 'en' });
      addLocalizedHyperlinkStyle(workbook, false);
      const mounted = await mountStubSheet({ workbook, locale: 'en' });
      const controller = interactionControllerFor(mounted.instance.store);
      let reentered = false;
      let policy: InteractionPolicy;
      policy = {
        operations: { valueEdit: true, format: true, hyperlink: true },
        restrict: ({ commandId, operation }) => {
          if (commandId !== 'editor.autoHyperlink' || operation !== 'format' || reentered) return;
          reentered = true;
          if (reentrantChange === 'store') {
            mutators.setCellFormat(mounted.instance.store, A1, { color: '#123456' });
          } else if (reentrantChange === 'policy') {
            controller?.setPolicy(policy);
          } else {
            workbook.setText(A1, 'host changed the cell');
          }
        },
      };
      try {
        controller?.setPolicy(policy);
        const result = commitCellInput({
          store: mounted.instance.store,
          wb: workbook,
          addr: A1,
          raw: URL,
          origin: 'editor',
        });

        expect(result.status).toBe('applied');
        expect(reentered).toBe(true);
        expect(hyperlinkAt(mounted.instance.store.getState(), A1)).toBeNull();
        if (reentrantChange === 'workbook') {
          expect(workbook.getValue(A1)).toEqual({ kind: 'text', value: 'host changed the cell' });
        } else {
          expect(workbook.getValue(A1)).toEqual({ kind: 'text', value: URL });
        }
        if (reentrantChange === 'store') {
          expect(mounted.instance.store.getState().format.formats.get('0:0:0')?.color).toBe(
            '#123456',
          );
        }
      } finally {
        mounted.dispose();
        workbook.dispose();
      }
    },
  );

  it.each([false, true])(
    'keeps controllerless URL commits format-neutral with an unhydrated store bridge=%s',
    async (attachBridge) => {
      const workbook = await WorkbookHandle.createDefault({ locale: 'en' });
      const store = createSpreadsheetStore();
      const baseXf = workbook.getCellXf(0);
      const baseFont = workbook.getFontRecord(0);
      if (!baseXf || !baseFont) throw new Error('default XF/font is unavailable');
      const fontIndex = workbook.addFontRecord({ ...baseFont, size: 17 });
      const originalXf = workbook.addXfRecord({ ...baseXf, fontIndex, applyFont: true });
      if (fontIndex < 0 || originalXf < 0 || !workbook.setCellXfIndex(0, 0, 0, originalXf)) {
        throw new Error('could not create A1 native font');
      }
      workbook.setText(B1, 'existing native link');
      expect(workbook.addHyperlink(0, 0, 1, 'https://existing.example/', 'Existing', 'Tip')).toBe(
        true,
      );
      expect(
        workbook.addValidationEntry(0, {
          ranges: [{ sheet: 0, r0: 0, c0: 1, r1: 0, c1: 1 }],
          type: 3,
          op: 0,
          allowBlank: true,
          showInputMessage: true,
          showErrorMessage: true,
          formula1: '"alpha,beta"',
          promptTitle: 'Keep this rule',
        }),
      ).toBe(true);
      if (attachBridge) workbook.attachStore(store);

      const beforeLinks = workbook.getHyperlinksFull(0);
      const beforeValidations = workbook.getValidationsForSheet(0);
      const clearHyperlinks = vi.spyOn(workbook, 'clearHyperlinks');
      const clearValidations = vi.spyOn(workbook, 'clearValidations');
      const clearCellXf = vi.spyOn(workbook, 'clearCellXfIndex');
      let restored: WorkbookHandle | undefined;
      try {
        expect(workbook.getCellXfIndex(A1.sheet, A1.row, A1.col)).toBe(originalXf);
        expect(fontAt(workbook, A1).size).toBe(17);
        expect(store.getState().format.formats.size).toBe(0);

        const result = commitCellInput({
          store,
          wb: workbook,
          addr: A1,
          raw: URL,
          origin: 'editor',
        });

        expect(result.status).toBe('applied');
        expect(workbook.getValue(A1)).toEqual({ kind: 'text', value: URL });
        expect(workbook.getCellXfIndex(A1.sheet, A1.row, A1.col)).toBe(originalXf);
        expect(fontAt(workbook, A1).size).toBe(17);
        expect(workbook.getHyperlinksFull(0)).toEqual(beforeLinks);
        expect(workbook.getHyperlinksFull(0)).toMatchObject([
          { row: 0, col: 1, target: 'https://existing.example/' },
        ]);
        expect(workbook.getValidationsForSheet(0)).toEqual(beforeValidations);
        expect(hyperlinkAt(store.getState(), A1)).toBeNull();
        expect(clearHyperlinks).not.toHaveBeenCalled();
        expect(clearValidations).not.toHaveBeenCalled();
        expect(clearCellXf).not.toHaveBeenCalled();

        restored = await WorkbookHandle.loadBytes(workbook.save());
        expect(restored.getValue(A1)).toEqual({ kind: 'text', value: URL });
        expect(fontAt(restored, A1).size).toBe(17);
        expect(restored.getHyperlinksFull(0)).toEqual(beforeLinks);
        expect(restored.getValidationsForSheet(0)).toEqual(beforeValidations);
        expect(clearHyperlinks).not.toHaveBeenCalled();
        expect(clearValidations).not.toHaveBeenCalled();
        expect(clearCellXf).not.toHaveBeenCalled();
      } finally {
        vi.restoreAllMocks();
        restored?.dispose();
        if (attachBridge) workbook.detachStore(store);
        workbook.dispose();
      }
    },
  );

  it('rechecks policy and protected-sheet permissions before history replay', async () => {
    const workbook = await WorkbookHandle.createDefault({ locale: 'en' });
    addLocalizedHyperlinkStyle(workbook, false);
    const mounted = await mountStubSheet({ workbook, locale: 'en' });
    const controller = interactionControllerFor(mounted.instance.store);
    try {
      controller?.setPolicy({ operations: { valueEdit: true, format: true, hyperlink: true } });
      expect(
        commitCellInput({
          store: mounted.instance.store,
          wb: workbook,
          addr: A1,
          raw: URL,
          origin: 'editor',
        }).status,
      ).toBe('applied');
      controller?.setPolicy({ operations: { valueEdit: true, format: true, hyperlink: false } });
      expect(mounted.instance.history.undo()).toBe(false);
      expect(hyperlinkAt(mounted.instance.store.getState(), A1)).toBe(URL);
      mounted.instance.history.clear();

      controller?.setPolicy(undefined);
      mutators.setCellFormat(mounted.instance.store, B1, { locked: false });
      expect(
        commitCellInput({
          store: mounted.instance.store,
          wb: workbook,
          addr: B1,
          raw: URL,
          origin: 'editor',
        }).status,
      ).toBe('applied');
      setProtectedSheet(mounted.instance.store, 0, true, {
        permissions: { formatCells: true, insertHyperlinks: true },
      });
      setProtectedSheet(mounted.instance.store, 0, true, {
        permissions: { formatCells: true, insertHyperlinks: false },
      });
      expect(mounted.instance.history.undo()).toBe(false);
      expect(hyperlinkAt(mounted.instance.store.getState(), B1)).toBe(URL);
    } finally {
      mounted.dispose();
      workbook.dispose();
    }
  });

  it.each([
    { otherCell: 'blank', populateOther: false },
    { otherCell: 'populated', populateOther: true },
  ])(
    'persists whole-sheet font exceptions after an unrelated fill flush when B1 is $otherCell',
    async ({ populateOther }) => {
      const workbook = await WorkbookHandle.createDefault({ locale: 'en' });
      addLocalizedHyperlinkStyle(workbook, true);
      workbook.setText(A1, 'before');
      if (populateOther) workbook.setText(B1, 'other');
      workbook.setText(FAR, 'far');
      const mounted = await mountStubSheet({ workbook, locale: 'en' });
      let restored: WorkbookHandle | undefined;
      let restoredMount: Awaited<ReturnType<typeof mountStubSheet>> | undefined;
      let restoredAgain: WorkbookHandle | undefined;
      try {
        mutators.selectAll(mounted.instance.store);
        expect(
          setFont(mounted.instance.store.getState(), mounted.instance.store, { fontSize: 32 }),
        ).toBe(true);
        expect(fontAt(workbook, A1).size).toBe(24);
        expect(fontAt(workbook, FAR).size).toBe(24);

        expect(
          commitCellInput({
            store: mounted.instance.store,
            wb: workbook,
            addr: A1,
            raw: URL,
            origin: 'editor',
          }).status,
        ).toBe('applied');
        expect(fontAt(workbook, A1).size).toBe(11);
        expect(fontAt(workbook, FAR).size).toBe(24);

        const firstReload = await WorkbookHandle.loadBytes(workbook.save());
        restored = firstReload;
        restoredMount = await mountStubSheet({ workbook: firstReload, locale: 'en' });
        expect(fontAt(firstReload, A1).size).toBe(11);
        expect(fontAt(firstReload, FAR).size).toBe(24);
        expect(hyperlinkAt(restoredMount.instance.store.getState(), A1)).toBe(URL);

        mutators.setActive(restoredMount.instance.store, B1);
        expect(
          applySelectionFormatAction(
            restoredMount.instance.store.getState(),
            restoredMount.instance.store,
            { patch: { fill: '#fff2cc' } },
            { allowPending: false },
          ),
        ).toBe(true);
        flushFormatToEngine(firstReload, restoredMount.instance.store, 0);
        const secondReload = await WorkbookHandle.loadBytes(firstReload.save());
        restoredAgain = secondReload;
        const finalMount = await mountStubSheet({ workbook: secondReload, locale: 'en' });
        try {
          expect(fontAt(secondReload, A1).size).toBe(11);
          expect(fontAt(secondReload, FAR).size).toBe(24);
          expect(fontAt(secondReload, B1).size).toBe(24);
          expect(hyperlinkAt(finalMount.instance.store.getState(), A1)).toBe(URL);
          expect(rawNamedStyleMetadata(secondReload)).toEqual(rawNamedStyleMetadata(workbook));
        } finally {
          finalMount.dispose();
        }
      } finally {
        restoredMount?.dispose();
        restoredAgain?.dispose();
        restored?.dispose();
        mounted.dispose();
        workbook.dispose();
      }
    },
  );
});
