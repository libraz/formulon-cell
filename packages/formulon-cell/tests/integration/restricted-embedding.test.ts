import { afterEach, describe, expect, it } from 'vitest';
import { fixedFormPolicy, viewerPolicy } from '../../src/commands/interaction-policy.js';
import { WorkbookHandle } from '../../src/engine/workbook-handle.js';
import { mountToolbar } from '../../src/mount/toolbar.js';
import { createDialogShell } from '../../src/toolbar/dialogs/shell.js';
import { type MountedStubSheet, mountStubSheet } from '../test-utils/mount.js';

const a1 = { sheet: 0, row: 0, col: 0 };
const b1 = { sheet: 0, row: 0, col: 1 };
const range = { sheet: 0, r0: 0, c0: 0, r1: 1, c1: 1 };
let sheet: MountedStubSheet | undefined;
afterEach(() => {
  sheet?.dispose();
  sheet = undefined;
});

describe('restricted embedding public API', () => {
  it('prefills readonly cells through the host while rejecting user edits', async () => {
    sheet = await mountStubSheet({
      ui: { profile: 'embedded' },
      policy: viewerPolicy(),
      viewport: { range },
    });
    const { instance, workbook, host } = sheet;
    expect(host.querySelector('.fc-tb')).toBeNull();
    expect(instance.applyChanges([{ addr: a1, input: 'Prefilled' }]).status).toBe('applied');
    expect(
      instance.commands.execute({
        type: 'cellBatch',
        operation: 'valueEdit',
        origin: 'editor',
        changes: [{ addr: a1, input: 'Changed' }],
      }).status,
    ).toBe('rejected');
    expect(workbook.getValue(a1)).toEqual({ kind: 'text', value: 'Prefilled' });
    expect(instance.history.canUndo()).toBe(false);
    expect(
      instance.applyChanges([{ addr: { ...a1, row: 20 }, input: 'Hidden dependency' }]).status,
    ).toBe('applied');
    expect(workbook.getValue({ ...a1, row: 20 })).toEqual({
      kind: 'text',
      value: 'Hidden dependency',
    });
  });

  it('keeps a mixed paste atomic and supports an explicit skip without shifting addresses', async () => {
    sheet = await mountStubSheet({
      policy: fixedFormPolicy([{ ...range, c1: 0 }]),
      viewport: { range },
    });
    const { instance, workbook } = sheet;
    const changes = [
      { addr: a1, input: 'Editable' },
      { addr: b1, input: 'Locked' },
    ];
    const command = {
      type: 'cellBatch' as const,
      operation: 'paste' as const,
      origin: 'clipboard' as const,
      changes,
    };
    expect(instance.commands.execute(command).status).toBe('rejected');
    expect(workbook.getValue(a1).kind).toBe('blank');
    const result = instance.commands.execute({ ...command, denied: 'skipIneligible' });
    expect(result.applied).toEqual([a1]);
    expect(result.rejected[0]?.addr).toEqual(b1);
    expect(workbook.getValue(b1).kind).toBe('blank');
    expect(instance.undo()).toBe(true);
    expect(workbook.getValue(a1).kind).toBe('blank');
  });

  it('rechecks undo after policy changes and resumes it when permitted again', async () => {
    sheet = await mountStubSheet({ policy: fixedFormPolicy([range]) });
    const { instance, workbook } = sheet;
    instance.commands.execute({
      type: 'cellBatch',
      operation: 'valueEdit',
      origin: 'editor',
      changes: [{ addr: a1, input: '42' }],
    });
    instance.setPolicy(viewerPolicy());
    expect(instance.undo()).toBe(false);
    expect(workbook.getValue(a1)).toEqual({ kind: 'number', value: 42 });
    instance.setPolicy(fixedFormPolicy([range]));
    expect(instance.undo()).toBe(true);
    expect(workbook.getValue(a1).kind).toBe('blank');
  });

  it('emits one batch event only for a successful host update', async () => {
    sheet = await mountStubSheet({ policy: viewerPolicy(), viewport: { range } });
    const batches: unknown[] = [];
    sheet.instance.on('changeBatch', (event) => batches.push(event));
    expect(
      sheet.instance.applyChanges([
        { addr: a1, input: '1' },
        { addr: b1, input: '2' },
      ]).status,
    ).toBe('applied');
    expect(sheet.instance.applyChanges([{ addr: { ...a1, sheet: 5 }, input: '3' }]).status).toBe(
      'rejected',
    );
    expect(batches).toHaveLength(1);
  });
  it('mounts a fixed range on another sheet and rejects an incompatible replacement before detaching', async () => {
    const original = await WorkbookHandle.createDefault({ preferStub: true });
    original.addSheet('Form');
    const otherSheetRange = { ...range, sheet: 1 };
    const next = await WorkbookHandle.createDefault({ preferStub: true });
    try {
      sheet = await mountStubSheet({
        workbook: original,
        policy: viewerPolicy(),
        viewport: { range: otherSheetRange },
      });
      expect(sheet.instance.store.getState().data.sheetIndex).toBe(1);
      sheet.instance.applyChanges([{ addr: { ...a1, sheet: 1 }, input: 'Original' }]);
      await expect(sheet.instance.setWorkbook(next)).rejects.toThrow('outside the workbook');
      expect(sheet.instance.workbook).toBe(original);
      expect(
        sheet.instance.applyChanges([{ addr: { ...b1, sheet: 1 }, input: 'Still attached' }])
          .status,
      ).toBe('applied');
      expect(original.getValue({ ...b1, sheet: 1 })).toEqual({
        kind: 'text',
        value: 'Still attached',
      });
    } finally {
      next.dispose();
      sheet?.dispose();
      sheet = undefined;
      original.dispose();
    }
  });
  it('updates a separately mounted ribbon when policy is added and removed', async () => {
    sheet = await mountStubSheet();
    sheet.instance.applyChanges([{ addr: a1, input: 'Seeded' }]);
    const toolbarHost = document.createElement('div');
    document.body.append(toolbarHost);
    const toolbar = mountToolbar(toolbarHost, sheet.instance, { dynamicDropdowns: true });
    try {
      sheet.instance.setPolicy(viewerPolicy());
      const bold = toolbarHost.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]');
      expect(bold?.disabled).toBe(true);
      bold?.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      expect(sheet.instance.store.getState().format.formats.get('0:0:0')?.bold).not.toBe(true);
      sheet.instance.setPolicy(undefined);
      expect(
        toolbarHost.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]')?.disabled,
      ).toBe(false);
      toolbarHost.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]')?.click();
      expect(sheet.instance.store.getState().format.formats.get('0:0:0')?.bold).toBe(true);
      const overlayRoot = document.createElement('div');
      document.body.append(overlayRoot);
      sheet.instance.setOverlayOptions({ root: overlayRoot });
      toolbarHost.querySelector<HTMLButtonElement>('[data-ribbon-command="bold"]')?.focus();
      const dialog = createDialogShell({ title: 'External toolbar dialog' });
      expect(overlayRoot.contains(dialog.overlay)).toBe(true);
      dialog.overlay.remove();
      sheet.instance.setOverlayOptions(undefined);
      overlayRoot.remove();
    } finally {
      toolbar.dispose();
      toolbarHost.remove();
    }
  });
  it('switches and hydrates the requested sheet when changing a fixed viewport', async () => {
    const workbook = await WorkbookHandle.createDefault({ preferStub: true });
    workbook.addSheet('Another form');
    workbook.setText({ ...a1, sheet: 1 }, 'Other sheet');
    sheet = await mountStubSheet({ workbook, viewport: { range }, policy: viewerPolicy() });
    try {
      sheet.instance.setViewportOptions({ range: { ...range, sheet: 1 } });
      expect(sheet.instance.store.getState().data.sheetIndex).toBe(1);
      expect(sheet.instance.store.getState().data.cells.get('1:0:0')?.value).toEqual({
        kind: 'text',
        value: 'Other sheet',
      });
      expect(() => sheet?.instance.setViewportOptions({ range: { ...range, sheet: 5 } })).toThrow(
        'outside the workbook',
      );
      expect(sheet.instance.store.getState().data.sheetIndex).toBe(1);
    } finally {
      sheet.dispose();
      sheet = undefined;
      workbook.dispose();
    }
  });

  it('completes a workbook swap and later hooks when one extension hook throws', async () => {
    let notified = 0;
    sheet = await mountStubSheet({
      extensions: [
        {
          id: 'host-hooks',
          setup(context) {
            const stopThrowing = context.onWorkbookChange(() => {
              throw new Error('Host hook error');
            });
            const stopCounting = context.onWorkbookChange(() => {
              notified += 1;
            });
            return {
              dispose() {
                stopThrowing();
                stopCounting();
              },
              rebindWorkbook() {
                throw new Error('Extension hook error');
              },
            };
          },
        },
      ],
    });
    const next = await WorkbookHandle.createDefault({ preferStub: true });
    next.setText(a1, 'Replacement');
    await expect(sheet.instance.setWorkbook(next)).resolves.toBeUndefined();
    expect(notified).toBe(1);
    expect(sheet.instance.workbook).toBe(next);
    expect(sheet.instance.store.getState().data.cells.get('0:0:0')?.value).toEqual({
      kind: 'text',
      value: 'Replacement',
    });
  });

  it('commits real ACOS through the mounted Mac dialog and projects it into Recent', async () => {
    const workbook = await WorkbookHandle.createDefault();
    expect(workbook.isStub).toBe(false);
    const editable = { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 };
    sheet = await mountStubSheet({
      workbook,
      locale: 'en',
      ui: { profile: 'minimal' },
      features: { fxDialog: true },
    });
    let toolbarHost: HTMLDivElement | undefined;
    let toolbar: ReturnType<typeof mountToolbar> | undefined;
    try {
      sheet.instance.openFunctionArguments('ACOS');
      const input = document.querySelector<HTMLInputElement>('.fc-fxdialog__arg-input');
      if (!input) throw new Error('expected ACOS argument input');
      input.value = '0';
      input.dispatchEvent(new Event('input'));
      document.querySelector<HTMLButtonElement>('.fc-fxdialog .fc-fmtdlg__btn--primary')?.click();
      expect(workbook.cellFormula(a1)).toBe('=ACOS(0)');
      sheet.instance.setPolicy({
        editable: [editable],
        operations: { formulaEdit: true },
        defaultOperation: 'deny',
        selection: true,
      });
      toolbarHost = document.createElement('div');
      document.body.append(toolbarHost);
      toolbar = mountToolbar(toolbarHost, sheet.instance, { platform: 'mac', lang: 'en' });
      expect(toolbarHost.querySelector('[data-ribbon-command="mac.function.ACOS"]')).not.toBeNull();
      expect(toolbar.applyCommand('mac.function.NOT_A_FUNCTION')).toBe(true);
    } finally {
      toolbar?.dispose();
      toolbarHost?.remove();
    }
  });

  it('removes an engine-only Recent leaf after swapping to a static-catalog workbook', async () => {
    const workbook = await WorkbookHandle.createDefault();
    expect(workbook.isStub).toBe(false);
    expect(workbook.functionNames()).toContain('ACCRINT');
    sheet = await mountStubSheet({
      workbook,
      locale: 'en',
      ui: { profile: 'minimal' },
      features: { fxDialog: true },
    });
    let toolbarHost: HTMLDivElement | undefined;
    let toolbar: ReturnType<typeof mountToolbar> | undefined;
    try {
      sheet.instance.openFunctionArguments('ACCRINT');
      const input = document.querySelector<HTMLInputElement>('.fc-fxdialog__arg-input');
      if (!input) throw new Error('expected ACCRINT argument input');
      input.value = '0';
      input.dispatchEvent(new Event('input'));
      document.querySelector<HTMLButtonElement>('.fc-fxdialog .fc-fmtdlg__btn--primary')?.click();
      expect(workbook.cellFormula(a1)).toBe('=ACCRINT(0)');
      toolbarHost = document.createElement('div');
      document.body.append(toolbarHost);
      toolbar = mountToolbar(toolbarHost, sheet.instance, { platform: 'mac', lang: 'en' });
      expect(
        toolbarHost.querySelector('[data-ribbon-command="mac.function.ACCRINT"]'),
      ).not.toBeNull();

      const fallback = await WorkbookHandle.createDefault({ preferStub: true });
      await sheet.instance.setWorkbook(fallback);
      toolbar.rerender();
      expect(fallback.functionNames()).toBeNull();
      expect(toolbarHost.querySelector('[data-ribbon-command="mac.function.ACCRINT"]')).toBeNull();
    } finally {
      toolbar?.dispose();
      toolbarHost?.remove();
    }
  });
});
