import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../../src/engine/workbook-handle.js';
import { createEditingDropdownDefaults } from '../../../../src/mount/dynamic-dropdown-defaults/editing.js';
import { mutators } from '../../../../src/store/store.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';

describe('Fill > Justify from the Home dropdown', () => {
  let sheet: MountedStubSheet;

  beforeEach(async () => {
    sheet = await mountStubSheet({
      workbook: await WorkbookHandle.createDefault({ preferStub: true }),
      locale: 'en',
    });
  });
  afterEach(() => sheet.dispose());

  it('re-flows the selected text down its first column and undoes in one step', () => {
    sheet.workbook.setText({ sheet: 0, row: 0, col: 0 }, 'one two three four five six seven');
    sheet.instance.store.setState((s) => ({ ...s, layout: { ...s.layout, defaultColWidth: 40 } }));
    mutators.setRange(sheet.instance.store, { sheet: 0, r0: 0, c0: 0, r1: 4, c1: 1 });
    createEditingDropdownDefaults(sheet.instance).applyFillDirection('justify');
    const first = sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 });
    const second = sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 });
    expect(first.kind).toBe('text');
    expect(second.kind).toBe('text');
    expect(sheet.instance.history.undo()).toBe(true);
    expect(sheet.workbook.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({
      kind: 'text',
      value: 'one two three four five six seven',
    });
    expect(sheet.workbook.getValue({ sheet: 0, row: 1, col: 0 }).kind).toBe('blank');
  });
});
