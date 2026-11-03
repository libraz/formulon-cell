import { describe, expect, it } from 'vitest';
import { isValidDefinedName } from '../../../src/commands/named-ranges.js';

// Excel Names.Add results.
const REJECTED = [
  'A1',
  'R1C1',
  'R',
  'C',
  'r',
  'c',
  'R1',
  'C1',
  'RC',
  'R1C',
  'R2C',
  'RC1',
  'R1C1x',
  'R[1]C',
  '1abc',
  '1',
  'a b',
  'a-b',
  'Z$',
  '.a',
  '?a',
  'TRUE',
  '_xlnm.Print_Area',
  'ok1',
  'AB1',
  'XFD1',
  'a'.repeat(256),
];
const ACCEPTED = [
  '日本語名',
  'ｱｲｳ',
  '＿a',
  'é',
  'ä1',
  '_ok',
  '_1',
  '\\ok',
  '\\',
  'a.b',
  'a?b',
  'a1b',
  'a_1',
  'a',
  'Rx',
  'Cx',
  'ABCD1',
  'XFE1',
  'Sheet1',
  'a'.repeat(255),
  'Print_Area',
];

describe('isValidDefinedName as Excel measures it', () => {
  it.each(REJECTED.map((n) => [n.length > 20 ? `${n.length} chars` : n, n]))(
    'rejects %s',
    (_label, name) => {
      expect(isValidDefinedName(name)).toBe(false);
    },
  );
  it.each(ACCEPTED.map((n) => [n.length > 20 ? `${n.length} chars` : n, n]))(
    'accepts %s',
    (_label, name) => {
      expect(isValidDefinedName(name)).toBe(true);
    },
  );
});

describe('upsertDefinedName as Excel measures it', () => {
  it('replaces an existing name case-insensitively, taking the new casing', async () => {
    const { WorkbookHandle } = await import('../../../src/engine/workbook-handle.js');
    const { upsertDefinedName, listDefinedNames } = await import(
      '../../../src/commands/named-ranges.js'
    );
    const wb = await WorkbookHandle.createDefault();
    expect(upsertDefinedName(wb, 'dup', '=Sheet1!$A$1').ok).toBe(true);
    expect(upsertDefinedName(wb, 'DUP', '=Sheet1!$B$2').ok).toBe(true);
    const workbookScoped = listDefinedNames(wb).filter((n) => n.localSheetId === -1);
    expect(workbookScoped.map((n) => [n.name, n.formula])).toEqual([['DUP', 'Sheet1!$B$2']]);
  });

  it('keeps a sheet-scoped name beside a workbook-scoped one of the same spelling', async () => {
    const { WorkbookHandle } = await import('../../../src/engine/workbook-handle.js');
    const { upsertDefinedName, listDefinedNames } = await import(
      '../../../src/commands/named-ranges.js'
    );
    const wb = await WorkbookHandle.createDefault();
    expect(upsertDefinedName(wb, 'DUP', '=Sheet1!$A$1').ok).toBe(true);
    expect(upsertDefinedName(wb, 'dup', '=Sheet1!$B$2', 0).ok).toBe(true);
    expect(listDefinedNames(wb)).toHaveLength(2);
  });

  it('accepts constant refers-to and refuses a syntax error', async () => {
    const { WorkbookHandle } = await import('../../../src/engine/workbook-handle.js');
    const { upsertDefinedName } = await import('../../../src/commands/named-ranges.js');
    const wb = await WorkbookHandle.createDefault();
    expect(upsertDefinedName(wb, 'k1x', '=1+2').ok).toBe(true);
    expect(upsertDefinedName(wb, 'k2x', '="txt"').ok).toBe(true);
    expect(upsertDefinedName(wb, 'k3x', '=1+').ok).toBe(false);
  });
});
