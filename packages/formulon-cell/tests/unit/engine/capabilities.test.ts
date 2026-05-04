import { describe, expect, it } from 'vitest';
import { detectCapabilities } from '../../../src/engine/capabilities.js';
import type { Workbook } from '../../../src/engine/types.js';

const makeWb = (methods: string[]): Workbook => {
  const obj: Record<string, () => void> = {};
  for (const m of methods) obj[m] = () => {};
  return obj as unknown as Workbook;
};

describe('detectCapabilities', () => {
  it('cellFormatting requires the full XF resolver/dedup surface', () => {
    const partial = makeWb(['getCellXfIndex', 'setCellXfIndex', 'getCellXf']);
    expect(detectCapabilities(partial).cellFormatting).toBe(false);

    const full = makeWb([
      'getCellXfIndex',
      'setCellXfIndex',
      'getCellXf',
      'addFont',
      'addFill',
      'addBorder',
      'addNumFmt',
      'addXf',
      'getFont',
      'getFill',
      'getBorder',
      'getNumFmt',
    ]);
    expect(detectCapabilities(full).cellFormatting).toBe(true);
  });

  it('dataValidation requires read + write surface', () => {
    const readOnly = makeWb(['getValidations']);
    expect(detectCapabilities(readOnly).dataValidation).toBe(false);

    const full = makeWb(['getValidations', 'addValidation', 'clearValidations']);
    expect(detectCapabilities(full).dataValidation).toBe(true);
  });

  it('hyperlinks requires read + write surface', () => {
    const readOnly = makeWb(['getHyperlinks']);
    expect(detectCapabilities(readOnly).hyperlinks).toBe(false);

    const full = makeWb(['getHyperlinks', 'addHyperlink', 'clearHyperlinks']);
    expect(detectCapabilities(full).hyperlinks).toBe(true);
  });

  it('every flag is false for an empty wb', () => {
    const empty = makeWb([]);
    const caps = detectCapabilities(empty);
    expect(caps.merges).toBe(false);
    expect(caps.cellFormatting).toBe(false);
    expect(caps.dataValidation).toBe(false);
    expect(caps.hyperlinks).toBe(false);
    expect(caps.sheetMutate).toBe(false);
    expect(caps.insertDeleteRowsCols).toBe(false);
    expect(caps.hiddenRowsCols).toBe(false);
    expect(caps.colRowSize).toBe(false);
    expect(caps.freeze).toBe(false);
    expect(caps.sheetZoom).toBe(false);
    expect(caps.sheetTabHidden).toBe(false);
    expect(caps.outlines).toBe(false);
    expect(caps.comments).toBe(false);
    expect(caps.definedNameMutate).toBe(false);
    expect(caps.partialRecalc).toBe(false);
    expect(caps.iterativeProgress).toBe(false);
  });
});
