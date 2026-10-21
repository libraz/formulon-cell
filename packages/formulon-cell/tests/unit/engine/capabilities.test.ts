import { describe, expect, it } from 'vitest';
import { detectCapabilities } from '../../../src/engine/capabilities.js';
import type { Workbook } from '../../../src/engine/types.js';

const makeWb = (methods: string[]): Workbook => {
  const obj: Record<string, () => void> = {};
  for (const m of methods) obj[m] = () => {};
  return obj as unknown as Workbook;
};

describe('detectCapabilities', () => {
  it('requires the paired native range-style operations independently', () => {
    expect(detectCapabilities(makeWb(['rangeStyleBases'])).rangeStylePatch).toBe(false);
    expect(detectCapabilities(makeWb(['rangeStyleBases', 'remapRangeStyle'])).rangeStylePatch).toBe(
      true,
    );
    expect(detectCapabilities(makeWb(['snapshotRangeStyle'])).rangeStyleSnapshot).toBe(false);
    expect(
      detectCapabilities(makeWb(['snapshotRangeStyle', 'restoreRangeStyle'])).rangeStyleSnapshot,
    ).toBe(true);
    expect(detectCapabilities(makeWb(['setRowStyle', 'clearRowStyle'])).rowStyleRanges).toBe(false);
    expect(
      detectCapabilities(makeWb(['rangeStyleBases', 'setRowStyle', 'clearRowStyle']))
        .rowStyleRanges,
    ).toBe(true);
  });
  it('detects clearing direct XF independently for older engine bundles', () => {
    expect(detectCapabilities(makeWb(['setCellXfIndex'])).cellFormattingClear).toBe(false);
    expect(detectCapabilities(makeWb(['clearCellXfIndex'])).cellFormattingClear).toBe(true);
  });
  it('requires both conversion directions and sheet defaults for pixel geometry', () => {
    const methods = ['getSheetFormatDefaults', 'columnCharsToPt', 'columnPtToChars'];
    expect(detectCapabilities(makeWb(methods)).geometry).toBe(true);
    for (const omitted of methods) {
      expect(
        detectCapabilities(makeWb(methods.filter((method) => method !== omitted))).geometry,
      ).toBe(false);
    }
    const caps = detectCapabilities(makeWb(['clearRowHeight', 'resolveColor']));
    expect(caps.geometry).toBe(false);
    expect(caps.rowHeightClear).toBe(true);
    expect(caps.colorResolution).toBe(true);
  });

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

  it('comment enumeration is independent from point comment read/write', () => {
    const pointOnly = makeWb(['getComment', 'setComment']);
    expect(detectCapabilities(pointOnly).comments).toBe(true);
    expect(detectCapabilities(pointOnly).commentsEnumerable).toBe(false);

    const enumerable = makeWb(['getComment', 'setComment', 'getComments']);
    expect(detectCapabilities(enumerable).comments).toBe(true);
    expect(detectCapabilities(enumerable).commentsEnumerable).toBe(true);
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
    expect(caps.sheetView).toBe(false);
    expect(caps.sheetZoom).toBe(false);
    expect(caps.sheetTabHidden).toBe(false);
    expect(caps.sheetViewFlags).toBe(false);
    expect(caps.outlines).toBe(false);
    expect(caps.comments).toBe(false);
    expect(caps.commentsEnumerable).toBe(false);
    expect(caps.definedNameMutate).toBe(false);
    expect(caps.partialRecalc).toBe(false);
    expect(caps.iterativeProgress).toBe(false);
    expect(caps.spillInfo).toBe(false);
    expect(caps.traceArrows).toBe(false);
    expect(caps.functionMetadata).toBe(false);
    expect(caps.functionLocale).toBe(false);
    expect(caps.formulaLocale).toBe(false);
    expect(caps.localeFacts).toBe(false);
    expect(caps.arrayFormulaEvaluation).toBe(false);
    expect(caps.calcMode).toBe(false);
    expect(caps.spreadsheetProfile).toBe(false);
    expect(caps.sheetProtectionRoundtrip).toBe(false);
    expect(caps.externalLinks).toBe(false);
    expect(caps.lambdaText).toBe(false);
    expect(caps.cellStyles).toBe(false);
    expect(caps.conditionalFormatMutate).toBe(false);
    expect(caps.pivotTables).toBe(false);
    expect(caps.pivotTableMutate).toBe(false);
  });

  it('keeps full formula localization separate from function aliases', () => {
    const aliasesOnly = makeWb(['localizeFunctionName', 'canonicalizeFunctionName']);
    expect(detectCapabilities(aliasesOnly).functionLocale).toBe(true);
    expect(detectCapabilities(aliasesOnly).formulaLocale).toBe(false);
    expect(detectCapabilities(aliasesOnly).localeFacts).toBe(false);

    const formulaOnly = makeWb(['localizeFormula', 'canonicalizeFormula']);
    expect(detectCapabilities(formulaOnly).formulaLocale).toBe(true);
    expect(detectCapabilities(formulaOnly).localeFacts).toBe(false);

    const full = makeWb([
      'localizeFunctionName',
      'canonicalizeFunctionName',
      'localizeFormula',
      'canonicalizeFormula',
      'localeFacts',
    ]);
    expect(detectCapabilities(full).formulaLocale).toBe(true);
    expect(detectCapabilities(full).localeFacts).toBe(true);
  });

  it('sheet view readback is independent from zoom writeback', () => {
    const readOnly = makeWb(['getSheetView']);
    const caps = detectCapabilities(readOnly);
    expect(caps.sheetView).toBe(true);
    expect(caps.sheetZoom).toBe(false);
    expect(caps.freeze).toBe(false);
    expect(caps.sheetTabHidden).toBe(false);
  });

  it('sheet view display flags require every paired setter', () => {
    const partial = makeWb(['setSheetShowGridLines', 'setSheetShowRowColHeaders']);
    expect(detectCapabilities(partial).sheetViewFlags).toBe(false);

    const full = makeWb([
      'setSheetShowGridLines',
      'setSheetShowRowColHeaders',
      'setSheetShowZeros',
      'setSheetRightToLeft',
    ]);
    expect(detectCapabilities(full).sheetViewFlags).toBe(true);
  });

  it('arrayFormulaEvaluation is independent from scalar formula evaluation', () => {
    const scalarOnly = makeWb(['evaluateFormulaText']);
    expect(detectCapabilities(scalarOnly).formulaTextEvaluation).toBe(true);
    expect(detectCapabilities(scalarOnly).arrayFormulaEvaluation).toBe(false);

    const both = makeWb(['evaluateFormulaText', 'evaluateFormulaArray']);
    expect(detectCapabilities(both).arrayFormulaEvaluation).toBe(true);
  });

  it('traceArrows requires both precedents and dependents', () => {
    const onlyPrec = makeWb(['precedents']);
    expect(detectCapabilities(onlyPrec).traceArrows).toBe(false);
    const both = makeWb(['precedents', 'dependents']);
    expect(detectCapabilities(both).traceArrows).toBe(true);
  });

  it('cellStyles requires the full named-style surface', () => {
    const partial = makeWb(['cellStyleCount', 'getCellStyle']);
    expect(detectCapabilities(partial).cellStyles).toBe(false);
    const full = makeWb(['cellStyleCount', 'cellStyleXfCount', 'getCellStyle', 'getCellStyleXf']);
    expect(detectCapabilities(full).cellStyles).toBe(true);
  });

  it('conditionalFormatMutate requires read + write + remove + clear', () => {
    const readOnly = makeWb(['getConditionalFormats']);
    expect(detectCapabilities(readOnly).conditionalFormatMutate).toBe(false);
    const full = makeWb([
      'getConditionalFormats',
      'addConditionalFormat',
      'removeConditionalFormatAt',
      'clearConditionalFormats',
    ]);
    expect(detectCapabilities(full).conditionalFormatMutate).toBe(true);
  });

  it('pivotTables requires count + layout projection', () => {
    const countOnly = makeWb(['pivotCount']);
    expect(detectCapabilities(countOnly).pivotTables).toBe(false);
    const full = makeWb(['pivotCount', 'pivotLayout']);
    expect(detectCapabilities(full).pivotTables).toBe(true);
  });

  it('pivotDefinitionRead requires every definition read-back', () => {
    const keys = [
      'pivotCacheId',
      'pivotName',
      'pivotGetGrandTotals',
      'pivotRowFieldOrder',
      'pivotColFieldOrder',
      'pivotFieldAt',
      'pivotFieldItemCount',
      'pivotFieldItemAt',
      'pivotDataFieldAt',
    ];
    expect(detectCapabilities(makeWb(keys.slice(1))).pivotDefinitionRead).toBe(false);
    expect(detectCapabilities(makeWb(keys)).pivotDefinitionRead).toBe(true);
  });

  it('pivotTableMutate requires the low-level cache, table, field, data-field, and filter surface', () => {
    const partial = makeWb(['pivotCacheCount', 'pivotCreate']);
    expect(detectCapabilities(partial).pivotTableMutate).toBe(false);
    const full = makeWb([
      'pivotCacheCount',
      'pivotCacheIdAt',
      'pivotCacheCreate',
      'pivotCacheRemove',
      'pivotCacheFieldCount',
      'pivotCacheFieldName',
      'pivotCacheFieldAdd',
      'pivotCacheFieldClear',
      'pivotCacheRecordCount',
      'pivotCacheRecordAdd',
      'pivotCacheRecordClear',
      'pivotCreate',
      'pivotRemove',
      'pivotSetName',
      'pivotSetAnchor',
      'pivotSetGrandTotals',
      'pivotFieldCount',
      'pivotFieldAdd',
      'pivotFieldClear',
      'pivotFieldSetAxis',
      'pivotFieldSetSort',
      'pivotFieldSetSubtotalTop',
      'pivotFieldAddItem',
      'pivotFieldClearItems',
      'pivotFieldSetItemVisible',
      'pivotFieldAddSubtotalFn',
      'pivotFieldClearSubtotalFns',
      'pivotFieldSetDateGroup',
      'pivotFieldClearDateGroup',
      'pivotFieldSetNumberFormat',
      'pivotSetRowFieldOrder',
      'pivotSetColFieldOrder',
      'pivotDataFieldCount',
      'pivotDataFieldAdd',
      'pivotDataFieldClear',
      'pivotDataFieldSet',
      'pivotFilterCount',
      'pivotFilterAt',
      'pivotFilterAdd',
      'pivotFilterClear',
      'pivotFilterRemoveAt',
    ]);
    expect(detectCapabilities(full).pivotTableMutate).toBe(true);
    const withoutReadback = { ...full } as unknown as Record<string, unknown>;
    delete withoutReadback.pivotFilterAt;
    expect(detectCapabilities(withoutReadback as unknown as Workbook).pivotTableMutate).toBe(false);
  });

  it('spreadsheetProfile requires get + set profile ids', () => {
    const getOnly = makeWb(['excelProfileId']);
    expect(detectCapabilities(getOnly).spreadsheetProfile).toBe(false);
    const full = makeWb(['excelProfileId', 'setExcelProfileId']);
    expect(detectCapabilities(full).spreadsheetProfile).toBe(true);
  });
});
