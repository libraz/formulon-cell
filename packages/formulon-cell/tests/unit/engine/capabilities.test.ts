import type { Formulon } from '@libraz/formulon';
import { describe, expect, it } from 'vitest';
import { detectCapabilities } from '../../../src/engine/capabilities.js';
import { loadFormulon } from '../../../src/engine/loader.js';
import type { EngineCapabilities } from '../../../src/engine/types.js';

// Any module the loader did not hand out as the stub is the pinned engine.
const engineModule = {} as Formulon;

const expectEngineOnly = async (flags: readonly (keyof EngineCapabilities)[]): Promise<void> => {
  const engine = detectCapabilities(engineModule);
  const stub = detectCapabilities(await loadFormulon({ preferStub: true }));
  for (const flag of flags) {
    expect(engine[flag], flag).toBe(true);
    expect(stub[flag], flag).toBe(false);
  }
};

describe('detectCapabilities', () => {
  it('reports the paired native range-style operations together', async () => {
    await expectEngineOnly(['rangeStylePatch', 'rangeStyleSnapshot', 'rowStyleRanges']);
  });

  it('reports shapes, sparklines and slicers', async () => {
    await expectEngineOnly(['shapes', 'slicers', 'sparklines']);
  });

  it('reports clearing direct XF', async () => {
    await expectEngineOnly(['cellFormattingClear']);
  });

  it('reports pixel geometry, row-height clearing and colour resolution', async () => {
    await expectEngineOnly(['colorResolution', 'geometry', 'rowHeightClear']);
  });

  it('cellFormatting is reported for the full XF resolver/dedup surface', async () => {
    await expectEngineOnly(['cellFormatting']);
  });

  it('dataValidation is reported for read + write', async () => {
    await expectEngineOnly(['dataValidation']);
  });

  it('hyperlinks is reported for read + write', async () => {
    await expectEngineOnly(['hyperlinks']);
  });

  it('comment enumeration and point comment read/write are both reported', async () => {
    await expectEngineOnly(['comments', 'commentsEnumerable']);
  });

  it('keeps full formula localization and function aliases as separate flags', async () => {
    await expectEngineOnly(['formulaLocale', 'functionLocale', 'localeFacts']);
  });

  it('sheet view readback and zoom writeback are both reported', async () => {
    await expectEngineOnly(['freeze', 'sheetTabHidden', 'sheetView', 'sheetZoom']);
  });

  it('sheet view display flags are reported', async () => {
    await expectEngineOnly(['sheetViewFlags']);
  });

  it('arrayFormulaEvaluation and scalar formula evaluation are both reported', async () => {
    await expectEngineOnly(['arrayFormulaEvaluation', 'formulaTextEvaluation']);
  });

  it('traceArrows is reported', async () => {
    await expectEngineOnly(['traceArrows']);
  });

  it('cellStyles is reported for the named-style surface', async () => {
    await expectEngineOnly(['cellStyles']);
  });

  it('conditionalFormatMutate is reported', async () => {
    await expectEngineOnly(['conditionalFormatMutate']);
  });

  it('pivotTables is reported', async () => {
    await expectEngineOnly(['pivotTables']);
  });

  it('pivotDefinitionRead is reported', async () => {
    await expectEngineOnly(['pivotDefinitionRead']);
  });

  it('pivotTableMutate is reported', async () => {
    await expectEngineOnly(['pivotTableMutate']);
  });

  it('calc settings and data tables are reported', async () => {
    await expectEngineOnly(['calcSettings', 'dataTables']);
  });

  it('spreadsheetProfile is reported', async () => {
    await expectEngineOnly(['spreadsheetProfile']);
  });

  it('every engine-only flag is true for the engine and false for the stub', async () => {
    await expectEngineOnly([
      'merges',
      'cellFormatting',
      'dataValidation',
      'hyperlinks',
      'sheetMutate',
      'insertDeleteRowsCols',
      'hiddenRowsCols',
      'colRowSize',
      'freeze',
      'sheetView',
      'sheetZoom',
      'sheetTabHidden',
      'sheetViewFlags',
      'outlines',
      'comments',
      'commentsEnumerable',
      'definedNameMutate',
      'partialRecalc',
      'iterativeProgress',
      'spillInfo',
      'traceArrows',
      'functionMetadata',
      'functionLocale',
      'formulaLocale',
      'localeFacts',
      'arrayFormulaEvaluation',
      'calcMode',
      'spreadsheetProfile',
      'sheetProtectionRoundtrip',
      'externalLinks',
      'lambdaText',
      'cellStyles',
      'conditionalFormatMutate',
      'pivotTables',
      'pivotTableMutate',
    ]);
  });
});
