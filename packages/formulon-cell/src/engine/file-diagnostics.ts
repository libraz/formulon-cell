import type { ReadDiagnosticsResult, SaveDiagnosticsResult } from '@libraz/formulon';

/** Loss counters the engine captured while reading a file. */
export interface LoadDiagnostics {
  undecodedFormulaCount: number;
  undecodedDefinedNameCount: number;
  undecodedPartCount: number;
  skippedFeatureCount: number;
  unknownContentTypeCount: number;
}

/** Loss counters the engine raised while writing a file. */
export interface SaveDiagnostics {
  downgradedFormulaCount: number;
  deferredFeatureCount: number;
  droppedPartCount: number;
  droppedRelationshipCount: number;
  renumberedPartCount: number;
}

export type FileLossPhase = 'load' | 'save';

export type FileLossId =
  | 'undecodedFormula'
  | 'undecodedDefinedName'
  | 'undecodedPart'
  | 'skippedFeature'
  | 'unknownContentType'
  | 'downgradedFormula'
  | 'deferredFeature'
  | 'droppedPart'
  | 'droppedRelationship'
  | 'renumberedPart';

/** One concrete issue: what the engine lost or deferred and how many times. */
export interface FileLossEntry {
  id: FileLossId;
  phase: FileLossPhase;
  count: number;
}

export const emptyLoadDiagnostics = (): LoadDiagnostics => ({
  undecodedFormulaCount: 0,
  undecodedDefinedNameCount: 0,
  undecodedPartCount: 0,
  skippedFeatureCount: 0,
  unknownContentTypeCount: 0,
});

export const emptySaveDiagnostics = (): SaveDiagnostics => ({
  downgradedFormulaCount: 0,
  deferredFeatureCount: 0,
  droppedPartCount: 0,
  droppedRelationshipCount: 0,
  renumberedPartCount: 0,
});

const count = (n: number | undefined): number => (typeof n === 'number' && n > 0 ? n : 0);

export const loadDiagnosticsFromEngine = (r: Partial<ReadDiagnosticsResult>): LoadDiagnostics => ({
  undecodedFormulaCount: count(r.undecodedFormulaCount),
  undecodedDefinedNameCount: count(r.undecodedDefinedNameCount),
  undecodedPartCount: count(r.undecodedPartCount),
  skippedFeatureCount: count(r.skippedFeatureCount),
  unknownContentTypeCount: count(r.unknownContentTypeCount),
});

export const saveDiagnosticsFromEngine = (r: Partial<SaveDiagnosticsResult>): SaveDiagnostics => ({
  downgradedFormulaCount: count(r.downgradedFormulaCount),
  deferredFeatureCount: count(r.deferredFeatureCount),
  droppedPartCount: count(r.droppedPartCount),
  droppedRelationshipCount: count(r.droppedRelationshipCount),
  renumberedPartCount: count(r.renumberedPartCount),
});

/** Non-zero counters as rows, load first. Counters are reported one per row and
 *  never summed: dropped parts and dropped relationships overlap by design. */
export function fileLossEntries(
  load: LoadDiagnostics,
  save: SaveDiagnostics = emptySaveDiagnostics(),
): FileLossEntry[] {
  const rows: FileLossEntry[] = [
    { id: 'undecodedFormula', phase: 'load', count: load.undecodedFormulaCount },
    { id: 'undecodedDefinedName', phase: 'load', count: load.undecodedDefinedNameCount },
    { id: 'undecodedPart', phase: 'load', count: load.undecodedPartCount },
    { id: 'skippedFeature', phase: 'load', count: load.skippedFeatureCount },
    { id: 'unknownContentType', phase: 'load', count: load.unknownContentTypeCount },
    { id: 'downgradedFormula', phase: 'save', count: save.downgradedFormulaCount },
    { id: 'deferredFeature', phase: 'save', count: save.deferredFeatureCount },
    { id: 'droppedPart', phase: 'save', count: save.droppedPartCount },
    { id: 'droppedRelationship', phase: 'save', count: save.droppedRelationshipCount },
    { id: 'renumberedPart', phase: 'save', count: save.renumberedPartCount },
  ];
  return rows.filter((r) => r.count > 0);
}

/** True when the engine reported any loss while reading. */
export const hasLoadLosses = (load: LoadDiagnostics): boolean => fileLossEntries(load).length > 0;
