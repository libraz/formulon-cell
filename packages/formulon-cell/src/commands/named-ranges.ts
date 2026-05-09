import type { WorkbookHandle } from '../engine/workbook-handle.js';

export interface DefinedNameEntry {
  name: string;
  formula: string;
}

export type DefinedNameMutationResult =
  | { ok: true; entry: DefinedNameEntry }
  | { ok: false; reason: 'empty-name' | 'empty-formula' | 'unsupported' | 'engine-failed' };

export type DefinedNameDeleteResult =
  | { ok: true; entry: DefinedNameEntry }
  | { ok: false; reason: 'empty-name' | 'unsupported' | 'engine-failed' };

/** Snapshot workbook-scoped defined names into a stable array. */
export function listDefinedNames(wb: WorkbookHandle): DefinedNameEntry[] {
  return [...wb.definedNames()];
}

/** Add or replace a workbook-scoped defined name. */
export function upsertDefinedName(
  wb: WorkbookHandle,
  name: string,
  formula: string,
): DefinedNameMutationResult {
  const trimmedName = name.trim();
  const trimmedFormula = formula.trim();
  if (!trimmedName) return { ok: false, reason: 'empty-name' };
  if (!trimmedFormula) return { ok: false, reason: 'empty-formula' };
  if (!wb.capabilities.definedNameMutate) return { ok: false, reason: 'unsupported' };
  if (!wb.setDefinedNameEntry(trimmedName, trimmedFormula)) {
    return { ok: false, reason: 'engine-failed' };
  }
  return { ok: true, entry: { name: trimmedName, formula: trimmedFormula } };
}

/** Remove a workbook-scoped defined name. */
export function deleteDefinedName(wb: WorkbookHandle, name: string): DefinedNameDeleteResult {
  const trimmedName = name.trim();
  if (!trimmedName) return { ok: false, reason: 'empty-name' };
  if (!wb.capabilities.definedNameMutate) return { ok: false, reason: 'unsupported' };
  if (!wb.setDefinedNameEntry(trimmedName, '')) return { ok: false, reason: 'engine-failed' };
  return { ok: true, entry: { name: trimmedName, formula: '' } };
}
