import type { CellRect, Workbook as EngineWorkbook } from '@libraz/formulon';
import type { FacadeChangeKind, WorkbookHandle } from '../../engine/workbook-handle.js';
import { FORMAT_JOURNAL } from './format.js';
import { LAYOUT_JOURNAL } from './layout.js';
import { OBJECTS_JOURNAL } from './objects.js';
import { RECORDS_JOURNAL } from './records.js';
import { STRUCTURE_JOURNAL } from './structure.js';

/** Writes a captured slice of engine state back. */
export type Restore = () => void;

/** What a codec reads to snapshot the slice one facade write touches. */
export interface JournalScope {
  readonly handle: WorkbookHandle;
  readonly engine: EngineWorkbook;
  /** The sheet the write targets; `null` for a workbook-level write. */
  readonly sheet: number | null;
  readonly range?: CellRect;
  readonly hostId: string;
  readonly name: string;
  readonly args: readonly unknown[];
}

/** Snapshots one slice of engine state and returns the function that restores it. */
export interface JournalCodec {
  readonly id: string;
  /** Identifies the slice, so a slice written twice in a segment is captured once. */
  scopeKey(scope: JournalScope): string;
  /** `null` when the slice cannot be captured faithfully; the batch then resets History. */
  capture(scope: JournalScope): Restore | null;
}

/** The facade member a write comes from. */
export interface JournalMember {
  readonly hostId: string;
  readonly name: string;
  readonly args: readonly unknown[];
}

/** How a write is journaled: not at all, as a History reset, or per the registry. */
export type JournalSpec = 'none' | 'reset' | JournalMember;

/** A resolved journal decision. */
export type JournalChoice = 'none' | 'reset' | JournalCodec;

/** Where a kind-default codec applies. */
export type JournalScopeKind = 'workbook' | 'sheet' | 'range';

/** The codecs one family of writes registers. */
export interface JournalFamily {
  /** Per-member overrides: `members[hostId][memberName]`. */
  readonly members?: Readonly<Record<string, Readonly<Record<string, JournalChoice>>>>;
  /** Kind defaults: `kinds['kind/scope']`. */
  readonly kinds?: Readonly<
    Partial<Record<`${FacadeChangeKind}/${JournalScopeKind}`, JournalChoice>>
  >;
}

const FAMILIES: readonly JournalFamily[] = [
  RECORDS_JOURNAL,
  FORMAT_JOURNAL,
  LAYOUT_JOURNAL,
  OBJECTS_JOURNAL,
  STRUCTURE_JOURNAL,
];

const overrides = new Map<string, JournalChoice>();

/** @internal Registers a member override ahead of the families; returns its removal. */
export const registerJournalOverride = (
  hostId: string,
  name: string,
  choice: JournalChoice,
): (() => void) => {
  const key = `${hostId}/${name}`;
  overrides.set(key, choice);
  return () => {
    overrides.delete(key);
  };
};

/** The journal decision for a write: the member override, else the kind default of its scope. */
export const resolveJournal = (
  kind: FacadeChangeKind | null,
  sheet: number | null,
  range: CellRect | undefined,
  spec: JournalSpec | undefined,
): JournalChoice => {
  if (spec === 'none' || spec === 'reset') return spec;
  if (spec) {
    const override = overrides.get(`${spec.hostId}/${spec.name}`);
    if (override) return override;
    for (const family of FAMILIES) {
      const choice = family.members?.[spec.hostId]?.[spec.name];
      if (choice) return choice;
    }
  }
  // A cell-side record has no change kind: only a member override journals it.
  if (kind === null) return 'none';
  const scopes: JournalScopeKind[] =
    sheet === null ? ['workbook'] : range ? ['range', 'sheet'] : ['sheet'];
  for (const scope of scopes) {
    for (const family of FAMILIES) {
      const choice = family.kinds?.[`${kind}/${scope}`];
      if (choice) return choice;
    }
  }
  return 'none';
};
