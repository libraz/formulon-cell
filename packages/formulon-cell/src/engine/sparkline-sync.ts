import type { SheetRect } from '@libraz/formulon';
import { newSparklineGroupId, splitSourceQualifier } from '../commands/sparkline.js';
import { sameSparklineSettings, withSparklineMembers } from '../store/sparkline-slice.js';
import type { Sparkline, SparklineGroupSettings, SpreadsheetStore, State } from '../store/store.js';
import { addrKey, formatA1Cell, formatA1Range, parseAddrKey, quoteSheetName } from './address.js';
import type { WorkbookHandle } from './workbook-handle.js';
import type { EngineSparklineInput } from './workbook-handle-drawing-objects.js';

/** Sheet-qualified source the engine stores. */
const engineFormula = (source: string, hostSheet: string): string =>
  splitSourceQualifier(source).qualifier ? source : `${quoteSheetName(hostSheet)}!${source}`;

/** Source as the store keeps it: bare when it is on the host sheet. */
const storeSource = (wb: WorkbookHandle, source: SheetRect, hostSheet: number): string => {
  const { rect } = source;
  const range = formatA1Range({
    r0: rect.firstRow,
    c0: rect.firstCol,
    r1: rect.lastRow,
    c1: rect.lastCol,
  });
  return source.sheet === hostSheet
    ? range
    : `${quoteSheetName(wb.sheetName(source.sheet))}!${range}`;
};

/** Hydrate the sparkline groups of `sheet` (every sheet when omitted) into the
 *  store, replacing what the store held for them. */
export function hydrateSparklinesFromEngine(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
  sheet?: number,
): void {
  if (!wb.capabilities.sparklines) return;
  const sheets = sheet === undefined ? [...Array(wb.sheetCount).keys()] : [sheet];
  const hydrate = (): void => {
    const loaded = new Map<string, Sparkline>();
    const loadedGroups = new Map<string, SparklineGroupSettings>();
    for (const index of sheets) {
      for (const group of wb.listSparklineGroups(index)) {
        const { sparklines, id: _engineId, ...settings } = group;
        const id = newSparklineGroupId();
        loadedGroups.set(id, settings);
        for (const member of sparklines) {
          if (!member.source) continue;
          const { row, col } = member.location;
          loaded.set(addrKey({ sheet: index, row, col }), {
            source: storeSource(wb, member.source, index),
            group: id,
          });
        }
      }
    }
    store.setState((state) => {
      const sparklines = new Map(
        [...state.sparkline.sparklines].filter(([key]) => {
          const at = parseAddrKey(key);
          return !at || !sheets.includes(at.sheet);
        }),
      );
      for (const [key, spec] of loaded) sparklines.set(key, spec);
      const groups = new Map([...state.sparkline.groups, ...loadedGroups]);
      return { ...state, sparkline: withSparklineMembers({ sparklines, groups }, sparklines) };
    });
  };
  wb.withEngineSyncMuted(hydrate);
}

interface Member {
  sheet: number;
  row: number;
  col: number;
  source: string;
}

interface Group {
  id: string;
  sheet: number;
  members: Member[];
  settings: SparklineGroupSettings;
}

/** The slice's groups with their members in row-major order; a member whose group has no
 *  settings is left out. */
const groupsOf = (slice: State['sparkline']): Map<string, Group> => {
  const groups = new Map<string, Group>();
  for (const [key, spec] of slice.sparklines) {
    const at = parseAddrKey(key);
    const settings = slice.groups.get(spec.group);
    if (!at || !settings) continue;
    const member = { ...at, source: spec.source };
    const group = groups.get(spec.group);
    if (group) group.members.push(member);
    else groups.set(spec.group, { id: spec.group, sheet: at.sheet, members: [member], settings });
  }
  for (const group of groups.values()) {
    group.members.sort((a, b) => a.row - b.row || a.col - b.col);
  }
  return groups;
};

const locationOf = (member: Member): string => formatA1Cell(member.row, member.col);

/** Whether the members in `a` and `b` are the same cells with the same sources. */
const sameMembers = (a: readonly Member[], b: readonly Member[]): boolean =>
  a.length === b.length &&
  a.every(
    (member, i) =>
      member.row === b[i]?.row && member.col === b[i]?.col && member.source === b[i]?.source,
  );

/** Overlay sparklines mirrored into the engine's sparkline groups. A group is a store group
 *  id with its settings; the engine group of one is found again by its cells for every
 *  write, as a removal shifts positions. */
export class SparklineMirror {
  constructor(private readonly handle: WorkbookHandle) {}

  private last: State['sparkline'] | null = null;

  /** Adopt `state` as what the engine already carries. */
  rebase(state: State | null): void {
    this.last = state?.sparkline ?? null;
  }

  sync(store: SpreadsheetStore): void {
    const slice = store.getState().sparkline;
    if (slice === this.last) return;
    const previous = this.last;
    this.last = slice;
    const wb = this.handle;
    if (!wb.capabilities.sparklines) return;
    const before = previous ? groupsOf(previous) : new Map<string, Group>();
    const after = groupsOf(slice);
    const rejected: string[] = [];
    const refused = new Set<number>();
    // Engine ids are read up front: a write shifts positions, an id stays.
    const engineIds = new Map<string, string>();
    for (const [id, group] of before) {
      const engineId = this.engineIdOf(group);
      if (engineId !== undefined) engineIds.set(id, engineId);
    }
    for (const [id, group] of before) {
      const engineId = engineIds.get(id);
      if (engineId === undefined || after.has(id)) continue;
      const index = this.indexOfId(group.sheet, engineId);
      if (index >= 0) wb.removeSparklineGroup(group.sheet, index);
    }
    const additions: Group[] = [];
    const updates: { group: Group; old: Group; engineId: string }[] = [];
    for (const [id, group] of after) {
      const old = before.get(id);
      const engineId = engineIds.get(id);
      if (!old || engineId === undefined) additions.push(group);
      else updates.push({ group, old, engineId });
    }
    // Shrinking groups first, so a cell never sits in two engine groups at once.
    updates.sort(
      (a, b) =>
        a.group.members.length -
        a.old.members.length -
        (b.group.members.length - b.old.members.length),
    );
    for (const { group, old, engineId } of updates) {
      const index = this.indexOfId(group.sheet, engineId);
      if (index < 0) continue;
      if (!sameMembers(old.members, group.members) && !this.setMembers(group, index)) {
        refused.add(group.sheet);
        continue;
      }
      if (!sameSparklineSettings(old.settings, group.settings)) {
        if (!wb.updateSparklineGroup(group.sheet, index, group.settings)) {
          refused.add(group.sheet);
        }
      }
    }
    for (const group of additions) {
      if (!this.add(group)) rejected.push(...group.members.map((m) => addrKey(m)));
    }
    if (rejected.length > 0) {
      // Drop what the engine refused.
      store.setState((state) => {
        const sparklines = new Map(state.sparkline.sparklines);
        for (const key of rejected) sparklines.delete(key);
        const next = {
          ...state,
          sparkline: withSparklineMembers(state.sparkline, sparklines),
        };
        this.last = next.sparkline;
        return next;
      });
    }
    // The engine kept the group as it was; read it back instead of guessing.
    for (const sheet of refused) hydrateSparklinesFromEngine(wb, store, sheet);
  }

  /** Id of the engine group holding any of `group`'s cells; undefined if none. */
  private engineIdOf(group: Group): string | undefined {
    const cells = new Set(group.members.map(locationOf));
    return this.handle
      .listSparklineGroups(group.sheet)
      .find((candidate) =>
        candidate.sparklines.some((sparkline) =>
          cells.has(formatA1Cell(sparkline.location.row, sparkline.location.col)),
        ),
      )?.id;
  }

  private indexOfId(sheet: number, engineId: string): number {
    return this.handle.listSparklineGroups(sheet).findIndex((group) => group.id === engineId);
  }

  private engineMembers(group: Group): EngineSparklineInput[] {
    const hostSheet = this.handle.sheetName(group.sheet);
    return group.members.map((member) => ({
      source: engineFormula(member.source, hostSheet),
      location: locationOf(member),
    }));
  }

  private setMembers(group: Group, index: number): boolean {
    return this.handle.setSparklineGroupSparklines(group.sheet, index, this.engineMembers(group));
  }

  private add(group: Group): boolean {
    return this.handle.addSparklineGroup(group.sheet, {
      ...group.settings,
      sparklines: this.engineMembers(group),
    });
  }
}

/** Capture every sheet's sparkline groups ahead of a delete that may drop some.
 *  The restorer puts the captured groups back and re-reads the store. */
export function captureSparklineGroups(
  wb: WorkbookHandle,
  store: SpreadsheetStore,
): (() => void) | null {
  if (!wb.capabilities.sparklines) return null;
  const captured = [...Array(wb.sheetCount).keys()].map((sheet) => wb.listSparklineGroups(sheet));
  if (captured.every((groups) => groups.length === 0)) return null;
  return () => {
    captured.forEach((groups, sheet) => {
      const count = wb.listSparklineGroups(sheet).length;
      for (let index = count - 1; index >= 0; index -= 1) wb.removeSparklineGroup(sheet, index);
      for (const group of groups) wb.addSparklineGroup(sheet, group);
    });
    hydrateSparklinesFromEngine(wb, store);
  };
}
