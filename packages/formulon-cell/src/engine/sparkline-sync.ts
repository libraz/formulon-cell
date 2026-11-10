import type { Color, ColorInput, SheetRect } from '@libraz/formulon';
import { ColorKind, SparklineType } from '@libraz/formulon';
import { newSparklineGroupId, splitSourceQualifier } from '../commands/sparkline.js';
import { DEFAULT_NEGATIVE_COLOR, DEFAULT_SPARKLINE_COLOR } from '../render/painters/controls.js';
import type { Sparkline, SparklineKind, SpreadsheetStore, State } from '../store/store.js';
import { addrKey, formatA1Cell, formatA1Range, parseAddrKey, quoteSheetName } from './address.js';
import { argbFromCssColor, cssColorFromArgb } from './theme-color.js';
import type { WorkbookHandle } from './workbook-handle.js';
import type {
  EngineSparklineGroupPatch,
  EngineSparklineInput,
} from './workbook-handle-drawing-objects.js';

const TYPE_OF_KIND: Record<SparklineKind, SparklineType> = {
  line: SparklineType.Line,
  column: SparklineType.Column,
  'win-loss': SparklineType.Stacked,
};

const KIND_OF_TYPE: Record<SparklineType, SparklineKind> = {
  [SparklineType.Line]: 'line',
  [SparklineType.Column]: 'column',
  [SparklineType.Stacked]: 'win-loss',
};

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

const colorSpecOf = (css: string): ColorInput =>
  cssColorFromArgb(argbFromCssColor(css) ?? 0xff000000);

/** CSS colour of an engine colour; undefined when it is absent. */
const cssOfSpec = (wb: WorkbookHandle, spec: Color | null): string | undefined => {
  if (spec === null || spec.kind === ColorKind.None) return undefined;
  const argb = wb.resolveColorArgb(spec) ?? (spec.kind === ColorKind.Rgb ? spec.argb : null);
  return argb === null ? undefined : cssColorFromArgb(argb);
};

/** `css` as an engine colour, keeping `base` while it still shows as `css`; with
 *  no `css`, `base` or else `fallback`. */
const colorFor = (
  wb: WorkbookHandle,
  css: string | undefined,
  base: Color | null | undefined,
  fallback: string,
): ColorInput => {
  if (css === undefined) return base ?? colorSpecOf(fallback);
  return base && cssOfSpec(wb, base) === css ? base : colorSpecOf(css);
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
    for (const index of sheets) {
      for (const group of wb.listSparklineGroups(index)) {
        const { sparklines, ...settings } = group;
        const id = newSparklineGroupId();
        const color = cssOfSpec(wb, group.colorSeries);
        const negativeColor = cssOfSpec(wb, group.colorNegative);
        for (const member of sparklines) {
          if (!member.source) continue;
          const { row, col } = member.location;
          loaded.set(addrKey({ sheet: index, row, col }), {
            kind: KIND_OF_TYPE[group.type] ?? 'line',
            source: storeSource(wb, member.source, index),
            ...(color ? { color } : {}),
            showNegative: group.negative,
            ...(negativeColor ? { negativeColor } : {}),
            group: id,
            nativeGroup: settings,
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
      return { ...state, sparkline: { sparklines } };
    });
  };
  if (typeof wb.withEngineSyncMuted === 'function') wb.withEngineSyncMuted(hydrate);
  else hydrate();
}

interface Member {
  sheet: number;
  row: number;
  col: number;
  spec: Sparkline;
}

/** Sparklines of one sheet by group: the sparklines created together, or a
 *  cell on its own. */
const groupsOf = (map: ReadonlyMap<string, Sparkline>): Map<string, Member[]> => {
  const groups = new Map<string, Member[]>();
  for (const [key, spec] of map) {
    const at = parseAddrKey(key);
    if (!at) continue;
    const id = `${at.sheet}|${spec.group ?? `cell:${key}`}`;
    const members = groups.get(id);
    if (members) members.push({ ...at, spec });
    else groups.set(id, [{ ...at, spec }]);
  }
  for (const members of groups.values()) {
    members.sort((a, b) => a.row - b.row || a.col - b.col);
  }
  return groups;
};

const locationOf = (member: Member): string => formatA1Cell(member.row, member.col);

/** Whether the members in `a` and `b` are the same cells with the same sources. */
const sameMembers = (a: readonly Member[], b: readonly Member[]): boolean =>
  a.length === b.length &&
  a.every(
    (member, i) =>
      member.row === b[i]?.row &&
      member.col === b[i]?.col &&
      member.spec.source === b[i]?.spec.source,
  );

const styleOf = (members: readonly Member[]): Sparkline => members[0]?.spec as Sparkline;

const sameStyle = (a: Sparkline, b: Sparkline): boolean =>
  a.kind === b.kind &&
  a.color === b.color &&
  !!a.showNegative === !!b.showNegative &&
  a.negativeColor === b.negativeColor;

/** Overlay sparklines mirrored into the engine's sparkline groups. A group is
 *  the sparklines that share a `group` id; the engine position of a group is
 *  found again by its cells for every write, as a removal shifts it. */
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
    const previous = this.last?.sparklines ?? new Map<string, Sparkline>();
    this.last = slice;
    const wb = this.handle;
    if (!wb.capabilities.sparklines) return;
    const before = groupsOf(previous);
    const after = groupsOf(slice.sparklines);
    const rejected: string[] = [];
    const refused = new Set<number>();
    // Style a group's members take on, by cell.
    const restyled = new Map<string, Partial<Sparkline>>();
    const unify = (members: readonly Member[], style: Sparkline): void => {
      for (const member of members) {
        if (sameStyle(member.spec, style)) continue;
        restyled.set(addrKey(member), {
          kind: style.kind,
          color: style.color,
          showNegative: style.showNegative,
          negativeColor: style.negativeColor,
        });
      }
    };
    for (const [id, members] of before) {
      if (after.has(id)) continue;
      const index = this.indexOf(members);
      if (index >= 0) wb.removeSparklineGroup(members[0]?.sheet as number, index);
    }
    const additions: Member[][] = [];
    for (const [id, members] of after) {
      const old = before.get(id);
      const index = old ? this.indexOf(old) : -1;
      if (!old || index < 0) {
        additions.push(members);
        continue;
      }
      const sheet = members[0]?.sheet as number;
      if (!sameMembers(old, members) && !this.setMembers(sheet, index, members)) {
        refused.add(sheet);
        continue;
      }
      // A style is the group's: the member that changed sets it for the rest.
      const was = styleOf(old);
      const changed = members.find((member) => {
        const prior = previous.get(addrKey(member));
        return prior !== undefined && !sameStyle(prior, member.spec);
      });
      const now = changed?.spec ?? was;
      if (!sameStyle(was, now)) this.restyle(sheet, index, was, now);
      unify(members, now);
    }
    for (const members of additions) {
      if (this.add(members)) unify(members, styleOf(members));
      else rejected.push(...members.map((m) => addrKey(m)));
    }
    if (rejected.length > 0 || restyled.size > 0) {
      // Drop what the engine refused and show every member in its group's style.
      store.setState((state) => {
        const sparklines = new Map(state.sparkline.sparklines);
        for (const key of rejected) sparklines.delete(key);
        for (const [key, patch] of restyled) {
          const spec = sparklines.get(key);
          if (spec) sparklines.set(key, { ...spec, ...patch });
        }
        const next = { ...state, sparkline: { sparklines } };
        this.last = next.sparkline;
        return next;
      });
    }
    // The engine kept the group as it was; read it back instead of guessing.
    for (const sheet of refused) hydrateSparklinesFromEngine(wb, store, sheet);
  }

  /** Engine index of the group holding any of `members` on their sheet; -1 if none. */
  private indexOf(members: readonly Member[]): number {
    const sheet = members[0]?.sheet;
    if (sheet === undefined) return -1;
    const cells = new Set(members.map(locationOf));
    return this.handle
      .listSparklineGroups(sheet)
      .findIndex((group) =>
        group.sparklines.some((sparkline) =>
          cells.has(formatA1Cell(sparkline.location.row, sparkline.location.col)),
        ),
      );
  }

  private engineMembers(members: readonly Member[]): EngineSparklineInput[] {
    const hostSheet = this.handle.sheetName(members[0]?.sheet as number);
    return members.map((member) => ({
      source: engineFormula(member.spec.source, hostSheet),
      location: locationOf(member),
    }));
  }

  private setMembers(sheet: number, index: number, members: readonly Member[]): boolean {
    return this.handle.setSparklineGroupSparklines(sheet, index, this.engineMembers(members));
  }

  /** Apply the style change from `from` to `to` on group `index`; the group's other settings stay. */
  private restyle(sheet: number, index: number, from: Sparkline, to: Sparkline): void {
    this.handle.updateSparklineGroup(sheet, index, this.styleChange(from, to));
  }

  /** The group fields that `to` changes from `from`. */
  private styleChange(from: Sparkline, to: Sparkline): EngineSparklineGroupPatch {
    const change: {
      -readonly [K in keyof EngineSparklineGroupPatch]: EngineSparklineGroupPatch[K];
    } = {};
    if (from.kind !== to.kind) change.type = TYPE_OF_KIND[to.kind];
    if (from.color !== to.color) {
      change.colorSeries = colorSpecOf(to.color ?? DEFAULT_SPARKLINE_COLOR);
    }
    if (!!from.showNegative !== !!to.showNegative) change.negative = !!to.showNegative;
    if (from.negativeColor !== to.negativeColor) {
      change.colorNegative = colorSpecOf(to.negativeColor ?? DEFAULT_NEGATIVE_COLOR);
    }
    return change;
  }

  private add(members: readonly Member[]): boolean {
    const wb = this.handle;
    const style = styleOf(members);
    // The id is the read-back group's own; a new group gets one of its own.
    const { id: _id, ...base }: Partial<NonNullable<Sparkline['nativeGroup']>> =
      style.nativeGroup ?? {};
    return wb.addSparklineGroup(members[0]?.sheet as number, {
      ...base,
      type: TYPE_OF_KIND[style.kind],
      colorSeries: colorFor(wb, style.color, base.colorSeries, DEFAULT_SPARKLINE_COLOR),
      colorNegative: colorFor(wb, style.negativeColor, base.colorNegative, DEFAULT_NEGATIVE_COLOR),
      negative: !!style.showNegative,
      sparklines: this.engineMembers(members),
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
