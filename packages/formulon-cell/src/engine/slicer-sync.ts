import { AnchorEditAs, AnchorKind, FilterKind } from '@libraz/formulon';
import type { SlicerSpec, SpreadsheetStore, State } from '../store/store.js';
import { sessionPosition, twoCellAnchorForPixels } from './drawing-geometry.js';
import type { EngineDrawingAnchor } from './types.js';
import type { WorkbookHandle } from './workbook-handle.js';
import type { EngineSlicer, EngineSlicerInput } from './workbook-handle-drawing-objects.js';

/** Where and how large a slicer panel is when its spec carries no frame, in grid pixels. */
export const DEFAULT_SLICER_FRAME = { x: 16, y: 16, w: 192, h: 240 } as const;

const nameKey = (name: string): string => name.toLocaleLowerCase();

const tableOf = (wb: WorkbookHandle, tableName: string) =>
  wb
    .getTables()
    .find(
      (table) =>
        nameKey(table.name) === nameKey(tableName) ||
        nameKey(table.displayName) === nameKey(tableName),
    ) ?? null;

/** The items a table's AutoFilter keeps for `column`; empty when it does not
 *  filter that column by value. */
const selectedItems = (wb: WorkbookHandle, tableIndex: number, column: number): string[] => {
  const filter = wb.getTableAutoFilter(tableIndex);
  const entry = filter?.columns.find((candidate) => candidate.colId === column);
  if (!entry || entry.kind !== FilterKind.Values || entry.dateGroups.length > 0) return [];
  // A slicer lists only the items the table holds; blanks matter alone.
  return entry.values.length === 0 && entry.filterBlank ? [''] : [...entry.values];
};

/** The anchor of a slicer's frame; null unless it is a two-cell anchor. */
const anchorOf = (slicer: EngineSlicer): EngineDrawingAnchor | null =>
  slicer.anchorKind === AnchorKind.TwoCell
    ? {
        anchorKind: slicer.anchorKind,
        editAs: AnchorEditAs.TwoCell,
        fromRow: slicer.fromRow,
        fromCol: slicer.fromCol,
        fromRowOff: slicer.fromRowOff,
        fromColOff: slicer.fromColOff,
        toRow: slicer.toRow,
        toCol: slicer.toCol,
        toRowOff: slicer.toRowOff,
        toColOff: slicer.toColOff,
        cx: 0,
        cy: 0,
      }
    : null;

const sameAnchor = (a: EngineDrawingAnchor | undefined, b: EngineDrawingAnchor | undefined) =>
  a === b || (!!a && !!b && JSON.stringify(a) === JSON.stringify(b));

const anchorFields = (anchor: EngineDrawingAnchor): EngineSlicerInput => ({
  fromRow: anchor.fromRow,
  fromCol: anchor.fromCol,
  fromRowOff: anchor.fromRowOff,
  fromColOff: anchor.fromColOff,
  toRow: anchor.toRow,
  toCol: anchor.toCol,
  toRowOff: anchor.toRowOff,
  toColOff: anchor.toColOff,
});

const frameOf = (store: SpreadsheetStore, slicer: EngineSlicer) => {
  const anchor = anchorOf(slicer);
  if (!anchor) return { ...DEFAULT_SLICER_FRAME };
  const position = sessionPosition(store, anchor);
  return { x: position.x, y: position.y, w: position.w, h: position.h };
};

const specOf = (
  store: SpreadsheetStore,
  wb: WorkbookHandle,
  sheet: number,
  slicer: EngineSlicer,
) => {
  const table = tableOf(wb, slicer.tableName);
  const column = table ? table.columns.indexOf(slicer.columnName) : -1;
  const spec: SlicerSpec = {
    id: `slicer-native-${encodeURIComponent(slicer.name)}`,
    tableName: table?.name ?? slicer.tableName,
    column: slicer.columnName,
    selected: table && column >= 0 ? selectedItems(wb, table.index, column) : [],
    ...frameOf(store, slicer),
    ...(anchorOf(slicer) ? { nativeAnchor: anchorOf(slicer) as EngineDrawingAnchor } : {}),
    sheet,
    caption: slicer.caption,
    columnCount: slicer.columnCount,
    native: {
      name: slicer.name,
      rowHeight: slicer.rowHeight,
      sortOrder: slicer.sortOrder,
      crossFilter: slicer.crossFilter,
    },
  };
  if (slicer.style) spec.style = slicer.style;
  return spec;
};

/** Hydrate the table slicers of every sheet into the store, replacing the
 *  slicers the store holds for the engine. Slicers without an engine slicer
 *  stay. */
export function hydrateSlicersFromEngine(wb: WorkbookHandle, store: SpreadsheetStore): void {
  if (!wb.capabilities.slicers) return;
  const hydrate = (): void => {
    const loaded: SlicerSpec[] = [];
    for (let sheet = 0; sheet < wb.sheetCount; sheet += 1) {
      for (const slicer of wb.listSlicers(sheet)) loaded.push(specOf(store, wb, sheet, slicer));
    }
    store.setState((state) => ({
      ...state,
      slicers: {
        slicers: [...state.slicers.slicers.filter((spec) => spec.native === undefined), ...loaded],
      },
    }));
  };
  if (typeof wb.withEngineSyncMuted === 'function') wb.withEngineSyncMuted(hydrate);
  else hydrate();
}

/** Capture slicer `objectId` so a row or column delete that drops it can be
 *  undone. The restorer adds it again; null when it is not a table slicer. */
export function captureSlicerRestore(
  wb: WorkbookHandle,
  sheet: number,
  objectId: number,
): (() => boolean) | null {
  const slicer = wb.listSlicers(sheet).find((candidate) => candidate.objectId === objectId);
  const table = slicer ? tableOf(wb, slicer.tableName) : null;
  const column = slicer && table ? table.columns.indexOf(slicer.columnName) : -1;
  if (!slicer || !table || column < 0) return null;
  const input: EngineSlicerInput = {
    name: slicer.name,
    caption: slicer.caption,
    style: slicer.style,
    columnCount: slicer.columnCount,
    rowHeight: slicer.rowHeight,
    fromRow: slicer.fromRow,
    fromCol: slicer.fromCol,
    fromRowOff: slicer.fromRowOff,
    fromColOff: slicer.fromColOff,
    toRow: slicer.toRow,
    toCol: slicer.toCol,
    toRowOff: slicer.toRowOff,
    toColOff: slicer.toColOff,
    sortOrder: slicer.sortOrder,
    crossFilter: slicer.crossFilter,
  };
  return () => wb.addSlicer(sheet, table.index, column, input) !== null;
}

const anchorInput = (slicer: EngineSlicer): EngineSlicerInput => ({
  fromRow: slicer.fromRow,
  fromCol: slicer.fromCol,
  fromRowOff: slicer.fromRowOff,
  fromColOff: slicer.fromColOff,
  toRow: slicer.toRow,
  toCol: slicer.toCol,
  toRowOff: slicer.toRowOff,
  toColOff: slicer.toColOff,
});

/** Capture the anchors of `sheet`'s slicers ahead of a row or column delete,
 *  which collapses the slicers it does not remove. The restorer puts the
 *  survivors back where they were once the delete is undone. */
export function captureSlicerAnchors(wb: WorkbookHandle, sheet: number): () => void {
  if (!wb.capabilities.slicers) return () => undefined;
  const anchors = new Map(wb.listSlicers(sheet).map((slicer) => [nameKey(slicer.name), slicer]));
  return () => {
    wb.listSlicers(sheet).forEach((slicer, index) => {
      const old = anchors.get(nameKey(slicer.name));
      if (old && JSON.stringify(anchorInput(old)) !== JSON.stringify(anchorInput(slicer))) {
        wb.updateSlicer(sheet, index, anchorInput(old));
      }
    });
  };
}

/** Overlay slicers mirrored into the engine's table slicers. Slicers are
 *  matched by overlay id and found in the engine by name; the chip selection
 *  is not written here, it reaches the file through the table's AutoFilter. */
export class SlicerMirror {
  constructor(private readonly handle: WorkbookHandle) {}

  private last: State['slicers'] | null = null;

  /** Adopt `state` as what the engine already carries. */
  rebase(state: State | null): void {
    this.last = state?.slicers ?? null;
  }

  sync(store: SpreadsheetStore): void {
    const slice = store.getState().slicers;
    if (slice === this.last) return;
    const previous = this.last?.slicers ?? [];
    this.last = slice;
    const wb = this.handle;
    if (!wb.capabilities.slicers) return;
    const before = new Map(previous.map((spec) => [spec.id, spec] as const));
    const after = new Set(slice.slicers.map((spec) => spec.id));
    for (const spec of previous) {
      if (after.has(spec.id)) continue;
      const index = this.indexOf(spec);
      if (index >= 0 && spec.sheet !== undefined) wb.removeSlicer(spec.sheet, index);
    }
    // null drops the slicer: the engine refused it.
    const patches = new Map<string, Partial<SlicerSpec> | null>();
    for (const spec of slice.slicers) {
      const old = before.get(spec.id);
      if (old === spec) continue;
      const moved =
        old &&
        (old.sheet !== spec.sheet ||
          old.tableName !== spec.tableName ||
          old.column !== spec.column);
      if (!old || moved) {
        if (old && moved) {
          const index = this.indexOf(old);
          if (index >= 0 && old.sheet !== undefined) wb.removeSlicer(old.sheet, index);
        }
        patches.set(spec.id, this.add(store, spec));
      } else {
        const patch = this.update(store, old, spec);
        if (patch) patches.set(spec.id, patch);
      }
    }
    if (patches.size === 0) return;
    store.setState((state) => {
      const slicers = state.slicers.slicers.flatMap((spec) => {
        const patch = patches.get(spec.id);
        if (patch === undefined) return [spec];
        return patch === null ? [] : [{ ...spec, ...patch }];
      });
      const next = { ...state, slicers: { slicers } };
      this.last = next.slicers;
      return next;
    });
  }

  /** Engine position of `spec` on its sheet; -1 when the engine has no such slicer. */
  private indexOf(spec: SlicerSpec): number {
    if (spec.sheet === undefined || !spec.native) return -1;
    const key = nameKey(spec.native.name);
    return this.handle.listSlicers(spec.sheet).findIndex((slicer) => nameKey(slicer.name) === key);
  }

  /** Add `spec`; returns the fields that tie it to the new slicer, or null. */
  private add(store: SpreadsheetStore, spec: SlicerSpec): Partial<SlicerSpec> | null {
    const wb = this.handle;
    const table = tableOf(wb, spec.tableName);
    const column = table ? table.columns.indexOf(spec.column) : -1;
    if (!table || column < 0) return null;
    const sheet = spec.sheet ?? table.sheetIndex;
    const anchor =
      spec.nativeAnchor ?? twoCellAnchorForPixels(store, this.frameOf(spec), AnchorEditAs.TwoCell);
    const objectId = wb.addSlicer(sheet, table.index, column, {
      ...(spec.native ? { name: spec.native.name } : {}),
      ...(spec.caption !== undefined ? { caption: spec.caption } : {}),
      ...(spec.style ? { style: spec.style } : {}),
      ...(spec.columnCount !== undefined ? { columnCount: spec.columnCount } : {}),
      ...(spec.native
        ? {
            rowHeight: spec.native.rowHeight,
            sortOrder: spec.native.sortOrder,
            crossFilter: spec.native.crossFilter,
          }
        : {}),
      ...anchorFields(anchor),
    });
    if (objectId === null) return null;
    const slicer = wb.listSlicers(sheet).find((candidate) => candidate.objectId === objectId);
    if (!slicer) return null;
    const placed = anchorOf(slicer);
    return {
      sheet,
      ...(placed ? { nativeAnchor: placed } : {}),
      caption: slicer.caption,
      columnCount: slicer.columnCount,
      ...(slicer.style ? { style: slicer.style } : {}),
      native: {
        name: slicer.name,
        rowHeight: slicer.rowHeight,
        sortOrder: slicer.sortOrder,
        crossFilter: slicer.crossFilter,
      },
    };
  }

  /** Write what changed between `old` and `spec`; returns the fields that tie
   *  `spec` to the slicer's anchor as the engine holds it now. */
  private update(
    store: SpreadsheetStore,
    old: SlicerSpec,
    spec: SlicerSpec,
  ): Partial<SlicerSpec> | null {
    const index = this.indexOf(spec);
    if (index < 0 || spec.sheet === undefined) return null;
    const input: EngineSlicerInput = {};
    if (spec.caption !== old.caption && spec.caption !== undefined) input.caption = spec.caption;
    if (spec.style !== old.style) input.style = spec.style ?? '';
    if (spec.columnCount !== old.columnCount && spec.columnCount !== undefined) {
      input.columnCount = spec.columnCount;
    }
    const native = spec.native;
    if (native && native.rowHeight !== old.native?.rowHeight) input.rowHeight = native.rowHeight;
    if (native && native.sortOrder !== old.native?.sortOrder) input.sortOrder = native.sortOrder;
    if (native && native.crossFilter !== old.native?.crossFilter) {
      input.crossFilter = native.crossFilter;
    }
    const anchor = this.anchorChange(store, old, spec);
    if (anchor) Object.assign(input, anchorFields(anchor));
    if (Object.keys(input).length === 0) return null;
    if (!this.handle.updateSlicer(spec.sheet, index, input)) return old;
    const written = this.handle.listSlicers(spec.sheet)[index];
    const placed = written ? anchorOf(written) : null;
    return placed && !sameAnchor(placed, spec.nativeAnchor) ? { nativeAnchor: placed } : null;
  }

  /** The anchor `spec` moved or resized to; null when its frame is unchanged. */
  private anchorChange(
    store: SpreadsheetStore,
    old: SlicerSpec,
    spec: SlicerSpec,
  ): EngineDrawingAnchor | null {
    if (spec.nativeAnchor && !sameAnchor(old.nativeAnchor, spec.nativeAnchor)) {
      return spec.nativeAnchor;
    }
    const changed: Partial<Pick<SlicerSpec, 'x' | 'y' | 'w' | 'h'>> = {};
    for (const key of ['x', 'y', 'w', 'h'] as const) {
      if (spec[key] !== old[key] && spec[key] !== undefined) changed[key] = spec[key];
    }
    if (Object.keys(changed).length === 0) return null;
    // Unchanged sides keep the anchor's frame at the current view.
    const base = old.nativeAnchor ? sessionPosition(store, old.nativeAnchor) : this.frameOf(old);
    return twoCellAnchorForPixels(store, { ...base, ...changed }, AnchorEditAs.TwoCell);
  }

  private frameOf(spec: SlicerSpec) {
    return {
      x: spec.x ?? DEFAULT_SLICER_FRAME.x,
      y: spec.y ?? DEFAULT_SLICER_FRAME.y,
      w: spec.w ?? DEFAULT_SLICER_FRAME.w,
      h: spec.h ?? DEFAULT_SLICER_FRAME.h,
    };
  }
}
