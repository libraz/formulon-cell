import type {
  CustomTableStyle,
  PivotTableStyleAssignment,
  TableOverlay,
} from '../commands/format-as-table.js';
import type { SheetView, SheetViewPatch } from '../commands/sheet-views.js';
import { addrKey, parseAddrKey } from '../engine/address.js';
import type { Addr, Range } from '../engine/types.js';
import { rangesIntersect } from './selection-geometry.js';
import type { SpreadsheetStore } from './store.js';
import type {
  ConditionalRule,
  SessionChart,
  SessionIllustration,
  SlicerSpec,
  Sparkline,
} from './types.js';

function keyInRange(key: string, range: Range): boolean {
  const addr = parseAddrKey(key);
  return (
    addr !== null &&
    addr.sheet === range.sheet &&
    addr.row >= range.r0 &&
    addr.row <= range.r1 &&
    addr.col >= range.c0 &&
    addr.col <= range.c1
  );
}

/** Session objects layered over the grid: conditional rules, sparklines,
 *  charts, illustrations, slicers, table overlays/styles and sheet views. */
export const sheetObjectMutators = {
  addConditionalRule(store: SpreadsheetStore, rule: ConditionalRule): void {
    store.setState((s) => ({
      ...s,
      conditional: { rules: [...s.conditional.rules, rule] },
    }));
  },

  removeConditionalRuleAt(store: SpreadsheetStore, idx: number): void {
    store.setState((s) => ({
      ...s,
      conditional: { rules: s.conditional.rules.filter((_, i) => i !== idx) },
    }));
  },

  clearConditionalRules(store: SpreadsheetStore): void {
    store.setState((s) => ({ ...s, conditional: { rules: [] } }));
  },

  clearConditionalRulesInRange(store: SpreadsheetStore, range: Range): void {
    store.setState((s) => {
      const rules = s.conditional.rules.filter((rule) => !rangesIntersect(rule.range, range));
      if (rules.length === s.conditional.rules.length) return s;
      return { ...s, conditional: { rules } };
    });
  },

  /** Attach a sparkline spec to `addr`. Pass `null` to remove. */
  setSparkline(store: SpreadsheetStore, addr: Addr, spec: Sparkline | null): void {
    store.setState((s) => {
      const sparklines = new Map(s.sparkline.sparklines);
      const key = addrKey(addr);
      if (spec === null) sparklines.delete(key);
      else sparklines.set(key, { ...spec });
      return { ...s, sparkline: { sparklines } };
    });
  },

  clearSparkline(store: SpreadsheetStore, addr: Addr): void {
    store.setState((s) => {
      const sparklines = new Map(s.sparkline.sparklines);
      sparklines.delete(addrKey(addr));
      return { ...s, sparkline: { sparklines } };
    });
  },

  clearSparklinesInRange(store: SpreadsheetStore, range: Range): void {
    store.setState((s) => {
      const sparklines = new Map(s.sparkline.sparklines);
      for (const key of sparklines.keys()) {
        if (keyInRange(key, range)) sparklines.delete(key);
      }
      if (sparklines.size === s.sparkline.sparklines.size) return s;
      return { ...s, sparkline: { sparklines } };
    });
  },

  upsertChart(store: SpreadsheetStore, chart: SessionChart): void {
    store.setState((s) => {
      const next = s.charts.charts.filter((c) => c.id !== chart.id);
      return { ...s, charts: { charts: [...next, { ...chart, source: { ...chart.source } }] } };
    });
  },

  removeChart(store: SpreadsheetStore, id: string): void {
    store.setState((s) => {
      const next = s.charts.charts.filter((c) => c.id !== id);
      if (next.length === s.charts.charts.length) return s;
      return { ...s, charts: { charts: next } };
    });
  },

  updateChart(store: SpreadsheetStore, id: string, patch: Partial<Omit<SessionChart, 'id'>>): void {
    store.setState((s) => {
      let changed = false;
      const next = s.charts.charts.map((chart) => {
        if (chart.id !== id) return chart;
        changed = true;
        return { ...chart, ...patch, source: patch.source ? { ...patch.source } : chart.source };
      });
      if (!changed) return s;
      return { ...s, charts: { charts: next } };
    });
  },

  clearChartsInRange(store: SpreadsheetStore, range: Range): void {
    store.setState((s) => {
      const next = s.charts.charts.filter((chart) => !rangesIntersect(chart.source, range));
      if (next.length === s.charts.charts.length) return s;
      return { ...s, charts: { charts: next } };
    });
  },

  upsertIllustration(store: SpreadsheetStore, illustration: SessionIllustration): void {
    store.setState((s) => {
      const next = s.illustrations.illustrations.filter((item) => item.id !== illustration.id);
      return { ...s, illustrations: { illustrations: [...next, { ...illustration }] } };
    });
  },

  removeIllustration(store: SpreadsheetStore, id: string): void {
    store.setState((s) => {
      const next = s.illustrations.illustrations.filter((item) => item.id !== id);
      if (next.length === s.illustrations.illustrations.length) return s;
      return { ...s, illustrations: { illustrations: next } };
    });
  },

  clearIllustrations(store: SpreadsheetStore): void {
    store.setState((s) =>
      s.illustrations.illustrations.length === 0
        ? s
        : { ...s, illustrations: { illustrations: [] } },
    );
  },

  updateIllustration(
    store: SpreadsheetStore,
    id: string,
    patch: Partial<Omit<SessionIllustration, 'id'>>,
  ): void {
    store.setState((s) => {
      let changed = false;
      const next = s.illustrations.illustrations.map((item) => {
        if (item.id !== id) return item;
        changed = true;
        return { ...item, ...patch };
      });
      if (!changed) return s;
      return { ...s, illustrations: { illustrations: next } };
    });
  },

  /** Append a fresh slicer to the slice. Caller is responsible for picking a
   *  unique `id` — duplicates are rejected (the older spec wins). */
  addSlicer(store: SpreadsheetStore, spec: SlicerSpec): void {
    store.setState((s) => {
      if (s.slicers.slicers.some((sp) => sp.id === spec.id)) return s;
      return {
        ...s,
        slicers: { slicers: [...s.slicers.slicers, { ...spec, selected: [...spec.selected] }] },
      };
    });
  },

  /** Remove the slicer with `id`. No-op when not present. */
  removeSlicer(store: SpreadsheetStore, id: string): void {
    store.setState((s) => {
      const next = s.slicers.slicers.filter((sp) => sp.id !== id);
      if (next.length === s.slicers.slicers.length) return s;
      return { ...s, slicers: { slicers: next } };
    });
  },

  /** Merge a partial patch onto the slicer with `id`. Skips the `id` field —
   *  ids stay immutable. */
  updateSlicer(store: SpreadsheetStore, id: string, patch: Partial<Omit<SlicerSpec, 'id'>>): void {
    store.setState((s) => {
      let changed = false;
      const next = s.slicers.slicers.map((sp) => {
        if (sp.id !== id) return sp;
        changed = true;
        return {
          ...sp,
          ...patch,
          selected: patch.selected ? [...patch.selected] : sp.selected,
        };
      });
      if (!changed) return s;
      return { ...s, slicers: { slicers: next } };
    });
  },

  /** Replace the chip selection for slicer `id`. Empty array = "include all". */
  setSlicerSelected(store: SpreadsheetStore, id: string, values: readonly string[]): void {
    store.setState((s) => {
      let changed = false;
      const next = s.slicers.slicers.map((sp) => {
        if (sp.id !== id) return sp;
        changed = true;
        return { ...sp, selected: [...values] };
      });
      if (!changed) return s;
      return { ...s, slicers: { slicers: next } };
    });
  },

  upsertTableOverlay(store: SpreadsheetStore, next: TableOverlay): void {
    store.setState((s) => {
      const filtered = s.tables.tables.filter((t) => t.id !== next.id);
      return { ...s, tables: { ...s.tables, tables: [...filtered, next] } };
    });
  },

  removeTableOverlay(store: SpreadsheetStore, id: string): void {
    store.setState((s) => {
      const next = s.tables.tables.filter((t) => t.source === 'engine' || t.id !== id);
      if (next.length === s.tables.tables.length) return s;
      return { ...s, tables: { ...s.tables, tables: next } };
    });
  },

  clearTableOverlaysInRange(store: SpreadsheetStore, range: Range): void {
    store.setState((s) => {
      const next = s.tables.tables.filter(
        (t) => t.source === 'engine' || !rangesIntersect(t.range, range),
      );
      if (next.length === s.tables.tables.length) return s;
      return { ...s, tables: { ...s.tables, tables: next } };
    });
  },

  replaceEngineTableOverlays(store: SpreadsheetStore, tables: readonly TableOverlay[]): void {
    store.setState((s) => {
      const session = s.tables.tables.filter((t) => t.source !== 'engine');
      return { ...s, tables: { ...s.tables, tables: [...tables, ...session] } };
    });
  },

  upsertCustomTableStyle(store: SpreadsheetStore, style: CustomTableStyle): void {
    store.setState((s) => {
      const next = (s.tables.customTableStyles ?? []).filter((item) => item.id !== style.id);
      return {
        ...s,
        tables: {
          ...s.tables,
          customTableStyles: [...next, { ...style }],
        },
      };
    });
  },

  upsertCustomPivotTableStyle(store: SpreadsheetStore, style: CustomTableStyle): void {
    store.setState((s) => {
      const next = (s.tables.customPivotTableStyles ?? []).filter((item) => item.id !== style.id);
      return {
        ...s,
        tables: { ...s.tables, customPivotTableStyles: [...next, { ...style }] },
      };
    });
  },

  upsertPivotTableStyle(store: SpreadsheetStore, style: PivotTableStyleAssignment): void {
    store.setState((s) => {
      const next = (s.tables.pivotTableStyles ?? []).filter(
        (item) => item.sheetIndex !== style.sheetIndex || item.pivotIndex !== style.pivotIndex,
      );
      return {
        ...s,
        tables: { ...s.tables, pivotTableStyles: [...next, { ...style }] },
      };
    });
  },

  upsertSheetView(store: SpreadsheetStore, view: SheetView): void {
    store.setState((s) => {
      const next = s.sheetViews.views.filter((v) => v.id !== view.id);
      return { ...s, sheetViews: { ...s.sheetViews, views: [...next, view] } };
    });
  },

  removeSheetView(store: SpreadsheetStore, id: string): void {
    store.setState((s) => {
      const views = s.sheetViews.views.filter((v) => v.id !== id);
      if (views.length === s.sheetViews.views.length) return s;
      return {
        ...s,
        sheetViews: {
          views,
          activeViewId: s.sheetViews.activeViewId === id ? null : s.sheetViews.activeViewId,
        },
      };
    });
  },

  applySheetViewPatch(
    store: SpreadsheetStore,
    patch: SheetViewPatch,
    activeViewId: string | null = null,
  ): void {
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        freezeRows: Math.max(0, Math.floor(patch.freezeRows)),
        freezeCols: Math.max(0, Math.floor(patch.freezeCols)),
        hiddenRows: new Set(patch.hiddenRows),
        hiddenCols: new Set(patch.hiddenCols),
      },
      viewport: {
        ...s.viewport,
        rowStart: Math.max(Math.max(0, Math.floor(patch.freezeRows)), s.viewport.rowStart),
        colStart: Math.max(Math.max(0, Math.floor(patch.freezeCols)), s.viewport.colStart),
      },
      ui: { ...s.ui, filterRange: patch.filterRange, filterCriteria: patch.filterCriteria },
      sheetViews: { ...s.sheetViews, activeViewId },
    }));
  },
};
