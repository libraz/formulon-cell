import { type ChartInput, DrawingObjectKind, type Chart as EngineChart } from '@libraz/formulon';
import { DRAWING_CLASSES, Drawing, type DrawingRef } from './drawing.js';
import type { HandleContext, MemberSpec, SheetContext } from './facade-context.js';

/**
 * A chart of a sheet's drawing. Add one with `Sheet.addChart`. Members are the engine `Chart`
 * members of the same name, plus those of {@link Drawing}.
 */
// biome-ignore lint/suspicious/noUnsafeDeclarationMerging: the interface below declares the engine members installed on the prototype
export class Chart extends Drawing {}

DRAWING_CLASSES.set(DrawingObjectKind.Chart, Chart);

export const CHART_MEMBERS = {
  read: ['getProperties', 'getData', 'getPreview', 'snapshot'],
  change: {
    drawing: [
      'setZOrder',
      'setType',
      'setSource',
      'setTitle',
      'setLegend',
      'updateLegend',
      'setFilter',
      'addSeries',
      'updateSeries',
      'removeSeries',
      'setPoint',
      'updatePoint',
      'updateAxis',
      'updateDisplay',
      'updatePlot',
    ],
  },
} as const satisfies MemberSpec<HandleContext<EngineChart>>;

export const SHEET_CHART_MEMBERS = {
  methods: {
    listCharts: (ctx: SheetContext): Chart[] =>
      ctx.sheetEngine.listCharts().map((engine) => ctx.facade(engine, Chart)),
    chart: (ctx: SheetContext, ref: DrawingRef): Chart =>
      ctx.facade(ctx.sheetEngine.chart(ctx.engineOf(ref) as number), Chart),
    addChart: (ctx: SheetContext, input: ChartInput): Chart =>
      ctx.sheetChange('drawing', () => ctx.facade(ctx.sheetEngine.addChart(input), Chart)),
  },
} as const satisfies MemberSpec<SheetContext>;

type ChartMemberName =
  | (typeof CHART_MEMBERS.read)[number]
  | (typeof CHART_MEMBERS.change.drawing)[number];

export interface Chart extends Pick<EngineChart, ChartMemberName> {}

declare module './workbook.js' {
  interface Sheet {
    listCharts(): Chart[];
    /** The chart of an id; throws `NotFound` when there is none. */
    chart(ref: DrawingRef): Chart;
    /** Adds a chart over `input`'s source range and returns it. */
    addChart(input: ChartInput): Chart;
  }
}
