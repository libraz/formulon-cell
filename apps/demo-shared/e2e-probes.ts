/**
 * Exposes internal modules to the shared e2e scenarios on `window.__fcProbes`,
 * so they run the same against the dev server and a production build.
 * Demo apps only.
 */
import * as apiWorkbook from '../../packages/formulon-cell/src/api/workbook.js';
import * as chartsLayout from '../../packages/formulon-cell/src/charts/layout.js';
import * as commandsAutofitMeasurement from '../../packages/formulon-cell/src/commands/autofit-measurement.js';
import * as commandsFormatAsTable from '../../packages/formulon-cell/src/commands/format-as-table.js';
import * as commandsPivotTable from '../../packages/formulon-cell/src/commands/pivot-table.js';
import * as commandsRowColLayout from '../../packages/formulon-cell/src/commands/row-col-layout.js';
import * as commandsSessionChart from '../../packages/formulon-cell/src/commands/session-chart.js';
import * as engineAddress from '../../packages/formulon-cell/src/engine/address.js';
import * as engineAutoFilterSync from '../../packages/formulon-cell/src/engine/auto-filter-sync.js';
import * as engineWorkbookHandle from '../../packages/formulon-cell/src/engine/workbook-handle.js';
import * as interactContextMenu from '../../packages/formulon-cell/src/interact/context-menu.js';
import * as interactFormulaHelperPosition from '../../packages/formulon-cell/src/interact/formula-helper-position.js';
import * as interactObjectMenu from '../../packages/formulon-cell/src/interact/object-menu.js';
import * as mountInternals from '../../packages/formulon-cell/src/mount/internals.js';
import * as renderGeometry from '../../packages/formulon-cell/src/render/geometry.js';
import * as renderGridHitState from '../../packages/formulon-cell/src/render/grid/hit-state.js';
import * as renderPaintersText from '../../packages/formulon-cell/src/render/painters/text.js';
import * as storeStore from '../../packages/formulon-cell/src/store/store.js';
import * as themeResolve from '../../packages/formulon-cell/src/theme/resolve.js';
import * as demoSharedDemoStrings from './demo-strings.js';

export const e2eProbes = {
  'api/workbook': apiWorkbook,
  'charts/layout': chartsLayout,
  'commands/autofit-measurement': commandsAutofitMeasurement,
  'commands/format-as-table': commandsFormatAsTable,
  'commands/pivot-table': commandsPivotTable,
  'commands/row-col-layout': commandsRowColLayout,
  'commands/session-chart': commandsSessionChart,
  'engine/address': engineAddress,
  'engine/auto-filter-sync': engineAutoFilterSync,
  'engine/workbook-handle': engineWorkbookHandle,
  'interact/context-menu': interactContextMenu,
  'interact/formula-helper-position': interactFormulaHelperPosition,
  'interact/object-menu': interactObjectMenu,
  'mount/internals': mountInternals,
  'render/geometry': renderGeometry,
  'render/grid/hit-state': renderGridHitState,
  'render/painters/text': renderPaintersText,
  'store/store': storeStore,
  'theme/resolve': themeResolve,
  'demo-shared/demo-strings': demoSharedDemoStrings,
} as const;

export type E2eProbes = typeof e2eProbes;

export function installE2eProbes(): void {
  (window as unknown as { __fcProbes?: E2eProbes }).__fcProbes = e2eProbes;
}
