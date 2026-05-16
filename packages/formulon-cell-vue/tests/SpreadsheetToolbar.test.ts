import { existsSync, readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import {
  buildRibbonModel,
  EMPTY_ACTIVE_STATE,
  mutators,
  projectActiveState,
  RIBBON_TAB_LABELS,
  type RibbonTab,
  toggleBold,
} from '@libraz/formulon-cell';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createApp, defineComponent, h, nextTick, type Ref, shallowRef } from 'vue';
import { useToolbarActive } from '../src/toolbar/active';
import { useToolbarDropdown } from '../src/toolbar/dropdown';
import { toolbarTabs } from '../src/toolbar/tabs';
import {
  installVueDomStubs,
  type MountedVueSpreadsheet,
  mountVueSpreadsheet,
  uninstallVueDomStubs,
} from './test-utils/mount';

/**
 * The Vue ribbon ships as a Single File Component (`SpreadsheetToolbar.vue`)
 * which our happy-dom vitest config can't parse without a `@vitejs/plugin-vue`
 * dependency the project deliberately avoids. Instead we test the building
 * blocks the SFC composes — `toolbarTabs`, `useToolbarActive`,
 * `useToolbarDropdown` — and we mount a minimal `<RibbonProbe>` Vue component
 * so all of it is exercised through Vue's reactivity, not in isolation.
 */

const flush = async (): Promise<void> => {
  for (let i = 0; i < 8; i += 1) await Promise.resolve();
  await nextTick();
};

const readToolbarSource = (): string => {
  const sourcePath = [
    resolve(process.cwd(), 'src/SpreadsheetToolbar.vue'),
    resolve(process.cwd(), '../formulon-cell-vue/src/SpreadsheetToolbar.vue'),
    resolve(process.cwd(), 'packages/formulon-cell-vue/src/SpreadsheetToolbar.vue'),
  ].find((candidate) => existsSync(candidate));
  if (!sourcePath) throw new Error('SpreadsheetToolbar.vue source not found');
  return readFileSync(sourcePath, 'utf8');
};

interface RibbonProbeHandle {
  host: HTMLElement;
  /** Active state ref returned by `useToolbarActive`. */
  active: Ref<ReturnType<typeof projectActiveState>>;
  /** Open dropdown name. */
  openDropdown: Ref<string | null>;
  keydownDropdown: (event: KeyboardEvent) => void;
  toggleDropdown: (name: string) => void;
  pickDropdown: (name: string, value: string | number) => void;
  unmount: () => Promise<void>;
}

interface DropdownLogEntry {
  kind: string;
  value: unknown;
}

async function mountRibbonProbe(
  instance: MountedVueSpreadsheet['instance'],
  log: DropdownLogEntry[],
): Promise<RibbonProbeHandle> {
  installVueDomStubs();
  const host = document.createElement('div');
  document.body.appendChild(host);

  const instanceRef = shallowRef(instance);
  const activeRef = shallowRef<ReturnType<typeof projectActiveState>>(EMPTY_ACTIVE_STATE);
  const openRef = shallowRef<string | null>(null);
  let keydown: ((event: KeyboardEvent) => void) | null = null;
  let toggle: ((name: string) => void) | null = null;
  let pick: ((name: string, value: string | number) => void) | null = null;

  const Probe = defineComponent({
    setup() {
      const active = useToolbarActive(() => instanceRef.value);
      activeRef.value = active.value;

      const dd = useToolbarDropdown({
        onBorderPreset: (v) => log.push({ kind: 'borderPreset', value: v }),
        onBorderStyle: (v) => log.push({ kind: 'borderStyle', value: v }),
        onFontFamily: (v) => log.push({ kind: 'fontFamily', value: v }),
        onFontSize: (v) => log.push({ kind: 'fontSize', value: v }),
        onMarginPreset: (v) => log.push({ kind: 'marginPreset', value: v }),
        onOpenPageSetup: () => log.push({ kind: 'openPageSetup', value: null }),
        onPageOrientation: (v) => log.push({ kind: 'pageOrientation', value: v }),
        onPaperSize: (v) => log.push({ kind: 'paperSize', value: v }),
      });

      keydown = dd.onDropdownKeydown;
      toggle = (name) => dd.toggleDropdown(name as never);
      pick = (name, value) => dd.onDropdownPick(name as never, value);

      return () => {
        // Mirror `active` and `openDropdown` into shallow refs the test reads.
        activeRef.value = active.value;
        openRef.value = dd.openDropdown.value;
        return h('div', { class: 'probe' }, [
          h('span', { 'data-testid': 'bold' }, String(active.value.bold)),
          h('span', { 'data-testid': 'fontSize' }, String(active.value.fontSize)),
          h('span', { 'data-testid': 'open' }, String(dd.openDropdown.value ?? '')),
        ]);
      };
    },
  });

  const app = createApp(Probe);
  app.mount(host);
  await flush();

  return {
    host,
    active: activeRef,
    openDropdown: openRef,
    keydownDropdown: (event) => {
      if (!keydown) throw new Error('keydown not yet bound');
      keydown(event);
    },
    toggleDropdown: (name) => {
      if (!toggle) throw new Error('toggle not yet bound');
      toggle(name);
    },
    pickDropdown: (name, value) => {
      if (!pick) throw new Error('pick not yet bound');
      pick(name, value);
    },
    async unmount() {
      app.unmount();
      await flush();
      host.remove();
      uninstallVueDomStubs();
    },
  };
}

describe('Vue toolbar — toolbarTabs builder', () => {
  it('returns one entry per ribbon tab using RIBBON_TAB_LABELS', () => {
    const en = toolbarTabs('en');
    const expectedIds = Object.keys(RIBBON_TAB_LABELS) as RibbonTab[];
    expect(en.map((t) => t.id)).toEqual(expectedIds);
    expect(en.map((t) => t.label)).toEqual(expectedIds.map((id) => RIBBON_TAB_LABELS[id].en));

    const ja = toolbarTabs('ja');
    expect(ja.map((t) => t.label)).toEqual(expectedIds.map((id) => RIBBON_TAB_LABELS[id].ja));
  });
});

describe('Vue <SpreadsheetToolbar> ribbon command surface', () => {
  it('exposes the shared core ribbon commands as DOM command ids', () => {
    const source = readToolbarSource();
    const exposed = new Set(
      Array.from(source.matchAll(/data-ribbon-command="([^"]+)"/g), (match) => match[1]),
    );
    const coreIds = new Set(
      buildRibbonModel('en')
        .flatMap((tab) => tab.groups)
        .flatMap((group) => group.commands)
        .map((command) => command.id),
    );

    const missing = Array.from(coreIds).filter((id) => !exposed.has(id));
    expect(missing).toEqual([]);
  });

  it('keeps chart and data-validation commands on their dedicated handlers', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-ribbon-command="chartInsert"');
    expect(source).toContain('data-dropdown-name="chartInsert"');
    expect(source).toContain('data-ribbon-command="viewNormal"');
    expect(source).toContain('data-ribbon-command="viewPageLayout"');
    expect(source).toContain('data-ribbon-command="viewPageBreakPreview"');
    expect(source).toContain('setWorkbookView(inst.store, mode)');
    expect(source).toContain(
      'data-ribbon-command="viewFormulaBar" :class="{ \'demo__rb--active\': formulaBarVisible }"',
    );
    expect(source).toContain('strings.value.ribbonDisplay');
    expect(source).toContain('strings.value.backstage');
    expect(source).toContain(
      'data-ribbon-command="pageSetup" type="button" :disabled="disabled" @click="props.instance?.openPageSetup()"',
    );
    expect(source).toContain('data-ribbon-command="protect" type="button"');
    expect(source).toContain('@click="onBackstageProtectWorkbook"');
    expect(source).toContain("'demo__backstage-command--active': workbookStructureProtected");
    expect(source).toContain(':aria-pressed="workbookStructureProtected ? \'true\' : undefined"');
    expect(source).toContain(
      'data-ribbon-command="inspect" type="button" :disabled="disabled" @click="onBackstageInspectWorkbook"',
    );
    expect(source).toContain(
      'setWorkbookStructureProtected(inst.store, !isWorkbookStructureProtected(inst.store.getState()))',
    );
    expect(source).toContain('summarizeSpreadsheetCompatibility(inst.workbook)');
    expect(source).toContain('const objectsText = strings.value.workbookObjects');
    expect(source).toContain('objectsText.compatibilityLabels.cellFormatting');
    expect(source).toContain('objectsText.compatibilityDetails.cellFormatting');
    expect(source).toContain('objectsText.sessionOnly');
    expect(source).toContain('strings.value.pageScale');
    expect(source).toContain('strings.value.viewToggle');
    expect(source).toContain('data-dropdown-name="textOrientation"');
    expect(source).toContain("'demo__rb--active': active.textOrientation !== 'horizontalText'");
    expect(source).toContain("'demo__rb--active': active.textOrientation === 'rotateTextDown'");
    expect(source).toContain('role="menuitemradio"');
    expect(source).toContain("onTextOrientationAction('rotateTextDown')");
    expect(source).toContain('cellText.orientationHorizontalText');
    expect(source).toContain("onTextOrientationAction('formatAlignment')");
    expect(source).toContain("props.instance?.openFormatDialog('align')");
    expect(source).toContain('cellText.orientationFormatAlignment');
    expect(source).toContain(':title="tr.number" :aria-label="tr.number"');
    expect(source).toContain('role="listbox" :aria-label="tr.number"');
    expect(source).toContain("inst.openFormatDialog('number')");
    expect(source).toContain("{ kind: 'fixed', decimals: 2, thousands: true }");
    expect(source).not.toContain('title="Number format"');
    expect(source).not.toContain('aria-label="Number format"');
    expect(source).toContain("onCreateChart('column')");
    expect(source).toContain("onCreateChart('bar')");
    expect(source).toContain("onCreateChart('line')");
    expect(source).toContain("onCreateChart('area')");
    expect(source).toContain("onCreateChart('pie')");
    expect(source).toContain("onCreateChart('scatter')");
    expect(source).toContain("onCreateChart('recommended')");
    expect(source).toContain("action === 'recommended'");
    expect(source).toContain('createSessionChart(');
    expect(source).toContain('inst.store');
    expect(source).toContain('range,');
    expect(source).toContain('inst.history');
    expect(source).not.toContain(
      'data-ribbon-command="chartInsert" type="button" :disabled="disabled" @click="props.instance?.openQuickAnalysis()"',
    );
    expect(source).toContain('data-dropdown-name="symbolInsert"');
    expect(source).toContain('const insertSymbols = [');
    expect(source).toContain("'π'");
    expect(source).toContain("'₹'");
    expect(source).toContain('[12, 24, 32].includes(index)');
    expect(source).toContain('onSymbolAction(symbol)');
    expect(source).toContain('MORE_SYMBOL_ACTION');
    expect(source).toContain(':data-insert-action="MORE_SYMBOL_ACTION"');
    expect(source).toContain('window.prompt(cellText.value.symbolPrompt');
    expect(source).toContain('isCellWritable(inst.store.getState(), addr)');
    expect(source).toContain('writableAddrs(inst.store.getState(), range)');
    expect(source).toContain('data-ribbon-command="dataValidation"');
    expect(source).toContain('data-dropdown-name="dataValidation"');
    expect(source).toContain("onDataValidationAction('settings')");
    expect(source).not.toContain('@click="props.instance?.openFormatDialog(\'more\')"');
  });

  it('wires Acrobat Add-ins and PDF commands to host callback and print', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-ribbon-command="addIn"');
    expect(source).toContain("onAddInAction('my')");
    expect(source).toContain('cellText.addInGet');
    expect(source).toContain('title: cellText.value.addInMy');
    expect(source).toContain('cellText.addInManage');
    expect(source).toContain('data-ribbon-command="pdf"');
    expect(source).toContain("onPdfAction('create')");
    expect(source).toContain("onPdfAction('share')");
    expect(source).toContain("onPdfAction('preferences')");
    expect(source).toContain('cellText.value.pdfCreateReady');
  });

  it('renders View > Window > Freeze Panes as an Excel-style menu', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-ribbon-command="freeze"');
    expect(source).toContain('data-dropdown-name="freeze"');
    expect(source).toContain('data-dropdown-name="windowVisibility"');
    expect(source).toContain("onWindowAction('hideRows')");
    expect(source).toContain("onWindowAction('showRows')");
    expect(source).toContain("onWindowAction('hideCols')");
    expect(source).toContain("onWindowAction('showCols')");
    expect(source).toContain('hideRows(inst.store, inst.history, r.r0, r.r1, inst.workbook)');
    expect(source).toContain(
      'showColsAroundSelection(inst.store, inst.history, r.c0, r.c1, inst.workbook)',
    );
    expect(source).toContain('data-ribbon-command="zoomSelection"');
    expect(source).toContain('data-ribbon-command="zoomDialog"');
    expect(source).toContain('data-ribbon-command="zoom75"');
    expect(source).toContain('data-ribbon-command="zoom125"');
    expect(source).toContain('@click="openZoomDialog"');
    expect(source).toContain('setSheetZoom(inst.store, clamped / 100, inst.workbook)');
    expect(source).toContain('v-model="zoomDialog"');
    expect(source).toContain('@click="onZoomSelection"');
    expect(source).toContain('@click="onZoom(0.75)"');
    expect(source).toContain('@click="onZoom(1.25)"');
    expect(source).toContain('state.viewport.zoom * Math.min(rowFit, colFit)');
    expect(source).toContain('const activeFreezeAction = computed<FreezeAction>');
    expect(source).toContain("onFreezeAction('none')");
    expect(source).toContain("onFreezeAction('panes')");
    expect(source).toContain("onFreezeAction('topRow')");
    expect(source).toContain("onFreezeAction('firstColumn')");
    // Freeze logic moved to the shared `handleFreezeAction` helper in core.
    expect(source).toContain('handleFreezeAction(props.instance, action)');
    expect(source).toContain('strings.value.viewToolbar');
    expect(source).toContain('{{ viewToolbarText.freezeNone }}');
    expect(source).toContain('{{ viewToolbarText.freezePanes }}');
    expect(source).toContain('{{ viewToolbarText.freezeTopRow }}');
    expect(source).toContain('{{ viewToolbarText.freezeFirstColumn }}');
  });

  it('renders conditional-formatting as a menu backed by preset actions', () => {
    const source = readToolbarSource();

    // Action dispatch moved to the shared `handleConditionalAction` helper in core.
    expect(source).toContain('handleConditionalAction(props.instance, action)');
    expect(source).toContain('active.conditionalFormatting');
    expect(source).toContain('data-cf-action="cell-greater"');
    expect(source).toContain('data-cf-action="cell-equal"');
    expect(source).toContain('data-cf-action="text-contains"');
    expect(source).toContain('data-cf-action="date-occurring"');
    // Submenu dispatch (open conditional dialogs for various kinds) is now
    // performed by `handleConditionalAction` inside core; the assertions
    // above (data-cf-action attributes and conditional menu helper export)
    // cover the wiring.
    expect(source).toContain('data-cf-action="unique"');
    expect(source).toContain('data-cf-action="top10-percent"');
    expect(source).toContain("'data-solid-green'");
    expect(source).toContain("'data-solid-gray'");
    expect(source).toContain("'scale-ryg'");
    expect(source).toContain("'scale-gwg'");
    expect(source).toContain(':title="cfDataBarLabel(action)"');
    expect(source).toContain(':aria-label="cfScaleLabel(action)"');
    expect(source).toContain('conditionalDataBarLabel');
    expect(source).toContain('conditionalColorScaleLabel');
    expect(source).toContain("'icons-arrows5'");
    expect(source).toContain('cfIconSetGroups');
    expect(source).toContain('cfIconSetLabel(item.action)');
    expect(source).toContain('demo__cf-menu__panel--icons');
    expect(source).toContain('demo__cf-menu__iconset');
    expect(source).toContain('data-cf-action="highlight-more"');
    expect(source).toContain('data-cf-action="top-bottom-more"');
    expect(source).toContain('data-cf-action="data-bars-more"');
    expect(source).toContain('data-cf-action="color-scales-more"');
    expect(source).toContain('data-cf-action="icon-sets-more"');
    // Data-bar/color-scale/icon-set submenu dispatch is performed by
    // `handleConditionalAction` (covered above); the data attributes here
    // assert the buttons are wired with the matching action ids.
    expect(source).toContain('{{ cfText.clear }}');
    expect(source).toContain('data-cf-action="clear-selection"');
    expect(source).toContain('data-cf-action="clear-sheet"');
  });

  it('applies the compact Office 365 Home ribbon layout only on the Home tab', () => {
    const source = readToolbarSource();

    expect(source).toContain("'demo__ribbon--office365-home': props.activeTab === 'home'");
    expect(source).toContain(":class=\"['demo__ribbon'");
  });

  it('wires Home alignment indent commands through range formatting', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-ribbon-command="indentDecrease"');
    expect(source).toContain('data-ribbon-command="indentIncrease"');
    expect(source).toContain('wrapFormat((s, st) => bumpIndent(s, st, -1))');
    expect(source).toContain('wrapFormat((s, st) => bumpIndent(s, st, 1))');
    expect(source).toContain('RibbonIcon name="indentDecrease"');
    expect(source).toContain('RibbonIcon name="indentIncrease"');
  });

  it('wires Home font, alignment, and number commands through range formatting', () => {
    const source = readToolbarSource();

    for (const command of [
      'fontGrow',
      'fontShrink',
      'italic',
      'underline',
      'strike',
      'top',
      'middle',
      'bottomAlign',
      'wrap',
      'alignL',
      'alignC',
      'alignR',
      'currency',
      'decDown',
      'decUp',
    ]) {
      expect(source).toContain(`data-ribbon-command="${command}"`);
    }

    expect(source).toContain('setFont(s, st, { fontSize: active.fontSize + 1 })');
    expect(source).toContain('setFont(s, st, { fontSize: Math.max(1, active.fontSize - 1) })');
    expect(source).toContain("'demo__rb--active': active.vAlignTop");
    expect(source).toContain("'demo__rb--active': active.vAlignMiddle");
    expect(source).toContain("'demo__rb--active': active.vAlignBottom");
    expect(source).toContain("'demo__rb--active': active.wrapText");
    expect(source).toContain("'demo__rb--active': active.merged");
    expect(source).toContain("'demo__rb--active': active.mergeCenter");
    expect(source).toContain("'demo__rb--active': active.commaStyle");
    expect(source).toContain('wrapFormat(toggleItalic)');
    expect(source).toContain('wrapFormat(toggleUnderline)');
    expect(source).toContain('wrapFormat(toggleStrike)');
    expect(source).toContain("setVAlign(s, st, 'middle')");
    expect(source).toContain("setVAlign(s, st, 'bottom')");
    expect(source).toContain("const onAlign = (kind: 'left' | 'center' | 'right'): void =>");
    expect(source).toContain('wrapFormat((s, st) => setAlign(s, st, kind))');
    expect(source).toContain('@click="onAlign(\'left\')"');
    expect(source).toContain('@click="onAlign(\'center\')"');
    expect(source).toContain('@click="onAlign(\'right\')"');
    expect(source).toContain('cycleCurrency(s, st, lang)');
    expect(source).toContain('const onBumpDecimals = (delta: 1 | -1): void =>');
    expect(source).toContain('wrapFormat((s, st) => bumpDecimals(s, st, delta))');
    expect(source).toContain('@click="onBumpDecimals(-1)"');
    expect(source).toContain('@click="onBumpDecimals(1)"');
  });

  it('renders Cells group insert/delete/format commands as menus', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-dropdown-name="cellInsert"');
    expect(source).toContain('data-dropdown-name="cellDelete"');
    expect(source).toContain('data-dropdown-name="cellFormat"');
    // Insert / delete dispatch moved to the shared core helpers.
    expect(source).toContain('handleInsertCellsAction(props.instance, action)');
    expect(source).toContain('handleDeleteCellsAction(props.instance, action)');
    expect(source).toContain("onInsertCellsAction('sheet')");
    expect(source).toContain("onDeleteCellsAction('sheet')");
    expect(source).toContain("onCellFormatAction('rowHeight')");
    expect(source).toContain("onCellFormatAction('autoFitRowHeight')");
    expect(source).toContain("onCellFormatAction('colWidth')");
    expect(source).toContain("onCellFormatAction('autoFitColWidth')");
    expect(source).toContain('const dimensionDialog = ref<DimensionDialogDraft | null>(null)');
    expect(source).toContain("dimensionDialog.value = { kind: 'rowHeight'");
    expect(source).toContain("dimensionDialog.value = { kind: 'colWidth'");
    expect(source).toContain('v-model="dimensionDialog.value"');
    expect(source).not.toContain('window.prompt(cellText.value.rowHeightPrompt');
    expect(source).not.toContain('window.prompt(cellText.value.colWidthPrompt');
    expect(source).toContain("onCellFormatAction('hideRows')");
    expect(source).toContain("onCellFormatAction('showCols')");
    expect(source).toContain("onCellFormatAction('renameSheet')");
    expect(source).toContain("onCellFormatAction('moveSheetLeft')");
    expect(source).toContain("onCellFormatAction('moveSheetRight')");
    expect(source).toContain("onCellFormatAction('hideSheet')");
    expect(source).toContain("onCellFormatAction('unhideSheet')");
    expect(source).toContain('const sheetRenameDialog = ref<SheetRenameDialogDraft | null>(null)');
    expect(source).toContain('renameSheet(inst.workbook, inst.store.getState().data.sheetIndex');
    expect(source).toContain(
      'moveSheet(inst.store, inst.workbook, sheet, sheet - 1, inst.history)',
    );
    expect(source).toContain(
      'moveSheet(inst.store, inst.workbook, sheet, sheet + 1, inst.history)',
    );
    expect(source).toContain('setSheetHidden(inst.store, inst.workbook, inst.history');
    expect(source).toContain('data-cell-action="tabColorNone"');
    expect(source).toContain('SHEET_TAB_COLOR_ACTIONS');
    expect(source).toContain('recordLayoutChange(inst.history, inst.store, () =>');
    expect(source).toContain('mutators.setSheetTabColor(inst.store, state.data.sheetIndex');
    expect(source).toContain("onCellFormatAction('protectSheet')");
    expect(source).toContain(':class="{ \'demo__rb--active\': active.protected }"');
    expect(source).toContain(':aria-checked="active.protected"');
    expect(source).toContain('setRowsHeight(inst.store, inst.history');
    expect(source).toContain('setColsWidth(inst.store, inst.history');
  });

  it('renders Editing group commands as Excel-style menus', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-ribbon-command="paste"');
    expect(source).toContain('data-ribbon-command="formatPainter"');
    expect(source).toContain('data-dropdown-name="paste"');
    expect(source).toContain('data-paste-action="paste"');
    expect(source).toContain('data-paste-action="pasteFormulas"');
    expect(source).toContain('data-paste-action="pasteFormulasNumFmt"');
    expect(source).toContain('data-paste-action="pasteValues"');
    expect(source).toContain('data-paste-action="pasteValuesNumFmt"');
    expect(source).toContain('data-paste-action="pasteFormatsOnly"');
    expect(source).toContain('data-paste-action="pasteTranspose"');
    expect(source).toContain('data-paste-action="insertCopiedCells"');
    // Paste dispatch is now performed by the shared `handlePasteAction`
    // helper in core; the data attributes above + helper call check
    // every action routes through it.
    expect(source).toContain('handlePasteAction(props.instance, action)');
    expect(source).toContain('data-paste-action="pasteSpecial"');
    expect(source).toContain('@click="onFormatPainter"');
    expect(source).toContain('data-dropdown-name="fillHome"');
    expect(source).toContain('data-dropdown-name="clearHome"');
    expect(source).toContain('data-dropdown-name="sortHome"');
    expect(source).toContain('data-dropdown-name="findHome"');
    expect(source.match(/data-ribbon-command="clearFormat"/g)).toHaveLength(1);
    expect(source).toContain("onFillAction('down')");
    expect(source).toContain("onFillAction('flash')");
    expect(source).toContain('inferFlashFillPattern(examples)');
    expect(source).toContain('applyFlashFill(');
    expect(source).toContain(
      'isCellWritable(inst.store.getState(), { sheet: range.sheet, row, col: range.c0 })',
    );
    expect(source).toContain("onFillAction('weekdays')");
    expect(source).toContain("onFillAction('months')");
    expect(source).toContain('dateUnit: isDateSeries ? action : undefined');
    expect(source).toContain('fillRange(inst.store.getState(), inst.workbook, src, range');
    expect(source).toContain("onClearAction('contents')");
    expect(source).toContain(
      'clearValidationInRangeWithEngine(inst.store, inst.history, inst.workbook, range)',
    );
    expect(source).toContain('wrapFormat(clearVisualFormat)');
    expect(source).toContain("onSortMenuAction('dedupe')");
    expect(source).toContain("onSortMenuAction('filter-reapply')");
    expect(source).toContain("onSortMenuAction('filter-advanced')");
    expect(source).toContain('reapplyFilters(inst.store.getState(), inst.store)');
    expect(source).toContain('advancedFilterDialog.value = {');
    expect(source).toContain('listRange: formatA1Range(s.selection.range)');
    expect(source).toContain('copyTo:');
    expect(source).toContain('data-ribbon-command="removeDupes"');
    expect(source).toContain('inst.history.begin()');
    expect(source).toContain('inst.history.end()');
    expect(source).toContain('ok = sortRange(s, inst.store, inst.workbook, range');
    expect(source).toContain(
      'const removeDuplicatesDialog = ref<RemoveDuplicatesDialogDraft | null>(null)',
    );
    expect(source).toContain('removeDuplicatesDialog.value = {');
    expect(source).toContain('cellText.removeDuplicatesDialogTitle');
    expect(source).toContain('cellText.removeDuplicatesColumns');
    expect(source).toContain('cellText.removeDuplicatesSelectAll');
    expect(source).toContain('removeDuplicates(s, inst.store, inst.workbook, s.selection.range, {');
    expect(source).toContain('columns: draft.columns');
    expect(source).toContain('hasHeader: draft.hasHeader');
    expect(source).toContain("else if (action === 'conditional') inst.openCfRulesDialog()");
    expect(source).toContain("onSortMenuAction('custom')");
    expect(source).toContain('cellText.sortCustom');
    expect(source).toContain('const sortDialog = ref<SortDialogDraft | null>(null)');
    expect(source).toContain('cellText.sortDialogTitle');
    expect(source).toContain('v-model.number="sortDialog.byCol"');
    expect(source).not.toContain('window.prompt(cellText.value.sortColumnPrompt');
    expect(source).toContain("onFindAction('go-to-special')");
    expect(source).toContain("onFindAction('formulas')");
    expect(source).toContain("onFindAction('constants')");
    expect(source).toContain("onFindAction('numbers')");
    expect(source).toContain("onFindAction('text')");
    expect(source).toContain("onFindAction('errors')");
    expect(source).toContain("onFindAction('data-validation')");
    expect(source).toContain("inst.openFindReplace('replace')");
    expect(source).toContain("action === 'go-to') inst.openGoTo()");
    expect(source).toContain("action === 'go-to-special') inst.openGoToSpecial()");
    expect(source).toContain("'numbers'");
    expect(source).toContain("'text'");
    expect(source).toContain("'errors'");
    expect(source).toContain("findMatchingCells(inst.workbook, inst.store, 'sheet', kind)");
    expect(source).toContain('label: cellText.value.findNoMatches');
    expect(source).toContain('label: cellText.value.commentNone');
    expect(source).toContain('selectionFromMatches(matches)');
    expect(source).not.toContain("action === 'conditional-format') inst.openCfRulesDialog()");
    expect(source).toContain('const comments = listComments(inst.store.getState())');
    expect(source).toContain('selectionFromMatches(comments.map((entry) => entry.addr))');
    expect(source).toContain('data-dropdown-name="deleteCommentReview"');
    expect(source).toContain("onCommentAction('delete-active')");
    expect(source).toContain("onCommentAction('delete-all')");
  });

  it('wires Review spelling, translation, and accessibility buttons to built-in reports when host callbacks are absent', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-ribbon-command="spellingReview"');
    expect(source).toContain('data-ribbon-command="translateReview"');
    expect(source).toContain('data-ribbon-command="accessibility"');
    expect(source).toContain(
      'analyzeSpellingCells(reviewCellsFromState(inst.store.getState()), lang.value)',
    );
    expect(source).toContain(
      'analyzeAccessibilityCells(reviewCellsFromState(inst.store.getState()), lang.value)',
    );
    expect(source).toContain('ribbonReportDialog.value = { title, items }');
    expect(source).toContain('v-if="ribbonReportDialog"');
    expect(source).toContain('const onTranslateReview = (): void =>');
    expect(source).toContain('buildTranslationReviewItems(');
    expect(source).toContain(
      'reviewCellsFromState(state, state.data.sheetIndex, state.selection.range)',
    );
    expect(source).toContain('@click="onSpellingReview"');
    expect(source).toContain('@click="onTranslateReview"');
    expect(source).toContain('@click="onAccessibilityReview"');
    expect(source).toContain(':disabled="disabled && !props.onSpellingReview"');
    expect(source).toContain(':disabled="disabled && !props.onTranslate"');
    expect(source).toContain(':disabled="disabled && !props.onAccessibilityCheck"');
  });

  it('wires Review comments to dialog, delete, and previous/next navigation commands', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-ribbon-command="newCommentReview"');
    expect(source).toContain('props.instance?.openCommentDialog()');
    expect(source).toContain('data-ribbon-command="deleteCommentReview"');
    expect(source).toContain('data-dropdown-name="deleteCommentReview"');
    expect(source).toContain("onCommentAction('delete-active')");
    expect(source).toContain("onCommentAction('delete-all')");
    expect(source).toContain('commentAt(state, state.selection.active)');
    expect(source).toContain('recordFormatChange(inst.history, inst.store, () =>');
    expect(source).toContain(
      'for (const entry of comments) clearComment(inst.store, entry.addr, inst.workbook)',
    );
    expect(source).toContain('data-ribbon-command="previousCommentReview"');
    expect(source).toContain('data-ribbon-command="nextCommentReview"');
    expect(source).toContain('@click="onSelectComment(-1)"');
    expect(source).toContain('@click="onSelectComment(1)"');
    expect(source).toContain('const comments = listComments(state)');
    expect(source).toContain('(current + direction + comments.length) % comments.length');
    expect(source).toContain('if (next) mutators.setActive(inst.store, next)');
  });

  it('wires Review and View protection buttons to shared sheet protection state', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-ribbon-command="protectReview"');
    expect(source).toContain('data-ribbon-command="protectWorkbookReview"');
    expect(source).toContain('data-ribbon-command="protect"');
    expect(source).toContain(':class="{ \'demo__rb--active\': active.protected }"');
    expect(source).toContain('@click="props.instance?.toggleSheetProtection()"');
    expect(source).toContain(':class="{ \'demo__rb--active\': workbookStructureProtected }"');
    expect(source).toContain('@click="onBackstageProtectWorkbook"');
    expect(source).toContain('cellText.protectWorkbookCommand');
    expect(source).toContain('cellText.unprotectWorkbookCommand');
    expect(source).toContain('data-dropdown-name="protectionReview"');
    expect(source).toContain("onProtectionAction('allow-edit-range')");
    expect(source).toContain("onProtectionAction('clear-allowed-edit-ranges')");
    expect(source).toContain('addAllowedEditRange(inst.store, range, { title: rangeText })');
    expect(source).toContain('clearAllowedEditRanges(inst.store, state.data.sheetIndex)');
    expect(source).toContain('{{ active.protected ? tr.unprotect : tr.protect }}');
  });

  it('wires Review Find to the find tab of the Find and Replace dialog', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-ribbon-command="findReview"');
    expect(source).toContain("props.instance?.openFindReplace('find')");
    expect(source).toContain(':aria-keyshortcuts="keyShortcuts(\'findReview\')"');
  });

  it('wires Automate script to the built-in selected-range text script when callback is absent', () => {
    const source = readToolbarSource();

    expect(source).toContain('const onRunScript = (): void =>');
    expect(source).toContain("scriptDialog.value = { command: 'uppercase' }");
    expect(source).toContain('const applyScriptDialog = (): void =>');
    expect(source).toContain('v-model="scriptDialog.command"');
    expect(source).not.toContain('window.prompt(cellText.value.scriptCommandPrompt)');
    expect(source).not.toContain('parseScriptCommand(raw)');
    expect(source).toContain('applyTextScriptToRange(');
    expect(source).toContain('draft.command');
    expect(source).toContain(
      'lastAutomationRun.value = { command: draft.command, range: formatA1Range(range), changed }',
    );
    expect(source).toContain('const recordedDetail = lastAutomationRun.value');
    expect(source).toContain('cellText.value.automationRunDetail');
    expect(source).toContain('cellText.value.automationRunStatus.replace(');
    expect(source).toContain(
      'mutators.replaceCells(inst.store, inst.workbook.cells(state.data.sheetIndex))',
    );
    expect(source).toContain('@click="onRunScript"');
    expect(source).toContain(':disabled="disabled && !props.onRunScript"');
  });

  it('wires Draw pen and eraser to built-in border draw modes when callbacks are absent', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-ribbon-command="drawPen"');
    expect(source).toContain('data-ribbon-command="drawGrid"');
    expect(source).toContain('data-ribbon-command="drawErase"');
    expect(source).toContain(
      "props.instance?.borderDraw?.activate('draw', borderStyle.value, borderColor.value)",
    );
    expect(source).toContain(
      "props.instance?.borderDraw?.activate('grid', borderStyle.value, borderColor.value)",
    );
    expect(source).toContain('props.instance?.borderDraw?.setStyle(next)');
    expect(source).toContain('props.instance?.borderDraw?.setColor(value)');
    expect(source).toContain("props.instance?.borderDraw?.activate('erase')");
    expect(source).toContain('@click="onDrawPen"');
    expect(source).toContain('@click="onDrawGrid"');
    expect(source).toContain('@click="onDrawEraser"');
    expect(source).toContain('data-ribbon-command="drawBorderGrid"');
    expect(source).toContain(':disabled="!props.onDrawPen && !props.instance?.borderDraw"');
    expect(source).toContain(':disabled="!props.onDrawEraser && !props.instance?.borderDraw"');
  });

  it('routes Insert tab dialog commands to workbook dialog methods', () => {
    const source = readToolbarSource();

    const expected = [
      ['linksInsert', 'props.instance?.openExternalLinksDialog()'],
      ['commentInsert', 'props.instance?.openCommentDialog()'],
      ['linksData', 'props.instance?.openExternalLinksDialog()'],
    ] as const;

    for (const [command, handler] of expected) {
      expect(source).toContain(`data-ribbon-command="${command}"`);
      expect(source).toContain(handler);
    }
    expect(source).toContain('data-dropdown-name="pivotTableInsert"');
    expect(source).toContain('data-dropdown-name="definedNamesInsert"');
    expect(source).toContain("onDefinedNameAction('define')");
    expect(source).toContain("onDefinedNameAction('createTopRow')");
    expect(source).toContain("onDefinedNameAction('manager')");
    expect(source).toContain('data-dropdown-name="hyperlinkInsert"');
    expect(source).toContain("onHyperlinkAction('edit')");
    expect(source).toContain("onHyperlinkAction('open')");
    expect(source).toContain("onHyperlinkAction('clear')");
    expect(source).toContain("onHyperlinkAction('external')");
    expect(source).toContain('hyperlinkAt(state, state.selection.active)');
    expect(source).toContain('clearHyperlink(inst.store, state.selection.active, inst.workbook)');
    expect(source).toContain("onPivotTableAction('dialog')");
    expect(source).toContain("onPivotTableAction('recommended')");
    expect(source).toContain('strings.value.workbookObjects.compatibilityDetails.pivotAuthoring');
    expect(source).toContain('data-dropdown-name="pictureInsert"');
    expect(source).toContain('data-dropdown-name="shapesInsert"');
    expect(source).toContain('data-dropdown-name="screenshotInsert"');
    expect(source).toContain('onIllustrationAction(cellText.shapeRoundedRectangle)');
    expect(source).toContain('strings.value.workbookObjects.compatibilityDetails.chartsDrawings');
  });

  it('renders AutoSum as a dropdown on Home and Formulas tabs', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-dropdown-name="autosum"');
    expect(source).toContain('data-dropdown-name="autosumFormula"');
    expect(source).toContain("onAutoSum('SUM')");
    expect(source).toContain("onAutoSum('AVERAGE')");
    expect(source).toContain("onAutoSum('COUNT')");
    expect(source).toContain("onAutoSum('MAX')");
    expect(source).toContain("onAutoSum('MIN')");
    expect(source).toContain("onAutoSum('MORE')");
    expect(source).toContain('cellText.autosumMoreFunctions');
    // Auto-sum dispatch (including history.begin/end and result handling)
    // moved to the shared `handleAutoSumAction` helper in core. The Vue
    // template just hands the function name to the helper.
    expect(source).toContain('handleAutoSumAction(props.instance, functionName)');
    expect(source).toContain('closeDropdown();');
  });

  it('wires View sheet views to save, activate, and delete state handlers', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-ribbon-command="sheetViewSelect"');
    expect(source).toContain('const onSheetViewSelect = (value: string): void =>');
    expect(source).toContain('activateSheetView(inst.store, value)');
    expect(source).toContain('data-ribbon-command="sheetViewSave"');
    expect(source).toContain(
      'saveSheetView(inst.store, id, `${viewToolbarText.value.views} ${count}`)',
    );
    expect(source).toContain('sheetViews: { ...state.sheetViews, activeViewId: id }');
    expect(source).toContain('data-ribbon-command="sheetViewDelete"');
    expect(source).toContain('deleteSheetView(inst.store, id)');
    expect(source).toContain(':disabled="disabled || activeSheetViewId === \'current\'"');
  });

  it('routes Formulas function-library buttons to the matching function argument helper', () => {
    const source = readToolbarSource();

    const expected = [
      ['fx', 'props.instance?.openFunctionArguments()'],
      ['sum', "props.instance?.openFunctionArguments('SUM')"],
      ['avg', "props.instance?.openFunctionArguments('AVERAGE')"],
    ] as const;

    for (const [command, handler] of expected) {
      expect(source).toContain(`data-ribbon-command="${command}"`);
      expect(source).toContain(handler);
    }
    for (const command of [
      'ifFormula',
      'xlookupFormula',
      'concatFormula',
      'todayFormula',
      'pmtFormula',
      'roundFormula',
    ]) {
      expect(source).toContain(`data-dropdown-name="${command}"`);
    }
    expect(source).toContain('tr.functionLogical');
    expect(source).toContain('tr.functionLookupReference');
    expect(source).toContain('tr.functionText');
    expect(source).toContain('tr.functionDateTime');
    expect(source).toContain('tr.functionFinancial');
    expect(source).toContain('tr.functionMathTrig');
    expect(source).toContain('onFunctionAction(name as FunctionAction)');
    expect(source).toContain("['XLOOKUP', 'VLOOKUP', 'INDEX', 'MATCH']");
    expect(source).toContain("['ROUND', 'SUMIF', 'COUNTIF', 'ABS']");
  });

  it('renders Defined Names as a menu with create and use actions', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-dropdown-name="definedNames"');
    expect(source).toContain("onDefinedNameAction('define')");
    expect(source).toContain("onDefinedNameAction('createTopRow')");
    expect(source).toContain("onDefinedNameAction('createBottomRow')");
    expect(source).toContain("onDefinedNameAction('createLeftColumn')");
    expect(source).toContain("onDefinedNameAction('createRightColumn')");
    expect(source).toContain("'bottom-row'");
    expect(source).toContain("'right-column'");
    expect(source).toContain("onDefinedNameAction('manager')");
    expect(source).toContain('createDefinedNamesFromSelection(');
    expect(source).toContain('insertDefinedNameFormula(');
    expect(source).toContain('recordDefinedNamesChange(inst.history, inst.workbook');
    expect(source).toContain("action.slice('use:'.length),\n      inst.store");
  });

  it('renders Calculation Options as an Excel-style calc mode menu', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-ribbon-command="recalcNow"');
    expect(source).toContain('props.instance?.recalc()');
    expect(source).toContain('data-dropdown-name="calcOptions"');
    expect(source).toContain("onCalculationAction('auto')");
    expect(source).toContain("onCalculationAction('autoNoTable')");
    expect(source).toContain("onCalculationAction('manual')");
    expect(source).toContain("onCalculationAction('iterative')");
    expect(source).toContain('data-dropdown-name="watch"');
    expect(source).toContain('data-dropdown-name="watchView"');
    expect(source).toContain("onWatchAction('open')");
    expect(source).toContain("onWatchAction('add')");
    expect(source).toContain("onWatchAction('delete')");
    expect(source).toContain("onWatchAction('delete-all')");
    expect(source).toContain('recordWatchesChange(inst.history, inst.store');
    expect(source).toContain('watchRange(inst.store, state.selection.range)');
    expect(source).toContain('unwatchCell(inst.store, state.selection.active)');
    expect(source).toContain('clearWatchedCells(inst.store)');
    expect(source).toContain("activeCalcAction === 'manual'");
    expect(source).toContain('role="menuitemradio"');
    expect(source).toContain('inst.workbook.setCalcMode(mode)');
    expect(source).toContain('active.value = projectActiveState(inst)');
    expect(source).toContain('inst.openIterativeDialog()');
  });

  it('renders Formula Auditing Show Formulas as a view toggle', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-ribbon-command="showFormulasFormula"');
    expect(source).toContain("'demo__rb--active': active.formulasVisible");
    expect(source).toContain("onViewFlag('formulas')");
    expect(source).toContain('{{ viewText.formulas }}');
  });

  it('wires View > Show commands to worksheet display flags', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-ribbon-command="viewGridlines"');
    expect(source).toContain('data-ribbon-command="viewHeadings"');
    expect(source).toContain('data-ribbon-command="viewFormulas"');
    expect(source).toContain('data-ribbon-command="viewR1C1"');
    expect(source).toContain("onViewFlag('gridlines')");
    expect(source).toContain("onViewFlag('headings')");
    expect(source).toContain("onViewFlag('formulas')");
    expect(source).toContain("onViewFlag('r1c1')");
    expect(source).toContain('setGridlinesVisible(inst.store, ui.showGridLines === false)');
    expect(source).toContain('setHeadingsVisible(inst.store, ui.showHeaders === false)');
    expect(source).toContain('setShowFormulas(inst.store, !ui.showFormulas)');
    expect(source).toContain('setR1C1ReferenceStyle(inst.store, !ui.r1c1)');
  });

  it('renders Formula Auditing trace arrows as direct instance commands', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-ribbon-command="precedents"');
    expect(source).toContain('data-ribbon-command="dependents"');
    expect(source).toContain('data-ribbon-command="clearArrows"');
    expect(source).toContain('data-dropdown-name="clearArrows"');
    expect(source).toContain('props.instance?.tracePrecedents()');
    expect(source).toContain('props.instance?.traceDependents()');
    expect(source).toContain("onClearArrowsAction('clear-all')");
    expect(source).toContain("onClearArrowsAction('clear-precedents')");
    expect(source).toContain("onClearArrowsAction('clear-dependents')");
    expect(source).toContain("clearTraceArrowsByKind(inst.store, 'precedent', inst.history)");
    expect(source).toContain("clearTraceArrowsByKind(inst.store, 'dependent', inst.history)");
    expect(source).toContain('clearTraceArrows(inst.store, inst.history)');
    expect(source).toContain('{{ tr.tracePrecedents }}');
    expect(source).toContain('{{ tr.traceDependents }}');
    expect(source).toContain('{{ tr.removeArrows }}');
  });

  it('renders Formula Auditing Error Checking as an Excel-style menu', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-dropdown-name="errorChecking"');
    expect(source).toContain("onFormulaAuditingAction('errorChecking')");
    expect(source).toContain("onFormulaAuditingAction('traceError')");
    expect(source).toContain("onFormulaAuditingAction('ignoreError')");
    expect(source).toContain('cellValueIsFormulaError(activeCell.value)');
    expect(source).toContain('recordIgnoredErrorsChange(inst.history, inst.store');
    expect(source).toContain('ignoreCellError(inst.store, state.selection.active)');
    expect(source).toContain("onFormulaAuditingAction('circleInvalid')");
    expect(source).toContain("onFormulaAuditingAction('clearCircles')");
    expect(source).toContain('selectNextFormulaError(inst.store)');
  });

  it('renders Formula Auditing Evaluate Formula as a dialog command', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-ribbon-command="evaluateFormula"');
    expect(source).toContain('props.instance?.openEvaluateFormulaDialog()');
    expect(source).toContain('{{ tr.evaluateFormula }}');
  });

  it('renders Data Filter as a clear/reapply/advanced menu', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-dropdown-name="dataFilter"');
    expect(source).toContain("onFilterDataAction('toggle')");
    expect(source).toContain("onFilterDataAction('clear')");
    expect(source).toContain("onFilterDataAction('reapply')");
    expect(source).toContain("onFilterDataAction('filter-by-selected')");
    expect(source).toContain("onFilterDataAction('advanced')");
    expect(source).toContain('recordFilterChange(inst.history, inst.store');
    expect(source).toContain('clearFilter(s, inst.store, s.ui.filterRange ?? undefined)');
    expect(source).toContain('reapplyFilters(inst.store.getState(), inst.store)');
    expect(source).toContain('filterBySelectedCellValue(inst.store.getState(), inst.store)');
    expect(source).toContain('advancedFilterDialog');
    expect(source).toContain('cellText.advancedFilterDialogTitle');
    expect(source).toContain('v-model="advancedFilterDialog.listRange"');
    expect(source).toContain('v-model="advancedFilterDialog.criteriaRange"');
    expect(source).toContain('v-model="advancedFilterDialog.copyTo"');
    expect(source).toContain('v-model="advancedFilterDialog.uniqueOnly"');
    expect(source).toContain('cellText.advancedFilterCopyTo');
    expect(source).toContain('cellText.advancedFilterUniqueOnly');
    expect(source).toContain('copyAdvancedFilterResult(');
    expect(source).toContain('uniqueOnly: draft.uniqueOnly');
    expect(source).toContain('cellText.value.advancedFilterCopiedStatus');
    expect(source).toContain(
      'applyAdvancedFilter(inst.store.getState(), inst.store, listRange, criteriaRange)',
    );
  });

  it('renders Data Sort as a custom sort menu', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-dropdown-name="sortData"');
    expect(source).toContain('data-ribbon-command="sortData"');
    expect(source).toContain("onSortMenuAction('custom')");
    expect(source).toContain("onSortMenuAction('asc')");
    expect(source).toContain("onSortMenuAction('desc')");
    expect(source).toContain('cellText.sortCustom');
  });

  it('renders Text to Columns as a delimiter menu with a custom dialog', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-ribbon-command="textToColumns"');
    expect(source).toContain('data-dropdown-name="textToColumns"');
    expect(source).toContain("onTextToColumnsAction('comma')");
    expect(source).toContain("onTextToColumnsAction('tab')");
    expect(source).toContain("onTextToColumnsAction('semicolon')");
    expect(source).toContain("onTextToColumnsAction('space')");
    expect(source).toContain("onTextToColumnsAction('custom')");
    expect(source).toContain('textToColumnsDialog');
    expect(source).toContain('cellText.textToColumnsCustom');
    expect(source).toContain('cellText.textToColumnsDialogTitle');
    expect(source).toContain('cellText.textToColumnsDialogDelimiters');
    expect(source).toContain('cellText.textToColumnsTreatConsecutive');
    expect(source).toContain('v-model="textToColumnsDialog.comma"');
    expect(source).toContain('v-model="textToColumnsDialog.semicolon"');
    expect(source).toContain('v-model="textToColumnsDialog.collapseConsecutive"');
    expect(source).toContain('inst.history.begin()');
    expect(source).toContain('inst.history.end()');
    expect(source).toContain(
      'textToColumns(state, inst.store, inst.workbook, state.selection.range, delimiters, {',
    );
    expect(source).toContain('collapseConsecutiveDelimiters: collapseConsecutive');
  });

  it('uses the live instance i18n strings when the toolbar locale matches', () => {
    const source = readToolbarSource();

    expect(source).toContain("import { useI18n } from './composables.js'");
    expect(source).toContain('const liveI18n = useI18n(instanceRef)');
    expect(source).toContain(
      'const liveLang = computed(() => dictionaryLocaleFor(liveI18n.locale.value))',
    );
    expect(source).toContain(
      'liveI18n.locale.value === props.locale || liveLang.value === lang.value',
    );
    expect(source).toContain('liveI18n.strings.value');
  });

  it('wires Data outline detail commands to explicit show/hide handlers', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-dropdown-name="outlineGroup"');
    expect(source).toContain('data-dropdown-name="outlineUngroup"');
    expect(source).toContain("onOutlineAction('group', 'rows')");
    expect(source).toContain("onOutlineAction('group', 'cols')");
    expect(source).toContain("onOutlineAction('ungroup', 'rows')");
    expect(source).toContain("onOutlineAction('ungroup', 'cols')");
    expect(source).toContain('data-ribbon-command="outlineUngroup"');
    expect(source).toContain('data-ribbon-command="outlineShowDetail"');
    expect(source).toContain('data-ribbon-command="outlineHideDetail"');
    expect(source).toContain("onOutlineAction('show-detail')");
    expect(source).toContain("onOutlineAction('hide-detail')");
    expect(source).toContain("action === 'show-detail') showRows");
    expect(source).toContain("action === 'show-detail') showCols");
    expect(source).toContain('else collapseRowGroup');
    expect(source).toContain('else collapseColGroup');
  });

  it('renders Data Validation as a validation actions menu', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-dropdown-name="dataValidation"');
    expect(source).toContain("onDataValidationAction('settings')");
    expect(source).toContain("onDataValidationAction('circleInvalid')");
    expect(source).toContain("onDataValidationAction('clearCircles')");
    expect(source).toContain("onDataValidationAction('clearValidation')");
    expect(source).toContain(
      'clearValidationInRangeWithEngine(inst.store, inst.history, inst.workbook',
    );
    expect(source).toContain('recordValidationCirclesChange(inst.history, inst.store');
    expect(source).toContain('circleInvalidValidationDataInSheet(');
    expect(source).toContain('state.selection.range.sheet');
    expect(source).toContain('clearValidationCircles(inst.store)');
    expect(source).toContain('makeRangeResolver(inst.workbook, state.data.sheetIndex)');
  });

  it('renders Cell Styles as an applying ribbon menu', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-dropdown-name="cellStyles"');
    expect(source).toContain('CELL_STYLE_GROUPS');
    expect(source).toContain('cellStyleGroups');
    expect(source).toContain('{{ group.label }}');
    expect(source).toContain('active.cellStyle !== null');
    expect(source).toContain('active.cellStyle === styleId');
    expect(source).toContain('role="menuitemradio"');
    expect(source).toContain('v-for="group in cellStyleGroups"');
    expect(source).toContain('onCellStyleAction(styleId as CellStyleAction)');
    expect(source).toContain('applyCellStyle(inst.store, inst.history, r, action)');
  });

  it('renders Format as Table as style-picking menus', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-dropdown-name="formatTableHome"');
    expect(source).toContain('data-dropdown-name="formatTableInsert"');
    expect(source).toContain('active.formatAsTable');
    expect(source).toContain("onFormatAsTable('light')");
    expect(source).toContain("onFormatAsTable('medium')");
    expect(source).toContain("onFormatAsTable('dark')");
    expect(source).toContain('recordTablesChange(inst.history, inst.store');
    expect(source).toContain('formatAsTable(inst.store, r, { style })');
  });

  it('renders Page Layout print-area as a set/clear menu', () => {
    const source = readToolbarSource();

    expect(source).toContain('data-dropdown-name="pageTheme"');
    expect(source).toContain("onThemeAction('paper')");
    expect(source).toContain("onThemeAction('ink')");
    expect(source).toContain("onThemeAction('contrast')");
    expect(source).toContain('inst.setTheme(action)');
    expect(source).toContain('cellText.themeContrast');
    expect(source).toContain('data-dropdown-name="printArea"');
    expect(source).toContain("onPrintAreaAction('set')");
    expect(source).toContain("onPrintAreaAction('clear')");
    expect(source).toContain('setPrintArea(inst.store, sheet');
    expect(source).toContain('clearPrintArea(inst.store, sheet)');
    expect(source).toContain('data-dropdown-name="pageBreaks"');
    expect(source).toContain("onPageBreakAction('insert-row')");
    expect(source).toContain("onPageBreakAction('insert-col')");
    expect(source).toContain("onPageBreakAction('reset')");
    expect(source).toContain('insertManualPageBreak(inst.store, sheet');
    expect(source).toContain('resetManualPageBreaks(inst.store, sheet)');
    expect(source).toContain('data-dropdown-name="sheetBackground"');
    expect(source).toContain("onSheetBackgroundAction('set')");
    expect(source).toContain("onSheetBackgroundAction('clear')");
    expect(source).toContain('setSheetBackgroundImage(inst.store, sheet');
    expect(source).toContain('clearSheetBackgroundImage(inst.store, sheet, inst.history)');
    expect(source).toContain('data-ribbon-file-input="sheetBackground"');
    expect(source).toContain('readAsDataURL(file)');
    expect(source).toContain('data-dropdown-name="printTitles"');
    expect(source).toContain("onPrintTitleAction('rows')");
    expect(source).toContain("onPrintTitleAction('cols')");
    expect(source).toContain("onPrintTitleAction('clear')");
    expect(source).toContain('setPrintTitleRows(inst.store, sheet');
    expect(source).toContain('setPrintTitleCols(inst.store, sheet');
    expect(source).toContain('clearPrintTitles(inst.store, sheet)');
    expect(source).toContain('data-dropdown-name="scaleWidth"');
    expect(source).toContain("onScaleFit('width', value)");
    expect(source).toContain('data-dropdown-name="scalePercent"');
    expect(source).toContain('setPageScale(inst.store, sheet');
    expect(source).toContain('data-ribbon-command="pageLayoutGridlinesView"');
    expect(source).toContain('data-ribbon-command="pageLayoutGridlinesPrint"');
    expect(source).toContain("onPrintSheetOption('gridlines')");
    expect(source).toContain("onPrintSheetOption('headings')");
    expect(source).toContain(
      'setPrintGridlines(inst.store, sheet, !active.value.printGridlines, inst.history)',
    );
    expect(source).toContain(
      'setPrintHeadings(inst.store, sheet, !active.value.printHeadings, inst.history)',
    );
    expect(source).toContain('data-ribbon-command="printPageLayout"');
    expect(source).toContain('@click="props.instance?.print(\'print\')"');
  });
});

describe('Vue <SpreadsheetToolbar> building blocks', () => {
  let mounted: MountedVueSpreadsheet | null = null;
  let probe: RibbonProbeHandle | null = null;

  afterEach(async () => {
    if (probe) {
      await probe.unmount();
      probe = null;
    }
    if (mounted) {
      await mounted.dispose();
      mounted = null;
    }
    document.body.replaceChildren();
  });

  it('useToolbarActive seeds from a live instance and reflects bold toggles', async () => {
    mounted = await mountVueSpreadsheet();
    const log: DropdownLogEntry[] = [];
    probe = await mountRibbonProbe(mounted.instance, log);

    expect(probe.active.value.bold).toBe(false);
    expect(probe.host.querySelector('[data-testid="bold"]')?.textContent).toBe('false');

    // Toggle bold via the same command the toolbar would invoke.
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 0, col: 0 });
    toggleBold(mounted.instance.store.getState(), mounted.instance.store);
    await flush();

    expect(projectActiveState(mounted.instance).bold).toBe(true);
    expect(probe.active.value.bold).toBe(true);
    expect(probe.host.querySelector('[data-testid="bold"]')?.textContent).toBe('true');
  });

  it('useToolbarDropdown opens, picks, and routes the right handler per dropdown name', async () => {
    mounted = await mountVueSpreadsheet();
    const log: DropdownLogEntry[] = [];
    probe = await mountRibbonProbe(mounted.instance, log);

    // Initially closed.
    expect(probe.openDropdown.value).toBeNull();

    probe.toggleDropdown('fontFamily');
    await flush();
    expect(probe.openDropdown.value).toBe('fontFamily');

    probe.pickDropdown('fontFamily', 'Calibri');
    await flush();
    expect(log).toContainEqual({ kind: 'fontFamily', value: 'Calibri' });
    // Picking auto-closes.
    expect(probe.openDropdown.value).toBeNull();

    // borderStyle is local state plus a brush-sync callback for active border drawing.
    probe.toggleDropdown('borderStyle');
    await flush();
    expect(probe.openDropdown.value).toBe('borderStyle');
    probe.pickDropdown('borderStyle', 'thick');
    await flush();
    expect(log).toContainEqual({ kind: 'borderStyle', value: 'thick' });
    expect(probe.openDropdown.value).toBeNull();

    // margins=custom is a special-case redirect to onOpenPageSetup.
    probe.toggleDropdown('margins');
    probe.pickDropdown('margins', 'custom');
    await flush();
    expect(log).toContainEqual({ kind: 'openPageSetup', value: null });
    expect(log.find((e) => e.kind === 'marginPreset')).toBeUndefined();

    // margins=normal routes to onMarginPreset.
    probe.toggleDropdown('margins');
    probe.pickDropdown('margins', 'normal');
    await flush();
    expect(log).toContainEqual({ kind: 'marginPreset', value: 'normal' });
  });

  it('useToolbarDropdown closes the open dropdown on Escape and on outside click', async () => {
    mounted = await mountVueSpreadsheet();
    const log: DropdownLogEntry[] = [];
    probe = await mountRibbonProbe(mounted.instance, log);

    probe.toggleDropdown('fontSize');
    await flush();
    expect(probe.openDropdown.value).toBe('fontSize');

    // Escape from anywhere closes the dropdown.
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await flush();
    expect(probe.openDropdown.value).toBeNull();

    // Re-open and verify outside pointerdown closes too.
    probe.toggleDropdown('paperSize');
    await flush();
    expect(probe.openDropdown.value).toBe('paperSize');

    const outside = document.createElement('div');
    document.body.appendChild(outside);
    outside.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    await flush();
    expect(probe.openDropdown.value).toBeNull();
    outside.remove();
  });

  it('useToolbarDropdown handles list keyboard navigation, pick, and focus return', async () => {
    mounted = await mountVueSpreadsheet();
    const log: DropdownLogEntry[] = [];
    probe = await mountRibbonProbe(mounted.instance, log);

    const root = document.createElement('div');
    root.className = 'demo__rb-dd';
    root.dataset.dropdownName = 'fontFamily';
    root.innerHTML = `
      <button class="demo__rb-dd__btn" type="button" aria-expanded="false">Font</button>
      <div class="demo__rb-dd__list" role="listbox">
        <button class="demo__rb-dd__opt" type="button" role="option" aria-selected="true">Aptos</button>
        <button class="demo__rb-dd__opt" type="button" role="option" aria-selected="false">Calibri</button>
        <button class="demo__rb-dd__opt" type="button" role="option" aria-selected="false">Consolas</button>
      </div>
    `;
    document.body.appendChild(root);
    root.addEventListener('keydown', probe.keydownDropdown);
    root.classList.add('demo__rb-dd--open');

    const button = root.querySelector<HTMLButtonElement>('.demo__rb-dd__btn');
    const options = root.querySelectorAll<HTMLButtonElement>('[role="option"]');
    expect(button).toBeTruthy();

    probe.toggleDropdown('fontFamily');
    await flush();
    options[0]?.focus();
    options[0]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    await flush();
    expect(document.activeElement).toBe(options[1]);

    options[1]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
    await flush();
    expect(document.activeElement).toBe(options[2]);

    options[2]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await flush();
    expect(probe.openDropdown.value).toBeNull();
    expect(document.activeElement).toBe(button);
    root.removeEventListener('keydown', probe.keydownDropdown);
    root.remove();
  });

  it('useToolbarActive cleans up its store subscription on unmount', async () => {
    mounted = await mountVueSpreadsheet();
    const log: DropdownLogEntry[] = [];
    probe = await mountRibbonProbe(mounted.instance, log);

    const beforeBold = probe.active.value.bold;
    expect(beforeBold).toBe(false);

    await probe.unmount();
    probe = null;

    // After unmount, store changes must not panic the (gone) subscriber.
    const errSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 0, col: 0 });
    toggleBold(mounted.instance.store.getState(), mounted.instance.store);
    await flush();
    expect(errSpy).not.toHaveBeenCalled();
    errSpy.mockRestore();
  });

  it('useToolbarDropdown stops listening on document events after unmount', async () => {
    mounted = await mountVueSpreadsheet();
    const log: DropdownLogEntry[] = [];
    probe = await mountRibbonProbe(mounted.instance, log);

    probe.toggleDropdown('fontFamily');
    await flush();
    expect(probe.openDropdown.value).toBe('fontFamily');

    await probe.unmount();
    probe = null;

    // Escape after unmount should NOT mutate any captured state — and must
    // not throw. We exercise the path simply by dispatching the event; if
    // the listener wasn't removed we'd see a console error from Vue trying
    // to update an unmounted ref.
    const errSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await flush();
    expect(errSpy).not.toHaveBeenCalled();
    errSpy.mockRestore();
  });
});
