import {
  addSheet,
  applyCellStyle,
  type CellStyleId,
  clearHyperlink,
  createDefinedNamesFromSelection,
  createPivotTableFromRange,
  formatAsTable,
  hyperlinkAt,
  inferPivotSourceFields,
  insertDefinedNameFormula,
  isWorkbookStructureProtected,
  listDefinedNames,
  mutators,
  PivotAggregation,
  type PivotSourceField,
  type Range,
  recordDefinedNamesChange,
  recordFormatChange,
  recordTablesChange,
  type SpreadsheetInstance,
  setNumFmt,
  type TableStyle,
  type TableVariantId,
  type ToolbarMenuText,
  type ToolbarText,
  tableVariantOptions,
} from '@libraz/formulon-cell';
import { showChoiceDialog, showFormatAsTableDialog, showMessage } from './dialogs.js';

export type { TableVariantId };

type RecommendedPivotSpec = {
  rowField: string;
  columnField?: string;
  valueField: string;
  aggregation: PivotAggregation;
  placement: 'existing' | 'new-sheet';
};

export interface WorkbookActionsCtx {
  getInst: () => SpreadsheetInstance | null;
  ribbonLang: 'ja' | 'en';
  ribbonText: ToolbarText;
  ribbonMenuText: ToolbarMenuText;
  refreshWorkbookCells: () => void;
  focusSheet: () => void;
  renderSheetTabs: () => void;
  switchSheet: (idx: number) => void;
  applyRibbonFormat: (
    fn: (
      state: ReturnType<SpreadsheetInstance['store']['getState']>,
      store: SpreadsheetInstance['store'],
    ) => void,
  ) => void;
  sortTargetRange: (state: ReturnType<SpreadsheetInstance['store']['getState']>) => Range;
  rangeRef: (range: Range) => string;
  parseA1Range: (raw: string, sheet: number) => Range | null;
  getStatusMetric: () => HTMLElement | null;
}

export interface WorkbookActionsApi {
  applyCellStyleFromRibbon: (id: CellStyleId) => void;
  applyCurrencyPreset: (symbol: string) => void;
  openCurrencyFooterAction: (action: string) => void;
  openCellStyleFooterAction: (action: string) => Promise<void>;
  openTableStyleFooterAction: (action: string) => Promise<void>;
  createTableFromSelection: (
    style?: TableStyle,
    color?: string,
    variant?: TableVariantId,
  ) => Promise<void>;
  applyPivotTableAction: (action: string) => void;
  applyDefinedNameAction: (action: string) => void;
  clearHyperlinksInSelection: (mode?: 'clear' | 'remove') => void;
  applyLinksAction: (action: string) => void;
}

export const createWorkbookActions = (ctx: WorkbookActionsCtx): WorkbookActionsApi => {
  const {
    getInst,
    ribbonLang,
    ribbonText,
    ribbonMenuText,
    refreshWorkbookCells,
    focusSheet,
    renderSheetTabs,
    switchSheet,
    applyRibbonFormat,
    sortTargetRange,
    rangeRef,
    parseA1Range,
    getStatusMetric,
  } = ctx;

  const applyCellStyleFromRibbon = (id: CellStyleId): void => {
    const i = getInst();
    if (!i) return;
    const range = i.store.getState().selection.range;
    applyCellStyle(i.store, i.history, range, id);
    refreshWorkbookCells();
    focusSheet();
  };

  const applyCurrencyPreset = (symbol: string): void => {
    applyRibbonFormat((state, store) =>
      setNumFmt(state, store, { kind: 'currency', decimals: 2, symbol }),
    );
  };

  const openCurrencyFooterAction = (action: string): void => {
    if (action === 'more') {
      getInst()?.openFormatDialog();
    }
  };

  const openCellStyleFooterAction = async (action: string): Promise<void> => {
    const ja = ribbonLang === 'ja';
    if (action === 'new-cell-style') {
      await showMessage({
        title: ja ? '新しいセルのスタイル' : 'New Cell Style',
        message: ja
          ? 'カスタム セル スタイルの作成は今後のリリースで対応予定です。'
          : 'Authoring custom cell styles is coming in a future release.',
      });
      focusSheet();
      return;
    }
    if (action === 'merge-cell-style') {
      await showMessage({
        title: ja ? 'スタイルの結合' : 'Merge Styles',
        message: ja
          ? '他のブックのスタイル結合は今後のリリースで対応予定です。'
          : 'Merging styles from another workbook is coming in a future release.',
      });
      focusSheet();
    }
  };

  const openTableStyleFooterAction = async (action: string): Promise<void> => {
    const ja = ribbonLang === 'ja';
    if (action === 'new-table-style') {
      await showMessage({
        title: ja ? '新しい表スタイル' : 'New Table Style',
        message: ja
          ? 'カスタム表スタイルの作成は今後のリリースで対応予定です。'
          : 'Authoring custom table styles is coming in a future release.',
      });
      focusSheet();
      return;
    }
    if (action === 'new-pivot-style') {
      await showMessage({
        title: ja ? '新しいピボットテーブル スタイル' : 'New PivotTable Style',
        message: ja
          ? 'カスタム ピボットテーブル スタイルの作成は今後のリリースで対応予定です。'
          : 'Authoring custom PivotTable styles is coming in a future release.',
      });
      focusSheet();
    }
  };

  const createTableFromSelection = async (
    style: TableStyle = 'medium',
    color?: string,
    variant: TableVariantId = 'banded',
  ): Promise<void> => {
    const i = getInst();
    if (!i) return;
    const state = i.store.getState();
    const result = await showFormatAsTableDialog({
      title: ribbonLang === 'ja' ? 'テーブルとして書式設定' : 'Format as Table',
      rangeLabel:
        ribbonLang === 'ja'
          ? 'テーブルに変換するデータ範囲を指定してください'
          : 'Where is the data for your table?',
      headersLabel:
        ribbonLang === 'ja' ? '先頭行をテーブルの見出しとして使用する' : 'My table has headers',
      initialRange: rangeRef(state.selection.range),
      initialHasHeaders: true,
      okLabel: 'OK',
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
      validateRange: (value) =>
        parseA1Range(value, state.selection.active.sheet)
          ? null
          : ribbonLang === 'ja'
            ? 'A1:B10 の形式で入力してください。'
            : 'Enter a range such as A1:B10.',
    });
    if (result === null) {
      focusSheet();
      return;
    }
    const r = parseA1Range(result.range, state.selection.active.sheet);
    if (!r) {
      focusSheet();
      return;
    }
    const variantOptions = tableVariantOptions(variant);
    recordTablesChange(i.history, i.store, () => {
      formatAsTable(i.store, r, {
        showHeader: result.hasHeaders,
        style,
        color,
        banded: variantOptions.banded,
        firstCol: variantOptions.firstCol,
      });
    });
    focusSheet();
  };

  const pivotSpecKey = (spec: RecommendedPivotSpec): string =>
    [spec.rowField, spec.columnField ?? '', spec.valueField, spec.aggregation, spec.placement].join(
      '',
    );

  const buildRecommendedPivotSpecs = (
    fields: readonly PivotSourceField[],
    placement: RecommendedPivotSpec['placement'],
  ): RecommendedPivotSpec[] => {
    const numeric = fields.filter((field) => field.numericCount > 0);
    const values = numeric.length > 0 ? numeric : fields;
    const categories = fields.filter((field) => field.numericCount === 0);
    const rows = categories.length > 0 ? categories : fields;
    const specs: RecommendedPivotSpec[] = [];
    const add = (rowField = rows[0], valueField = values[0], columnField?: PivotSourceField) => {
      if (!rowField || !valueField || rowField.name === valueField.name) return;
      if (
        columnField &&
        (columnField.name === rowField.name || columnField.name === valueField.name)
      )
        return;
      specs.push({
        rowField: rowField.name,
        columnField: columnField?.name,
        valueField: valueField.name,
        aggregation: valueField.numericCount > 0 ? PivotAggregation.Sum : PivotAggregation.Count,
        placement,
      });
    };
    add(
      rows[0],
      values[0],
      categories.find((field) => field.name !== rows[0]?.name),
    );
    add(rows[0], values[0]);
    add(rows[1], values[0]);
    add(rows[0], values[1] ?? values[0]);
    const seen = new Set<string>();
    return specs.filter((spec) => {
      const key = pivotSpecKey(spec);
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  };

  const pivotSpecLabel = (spec: RecommendedPivotSpec): string => {
    const valueLabel =
      spec.aggregation === PivotAggregation.Sum
        ? ribbonLang === 'ja'
          ? `合計 / ${spec.valueField}`
          : `Sum of ${spec.valueField}`
        : ribbonLang === 'ja'
          ? `データの個数 / ${spec.valueField}`
          : `Count of ${spec.valueField}`;
    const axisLabel = spec.columnField
      ? ribbonLang === 'ja'
        ? `${spec.rowField} x ${spec.columnField}`
        : `${spec.rowField} by ${spec.columnField}`
      : spec.rowField;
    return `${axisLabel} - ${valueLabel}`;
  };

  const createRecommendedPivotTable = (
    placement: 'existing' | 'new-sheet',
    sourceOverride?: Range,
    specOverride?: RecommendedPivotSpec,
  ): void => {
    const i = getInst();
    if (!i) return;
    const state = i.store.getState();
    const source = sourceOverride ?? sortTargetRange(state);
    const fields = inferPivotSourceFields(i.workbook, source);
    const spec = specOverride ?? buildRecommendedPivotSpecs(fields, placement)[0];
    if (!spec) {
      void showMessage({
        title: ribbonText.pivotTable,
        message:
          ribbonLang === 'ja'
            ? 'ピボットテーブルを作成できる見出し付きデータ範囲を選択してください。'
            : 'Select a labeled data range that can be used for a PivotTable.',
      });
      return;
    }
    let destinationSheet = source.sheet;
    if (placement === 'new-sheet') {
      const added = addSheet(i.store, i.workbook);
      if (added < 0) {
        const statusMetric = getStatusMetric();
        if (statusMetric && isWorkbookStructureProtected(i.store.getState())) {
          statusMetric.textContent = ribbonMenuText.workbookStructureProtectedBlocked;
        }
        return;
      }
      destinationSheet = added;
      renderSheetTabs();
    }
    const destination =
      placement === 'new-sheet'
        ? { sheet: destinationSheet, row: 0, col: 0 }
        : { sheet: destinationSheet, row: source.r1 + 3, col: source.c0 };
    const result = createPivotTableFromRange(i.workbook, {
      source,
      destination,
      name: `PivotTable${i.workbook.getPivotTables().length + 1}`,
      rowField: spec.rowField,
      columnField: spec.columnField,
      valueField: spec.valueField,
      aggregation: spec.aggregation,
    });
    if (!result.ok) {
      void showMessage({
        title: ribbonText.pivotTable,
        message:
          ribbonLang === 'ja'
            ? 'ピボットテーブルを作成できませんでした。'
            : 'Could not create a PivotTable from the selected range.',
      });
      return;
    }
    mutators.setActive(i.store, destination);
    if (placement === 'new-sheet') switchSheet(destinationSheet);
    else refreshWorkbookCells();
    const statusMetric = getStatusMetric();
    if (statusMetric) statusMetric.textContent = ribbonMenuText.pivotTableCreated;
    focusSheet();
  };

  const openRecommendedPivotTablesDialog = async (): Promise<void> => {
    const i = getInst();
    if (!i) return;
    const source = sortTargetRange(i.store.getState());
    const fields = inferPivotSourceFields(i.workbook, source);
    const specs = buildRecommendedPivotSpecs(fields, 'existing');
    if (specs.length === 0) {
      createRecommendedPivotTable('existing', source);
      return;
    }
    const options = specs.map((spec, index) => ({
      value: `pivot-${index}`,
      label: pivotSpecLabel(spec),
    }));
    const choice = await showChoiceDialog<string>({
      title: ribbonMenuText.recommendedPivotTables,
      label: ribbonText.pivotTable,
      initial: options[0]?.value,
      okLabel: 'OK',
      cancelLabel: ribbonLang === 'ja' ? 'キャンセル' : 'Cancel',
      options,
    });
    if (!choice) {
      focusSheet();
      return;
    }
    const index = Number(choice.replace('pivot-', ''));
    const spec = specs[index];
    if (spec) createRecommendedPivotTable(spec.placement, source, spec);
  };

  const applyPivotTableAction = (action: string): void => {
    const i = getInst();
    if (!i) return;
    if (action === 'recommended') {
      void openRecommendedPivotTablesDialog();
      return;
    }
    if (action === 'new-sheet') {
      createRecommendedPivotTable('new-sheet');
      return;
    }
    i.openPivotTableDialog();
  };

  const applyDefinedNameAction = (action: string): void => {
    const i = getInst();
    if (!i) return;
    if (action === 'define') {
      i.openDefineNameDialog();
      return;
    }
    if (action === 'manager') {
      i.openNamedRangeDialog();
      return;
    }
    if (
      action === 'create-top-row' ||
      action === 'create-bottom-row' ||
      action === 'create-left-column' ||
      action === 'create-right-column'
    ) {
      const source =
        action === 'create-top-row'
          ? 'top-row'
          : action === 'create-bottom-row'
            ? 'bottom-row'
            : action === 'create-left-column'
              ? 'left-column'
              : 'right-column';
      const result = recordDefinedNamesChange(i.history, i.workbook, () =>
        createDefinedNamesFromSelection(i.store.getState(), i.workbook, source),
      );
      if (!result.ok) {
        void showMessage({
          title: ribbonText.definedNames,
          message: ribbonMenuText.definedNamesCreateFailed,
        });
        return;
      }
      const statusMetric = getStatusMetric();
      if (statusMetric) {
        statusMetric.textContent = ribbonMenuText.definedNamesCreated.replace(
          '{count}',
          String(result.entries.length),
        );
      }
      focusSheet();
      return;
    }
    if (action === 'use-formula') {
      const names = listDefinedNames(i.workbook);
      const firstName = names[0];
      if (firstName) {
        insertDefinedNameFormula(i.store.getState(), i.workbook, firstName.name, i.store);
        refreshWorkbookCells();
        focusSheet();
        return;
      }
      void showMessage({
        title: ribbonText.definedNames,
        message: ribbonMenuText.noDefinedNames,
      });
      return;
    }
    if (action.startsWith('insert:')) {
      const name = action.slice('insert:'.length);
      insertDefinedNameFormula(i.store.getState(), i.workbook, name, i.store);
      refreshWorkbookCells();
      focusSheet();
    }
  };

  const clearHyperlinksInSelection = (mode: 'clear' | 'remove' = 'clear'): void => {
    const i = getInst();
    if (!i) return;
    const range = i.store.getState().selection.range;
    recordFormatChange(i.history, i.store, () => {
      for (let row = range.r0; row <= range.r1; row += 1) {
        for (let col = range.c0; col <= range.c1; col += 1) {
          const addr = { sheet: range.sheet, row, col };
          clearHyperlink(i.store, addr, i.workbook);
          if (mode === 'remove') {
            mutators.setCellFormat(i.store, addr, {
              color: undefined,
              underline: undefined,
            });
          }
        }
      }
    });
    refreshWorkbookCells();
    focusSheet();
  };

  const applyLinksAction = (action: string): void => {
    const i = getInst();
    if (!i) return;
    if (action === 'hyperlink') {
      i.openHyperlinkDialog();
      return;
    }
    if (action === 'external') {
      i.openExternalLinksDialog();
      return;
    }
    if (action === 'clear') {
      clearHyperlinksInSelection('clear');
      return;
    }
    if (action === 'open') {
      const target = hyperlinkAt(i.store.getState(), i.store.getState().selection.active);
      if (!target) {
        void showMessage({
          title: ribbonText.links,
          message: ribbonMenuText.linkNoHyperlink,
        });
        return;
      }
      window.open(target, '_blank', 'noopener,noreferrer');
    }
  };

  return {
    applyCellStyleFromRibbon,
    applyCurrencyPreset,
    openCurrencyFooterAction,
    openCellStyleFooterAction,
    openTableStyleFooterAction,
    createTableFromSelection,
    applyPivotTableAction,
    applyDefinedNameAction,
    clearHyperlinksInSelection,
    applyLinksAction,
  };
};
