import { cellFormatZh } from '../cell-format/zh.js';
import { dataObjectsZh } from '../data-objects/zh.js';
import { editDialogsZh } from '../edit-dialogs/zh.js';
import { formulaDialogsZh } from '../formula-dialogs/zh.js';
import { gridPopupsZh } from '../grid-popups/zh.js';
import { pageLayoutZh } from '../page-layout/zh.js';
import { ribbonZh } from '../ribbon/zh.js';
import { ribbonMenusZh } from '../ribbon-menus/zh.js';
import { workbookChromeZh } from '../workbook-chrome/zh.js';
import { macAccessibilityPaneZh } from './accessibility-pane.js';
import { macAdvancedFilterFrom, macAdvancedFilterZh } from './advanced-filter.js';
import { macAlertsFrom, macAlertsZh, mergeMacAlerts } from './alerts.js';
import { macCfRulesDialogFrom, macCfRulesDialogZh } from './cf-rules-dialog.js';
import { macChromeFrom, macChromeZh } from './chrome.js';
import { macErrorMenuZh } from './error-menu.js';
import { macExternalLinksPaneZh } from './external-links-pane.js';
import { macFilterPanelZh } from './filter-panel.js';
import { macFindPanelZh, macFindReplaceFrom, macFindReplaceZh } from './find-replace.js';
import { macFormatDialogFrom, macFormatDialogZh } from './format-dialog.js';
import { macGoToDialogFrom, macGoToDialogZh } from './goto-dialog.js';
import { macHyperlinkDialogFrom, macHyperlinkDialogZh } from './hyperlink-dialog.js';
import { macIterativeDialogFrom, macIterativeDialogZh } from './iterative-dialog.js';
import { macNamedRangeDialogFrom, macNamedRangeDialogZh } from './named-range-dialog.js';
import { macPageSetupFrom, macPageSetupZh } from './page-setup.js';
import { macPasteSpecialFrom, macPasteSpecialZh } from './paste-special.js';
import { macPivotFieldSettingsZh } from './pivot-field-settings.js';
import { macPivotOptionsZh } from './pivot-options.js';
import { macPivotTableDialogFrom, macPivotTableDialogZh } from './pivot-table-dialog.js';
import { macProtectSheetFrom, macProtectSheetZh } from './protect-sheet.js';
import { macRemoveDuplicatesFrom, macRemoveDuplicatesZh } from './remove-duplicates.js';
import { macSheetTabMenuZh } from './sheet-tab-menu.js';
import { macSortDialogFrom, macSortDialogZh } from './sort-dialog.js';
import { macStyleDialogZh } from './style-dialog.js';
import { macTextToColumnsWizardZh } from './text-to-columns-wizard.js';
import type { MacStrings } from './types.js';
import { macWatchWindowZh } from './watch-window.js';
import { macZoomDialogFrom, macZoomDialogZh } from './zoom-dialog.js';

export const macZh: MacStrings = {
  data: {
    goalSeek: {
      title: '单变量求解',
      formulaCell: '设置单元格',
      targetValue: '目标值',
      changingCell: '可变单元格',
      iterations: '最大迭代次数',
      run: '确定',
      cancel: '取消',
      result: '结果',
      invalidCell: '请选择一个单元格。',
      invalidTarget: '请输入有限目标值。',
      calculating: '正在计算…',
      converged: '单变量求解已收敛。',
    },
    consolidate: {
      title: '合并计算',
      sources: '引用位置',
      sourcesHint: '每行一个区域',
      destination: '将结果复制到',
      function: '函数',
      replace: '替换当前内容',
      labels: '使用首行和最左列中的标签',
      links: '创建指向源数据的链接',
      unsupportedOption: '此选项尚不受支持。',
      run: '确定',
      cancel: '取消',
      referenceList: '所有引用位置:',
      select: '浏览...',
      add: '添加',
      remove: '删除',
      criteria: '标签位置:',
      topRow: '首行',
      leftColumn: '最左列',
    },
    subtotal: {
      title: '分类汇总',
      range: '区域',
      groupBy: '每次更改时',
      columns: '汇总添加到',
      function: '使用函数',
      replace: '替换当前分类汇总',
      summaryBelow: '汇总显示在数据下方',
      run: '确定',
      cancel: '取消',
      groupTotal: '{group} 汇总',
      grandTotal: '总计',
      pageBreak: '每组数据分页',
      removeAll: '全部删除',
    },
    functions: { sum: '求和', average: '平均值', count: '计数', min: '最小值', max: '最大值' },
    errors: {
      invalidRange: '请输入有效区域。',
      invalidColumn: '请输入区域内的有效列。',
      targetNotFinite: '目标值必须为有限数值。',
      sameCell: '设置单元格和可变单元格必须不同。',
      noFormula: '设置单元格必须包含公式。',
      changingHasFormula: '可变单元格不能包含公式。',
      changingNotNumeric: '可变单元格必须为空或包含数字。',
      scratchUnsupported: '工作簿无法在临时计算中保留该公式。',
      invalidTolerance: '容差必须为有限的非负数。',
      invalidIterations: '最大迭代次数必须为有限的正数。',
      nonFiniteResult: '公式未生成有限数值。',
      notConverged: '单变量求解在迭代上限内未收敛。',
      stale: '应用操作前工作簿已发生更改。',
      notEditable: '目标单元格不可编辑。',
      writeFailed: '操作无法写入目标单元格。',
      failed: '操作无法完成。',
      noSources: '请至少添加一个源区域。',
      invalidFunction: '请选择受支持的函数。',
      invalidSource: '源区域无效。',
      areaLimit: '源区域超出受支持的限制。',
      dimensionMismatch: '所有源区域必须具有相同的维度。',
      invalidDestination: '目标必须为一个单元格或大小相同的区域。',
      destinationOverlapsSource: '目标不能与源区域重叠。',
      activeSheetOnly: '分类汇总只能在活动工作表中插入行。',
      needsDataRow: '分类汇总需要标题行和至少一行数据。',
      engineUnsupported: '工作簿引擎不支持 SUBTOTAL。',
      unsupportedObjects: '选定区域中的表格、数据透视表或合并单元格不支持分类汇总。',
      protectedSheet: '受保护的工作表无法使用分类汇总。',
      groupOutsideRange: '分组列必须位于选定区域内。',
      groupLimit: '分类汇总输出超出受支持的大小。',
      noGroups: '未找到连续组。',
      noRoom: '没有空间插入分类汇总行。',
    },
  },
  sparkline: {
    title: '插入迷你图',
    dataRange: '数据范围',
    location: '位置',
    type: '类型',
    ok: '确定',
    cancel: '取消',
    invalidSource: '请输入有效的源区域。',
    locationOneCell: '位置必须为一个单元格。',
    outOfBounds: '所选区域位于工作表之外。',
    chooseType: '请选择迷你图类型。',
    unavailable: '当前选择无法执行此操作。',
    protectedLocation: '迷你图位置受保护。',
    weightTitle: '迷你图粗细',
    weightLabel: '输入迷你图粗细:',
    dateTitle: '日期坐标轴',
    dateRangeLabel: '为迷你图选择数据区域:',
    axisHorizontal: '水平',
    axisVertical: '垂直',
    axisGeneral: '常规坐标轴',
    axisShow: '显示坐标轴',
    axisPlotRtl: '从右到左的绘图数据',
    axisMinimum: '最小值:',
    axisMaximum: '最大值:',
    axisAutoEach: '自动设置每个迷你图',
    axisSameAll: '适用于所有迷你图',
    axisTitle: '坐标轴',
    customValueLabel: '自定义值:',
    hiddenEmptyMenu: '隐藏的单元格和空单元格',
    locationNotLine:
      '位置引用无效，因为单元格没有全部在同一列或行中。请选择全部在单列或单行中的单元格。',
    dateAxisNotLine:
      '日期坐标轴引用无效，因为单元格没有全部在同一列或行中。请选择全部在单列或单行中的单元格。',
  },
  slicer: {
    ok: '确定',
    cancel: '取消',
    noTable: '没有可用的表格。',
    noColumns: '表格没有列。',
    insertFailed: '无法为此表格插入切片器。',
  },
  workbookStats: {
    title: '工作簿统计信息',
    close: '关闭',
    sheets: '工作表',
    populatedCells: '非空单元格',
    formulas: '公式单元格',
    numbers: '数字单元格',
    text: '文本单元格',
    booleans: '布尔值单元格',
    errors: '错误单元格',
    tables: '表格',
    comments: '批注/备注',
    hyperlinks: '超链接',
    usedRows: '已用行',
    usedColumns: '已用列',
  },
  ink: {
    strokeLabel: '墨迹笔画',
    drawLabel: '使用墨迹绘图',
  },
  contextMenu: {
    cut: '剪切',
    copy: '拷贝',
    paste: '粘贴',
    pasteSpecial: '选择性粘贴',
    pasteSpecialDialog: '选择性粘贴…',
    pasteFormatsOnly: '格式',
    insert: '插入…',
    delete: '删除…',
    clear: '清除',
    newComment: '新建批注',
    newNote: '新建便笺',
    editNote: '编辑备注',
    deleteNote: '删除备注',
    formatCells: '设置单元格格式…',
    pickFromList: '从下拉列表中选择…',
    insertCopiedCells: '插入复制的单元格…',
    insertBand: '插入',
    deleteBand: '删除',
    insertCopiedBand: '插入复制的单元格',
    rowHeight: '行高…',
    colWidth: '列宽…',
    hide: '隐藏',
    unhide: '取消隐藏',
    showPhonetic: '显示拼音字段',
    pasteKeepSource: '保留源格式',
    pasteNoBorders: '无边框',
    pasteColumnWidths: '保留源列宽',
    pasteLink: '粘贴链接',
    defineName: '定义名称…',
    hyperlink: '插入超链接…',
  },
  homeRibbon: {
    conditional: '条件格式',
    paste: '粘贴',
    cut: '剪切',
    copy: '复制',
    formatPainter: '格式刷',
    cellStyles: '单元格样式',
    autoSum: '自动求和',
  },
  formatDialog: {
    ...{ ...macFormatDialogFrom(cellFormatZh.formatDialog), ...macFormatDialogZh },
    ...macFormatDialogZh,
  },
  iterativeDialog: {
    ...{
      ...macIterativeDialogFrom(formulaDialogsZh.iterativeDialog, {
        calcAutoNoTable: ribbonMenusZh.ribbonMenu.calcAutoNoTable,
      }),
      ...macIterativeDialogZh,
    },
    ...macIterativeDialogZh,
  },
  namedRangeDialog: {
    ...{
      ...macNamedRangeDialogFrom(formulaDialogsZh.namedRangeDialog, {
        defineTitle: '定义名称',
        defineNames: '名称:',
        defineNamePrompt: '请输入数据区域的名称:',
        defineRangePrompt: '请选择单元格区域:',
      }),
      ...macNamedRangeDialogZh,
    },
    ...macNamedRangeDialogZh,
  },
  pasteSpecial: {
    ...{ ...macPasteSpecialFrom(editDialogsZh.pasteSpecialDialog), ...macPasteSpecialZh },
    ...macPasteSpecialZh,
  },
  zoomDialog: {
    ...{
      ...macZoomDialogFrom(ribbonZh.ribbon, cellFormatZh.formatDialog),
      ...macZoomDialogZh,
    },
    ...macZoomDialogZh,
  },
  alerts: mergeMacAlerts(
    mergeMacAlerts(
      macAlertsFrom(
        ribbonZh.ribbon,
        {
          message: '此工作表将被永久删除。是否继续?',
          confirm: '删除',
          cancel: '取消',
        },
        editDialogsZh.moveCells,
        {
          headline: '无法找到您搜索的数据。',
          message: '如果您确定该数据存在于当前工作表中，请检查您键入的内容并重试。',
          confirm: '确定',
        },
      ),
      macAlertsZh,
    ),
    macAlertsZh,
  ),
  protectSheet: {
    ...{
      ...macProtectSheetFrom(editDialogsZh.protection, {
        heading: '保护工作表及锁定的单元格内容',
        note: '默认情况下所有单元格均已锁定，但可以解除锁定。',
      }),
      ...macProtectSheetZh,
    },
    ...macProtectSheetZh,
  },
  advancedFilter: {
    ...{
      ...macAdvancedFilterFrom(ribbonMenusZh.ribbonMenu, {
        filterInPlace: '在原有区域显示筛选结果',
        copyToLocation: '将筛选结果复制到其他位置',
      }),
      ...macAdvancedFilterZh,
    },
    ...macAdvancedFilterZh,
  },
  chrome: {
    ...{ ...macChromeFrom(workbookChromeZh, formulaDialogsZh), ...macChromeZh },
    ...macChromeZh,
  },
  textToColumnsWizard: macTextToColumnsWizardZh,
  styleDialog: macStyleDialogZh,
  externalLinksPane: macExternalLinksPaneZh,
  watchWindow: macWatchWindowZh,
  accessibilityPane: macAccessibilityPaneZh,
  cfRulesDialog: {
    ...{
      ...macCfRulesDialogFrom(cellFormatZh.cfRulesDialog, {
        orderLabel: '更改规则顺序:',
        appliesToLabel: '应用于:',
      }),
      ...macCfRulesDialogZh,
    },
    ...macCfRulesDialogZh,
  },
  hyperlinkDialog: {
    ...{
      ...macHyperlinkDialogFrom(editDialogsZh.hyperlinkDialog, {
        displayText: '要显示的文字:',
        selectionInDocument: '<<文档中的选定内容>>',
        screenTip: '屏幕提示...',
        screenTipTitle: '设置超链接屏幕提示',
        screenTipLabel: '屏幕提示文字:',
        tabWeb: '网页或文件',
        tabDocument: '本文档',
        tabEmail: '电子邮件地址',
        webIntro: '链接到现有文件或网页。',
        browse: '选择...',
        browseUnavailable: '浏览器中无法选择文件。请输入地址。',
        cellReference: '请键入单元格引用:',
        documentPlace: '或在此文档中选择一个位置:',
        emailAddress: '电子邮件地址:',
        subject: '主题:',
        recentEmails: '最近用过的电子邮件地址:',
      }),
      ...macHyperlinkDialogZh,
    },
    ...macHyperlinkDialogZh,
  },
  pageSetup: {
    ...{ ...macPageSetupFrom(pageLayoutZh.pageSetup), ...macPageSetupZh },
    ...macPageSetupZh,
  },
  dataValidation: {
    inputCaption: '选定单元格时显示下列输入信息:',
    errorCaption: '输入无效数据时显示下列出错警告:',
  },
  createTable: { nameLabel: '表名称是什么?' },
  tableStyleDialog: {
    newTitle: '新建表样式',
    modifyTitle: '修改表样式',
    newPivotTitle: '新建数据透视表快速样式',
    modifyPivotTitle: '修改数据透视表快速样式',
    setDefault: '设置为此工作簿的默认表样式',
    name: '名称:',
    format: '格式',
    clear: '清除',
    preview: '预览',
    stripeSize: '条纹尺寸',
  },
  goToDialog: {
    ...macGoToDialogFrom(editDialogsZh.goToDialog, {
      selectCaption: '选择',
      optionsGroup: '选项',
      noOptions: '此选择项没有选项。',
    }),
    ...macGoToDialogZh,
  },
  pivotFieldSettings: macPivotFieldSettingsZh,
  pivotOptions: macPivotOptionsZh,
  pivotTableDialog: {
    ...{
      ...macPivotTableDialogFrom(dataObjectsZh.pivotTableDialog, {
        chooseConnection: '选择连接...',
        moveTitle: '移动数据透视表',
        moveLocation: '位置:',
      }),
      ...macPivotTableDialogZh,
    },
    ...macPivotTableDialogZh,
  },
  removeDuplicates: {
    ...macRemoveDuplicatesFrom(ribbonMenusZh.ribbonMenu),
    ...macRemoveDuplicatesZh,
  },
  sortDialog: {
    ...{
      ...macSortDialogFrom(ribbonMenusZh.ribbonMenu, {
        levelsHeading: '添加排序级别:',
        columnHeading: '列',
        sortOnHeading: '排序依据',
        sortOnValues: editDialogsZh.findReplace.values,
        colorHeading: '颜色/图标',
        copyLevel: gridPopupsZh.contextMenu.copy,
        options: '选项...',
        rowHeading: '行',
        orientation: '方向',
        sortTopToBottom: '从上到下排序',
        sortLeftToRight: '从左到右排序',
        method: '方法',
        usePhonetic: '使用拼音',
        noPhonetic: '不使用拼音',
        caseSensitive: '区分大小写',
      }),
      ...macSortDialogZh,
    },
    ...macSortDialogZh,
  },
  errorMenu: macErrorMenuZh,
  filterPanel: macFilterPanelZh(gridPopupsZh.filterDropdown),
  sheetTabMenu: macSheetTabMenuZh,
  findPanel: macFindPanelZh,
  findReplace: {
    ...{
      ...macFindReplaceFrom(editDialogsZh.findReplace, '区分全/半角'),
      ...macFindReplaceZh,
    },
    ...macFindReplaceZh,
  },
};
