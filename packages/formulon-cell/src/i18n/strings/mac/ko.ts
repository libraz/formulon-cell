import { cellFormatKo } from '../cell-format/ko.js';
import { dataObjectsKo } from '../data-objects/ko.js';
import { editDialogsKo } from '../edit-dialogs/ko.js';
import { formulaDialogsKo } from '../formula-dialogs/ko.js';
import { gridPopupsKo } from '../grid-popups/ko.js';
import { pageLayoutKo } from '../page-layout/ko.js';
import { ribbonKo } from '../ribbon/ko.js';
import { ribbonMenusKo } from '../ribbon-menus/ko.js';
import { workbookChromeKo } from '../workbook-chrome/ko.js';
import { macAccessibilityPaneKo } from './accessibility-pane.js';
import { macAdvancedFilterFrom, macAdvancedFilterKo } from './advanced-filter.js';
import { macAlertsFrom, macAlertsKo, mergeMacAlerts } from './alerts.js';
import { macCfRulesDialogFrom, macCfRulesDialogKo } from './cf-rules-dialog.js';
import { macChromeFrom, macChromeKo } from './chrome.js';
import { macErrorMenuKo } from './error-menu.js';
import { macExternalLinksPaneKo } from './external-links-pane.js';
import { macFilterPanelKo } from './filter-panel.js';
import { macFindReplaceFrom, macFindReplaceKo } from './find-replace.js';
import { macFormatDialogFrom, macFormatDialogKo } from './format-dialog.js';
import { macGoToDialogFrom, macGoToDialogKo } from './goto-dialog.js';
import { macHyperlinkDialogFrom, macHyperlinkDialogKo } from './hyperlink-dialog.js';
import { macIterativeDialogFrom, macIterativeDialogKo } from './iterative-dialog.js';
import { macNamedRangeDialogFrom, macNamedRangeDialogKo } from './named-range-dialog.js';
import { macPageSetupFrom, macPageSetupKo } from './page-setup.js';
import { macPasteSpecialFrom, macPasteSpecialKo } from './paste-special.js';
import { macPivotFieldSettingsKo } from './pivot-field-settings.js';
import { macPivotOptionsKo } from './pivot-options.js';
import { macPivotTableDialogFrom, macPivotTableDialogKo } from './pivot-table-dialog.js';
import { macProtectSheetFrom, macProtectSheetKo } from './protect-sheet.js';
import { macRemoveDuplicatesFrom, macRemoveDuplicatesKo } from './remove-duplicates.js';
import { macSheetTabMenuKo } from './sheet-tab-menu.js';
import { macSortDialogFrom, macSortDialogKo } from './sort-dialog.js';
import { macStyleDialogKo } from './style-dialog.js';
import { macTextToColumnsWizardKo } from './text-to-columns-wizard.js';
import type { MacStrings } from './types.js';
import { macWatchWindowKo } from './watch-window.js';
import { macZoomDialogFrom, macZoomDialogKo } from './zoom-dialog.js';

export const macKo: MacStrings = {
  data: {
    goalSeek: {
      title: '목표값 찾기',
      formulaCell: '수식 셀',
      targetValue: '찾는 값',
      changingCell: '변경할 셀',
      iterations: '최대 반복 횟수',
      run: '확인',
      cancel: '취소',
      result: '결과',
      invalidCell: '셀 하나를 선택하세요.',
      invalidTarget: '유한한 목표값을 입력하세요.',
      calculating: '계산 중…',
      converged: '목표값 찾기가 수렴했습니다.',
    },
    consolidate: {
      title: '통합',
      sources: '참조',
      sourcesHint: '한 줄에 하나의 범위',
      destination: '결과 복사 위치',
      function: '함수',
      replace: '현재 내용 바꾸기',
      labels: '맨 위 행과 왼쪽 열의 레이블 사용',
      links: '원본 데이터에 연결 만들기',
      unsupportedOption: '이 옵션은 아직 지원되지 않습니다.',
      run: '확인',
      cancel: '취소',
      referenceList: '모든 참조 영역:',
      select: '찾아보기...',
      add: '추가',
      remove: '삭제',
      criteria: '사용할 레이블:',
      topRow: '첫 행',
      leftColumn: '왼쪽 열',
    },
    subtotal: {
      title: '부분합',
      range: '범위',
      groupBy: '변경할 때마다',
      columns: '부분합 추가 위치',
      function: '함수 사용',
      replace: '현재 부분합 바꾸기',
      summaryBelow: '데이터 아래에 요약 표시',
      run: '확인',
      cancel: '취소',
      groupTotal: '{group} 요약',
      grandTotal: '총합계',
      pageBreak: '그룹 사이에서 페이지 나누기',
      removeAll: '모두 제거',
    },
    functions: { sum: '합계', average: '평균', count: '개수', min: '최소', max: '최대' },
    errors: {
      invalidRange: '유효한 범위를 입력하세요.',
      invalidColumn: '범위 안의 유효한 열을 입력하세요.',
      targetNotFinite: '목표값은 유한해야 합니다.',
      sameCell: '수식 셀과 변경할 셀은 서로 달라야 합니다.',
      noFormula: '수식 셀에 수식이 있어야 합니다.',
      changingHasFormula: '변경할 셀에는 수식을 포함할 수 없습니다.',
      changingNotNumeric: '변경할 셀은 비어 있거나 숫자를 포함해야 합니다.',
      scratchUnsupported: '통합 문서가 임시 계산에서 수식을 유지하지 못했습니다.',
      invalidTolerance: '허용 오차는 유한한 음수가 아닌 숫자여야 합니다.',
      invalidIterations: '최대 반복 횟수는 유한한 양수여야 합니다.',
      nonFiniteResult: '수식에서 유한한 숫자가 반환되지 않았습니다.',
      notConverged: '반복 제한 내에서 목표값 찾기가 수렴하지 않았습니다.',
      stale: '작업을 적용하기 전에 통합 문서가 변경되었습니다.',
      notEditable: '대상 셀을 편집할 수 없습니다.',
      writeFailed: '대상 셀에 쓸 수 없습니다.',
      failed: '작업을 완료할 수 없습니다.',
      noSources: '원본 범위를 하나 이상 추가하세요.',
      invalidFunction: '지원되는 함수를 선택하세요.',
      invalidSource: '원본 범위가 유효하지 않습니다.',
      areaLimit: '원본 영역이 지원되는 제한을 초과합니다.',
      dimensionMismatch: '모든 원본 범위의 크기가 같아야 합니다.',
      invalidDestination: '대상은 한 셀 또는 같은 크기의 범위여야 합니다.',
      destinationOverlapsSource: '대상이 원본 범위와 겹치면 안 됩니다.',
      activeSheetOnly: '부분합은 활성 시트에만 행을 삽입할 수 있습니다.',
      needsDataRow: '부분합에는 머리글 행과 하나 이상의 데이터 행이 필요합니다.',
      engineUnsupported: '통합 문서 엔진이 SUBTOTAL을 지원하지 않습니다.',
      unsupportedObjects:
        '선택한 범위의 표, 피벗 테이블 또는 병합된 셀은 부분합을 지원하지 않습니다.',
      protectedSheet: '보호된 시트에서는 부분합을 사용할 수 없습니다.',
      groupOutsideRange: '그룹 열은 선택한 범위 안에 있어야 합니다.',
      groupLimit: '부분합 출력이 지원되는 크기를 초과합니다.',
      noGroups: '연속된 그룹을 찾을 수 없습니다.',
      noRoom: '부분합 행을 삽입할 공간이 없습니다.',
    },
  },
  sparkline: {
    title: '스파크라인 삽입',
    dataRange: '데이터 범위',
    location: '위치',
    type: '유형',
    ok: '확인',
    cancel: '취소',
    invalidSource: '유효한 원본 범위를 입력하세요.',
    locationOneCell: '위치는 한 셀이어야 합니다.',
    outOfBounds: '선택한 범위가 워크시트 밖에 있습니다.',
    chooseType: '스파크라인 유형을 선택하세요.',
    unavailable: '현재 선택에서는 이 작업을 사용할 수 없습니다.',
    protectedLocation: '스파크라인 위치가 보호되어 있습니다.',
    weightTitle: '스파크라인 두께',
    weightLabel: '스파크라인 가중치 입력:',
    dateTitle: '날짜 축',
    dateRangeLabel: '스파크라인의 날짜 범위 선택:',
    axisHorizontal: '가로',
    axisVertical: '세로',
    axisGeneral: '일반 축',
    axisShow: '축 표시',
    axisPlotRtl: '오른쪽에서 왼쪽으로 데이터 표시',
    axisMinimum: '최소값:',
    axisMaximum: '최대값:',
    axisAutoEach: '각 스파크라인에 대해 자동',
    axisSameAll: '모든 스파크라인에 대해 동일하게',
    axisRtlUnsupported: '이 옵션은 아직 지원되지 않습니다.',
    axisTitle: '축',
    customValueLabel: '사용자 지정 값:',
    emptyCellsLabel: '빈 셀 표시 형식:',
    showHiddenLabel: '숨겨진 행 및 열에 데이터 표시',
  },
  slicer: {
    ok: '확인',
    cancel: '취소',
    noTable: '사용 가능한 표가 없습니다.',
    noColumns: '표에 열이 없습니다.',
    insertFailed: '이 표에 슬라이서를 삽입할 수 없습니다.',
  },
  workbookStats: {
    title: '통합 문서 통계',
    close: '닫기',
    sheets: '시트',
    populatedCells: '입력된 셀',
    formulas: '수식 셀',
    numbers: '숫자 셀',
    text: '텍스트 셀',
    booleans: '부울 셀',
    errors: '오류 셀',
    tables: '표',
    comments: '메모 / 노트',
    hyperlinks: '하이퍼링크',
    usedRows: '사용된 행',
    usedColumns: '사용된 열',
  },
  ink: {
    strokeLabel: '잉크 선',
    drawLabel: '잉크로 그리기',
  },
  contextMenu: {
    cut: '잘라내기',
    copy: '복사',
    paste: '붙여넣기',
    pasteSpecial: '선택하여 붙여넣기',
    pasteSpecialDialog: '선택하여 붙여넣기…',
    pasteFormatsOnly: '서식',
    insert: '삽입…',
    delete: '삭제…',
    clear: '지우기',
    newComment: '새 댓글',
    editNote: '노트 편집',
    deleteNote: '노트 삭제',
    formatCells: '셀 서식…',
    pickFromList: '드롭다운 목록에서 선택…',
    insertCopiedCells: '복사한 셀 삽입…',
    insertBand: '삽입',
    deleteBand: '삭제',
    insertCopiedBand: '복사한 셀 삽입',
    rowHeight: '행 높이…',
    colWidth: '열 너비…',
    hide: '숨기기',
    unhide: '숨기기 취소',
    showPhonetic: '윗주 필드 표시',
    pasteKeepSource: '원본 서식 유지',
    pasteNoBorders: '테두리 없음',
    pasteColumnWidths: '원본 열 너비 유지',
    pasteLink: '연결하여 붙여넣기',
    defineName: '이름 정의…',
    hyperlink: '하이퍼링크 삽입…',
  },
  homeRibbon: {
    conditional: '조건부 서식',
    paste: '붙여넣기',
    cut: '잘라내기',
    copy: '복사',
    formatPainter: '서식 복사',
    cellStyles: '셀 스타일',
    autoSum: '자동 합계',
  },
  formatDialog: {
    ...{ ...macFormatDialogFrom(cellFormatKo.formatDialog), ...macFormatDialogKo },
    ...macFormatDialogKo,
  },
  iterativeDialog: {
    ...{
      ...macIterativeDialogFrom(formulaDialogsKo.iterativeDialog, {
        calcAutoNoTable: ribbonMenusKo.ribbonMenu.calcAutoNoTable,
      }),
      ...macIterativeDialogKo,
    },
    ...macIterativeDialogKo,
  },
  namedRangeDialog: {
    ...{
      ...macNamedRangeDialogFrom(formulaDialogsKo.namedRangeDialog, {
        defineTitle: '이름 정의',
        defineNames: '이름:',
        defineNamePrompt: '데이터 범위의 이름을 입력하십시오:',
        defineRangePrompt: '셀 범위를 선택하십시오:',
      }),
      ...macNamedRangeDialogKo,
    },
    ...macNamedRangeDialogKo,
  },
  pasteSpecial: {
    ...{ ...macPasteSpecialFrom(editDialogsKo.pasteSpecialDialog), ...macPasteSpecialKo },
    ...macPasteSpecialKo,
  },
  zoomDialog: {
    ...{
      ...macZoomDialogFrom(ribbonKo.ribbon, cellFormatKo.formatDialog),
      ...macZoomDialogKo,
    },
    ...macZoomDialogKo,
  },
  alerts: mergeMacAlerts(
    mergeMacAlerts(
      macAlertsFrom(
        ribbonKo.ribbon,
        {
          message: '이 시트는 영구적으로 삭제됩니다. 계속하시겠습니까?',
          confirm: '삭제',
          cancel: '취소',
        },
        editDialogsKo.moveCells,
      ),
      macAlertsKo,
    ),
    macAlertsKo,
  ),
  protectSheet: {
    ...{
      ...macProtectSheetFrom(editDialogsKo.protection, {
        heading: '워크시트 및 잠긴 셀의 내용 보호',
        note: '기본적으로 모든 셀이 잠겨 있지만 잠금을 해제할 수 있습니다.',
      }),
      ...macProtectSheetKo,
    },
    ...macProtectSheetKo,
  },
  advancedFilter: {
    ...{
      ...macAdvancedFilterFrom(ribbonMenusKo.ribbonMenu, {
        filterInPlace: '현재 위치에 필터',
        copyToLocation: '다른 장소에 복사',
      }),
      ...macAdvancedFilterKo,
    },
    ...macAdvancedFilterKo,
  },
  chrome: {
    ...{ ...macChromeFrom(workbookChromeKo, formulaDialogsKo), ...macChromeKo },
    ...macChromeKo,
  },
  textToColumnsWizard: macTextToColumnsWizardKo,
  styleDialog: macStyleDialogKo,
  externalLinksPane: macExternalLinksPaneKo,
  watchWindow: macWatchWindowKo,
  accessibilityPane: macAccessibilityPaneKo,
  cfRulesDialog: {
    ...{
      ...macCfRulesDialogFrom(cellFormatKo.cfRulesDialog, {
        orderLabel: '규칙 순서 변경:',
        appliesToLabel: '적용 대상:',
      }),
      ...macCfRulesDialogKo,
    },
    ...macCfRulesDialogKo,
  },
  hyperlinkDialog: {
    ...{
      ...macHyperlinkDialogFrom(editDialogsKo.hyperlinkDialog, {
        displayText: '표시할 텍스트:',
        selectionInDocument: '<<문서에서 선택한 영역>>',
        screenTip: '화면 설명...',
        screenTipTitle: '하이퍼링크 화면 설명 설정',
        screenTipLabel: '화면 설명 텍스트:',
        tabWeb: '웹 페이지 또는 파일',
        tabDocument: '현재 문서',
        tabEmail: '전자 메일 주소',
        webIntro: '기존 파일 또는 웹 페이지에 연결합니다.',
        browse: '선택...',
        browseUnavailable: '브라우저에서는 파일을 선택할 수 없습니다. 주소를 입력하십시오.',
        cellReference: '셀 참조 입력:',
        documentPlace: '또는 문서에서 위치 선택:',
        emailAddress: '전자 메일 주소:',
        subject: '제목:',
        recentEmails: '최근에 사용한 전자 메일 주소:',
      }),
      ...macHyperlinkDialogKo,
    },
    ...macHyperlinkDialogKo,
  },
  pageSetup: {
    ...{ ...macPageSetupFrom(pageLayoutKo.pageSetup), ...macPageSetupKo },
    ...macPageSetupKo,
  },
  dataValidation: {
    inputCaption: '셀을 선택하면 나타낼 설명 메시지:',
    errorCaption: '유효하지 않은 데이터를 입력하면 표시할 오류 메시지:',
  },
  createTable: { nameLabel: '표 이름은 무엇입니까?' },
  tableStyleDialog: {
    newTitle: '표 스타일 새로 만들기',
    modifyTitle: '테이블 스타일 수정',
    newPivotTitle: '새 피벗 테이블 빠른 스타일',
    modifyPivotTitle: '피벗 테이블 빠른 스타일 수정',
    setDefault: '이 통합 문서의 기본 표 스타일로 설정',
    name: '이름:',
    format: '서식',
    clear: '지우기',
    preview: '미리 보기',
    stripeSize: '줄무늬 크기',
  },
  goToDialog: {
    ...macGoToDialogFrom(editDialogsKo.goToDialog, {
      selectCaption: '선택',
      optionsGroup: '옵션',
      noOptions: '이 선택 항목에는 옵션이 없습니다.',
    }),
    ...macGoToDialogKo,
  },
  pivotFieldSettings: macPivotFieldSettingsKo,
  pivotOptions: macPivotOptionsKo,
  pivotTableDialog: {
    ...{
      ...macPivotTableDialogFrom(dataObjectsKo.pivotTableDialog, {
        chooseConnection: '연결 선택...',
        moveTitle: '피벗 테이블 이동',
        moveLocation: '위치:',
      }),
      ...macPivotTableDialogKo,
    },
    ...macPivotTableDialogKo,
  },
  removeDuplicates: {
    ...macRemoveDuplicatesFrom(ribbonMenusKo.ribbonMenu),
    ...macRemoveDuplicatesKo,
  },
  sortDialog: {
    ...{
      ...macSortDialogFrom(ribbonMenusKo.ribbonMenu, {
        levelsHeading: '정렬 수준 추가:',
        columnHeading: '열',
        sortOnHeading: '정렬 기준',
        sortOnValues: editDialogsKo.findReplace.values,
        colorHeading: '색/아이콘',
        copyLevel: gridPopupsKo.contextMenu.copy,
        options: '옵션...',
        rowHeading: '행',
        orientation: '방향',
        sortTopToBottom: '위에서 아래로 정렬',
        sortLeftToRight: '왼쪽에서 오른쪽으로 정렬',
        method: '방법',
        usePhonetic: '발음 사용',
        noPhonetic: '발음 사용 안 함',
        caseSensitive: '대/소문자 구분',
      }),
      ...macSortDialogKo,
    },
    ...macSortDialogKo,
  },
  errorMenu: macErrorMenuKo,
  filterPanel: macFilterPanelKo(gridPopupsKo.filterDropdown),
  sheetTabMenu: macSheetTabMenuKo,
  findReplace: {
    ...{
      ...macFindReplaceFrom(editDialogsKo.findReplace, '반각과 전각 구분'),
      ...macFindReplaceKo,
    },
    ...macFindReplaceKo,
  },
};
