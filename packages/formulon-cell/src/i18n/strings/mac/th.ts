import { cellFormatTh } from '../cell-format/th.js';
import { dataObjectsTh } from '../data-objects/th.js';
import { editDialogsTh } from '../edit-dialogs/th.js';
import { formulaDialogsTh } from '../formula-dialogs/th.js';
import { gridPopupsTh } from '../grid-popups/th.js';
import { pageLayoutTh } from '../page-layout/th.js';
import { ribbonTh } from '../ribbon/th.js';
import { ribbonMenusTh } from '../ribbon-menus/th.js';
import { workbookChromeTh } from '../workbook-chrome/th.js';
import { macAccessibilityPaneTh } from './accessibility-pane.js';
import { macAdvancedFilterFrom } from './advanced-filter.js';
import { macAlertsFrom } from './alerts.js';
import { macCfRulesDialogFrom } from './cf-rules-dialog.js';
import { macChromeFrom } from './chrome.js';
import { macErrorMenuTh } from './error-menu.js';
import { macExternalLinksPaneTh } from './external-links-pane.js';
import { macFindReplaceFrom } from './find-replace.js';
import { macFormatDialogFrom } from './format-dialog.js';
import { macGoToDialogFrom } from './goto-dialog.js';
import { macHyperlinkDialogFrom } from './hyperlink-dialog.js';
import { macIterativeDialogFrom } from './iterative-dialog.js';
import { macNamedRangeDialogFrom } from './named-range-dialog.js';
import { macPageSetupFrom } from './page-setup.js';
import { macPasteSpecialFrom } from './paste-special.js';
import { macPivotFieldSettingsTh } from './pivot-field-settings.js';
import { macPivotOptionsTh } from './pivot-options.js';
import { macPivotTableDialogFrom } from './pivot-table-dialog.js';
import { macProtectSheetFrom } from './protect-sheet.js';
import { macRemoveDuplicatesFrom } from './remove-duplicates.js';
import { macSheetTabMenuTh } from './sheet-tab-menu.js';
import { macSortDialogFrom } from './sort-dialog.js';
import { macStyleDialogTh } from './style-dialog.js';
import { macTextToColumnsWizardTh } from './text-to-columns-wizard.js';
import type { MacStrings } from './types.js';
import { macWatchWindowTh } from './watch-window.js';
import { macZoomDialogFrom } from './zoom-dialog.js';

export const macTh: MacStrings = {
  macData: {
    goalSeek: {
      title: 'ค้นหาค่าเป้าหมาย',
      formulaCell: 'เซลล์สูตร',
      targetValue: 'ค่าเป้าหมาย',
      changingCell: 'เซลล์ที่เปลี่ยนแปลง',
      iterations: 'จํานวนการวนซ้ำสูงสุด',
      run: 'ตกลง',
      cancel: 'ยกเลิก',
      result: 'ผลลัพธ์',
      invalidCell: 'เลือกเซลล์เดียว',
      invalidTarget: 'ใส่ค่าเป้าหมายที่เป็นจำนวนจำกัด',
      calculating: 'กำลังคำนวณ…',
      converged: 'การค้นหาค่าเป้าหมายมาบรรจบกันแล้ว',
    },
    consolidate: {
      title: 'รวม',
      sources: 'การอ้างอิง',
      sourcesHint: 'หนึ่งช่วงต่อหนึ่งบรรทัด',
      destination: 'คัดลอกผลลัพธ์ไปยัง',
      function: 'ฟังก์ชัน',
      replace: 'แทนที่เนื้อหาปัจจุบัน',
      labels: 'ใช้ป้ายชื่อในแถวบนสุดและคอลัมน์ด้านซ้าย',
      links: 'สร้างลิงก์ไปยังข้อมูลต้นฉบับ',
      unsupportedOption: 'ยังไม่รองรับตัวเลือกนี้',
      run: 'ตกลง',
      cancel: 'ยกเลิก',
      referenceList: 'การอ้างอิงทั้งหมด:',
      select: 'เรียกดู...',
      add: 'เพิ่ม',
      remove: 'ลบ',
      criteria: 'ใช้ป้ายชื่อใน:',
      topRow: 'แถวบนสุด',
      leftColumn: 'คอลัมน์ซ้ายสุด',
    },
    subtotal: {
      title: 'ผลรวมย่อย',
      range: 'ช่วง',
      groupBy: 'เมื่อมีการเปลี่ยนแปลงในแต่ละ',
      columns: 'เพิ่มผลรวมย่อยลงใน',
      function: 'ใช้ฟังก์ชัน',
      replace: 'แทนที่ผลรวมย่อยปัจจุบัน',
      summaryBelow: 'สรุปด้านล่างข้อมูล',
      run: 'ตกลง',
      cancel: 'ยกเลิก',
      groupTotal: '{group} รวม',
      grandTotal: 'ผลรวมทั้งหมด',
      pageBreak: 'ตัวแบ่งหน้าระหว่างกลุ่ม',
      removeAll: 'เอาออกทั้งหมด',
    },
    functions: { sum: 'ผลรวม', average: 'ค่าเฉลี่ย', count: 'นับ', min: 'ต่ำสุด', max: 'สูงสุด' },
    errors: {
      invalidRange: 'ใส่ช่วงที่ถูกต้อง',
      invalidColumn: 'ใส่คอลัมน์ที่ถูกต้องภายในช่วง',
      targetNotFinite: 'ค่าเป้าหมายต้องเป็นจำนวนจำกัด',
      sameCell: 'เซลล์สูตรและเซลล์ที่เปลี่ยนแปลงต้องแตกต่างกัน',
      noFormula: 'เซลล์สูตรต้องมีสูตร',
      changingHasFormula: 'เซลล์ที่เปลี่ยนแปลงต้องไม่มีสูตร',
      changingNotNumeric: 'เซลล์ที่เปลี่ยนแปลงต้องว่างหรือมีตัวเลข',
      scratchUnsupported: 'เวิร์กบุ๊กไม่สามารถเก็บสูตรไว้ในการคำนวณชั่วคราวได้',
      invalidTolerance: 'ค่าความคลาดเคลื่อนต้องเป็นจำนวนจำกัดที่ไม่ติดลบ',
      invalidIterations: 'จำนวนการวนซ้ำสูงสุดต้องเป็นจำนวนจำกัดที่เป็นบวก',
      nonFiniteResult: 'สูตรไม่ได้สร้างตัวเลขที่เป็นจำนวนจำกัด',
      notConverged: 'การค้นหาค่าเป้าหมายไม่มาบรรจบกันภายในขีดจำกัดการวนซ้ำ',
      stale: 'เวิร์กบุ๊กมีการเปลี่ยนแปลงก่อนที่จะใช้การดำเนินการได้',
      notEditable: 'ไม่สามารถแก้ไขเซลล์เป้าหมายได้',
      writeFailed: 'การดำเนินการไม่สามารถเขียนลงในเซลล์เป้าหมายได้',
      failed: 'ไม่สามารถดำเนินการให้เสร็จสมบูรณ์ได้',
      noSources: 'เพิ่มช่วงต้นฉบับอย่างน้อยหนึ่งช่วง',
      invalidFunction: 'เลือกฟังก์ชันที่รองรับ',
      invalidSource: 'ช่วงต้นฉบับไม่ถูกต้อง',
      areaLimit: 'พื้นที่ต้นฉบับเกินขีดจำกัดที่รองรับ',
      dimensionMismatch: 'ช่วงต้นฉบับทั้งหมดต้องมีมิติเดียวกัน',
      invalidDestination: 'ปลายทางต้องเป็นเซลล์เดียวหรือช่วงที่มีขนาดเท่ากัน',
      destinationOverlapsSource: 'ปลายทางต้องไม่ทับซ้อนกับช่วงต้นฉบับ',
      activeSheetOnly: 'ผลรวมย่อยสามารถแทรกแถวได้เฉพาะในแผ่นงานที่ใช้งานอยู่',
      needsDataRow: 'ผลรวมย่อยต้องมีแถวส่วนหัวและแถวข้อมูลอย่างน้อยหนึ่งแถว',
      engineUnsupported: 'กลไกเวิร์กบุ๊กไม่สนับสนุน SUBTOTAL',
      unsupportedObjects: 'ผลรวมย่อยไม่สนับสนุนตาราง PivotTable หรือเซลล์ที่ผสานในช่วงที่เลือก',
      protectedSheet: 'ผลรวมย่อยไม่พร้อมใช้งานในแผ่นงานที่ได้รับการป้องกัน',
      groupOutsideRange: 'คอลัมน์กลุ่มต้องอยู่ภายในช่วงที่เลือก',
      groupLimit: 'ผลลัพธ์ผลรวมย่อยเกินขนาดที่รองรับ',
      noGroups: 'ไม่พบกลุ่มที่ต่อเนื่องกัน',
      noRoom: 'ไม่มีพื้นที่สำหรับแทรกแถวผลรวมย่อย',
    },
  },
  macSparkline: {
    title: 'แทรกเส้นประกายไฟ',
    dataRange: 'ช่วงข้อมูล',
    location: 'ตำแหน่งที่ตั้ง',
    type: 'ชนิด',
    ok: 'ตกลง',
    cancel: 'ยกเลิก',
    invalidSource: 'ใส่ช่วงต้นฉบับที่ถูกต้อง',
    locationOneCell: 'ตำแหน่งที่ตั้งต้องเป็นเซลล์เดียว',
    outOfBounds: 'ช่วงที่เลือกอยู่นอกเวิร์กชีต',
    chooseType: 'เลือกชนิดเส้นประกายไฟ',
    unavailable: 'การดำเนินการนี้ไม่พร้อมใช้งานสำหรับส่วนที่เลือกปัจจุบัน',
    protectedLocation: 'ตำแหน่งเส้นประกายไฟได้รับการป้องกัน',
  },
  macSlicer: {
    ok: 'ตกลง',
    cancel: 'ยกเลิก',
    noTable: 'ไม่มีตารางที่พร้อมใช้งาน',
    noColumns: 'ตารางไม่มีคอลัมน์',
    insertFailed: 'ไม่สามารถแทรกตัวแบ่งส่วนข้อมูลสำหรับตารางนี้ได้',
  },
  macWorkbookStats: {
    title: 'สถิติเวิร์กบุ๊ก',
    close: 'ปิด',
    sheets: 'แผ่นงาน',
    populatedCells: 'เซลล์ที่มีข้อมูล',
    formulas: 'เซลล์สูตร',
    numbers: 'เซลล์ตัวเลข',
    text: 'เซลล์ข้อความ',
    booleans: 'เซลล์บูลีน',
    errors: 'เซลล์ข้อผิดพลาด',
    tables: 'ตาราง',
    comments: 'ข้อคิดเห็น / บันทึกย่อ',
    hyperlinks: 'ไฮเปอร์ลิงก์',
    usedRows: 'แถวที่ใช้',
    usedColumns: 'คอลัมน์ที่ใช้',
  },
  macInk: {
    strokeLabel: 'เส้นหมึก',
    drawLabel: 'วาดด้วยหมึก',
  },
  macContextMenu: {
    cut: 'ตัด',
    copy: 'คัดลอก',
    paste: 'วาง',
    pasteSpecial: 'วางแบบพิเศษ',
    pasteSpecialDialog: 'วางแบบพิเศษ…',
    pasteFormatsOnly: 'การจัดรูปแบบ',
    insert: 'แทรก…',
    delete: 'ลบ…',
    clear: 'ล้าง',
    newNote: 'บันทึกย่อใหม่',
    editNote: 'แก้ไขบันทึกย่อ',
    deleteNote: 'ลบบันทึกย่อ',
    formatCells: 'จัดรูปแบบเซลล์…',
    editPhonetic: 'แก้ไขคำแนะนำการออกเสียง…',
    defineName: 'กำหนดชื่อ…',
    hyperlink: 'แทรกไฮเปอร์ลิงก์…',
  },
  macHomeRibbon: {
    conditional: 'การจัดรูปแบบตามเงื่อนไข',
    paste: 'วาง',
    cut: 'ตัด',
    copy: 'คัดลอก',
    formatPainter: 'ตัวคัดวางรูปแบบ',
    cellStyles: 'สไตล์เซลล์',
    autoSum: 'ผลรวมอัตโนมัติ',
  },
  macFormatDialog: macFormatDialogFrom(cellFormatTh.formatDialog),
  macIterativeDialog: macIterativeDialogFrom(formulaDialogsTh.iterativeDialog, {
    calcAutoNoTable: ribbonMenusTh.ribbonMenu.calcAutoNoTable,
  }),
  macNamedRangeDialog: macNamedRangeDialogFrom(formulaDialogsTh.namedRangeDialog, {
    defineTitle: 'กำหนดชื่อ',
    defineNames: 'ชื่อ:',
    defineNamePrompt: 'ใส่ชื่อสำหรับช่วงข้อมูล:',
    defineRangePrompt: 'เลือกช่วงของเซลล์:',
  }),
  macPasteSpecial: macPasteSpecialFrom(editDialogsTh.pasteSpecialDialog),
  macZoomDialog: macZoomDialogFrom(ribbonTh.ribbon, cellFormatTh.formatDialog),
  macAlerts: macAlertsFrom(ribbonTh.ribbon, {
    message: 'แผ่นงานนี้จะถูกลบอย่างถาวร คุณต้องการดำเนินการต่อหรือไม่',
    confirm: 'ลบ',
    cancel: 'ยกเลิก',
  }),
  macProtectSheet: macProtectSheetFrom(editDialogsTh.protection, {
    heading: 'ป้องกันเวิร์กชีตและเนื้อหาของเซลล์ที่ถูกล็อก',
    note: 'ตามค่าเริ่มต้น เซลล์ทั้งหมดถูกล็อก แต่สามารถปลดล็อกได้',
  }),
  macAdvancedFilter: macAdvancedFilterFrom(ribbonMenusTh.ribbonMenu, {
    filterInPlace: 'กรองรายการในตำแหน่งเดิม',
    copyToLocation: 'คัดลอกไปยังตำแหน่งอื่น',
  }),
  macChrome: macChromeFrom(workbookChromeTh, formulaDialogsTh),
  macTextToColumnsWizard: macTextToColumnsWizardTh,
  macStyleDialog: macStyleDialogTh,
  macExternalLinksPane: macExternalLinksPaneTh,
  macWatchWindow: macWatchWindowTh,
  macAccessibilityPane: macAccessibilityPaneTh,
  macCfRulesDialog: macCfRulesDialogFrom(cellFormatTh.cfRulesDialog, {
    orderLabel: 'เปลี่ยนลำดับกฎ:',
    appliesToLabel: 'นำไปใช้กับ:',
  }),
  macHyperlinkDialog: macHyperlinkDialogFrom(editDialogsTh.hyperlinkDialog, {
    displayText: 'ข้อความที่จะแสดง:',
    selectionInDocument: '<<ส่วนที่เลือกในเอกสาร>>',
    screenTip: 'คำแนะนำบนหน้าจอ...',
    screenTipTitle: 'ตั้งค่าคำแนะนำบนหน้าจอของไฮเปอร์ลิงก์',
    screenTipLabel: 'ข้อความคำแนะนำบนหน้าจอ:',
    tabWeb: 'เว็บเพจหรือไฟล์',
    tabDocument: 'เอกสารนี้',
    tabEmail: 'ที่อยู่อีเมล',
    webIntro: 'ลิงก์ไปยังไฟล์หรือเว็บเพจที่มีอยู่',
    browse: 'เลือก...',
    browseUnavailable: 'ไม่สามารถเลือกไฟล์ในเบราว์เซอร์ได้ โปรดพิมพ์ที่อยู่',
    cellReference: 'พิมพ์การอ้างอิงเซลล์:',
    documentPlace: 'หรือเลือกตำแหน่งในเอกสารนี้:',
    emailAddress: 'ที่อยู่อีเมล:',
    subject: 'เรื่อง:',
    recentEmails: 'ที่อยู่อีเมลที่ใช้ล่าสุด:',
  }),
  macPageSetup: macPageSetupFrom(pageLayoutTh.pageSetup),
  macDataValidation: {
    inputCaption: 'เมื่อเลือกเซลล์ ให้แสดงข้อความป้อนข้อมูลนี้:',
    errorCaption: 'เมื่อผู้ใช้ป้อนข้อมูลที่ไม่ถูกต้อง ให้แสดงการแจ้งเตือนข้อผิดพลาดนี้:',
  },
  macCreateTable: { nameLabel: 'ตารางของคุณชื่ออะไร' },
  macGoToDialog: macGoToDialogFrom(editDialogsTh.goToDialog, {
    selectCaption: 'เลือก',
    optionsGroup: 'ตัวเลือก',
    noOptions: 'รายการที่เลือกนี้ไม่มีตัวเลือก',
  }),
  macPivotFieldSettings: macPivotFieldSettingsTh,
  macPivotOptions: macPivotOptionsTh,
  macPivotTableDialog: macPivotTableDialogFrom(dataObjectsTh.pivotTableDialog, {
    chooseConnection: 'เลือกการเชื่อมต่อ...',
  }),
  macRemoveDuplicates: macRemoveDuplicatesFrom(ribbonMenusTh.ribbonMenu),
  macSortDialog: macSortDialogFrom(ribbonMenusTh.ribbonMenu, {
    levelsHeading: 'เพิ่มระดับการเรียงลำดับ:',
    columnHeading: 'คอลัมน์',
    sortOnHeading: 'เรียงลำดับตาม',
    sortOnValues: editDialogsTh.findReplace.values,
    colorHeading: 'สี/ไอคอน',
    copyLevel: gridPopupsTh.contextMenu.copy,
    options: 'ตัวเลือก...',
  }),
  macErrorMenu: macErrorMenuTh,
  macSheetTabMenu: macSheetTabMenuTh,
  macFindReplace: macFindReplaceFrom(
    editDialogsTh.findReplace,
    'แยกอักขระครึ่งความกว้างและเต็มความกว้าง',
  ),
};
