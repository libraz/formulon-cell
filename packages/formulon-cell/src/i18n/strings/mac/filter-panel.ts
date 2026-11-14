import type { DatePeriodKey } from '../data-objects/types.js';
import type { GridPopupsStrings } from '../grid-popups/types.js';

type FilterDropdownStrings = GridPopupsStrings['filterDropdown'];

/** Wording of the Mac column filter panel: the condition list, search and footer. */
export interface MacFilterPanelStrings {
  chooseOne: string;
  equals: string;
  notEquals: string;
  contains: string;
  notContains: string;
  beginsWith: string;
  notBeginsWith: string;
  endsWith: string;
  notEndsWith: string;
  greaterThan: string;
  greaterThanOrEqual: string;
  lessThan: string;
  lessThanOrEqual: string;
  between: string;
  top10: string;
  bottom10: string;
  aboveAverage: string;
  belowAverage: string;
  /** The join of a two-condition filter. */
  and: string;
  or: string;
  /** A date column's comparisons and relative periods. */
  before: string;
  beforeOrEqual: string;
  after: string;
  afterOrEqual: string;
  periods: Record<DatePeriodKey, string>;
  /** The date condition whose value is a quarter or month of any year. */
  allDatesInPeriod: string;
  /** Units of the Top/Bottom count. */
  items: string;
  percent: string;
  search: string;
  selectAll: string;
  autoApply: string;
  apply: string;
  clear: string;
  sortAsc: string;
  sortDesc: string;
  byColor: string;
  colorNone: string;
  cellColor: string;
  fontColor: string;
  cellIcon: string;
  /** The last cell colour (no fill) and font colour (no colour of its own) of the By color lists. */
  noFill: string;
  fontAutomatic: string;
  /** The sort side's last By color entry, which opens the Sort dialog. */
  customSort: string;
  /** Bold section labels above the sort and filter controls, and the close button's name. */
  sortSection: string;
  filterSection: string;
  close: string;
  /** Placeholder row for cells with no value. */
  blanks: string;
  /** Why the conditions the host has no filter for cannot be used. */
  unavailable: string;
}

type MacFilterPanelExtras = Pick<
  MacFilterPanelStrings,
  | 'sortSection'
  | 'filterSection'
  | 'close'
  | 'sortAsc'
  | 'sortDesc'
  | 'byColor'
  | 'colorNone'
  | 'cellColor'
  | 'fontColor'
  | 'cellIcon'
  | 'noFill'
  | 'fontAutomatic'
  | 'customSort'
  | 'autoApply'
  | 'between'
  | 'top10'
  | 'bottom10'
  | 'aboveAverage'
  | 'belowAverage'
  | 'and'
  | 'or'
  | 'before'
  | 'beforeOrEqual'
  | 'after'
  | 'afterOrEqual'
  | 'periods'
  | 'allDatesInPeriod'
  | 'beginsWith'
  | 'notBeginsWith'
  | 'endsWith'
  | 'notEndsWith'
  | 'items'
  | 'percent'
  | 'unavailable'
>;

/** A locale's panel wording: its shared filter-dropdown text plus the lines only the Mac panel has. */
export const macFilterPanelFrom = (
  t: FilterDropdownStrings,
  extras: MacFilterPanelExtras,
  overrides: Partial<MacFilterPanelStrings> = {},
): MacFilterPanelStrings => ({
  chooseOne: t.conditionNone,
  equals: t.conditionEquals,
  notEquals: t.conditionNotEquals,
  contains: t.conditionContains,
  notContains: t.conditionNotContains,
  greaterThan: t.conditionGreaterThan,
  greaterThanOrEqual: t.conditionGreaterThanOrEqual,
  lessThan: t.conditionLessThan,
  lessThanOrEqual: t.conditionLessThanOrEqual,
  search: t.searchPlaceholder,
  selectAll: t.selectAll,
  apply: t.apply,
  clear: t.clear,
  blanks: t.blanks,
  ...extras,
  ...overrides,
});

/** Shared-derived wording for this locale; the panel builder overrides the fields the Excel bundle answers. */
const extrasEn: MacFilterPanelExtras = {
  and: 'And',
  or: 'Or',
  sortSection: 'Sort',
  filterSection: 'Filter',
  close: 'Close',
  sortAsc: 'Ascending',
  sortDesc: 'Descending',
  byColor: 'By color:',
  colorNone: 'None',
  cellColor: 'Cell Color',
  fontColor: 'Font Color',
  cellIcon: 'Cell Icon',
  noFill: 'No Fill',
  fontAutomatic: 'Automatic',
  customSort: 'Custom Sort...',
  autoApply: 'Auto Apply',
  between: 'Between',
  top10: 'Top 10',
  bottom10: 'Bottom 10',
  aboveAverage: 'Above Average',
  belowAverage: 'Below Average',
  items: 'Items',
  percent: 'Percent',
  unavailable: 'This feature is not implemented yet.',
  before: 'Before',
  beforeOrEqual: 'Before or Equal To',
  after: 'After',
  afterOrEqual: 'After or Equal To',
  beginsWith: 'Begins With',
  notBeginsWith: 'Does Not Begin With',
  endsWith: 'Ends With',
  notEndsWith: 'Does Not End With',
  allDatesInPeriod: 'All Dates in the Period',
  periods: {
    tomorrow: 'Tomorrow',
    today: 'Today',
    yesterday: 'Yesterday',
    nextWeek: 'Next Week',
    thisWeek: 'This Week',
    lastWeek: 'Last Week',
    nextMonth: 'Next Month',
    thisMonth: 'This Month',
    lastMonth: 'Last Month',
    nextQuarter: 'Next Quarter',
    thisQuarter: 'This Quarter',
    lastQuarter: 'Last Quarter',
    nextYear: 'Next Year',
    thisYear: 'This Year',
    lastYear: 'Last Year',
    yearToDate: 'Year to Date',
    quarter1: 'Quarter 1',
    quarter2: 'Quarter 2',
    quarter3: 'Quarter 3',
    quarter4: 'Quarter 4',
    january: 'January',
    february: 'February',
    march: 'March',
    april: 'April',
    may: 'May',
    june: 'June',
    july: 'July',
    august: 'August',
    september: 'September',
    october: 'October',
    november: 'November',
    december: 'December',
  },
};

/** Japanese Mac app wording, measured from Excel for Mac. */
const extrasJa: MacFilterPanelExtras = {
  and: 'および',
  or: 'または',
  sortSection: '並べ替え',
  filterSection: 'フィルター',
  close: '閉じる',
  sortAsc: '昇順',
  sortDesc: '降順',
  byColor: '色別:',
  colorNone: 'なし',
  cellColor: 'セルの色',
  fontColor: 'フォントの色',
  cellIcon: 'セルのアイコン',
  noFill: '塗りつぶしなし',
  fontAutomatic: '自動',
  customSort: 'ユーザー設定の並べ替え...',
  autoApply: '自動適用',
  between: '指定の範囲内',
  top10: '上位 10 項目',
  bottom10: '下位 10 項目',
  aboveAverage: '平均より上',
  belowAverage: '平均より下',
  items: '項目',
  percent: 'パーセント',
  unavailable: 'この機能は未実装です。',
  before: 'より前',
  beforeOrEqual: '以前',
  after: 'より後',
  afterOrEqual: '以後',
  beginsWith: 'で始まる',
  notBeginsWith: 'で始まらない',
  endsWith: 'で終わる',
  notEndsWith: 'で終わらない',
  allDatesInPeriod: '期間内の全日付...',
  periods: {
    tomorrow: '明日',
    today: '今日',
    yesterday: '昨日',
    nextWeek: '来週',
    thisWeek: '今週',
    lastWeek: '先週',
    nextMonth: '来月',
    thisMonth: '今月',
    lastMonth: '先月',
    nextQuarter: '来四半期',
    thisQuarter: '今四半期',
    lastQuarter: '前四半期',
    nextYear: '来年',
    thisYear: '今年',
    lastYear: '昨年',
    yearToDate: '今年の初めから今日まで',
    quarter1: '四半期 1',
    quarter2: '四半期 2',
    quarter3: '四半期 3',
    quarter4: '四半期 4',
    january: '1 月',
    february: '2 月',
    march: '3 月',
    april: '4 月',
    may: '5 月',
    june: '6 月',
    july: '7 月',
    august: '8 月',
    september: '9 月',
    october: '10 月',
    november: '11 月',
    december: '12 月',
  },
};

/** Shared-derived wording for this locale; the panel builder overrides the fields the Excel bundle answers. */
const extrasDe: MacFilterPanelExtras = {
  and: 'Und',
  or: 'Oder',
  sortSection: 'Sortieren',
  filterSection: 'Filter',
  close: 'Schließen',
  sortAsc: 'Aufsteigend',
  sortDesc: 'Absteigend',
  byColor: 'Nach Farbe:',
  colorNone: 'Ohne',
  cellColor: 'Zellenfarbe',
  fontColor: 'Schriftfarbe',
  cellIcon: 'Zellensymbol',
  noFill: 'Keine Füllung',
  fontAutomatic: 'Automatisch',
  customSort: 'Benutzerdefiniertes Sortieren...',
  autoApply: 'Automatisch anwenden',
  between: 'zwischen',
  top10: 'Top 10',
  bottom10: 'Untere 10',
  aboveAverage: 'über dem Durchschnitt',
  belowAverage: 'unter dem Durchschnitt',
  items: 'Elemente',
  percent: 'Prozent',
  unavailable: 'Diese Funktion ist noch nicht implementiert.',
  before: 'vor',
  beforeOrEqual: 'vor oder gleich',
  after: 'nach',
  afterOrEqual: 'nach oder gleich',
  beginsWith: 'Beginnt mit',
  notBeginsWith: 'Beginnt nicht mit',
  endsWith: 'Endet mit',
  notEndsWith: 'Endet nicht mit',
  allDatesInPeriod: 'Alle Datumswerte im Zeitraum',
  periods: {
    tomorrow: 'Morgen',
    today: 'Heute',
    yesterday: 'Gestern',
    nextWeek: 'Nächste Woche',
    thisWeek: 'Diese Woche',
    lastWeek: 'Letzte Woche',
    nextMonth: 'Nächster Monat',
    thisMonth: 'Dieser Monat',
    lastMonth: 'Letzter Monat',
    nextQuarter: 'Nächstes Quartal',
    thisQuarter: 'Dieses Quartal',
    lastQuarter: 'Letztes Quartal',
    nextYear: 'Nächstes Jahr',
    thisYear: 'Dieses Jahr',
    lastYear: 'Letztes Jahr',
    yearToDate: 'Jahr bis zum aktuellen Datum',
    quarter1: 'Quartal 1',
    quarter2: 'Quartal 2',
    quarter3: 'Quartal 3',
    quarter4: 'Quartal 4',
    january: 'Januar',
    february: 'Februar',
    march: 'März',
    april: 'April',
    may: 'Mai',
    june: 'Juni',
    july: 'Juli',
    august: 'August',
    september: 'September',
    october: 'Oktober',
    november: 'November',
    december: 'Dezember',
  },
};

/** Shared-derived wording for this locale; the panel builder overrides the fields the Excel bundle answers. */
const extrasFr: MacFilterPanelExtras = {
  and: 'Et',
  or: 'Ou',
  sortSection: 'Trier',
  filterSection: 'Filtrer',
  close: 'Fermer',
  sortAsc: 'Croissant',
  sortDesc: 'Décroissant',
  byColor: 'Par couleur :',
  colorNone: 'Aucun',
  cellColor: 'Couleur de cellule',
  fontColor: 'Couleur de police',
  cellIcon: 'Icône de cellule',
  noFill: 'Aucun remplissage',
  fontAutomatic: 'Automatique\u00a0',
  customSort: 'Tri personnalisé...',
  autoApply: 'Appliquer automatiquement',
  between: 'Compris entre',
  top10: '10 premiers',
  bottom10: '10 derniers',
  aboveAverage: 'Au-dessus de la moyenne',
  belowAverage: 'En dessous de la moyenne',
  items: 'Éléments',
  percent: 'Pourcentage',
  unavailable: 'Cette fonctionnalité n’est pas encore implémentée.',
  before: 'Avant',
  beforeOrEqual: 'Antérieur ou égal au',
  after: 'Après',
  afterOrEqual: 'Postérieur ou égal au',
  beginsWith: 'Commence par',
  notBeginsWith: 'Ne commence pas par',
  endsWith: 'Se termine par',
  notEndsWith: 'Ne se termine pas par',
  allDatesInPeriod: 'Toutes les dates de cette période',
  periods: {
    tomorrow: 'Demain',
    today: 'Aujourd’hui',
    yesterday: 'Hier',
    nextWeek: 'La semaine prochaine',
    thisWeek: 'Cette semaine',
    lastWeek: 'La semaine dernière',
    nextMonth: 'Le mois prochain',
    thisMonth: 'Ce mois',
    lastMonth: 'Le mois dernier',
    nextQuarter: 'Le trimestre prochain',
    thisQuarter: 'Ce trimestre',
    lastQuarter: 'Le trimestre dernier',
    nextYear: 'L’année prochaine',
    thisYear: 'Cette année',
    lastYear: 'L’année dernière',
    yearToDate: 'Année à ce jour',
    quarter1: 'Trimestre 1',
    quarter2: 'Trimestre 2',
    quarter3: 'Trimestre 3',
    quarter4: 'Trimestre 4',
    january: 'janvier',
    february: 'février',
    march: 'mars',
    april: 'avril',
    may: 'mai',
    june: 'juin',
    july: 'juillet',
    august: 'août',
    september: 'septembre',
    october: 'octobre',
    november: 'novembre',
    december: 'décembre',
  },
};

/** Shared-derived wording for this locale; the panel builder overrides the fields the Excel bundle answers. */
const extrasZh: MacFilterPanelExtras = {
  and: '与',
  or: '或者',
  sortSection: '排序',
  filterSection: '筛选器',
  close: '关闭',
  sortAsc: '升序',
  sortDesc: '降序',
  byColor: '按颜色:',
  colorNone: '无',
  cellColor: '单元格颜色',
  fontColor: '字体颜色',
  cellIcon: '单元格图标',
  noFill: '无填充',
  fontAutomatic: '自动',
  customSort: '自定义排序...',
  autoApply: '自动应用',
  between: '之间',
  top10: '前 10 项',
  bottom10: '后 10 项',
  aboveAverage: '高于平均值',
  belowAverage: '低于平均值',
  items: '项目',
  percent: '百分比',
  unavailable: '此功能尚未实现。',
  before: '之前',
  beforeOrEqual: '早于或等于',
  after: '之后',
  afterOrEqual: '晚于或等于',
  beginsWith: '开头是',
  notBeginsWith: '开头不是',
  endsWith: '结尾是',
  notEndsWith: '结尾不是',
  allDatesInPeriod: '该期间的所有日期',
  periods: {
    tomorrow: '明天',
    today: '今天',
    yesterday: '昨天',
    nextWeek: '下周',
    thisWeek: '本周',
    lastWeek: '上周',
    nextMonth: '下月',
    thisMonth: '本月',
    lastMonth: '上个月',
    nextQuarter: '下季度',
    thisQuarter: '本季度',
    lastQuarter: '上季度',
    nextYear: '明年',
    thisYear: '今年',
    lastYear: '去年',
    yearToDate: '本年度截止到现在',
    quarter1: '第一季度',
    quarter2: '第二季度',
    quarter3: '第三季度',
    quarter4: '第四季度',
    january: '一月',
    february: '二月',
    march: '三月',
    april: '四月',
    may: '五月',
    june: '六月',
    july: '七月',
    august: '八月',
    september: '九月',
    october: '十月',
    november: '十一月',
    december: '十二月',
  },
};

/** Shared-derived wording for this locale; the panel builder overrides the fields the Excel bundle answers. */
const extrasKo: MacFilterPanelExtras = {
  and: '그리고',
  or: '논리합',
  sortSection: '정렬',
  filterSection: '필터',
  close: '닫기',
  sortAsc: '오름차순',
  sortDesc: '내림차순',
  byColor: '색 기준:',
  colorNone: '없음',
  cellColor: '셀 색',
  fontColor: '글꼴 색',
  cellIcon: '셀 아이콘',
  noFill: '채우기 없음',
  fontAutomatic: '자동',
  customSort: '사용자 지정 정렬...',
  autoApply: '자동 적용',
  between: '사이',
  top10: '상위 10개',
  bottom10: '하위 10개',
  aboveAverage: '평균 초과인 경우',
  belowAverage: '평균 미만인 경우',
  items: '항목',
  percent: '백분율',
  unavailable: '이 기능은 아직 구현되지 않았습니다.',
  before: '앞',
  beforeOrEqual: '이전 또는 같음',
  after: '이후',
  afterOrEqual: '이후 또는 같음',
  beginsWith: '시작 문자',
  notBeginsWith: '제외할 시작 문자',
  endsWith: '끝 문자',
  notEndsWith: '제외할 끝 문자',
  allDatesInPeriod: '기간 내의 모든 날짜',
  periods: {
    tomorrow: '내일',
    today: '오늘',
    yesterday: '어제',
    nextWeek: '다음 주',
    thisWeek: '이번 주',
    lastWeek: '지난 주',
    nextMonth: '다음 달',
    thisMonth: '이번 달',
    lastMonth: '지난 달',
    nextQuarter: '다음 분기',
    thisQuarter: '이번 분기',
    lastQuarter: '지난 분기',
    nextYear: '내년',
    thisYear: '올해',
    lastYear: '작년',
    yearToDate: '연간 누계',
    quarter1: '1분기',
    quarter2: '2분기',
    quarter3: '3분기',
    quarter4: '4분기',
    january: '1월',
    february: '2월',
    march: '3월',
    april: '4월',
    may: '5월',
    june: '6월',
    july: '7월',
    august: '8월',
    september: '9월',
    october: '10월',
    november: '11월',
    december: '12월',
  },
};

/** Shared-derived wording for this locale; the panel builder overrides the fields the Excel bundle answers. */
const extrasTh: MacFilterPanelExtras = {
  and: 'และ',
  or: 'หรือ',
  sortSection: 'เรียงลำดับ',
  filterSection: 'ตัวกรอง',
  close: 'ปิด',
  sortAsc: 'จากน้อยไปมาก',
  sortDesc: 'จากมากไปน้อย',
  byColor: 'ตามสี:',
  colorNone: 'ไม่มี',
  cellColor: 'สีเซลล์',
  fontColor: 'สีฟอนต์',
  cellIcon: 'ไอคอนเซลล์',
  noFill: 'ไม่เติม',
  fontAutomatic: 'อัตโนมัติ',
  customSort: 'เรียงลำดับแบบกำหนดเอง...',
  autoApply: 'ใช้อัตโนมัติ',
  between: 'ระหว่าง',
  top10: '10 อันดับแรก',
  bottom10: '10 อันดับท้าย',
  aboveAverage: 'สูงกว่าค่าเฉลี่ย',
  belowAverage: 'ต่ำกว่าค่าเฉลี่ย',
  items: 'รายการ',
  percent: 'เปอร์เซ็นต์',
  unavailable: 'ฟีเจอร์นี้ยังไม่ได้ใช้งาน',
  before: 'ก่อนหน้า',
  beforeOrEqual: 'ก่อนหรือเท่ากับ',
  after: 'หลัง',
  afterOrEqual: 'หลังหรือเท่ากับ',
  beginsWith: 'เริ่มต้นด้วย',
  notBeginsWith: 'ไม่ได้เริ่มต้นด้วย',
  endsWith: 'ลงท้ายด้วย',
  notEndsWith: 'ไม่ได้ลงท้ายด้วย',
  allDatesInPeriod: 'วันที่ทั้งหมดในช่วงเวลา',
  periods: {
    tomorrow: 'พรุ่งนี้',
    today: 'วันนี้',
    yesterday: 'เมื่อวานนี้',
    nextWeek: 'สัปดาห์หน้า',
    thisWeek: 'สัปดาห์นี้',
    lastWeek: 'สัปดาห์ที่แล้ว',
    nextMonth: 'เดือนถัดไป',
    thisMonth: 'เดือนนี้',
    lastMonth: 'เดือนที่แล้ว',
    nextQuarter: 'ไตรมาสถัดไป',
    thisQuarter: 'ไตรมาสนี้',
    lastQuarter: 'ไตรมาสที่แล้ว',
    nextYear: 'ปีหน้า',
    thisYear: 'ปีนี้',
    lastYear: 'ปีที่แล้ว',
    yearToDate: 'เริ่มต้นปีจนถึงปัจจุบัน',
    quarter1: 'ไตรมาส 1',
    quarter2: 'ไตรมาส 2',
    quarter3: 'ไตรมาส 3',
    quarter4: 'ไตรมาส 4',
    january: 'มกราคม',
    february: 'กุมภาพันธ์',
    march: 'มีนาคม',
    april: 'เมษายน',
    may: 'พค',
    june: 'มิถุนายน',
    july: 'กรกฎาคม',
    august: 'สิงหาคม',
    september: 'กันยายน',
    october: 'ตุลาคม',
    november: 'พฤศจิกายน',
    december: 'ธันวาคม',
  },
};

export const macFilterPanelEn = (t: FilterDropdownStrings): MacFilterPanelStrings =>
  macFilterPanelFrom(t, extrasEn, {
    chooseOne: 'Choose One',
    notContains: 'Does Not Contain',
    selectAll: '(Select All)',
    apply: 'Apply Filter',
    clear: 'Clear Filter',
  });

export const macFilterPanelJa = (t: FilterDropdownStrings): MacFilterPanelStrings =>
  macFilterPanelFrom(t, extrasJa, {
    chooseOne: '1 つ選択してください',
    equals: '指定の値に等しい',
    notEquals: '指定の値に等しくない',
    greaterThan: '指定の値より大きい',
    greaterThanOrEqual: '指定の値以上',
    lessThan: '指定の値より小さい',
    lessThanOrEqual: '指定の値以下',
    search: '検索',
    selectAll: '(すべて選択)',
    apply: 'フィルターの適用',
    clear: 'フィルターのクリア',
    contains: 'を含む',
    notContains: 'を含まない',
    blanks: '(空白セル)',
  });

export const macFilterPanelDe = (t: FilterDropdownStrings): MacFilterPanelStrings =>
  macFilterPanelFrom(t, extrasDe, {
    chooseOne: 'Eine Option auswählen',
    apply: 'Filter anwenden',
    clear: 'Filter entfernen',
    blanks: '(Leere)',
    close: 'Schließen',
    sortAsc: 'Aufsteigend',
    sortDesc: 'Absteigend',
    byColor: 'Nach Farbe:',
    cellColor: 'Zellenfarbe',
    cellIcon: 'Zellensymbol',
    autoApply: 'Automatisch anwenden',
    top10: 'oberer Bereich',
    bottom10: 'unterer Bereich',
    items: 'Elemente',
    percent: 'Prozent',
    equals: 'gleich',
    notEquals: 'nicht gleich',
    greaterThan: 'größer als',
    greaterThanOrEqual: 'größer als oder gleich',
    lessThan: 'kleiner als',
    lessThanOrEqual: 'kleiner als oder gleich',
  });

export const macFilterPanelFr = (t: FilterDropdownStrings): MacFilterPanelStrings =>
  macFilterPanelFrom(t, extrasFr, {
    chooseOne: 'Choisir',
    selectAll: '(Sélectionner tout)',
    apply: 'Appliquer le filtre',
    clear: 'Effacer le filtre',
    blanks: '(Vides)',
    sortAsc: 'Croissant',
    sortDesc: 'Décroissant',
    byColor: 'Par couleur :',
    cellColor: 'Couleur de cellule',
    cellIcon: 'Icône de cellule',
    autoApply: 'Appliquer automatiquement',
    top10: '10 premiers',
    bottom10: '10 derniers',
    items: 'Éléments',
    equals: 'Égal à',
    notEquals: 'Est différent de',
  });

export const macFilterPanelZh = (t: FilterDropdownStrings): MacFilterPanelStrings =>
  macFilterPanelFrom(t, extrasZh, {
    chooseOne: '选择一个',
    contains: '包含',
    notContains: '不包含',
    selectAll: '(全选)',
    apply: '应用筛选器',
    clear: '清除筛选器',
    blanks: '(空白)',
    close: '关闭',
    sortAsc: '升序',
    sortDesc: '降序',
    byColor: '按颜色:',
    colorNone: '无',
    cellColor: '单元格颜色',
    cellIcon: '单元格图标',
    autoApply: '自动应用',
    top10: '前 10 个',
    bottom10: '最小的 10 项',
    percent: '百分比',
  });

export const macFilterPanelKo = (t: FilterDropdownStrings): MacFilterPanelStrings =>
  macFilterPanelFrom(t, extrasKo, {
    chooseOne: '항목 선택',
    contains: '포함',
    apply: '필터 적용',
    clear: '필터 해제',
    blanks: '(필드 값 없음)',
    close: '닫기',
    sortAsc: '오름차순',
    sortDesc: '내림차순',
    byColor: '색상 기준:',
    cellColor: '셀 색',
    cellIcon: '셀 아이콘',
    autoApply: '자동 적용',
    top10: '상위 10개',
    bottom10: '하위 10개',
    selectAll: '[모두 선택]',
  });

export const macFilterPanelTh = (t: FilterDropdownStrings): MacFilterPanelStrings =>
  macFilterPanelFrom(t, extrasTh, {
    chooseOne: 'เลือกหนึ่งรายการ',
    contains: 'มี',
    notContains: 'ไม่มี',
    selectAll: '(เลือกทั้งหมด)',
    apply: 'ใช้ตัวกรอง',
    clear: 'ล้างตัวกรอง',
    blanks: '(ว่าง)',
    byColor: 'ตามสี:',
    cellColor: 'สีของเซลล์',
    cellIcon: 'ไอคอนของเซลล์',
    autoApply: 'นำไปใช้โดยอัตโนมัติ',
    top10: '10 อันดับแรก',
    bottom10: '10 รายการสุดท้าย',
    items: 'รายการ',
  });
