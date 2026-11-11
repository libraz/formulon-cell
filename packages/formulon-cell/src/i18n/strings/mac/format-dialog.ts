import type { CellFormatStrings } from '../cell-format/types.js';

type FormatDialogStrings = CellFormatStrings['formatDialog'];

/** The `formatDialog` keys the Mac host can override. Each locale starts from its own shared text and overrides the fields the Excel bundle answers. */
const MAC_FORMAT_DIALOG_KEYS = [
  'categoryLabel',
  'fontFamily',
  'fontStyle',
  'fontSize',
  'fontUnderline',
  'color',
  'fontEffects',
  'descGeneral',
  'descFixed',
  'descCurrency',
  'descAccounting',
  'descPercent',
  'descDate',
  'descTime',
  'descText',
  'descOther',
  'decimals',
  'thousandsSeparator',
  'negativeNumbers',
  'symbol',
  'patternType',
  'languageLocation',
  'calendarType',
  'specialFormatLabels',
  'horizontalAlign',
  'verticalAlign',
  'indent',
  'textControl',
  'direction',
  'rotation',
  'verticalText',
  'borderStyle',
  'borderColor',
  'borderStyleHair',
  'borderStyleDotted',
  'borderStyleDashDotDot',
  'borderStyleDashDot',
  'borderStyleDashed',
  'borderStyleThin',
  'borderStyleMediumDashDotDot',
  'borderStyleSlantDashDot',
  'borderStyleMediumDashDot',
  'borderStyleMediumDashed',
  'borderStyleMedium',
  'borderStyleThick',
  'borderStyleDouble',
  'fill',
  'fillPatternColor',
  'fillPatternStyle',
] as const satisfies readonly (keyof FormatDialogStrings)[];

export const macFormatDialogFrom = (dialog: FormatDialogStrings): Partial<FormatDialogStrings> =>
  Object.fromEntries(MAC_FORMAT_DIALOG_KEYS.map((key) => [key, dialog[key]]));

/** Mac app wording from the Excel bundle; not answered by it: categoryLabel, fontFamily, fontStyle, fontSize, fontUnderline, color, and 15 more. */
export const macFormatDialogDe: Partial<FormatDialogStrings> = {
  descGeneral: 'Standardzellen haben kein bestimmtes Zahlenformat.',
  descFixed:
    'Zahl wird für die allgemeine Anzeige von Zahlen verwendet. Währung und Buchhaltung bieten spezielle Formate für monetäre Werte.',
  descCurrency:
    'Währung wird für allgemeine monetäre Werte verwendet. Benutzen Sie Buchhaltung, um die Dezimalstellen in einer Spalte auszurichten.',
  descAccounting:
    'Buchhaltung richtet die Währungssymbole und die Dezimalstellen in einer Spalte aus.',
  descPercent:
    'Prozentformate multiplizieren den Zellenwert mit 100 und zeigen das Ergebnis mit einem Prozentzeichen an.',
  descDate:
    'Datumsformate zeigen die fortlaufenden Zahlen des Datums und der Uhrzeit als Datumswerte an. Datumsformate, die mit einem Sternchen (*) beginnen, reagieren auf Änderungen an regionalen Datums- und Uhrzeiteinstellungen, die für das Betriebssystem angegeben sind. Formate ohne Sternchen sind von Betriebssystemeinstellungen nicht betroffen.',
  descTime:
    'Zeitformate zeigen die fortlaufenden Zahlen des Datums und der Uhrzeit als Datumswerte an. Zeitformate, die mit einem Sternchen (*) beginnen, reagieren auf Änderungen an regionalen Datums- und Uhrzeiteinstellungen, die für das Betriebssystem angegeben sind. Formate ohne Sternchen sind von Betriebssystemeinstellungen nicht betroffen.',
  descText:
    'Als Text formatierte Zellen behandeln auch Zahlen als Text. Der Zellinhalt wird genauso angezeigt wie eingegeben.',
  descOther: 'Sonderformate sind hilfreich, um Listen und Datenbankwerte zu formatieren.',
  decimals: 'Dezimalstellen:',
  symbol: 'Symbol:',
  languageLocation: 'Sprache (Ort):',
  calendarType: 'Kalendertyp:',
  verticalAlign: 'Vertikale Ausrichtung:',
  textControl: 'Textsteuerung:',
  rotation: 'Grad:',
  borderStyleDashDotDot: 'Strich Punkt Punkt',
  borderStyleDashDot: 'Strich Punkt',
  borderStyleDashed: 'Gestrichelt',
  borderStyleThin: 'Dünn',
  borderStyleMediumDashDotDot: 'Strich Punkt Punkt Mittel',
  borderStyleSlantDashDot: 'Schrägstrich Strich Punkt',
  borderStyleMediumDashDot: 'Strich Punkt Mittel',
  borderStyleMediumDashed: 'Gestrichelt Mittel',
  borderStyleMedium: 'Mittel',
  borderStyleDouble: 'Doppelt',
  fillPatternColor: 'Musterfarbe:',
  fillPatternStyle: 'Musterformat:',
};

/** Mac app wording from the Excel bundle; not answered by it: categoryLabel, fontFamily, fontStyle, fontSize, fontUnderline, color, and 18 more. */
export const macFormatDialogFr: Partial<FormatDialogStrings> = {
  descGeneral: 'Les formules de format Standard n’ont pas de format de nombre spécifique.',
  descFixed:
    "La catégorie Nombre est utilisée pour l'affichage général des nombres. Les catégories Monétaire et Comptabilité offrent des formatages spécialisés pour des valeurs monétaires.",
  descCurrency:
    'Les formats Monétaire sont utilisés pour des valeurs monétaires générales. Utilisez les formats Comptabilité pour aligner les décimaux dans une colonne.',
  descAccounting:
    'Les formats Comptabilité alignent les symboles monétaires et les décimaux dans une colonne.',
  descPercent:
    'Les formats Pourcentage multiplient la valeur de la cellule par 100 et affichent le résultat avec le symbole pourcentage.',
  descDate:
    "Les formats Date affichent les numéros de série date et heure comme valeurs date. Les formats Date précédés d'un astérisque (*) répondent aux changements dans les paramètres régionaux de date et d'heure qui sont spécifiés pour le système d'exploitation. Les formats sans astérisque ne sont pas affectés par le système d'exploitation.",
  descTime:
    "Les formats Heure affichent les numéros de série date et heure comme valeurs date. Les formats Heure qui commencent par un astérisque (*) répondent aux changements dans les paramètres régionaux de date et d'heure qui sont spécifiés pour le système d'exploitation. Les formats sans astérisque ne sont pas affectés par le système d'exploitation.",
  descText:
    "Les cellules de format Texte sont traitées comme du texte même si c'est un nombre qui se trouve dans la cellule. La cellule est affichée exactement comme elle a été entrée.",
  descOther:
    'Les formats Spécial sont utilisés pour contrôler des valeurs de liste et de base de données.',
  decimals: 'Nombre de décimales :',
  symbol: 'Symbole :',
  languageLocation: 'Langue (emplacement) :',
  calendarType: 'Type de calendrier :',
  verticalAlign: 'Alignement vertical :',
  textControl: 'Contrôle du texte :',
  rotation: 'Degrés :',
  borderStyleDashDotDot: 'Tiret-point-point',
  borderStyleDashDot: 'Tiret-point',
  borderStyleMediumDashDotDot: 'Tiret-point-point moyen',
  borderStyleSlantDashDot: 'Tiret-point incliné',
  borderStyleMediumDashDot: 'Tiret-point moyen',
  borderStyleMediumDashed: 'Tiret moyen',
  borderStyleDouble: 'Double',
  fillPatternColor: 'Couleur de motif :',
  fillPatternStyle: 'Style de motif :',
};

/** Mac app wording from the Excel bundle; not answered by it: categoryLabel, fontFamily, fontStyle, fontSize, fontUnderline, color, and 18 more. */
export const macFormatDialogKo: Partial<FormatDialogStrings> = {
  descGeneral: '일반 셀 서식에서는 특정 번호 서식을 지정하지 않습니다.',
  descFixed:
    '숫자 서식은 일반적인 숫자를 나타내는 데 사용됩니다. 통화 및 회계 표시 형식에는 화폐 가치에 대한 특수 서식이 있습니다.',
  descCurrency:
    '통화 서식은 일반 통화 수치에 사용합니다. 회계 서식을 사용하면 소수점에 맞추어 열이 정렬됩니다.',
  descAccounting: '회계 서식을 사용하면 통화 기호와 소수점에 맞추어 열이 정렬됩니다.',
  descPercent: '백분율 서식을 사용하면 셀 값에 100을 곱한 값이 백분율 기호와 함께 나타납니다.',
  descDate:
    '날짜 서식으로 날짜와 시간에 해당하는 일련의 숫자를 날짜값으로 나타낼 수 있습니다. 별표(*)로 시작되는 날짜 서식은 운영 체제에 지정된 국가별 날짜 및 시간 설정에 따라 변경됩니다. 별표가 없는 서식은 운영 체제 설정의 영향을 받지 않습니다.',
  descTime:
    '시간 서식으로 날짜와 시간에 해당하는 일련의 숫자를 날짜값으로 나타낼 수 있습니다. 별표(*)로 시작되는 시간 서식은 운영 체제에 지정된 국가별 날짜 및 시간 설정에 따라 변경됩니다. 별표가 없는 서식은 운영 체제 설정의 영향을 받지 않습니다.',
  descText: '텍스트 서식으로 지정된 셀에 있는 숫자는 텍스트로 처리되므로 입력한 대로 표시됩니다.',
  descOther: '기타 서식은 목록 및 데이터베이스 값을 찾을 때에 사용합니다.',
  decimals: '소수 자릿수:',
  thousandsSeparator: '1000 단위 구분 기호(,) 사용',
  symbol: '기호:',
  languageLocation: '언어(위치):',
  verticalAlign: '세로 맞춤:',
  textControl: '텍스트 조정:',
  rotation: '차수:',
  borderStyleDashDotDot: '파선-점선-점선',
  borderStyleDashDot: '파선-점선',
  borderStyleDashed: '파선',
  borderStyleMediumDashDotDot: '중간 굵기 이점 쇄선',
  borderStyleSlantDashDot: '기울어진 일점 쇄선',
  borderStyleMediumDashDot: '중간 굵기 일점 쇄선',
  borderStyleMediumDashed: '중간 굵기 파선',
  fillPatternColor: '패턴 색:',
  fillPatternStyle: '무늬 스타일:',
};

/** Mac app wording from the Excel bundle; not answered by it: categoryLabel, fontFamily, fontStyle, fontSize, fontUnderline, color, and 14 more. */
export const macFormatDialogTh: Partial<FormatDialogStrings> = {
  descGeneral: 'เซลล์ที่มีรูปแบบทั่วไปไม่มีการระบุรูปแบบตัวเลขเฉพาะ',
  descFixed:
    "'ตัวเลข' ถูกใช้สำหรับแสดงผลตัวเลขทั่วไป สำหรับ 'สกุลเงิน' และ 'บัญชี' จะมีการจัดรูปแบบที่เฉพาะสำหรับค่าทางการเงิน",
  descCurrency:
    "รูปแบบ 'สกุลเงิน' ถูกใช้สำหรับค่าทางการเงินทั่วไป ให้ใช้รูปแบบ 'บัญชี' เพื่อจัดแนวจุดทศนิยมและสัญลักษณ์สกุลเงินในคอลัมน์",
  descAccounting: "รูปแบบ 'บัญชี' จัดตำแหน่งสัญลักษณ์สกุลเงินและจุดทศนิยมในคอลัมน์",
  descPercent: 'รูปแบบเปอร์เซ็นต์จะคูณค่าในเซลล์ด้วย 100 และแสดงผลลัพธ์ด้วยสัญลักษณ์เปอร์เซ็นต์',
  descDate:
    "รูปแบบ 'วันที่' จะแสดงตัวเลขอนุกรมของวันที่และเวลาเป็นค่าวันที่ รูปแบบ 'วันที่' ที่เริ่มต้นด้วยเครื่องหมายดอกจัน (*) จะตอบสนองต่อการเปลี่ยนแปลงในการตั้งค่าภูมิภาคของวันที่และเวลาที่ระบุไว้สำหรับระบบปฏิบัติการ รูปแบบที่ไม่มีเครื่องหมายดอกจันจะไม่ได้รับผลกระทบจากการตั้งค่าของระบบปฏิบัติการ",
  descTime:
    "รูปแบบ 'เวลา' แสดงตัวเลขอนุกรมของวันที่และเวลาเป็นค่าเวลา รูปแบบ 'เวลา' ที่เริ่มต้นด้วยเครื่องหมายดอกจัน (*) จะตอบสนองต่อการเปลี่ยนแปลงในการตั้งค่าภูมิภาคของวันที่และเวลาที่ระบุไว้สำหรับระบบปฏิบัติการ รูปแบบที่ไม่มีเครื่องหมายดอกจันจะไม่ได้รับผลกระทบจากการตั้งค่าของระบบปฏิบัติการ",
  descText:
    'เซลล์ที่มีรูปแบบเป็นข้อความ จะถือว่าทุกอย่างเป็นข้อความ ถึงแม้ว่าจะมีตัวเลขอยู่ในเซลล์ เซลล์จะแสดงผลเหมือนสิ่งที่ป้อนเข้าไป',
  descOther: "รูปแบบ 'พิเศษ' มีประโยชน์ในการติดตามรายการและค่าจากฐานข้อมูล",
  thousandsSeparator: 'ใช้ตัวคั่นหลักพัน (,)',
  symbol: 'สัญลักษณ์:',
  languageLocation: 'ภาษา (ตำแหน่งที่ตั้ง):',
  calendarType: 'ชนิดปฏิทิน:',
  verticalAlign: 'การจัดแนวตามแนวตั้ง',
  textControl: 'การควบคุมข้อความ:',
  rotation: 'องศา:',
  borderStyleDotted: 'จุด',
  borderStyleDashDotDot: 'เส้นประจุดจุด',
  borderStyleDashDot: 'เส้นประจุด',
  borderStyleDashed: 'เส้นประ',
  borderStyleThin: 'บาง',
  borderStyleMediumDashDotDot: 'เส้นประยาวปานกลางจุดจุด',
  borderStyleSlantDashDot: 'เส้นประเอียง-จุด',
  borderStyleMediumDashDot: 'เส้นประยาวปานกลางจุด',
  borderStyleMediumDashed: 'เส้นประยาวปานกลาง',
  borderStyleMedium: 'ปานกลาง',
  borderStyleThick: 'หนา',
  fillPatternColor: 'สีลวดลาย:',
  fillPatternStyle: 'สไตล์ลวดลาย:',
};

/** Mac app wording from the Excel bundle; not answered by it: categoryLabel, fontFamily, fontStyle, fontSize, fontUnderline, color, and 19 more. */
export const macFormatDialogZh: Partial<FormatDialogStrings> = {
  descGeneral: '常规单元格格式不包含任何特定的数字格式。',
  descFixed: '数值格式用于一般数字的表示。货币和会计格式则提供货币值计算的专用格式。',
  descCurrency: '货币格式用于表示一般货币数值。会计格式可以对一列数值进行小数点对齐。',
  descAccounting: '会计格式可对一列数值进行货币符号和小数点对齐。',
  descPercent: '百分比格式将单元格值乘以 100，并以百分数符号显示结果。',
  descDate:
    '日期格式将日期和时间系列数值显示为日期值。以星号(*)开头的日期格式响应操作系统特定的区域日期和时间设置的更改。不带星号的格式不受操作系统设置的影响。',
  descTime:
    '时间格式将日期和时间系列数值显示为时间值。以星号(*)开头的时间格式响应操作系统特定的区域日期和时间设置的更改。不带星号的格式不受操作系统设置的影响。',
  descText: '在文本单元格格式中，数字作为文本处理。单元格显示的内容与输入的内容完全一致。',
  descOther: '特殊格式可用于跟踪数据列表及数据库的值。',
  decimals: '小数位数:',
  thousandsSeparator: '使用千位分隔符(,)',
  languageLocation: '语言(位置):',
  calendarType: '日历类型:',
  verticalAlign: '垂直对齐方式:',
  textControl: '文本控制:',
  rotation: '度:',
  borderStyleDashDotDot: '划线-点-点',
  borderStyleDashDot: '划线-点',
  borderStyleMediumDashDotDot: '中划线-点-点',
  borderStyleSlantDashDot: '斜划线-点',
  borderStyleMediumDashDot: '中划线-点',
  borderStyleMediumDashed: '中划线',
  fillPatternColor: '图案颜色:',
  fillPatternStyle: '图案样式:',
};

/** Mac app wording from the Excel bundle. */
export const macFormatDialogEn: Partial<FormatDialogStrings> = {
  decimals: 'Decimal places:',
  symbol: 'Symbol:',
  languageLocation: 'Language (Location):',
  calendarType: 'Calendar type:',
  textControl: 'Text control:',
  borderStyleMediumDashed: 'Medium Dashed',
  fillPatternColor: 'Pattern color:',
  fillPatternStyle: 'Pattern style:',
  descGeneral: 'General format cells have no specific number format.',
  descFixed:
    'Number is used for general display of numbers.  Currency and Accounting offer specialized formatting for monetary value.',
  descCurrency:
    'Currency formats are used for general monetary values.  Use Accounting formats to align decimal points in a column.',
  descAccounting: 'Accounting formats line up the currency symbols and decimal points in a column.',
  descPercent:
    'Percentage formats multiply the cell value by 100 and display the result with a percent symbol.',
  descDate:
    'Date formats display date and time serial numbers as date values.  Date formats that begin with an asterisk (*) respond to changes in regional date and time settings that are specified for the operating system. Formats without an asterisk are not affected by operating system settings.',
  descTime:
    'Time formats display date and time serial numbers as date values.  Time formats that begin with an asterisk (*) respond to changes in regional date and time settings that are specified for the operating system. Formats without an asterisk are not affected by operating system settings.',
  descText:
    'Text format cells are treated as text even when a number is in the cell.  The cell is displayed exactly as entered.',
  descOther: 'Special formats are useful for tracking list and database values.',
  verticalAlign: 'Vertical alignment:',
  rotation: 'Degrees:',
  borderStyleDashDotDot: 'Dash Dot Dot',
  borderStyleDashDot: 'Dash Dot',
  borderStyleMediumDashDotDot: 'Medium Dash Dot Dot',
  borderStyleSlantDashDot: 'Slanted Dash Dot',
  borderStyleMediumDashDot: 'Medium Dash Dot',
};
