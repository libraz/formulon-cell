// Locale-dependent text shared by the function pickers: built-in function
// descriptions, the catalog locale ordinal, and category label keys.

import type {
  CatalogFunctionCategory,
  CatalogLocaleOrdinal,
  FunctionCatalogEntry,
  FunctionCatalogSnapshot,
} from '../commands/function-categories.js';

export type FunctionDescription = {
  en: string;
  ja: string;
  de: string;
  fr: string;
  zh: string;
  ko: string;
  th: string;
};

/** Concise "spreadsheet-style" descriptions for the most common functions. The
 *  catalog itself (FUNCTION_SIGNATURES) doesn't carry descriptions — anything
 *  not listed here renders without a description blurb. Keep the list small;
 *  exhaustive coverage isn't a goal. */
export const FUNCTION_DESCRIPTIONS: Readonly<Record<string, FunctionDescription>> = {
  SUM: {
    en: 'Adds its arguments.',
    ja: '引数の合計を返します。',
    de: 'Addiert die Argumente.',
    fr: 'Additionne ses arguments.',
    zh: '对参数求和。',
    ko: '인수의 합을 반환합니다.',
    th: 'หาผลรวมของอาร์กิวเมนต์',
  },
  ACOS: {
    en: 'Returns the arccosine of a number from -1 to 1, in radians.',
    ja: '-1 ～ 1 の数値のアークコサインをラジアンで返します。',
    de: 'Gibt den Arkuskosinus einer Zahl zwischen -1 und 1 im Bogenmaß zurück.',
    fr: "Renvoie l'arc cosinus d'un nombre compris entre -1 et 1, en radians.",
    zh: '返回 -1 到 1 之间数字的反余弦值（以弧度表示）。',
    ko: '-1에서 1 사이 숫자의 아크코사인(라디안)을 반환합니다.',
    th: 'ส่งกลับค่าอาร์กโคไซน์ของตัวเลขตั้งแต่ -1 ถึง 1 เป็นเรเดียน',
  },
  IF: {
    en: 'Returns one value when a condition is true and another when false.',
    ja: '条件が真のときと偽のときで異なる値を返します。',
    de: 'Gibt bei einer wahren bzw. falschen Bedingung unterschiedliche Werte zurück.',
    fr: 'Renvoie une valeur si la condition est vraie et une autre si elle est fausse.',
    zh: '条件为真时返回一个值，为假时返回另一个值。',
    ko: '조건이 참이면 한 값을, 거짓이면 다른 값을 반환합니다.',
    th: 'ส่งกลับค่าหนึ่งเมื่อเงื่อนไขเป็นจริง และอีกค่าเมื่อเป็นเท็จ',
  },
  VLOOKUP: {
    en: 'Looks up a value in the leftmost column of a table.',
    ja: '表の左端列で値を検索します。',
    de: 'Sucht einen Wert in der ersten Spalte einer Tabelle.',
    fr: "Recherche une valeur dans la première colonne d'un tableau.",
    zh: '在表格的最左列中查找值。',
    ko: '표의 가장 왼쪽 열에서 값을 조회합니다.',
    th: 'ค้นหาค่าในคอลัมน์ซ้ายสุดของตาราง',
  },
  COUNT: {
    en: 'Counts numeric cells in the range.',
    ja: '範囲内の数値セルの個数を返します。',
    de: 'Zählt die numerischen Zellen im Bereich.',
    fr: 'Compte les cellules numériques de la plage.',
    zh: '统计区域中的数字单元格数量。',
    ko: '범위에서 숫자가 있는 셀의 개수를 계산합니다.',
    th: 'นับจำนวนเซลล์ตัวเลขในช่วง',
  },
  COUNTA: {
    en: 'Counts non-empty cells in the range.',
    ja: '範囲内の空白でないセルの個数を返します。',
    de: 'Zählt die nicht leeren Zellen im Bereich.',
    fr: 'Compte les cellules non vides de la plage.',
    zh: '统计区域中的非空单元格数量。',
    ko: '범위에서 비어 있지 않은 셀의 개수를 계산합니다.',
    th: 'นับจำนวนเซลล์ที่ไม่ว่างในช่วง',
  },
  COUNTIF: {
    en: 'Counts cells in a range that match a condition.',
    ja: '条件を満たすセルの個数を返します。',
    de: 'Zählt die Zellen eines Bereichs, die eine Bedingung erfüllen.',
    fr: 'Compte les cellules de la plage qui respectent une condition.',
    zh: '统计区域中满足条件的单元格数量。',
    ko: '범위에서 조건을 충족하는 셀의 개수를 계산합니다.',
    th: 'นับเซลล์ในช่วงที่ตรงตามเงื่อนไข',
  },
  INDEX: {
    en: 'Returns a value at a given row/column in an array.',
    ja: '配列内の指定された行と列の値を返します。',
    de: 'Gibt den Wert an einer bestimmten Zeile und Spalte eines Arrays zurück.',
    fr: "Renvoie la valeur à l'intersection d'une ligne et d'une colonne d'un tableau.",
    zh: '返回数组中指定行和列位置的值。',
    ko: '배열에서 지정한 행과 열의 값을 반환합니다.',
    th: 'ส่งกลับค่าที่แถวหรือคอลัมน์ที่ระบุในอาร์เรย์',
  },
  MATCH: {
    en: 'Returns the position of a value in an array.',
    ja: '配列内で一致する値の位置を返します。',
    de: 'Gibt die Position eines Werts in einem Array zurück.',
    fr: "Renvoie la position d'une valeur dans un tableau.",
    zh: '返回值在数组中的位置。',
    ko: '배열에서 값의 위치를 반환합니다.',
    th: 'ส่งกลับตำแหน่งของค่าในอาร์เรย์',
  },
  AVERAGE: {
    en: 'Returns the arithmetic mean.',
    ja: '引数の平均値を返します。',
    de: 'Gibt den arithmetischen Mittelwert zurück.',
    fr: 'Renvoie la moyenne arithmétique.',
    zh: '返回算术平均值。',
    ko: '산술 평균을 반환합니다.',
    th: 'ส่งกลับค่าเฉลี่ยเลขคณิต',
  },
  MIN: {
    en: 'Returns the smallest argument.',
    ja: '引数の最小値を返します。',
    de: 'Gibt das kleinste Argument zurück.',
    fr: "Renvoie l'argument le plus petit.",
    zh: '返回最小参数。',
    ko: '가장 작은 인수를 반환합니다.',
    th: 'ส่งกลับอาร์กิวเมนต์ที่น้อยที่สุด',
  },
  MAX: {
    en: 'Returns the largest argument.',
    ja: '引数の最大値を返します。',
    de: 'Gibt das größte Argument zurück.',
    fr: "Renvoie l'argument le plus grand.",
    zh: '返回最大参数。',
    ko: '가장 큰 인수를 반환합니다.',
    th: 'ส่งกลับอาร์กิวเมนต์ที่มากที่สุด',
  },
  ROUND: {
    en: 'Rounds a number to a given precision.',
    ja: '数値を指定した桁数で四捨五入します。',
    de: 'Rundet eine Zahl auf die angegebene Genauigkeit.',
    fr: 'Arrondit un nombre à la précision indiquée.',
    zh: '将数字舍入到指定精度。',
    ko: '숫자를 지정한 자릿수로 반올림합니다.',
    th: 'ปัดเศษตัวเลขตามความแม่นยำที่ระบุ',
  },
  IFERROR: {
    en: 'Returns a fallback when the first argument is an error.',
    ja: '式がエラーの場合に代替値を返します。',
    de: 'Gibt einen Ersatzwert zurück, wenn das erste Argument ein Fehler ist.',
    fr: 'Renvoie une valeur de remplacement lorsque le premier argument est une erreur.',
    zh: '当第一个参数为错误时返回替代值。',
    ko: '첫 번째 인수가 오류이면 대체 값을 반환합니다.',
    th: 'ส่งกลับค่าแทนเมื่ออาร์กิวเมนต์แรกเป็นข้อผิดพลาด',
  },
  CONCAT: {
    en: 'Concatenates a list of texts.',
    ja: '複数の文字列を連結します。',
    de: 'Verkettet eine Liste von Texten.',
    fr: 'Concatène une liste de textes.',
    zh: '连接一组文本。',
    ko: '텍스트 목록을 연결합니다.',
    th: 'เชื่อมข้อความในรายการเข้าด้วยกัน',
  },
  TEXT: {
    en: 'Formats a value as text using a format code.',
    ja: '書式コードに従って数値を文字列に整形します。',
    de: 'Formatiert einen Wert mithilfe eines Formatcodes als Text.',
    fr: 'Formate une valeur en texte à l’aide d’un code de format.',
    zh: '使用格式代码将值设置为文本格式。',
    ko: '서식 코드를 사용하여 값을 텍스트로 서식 지정합니다.',
    th: 'จัดรูปแบบค่าเป็นข้อความโดยใช้รหัสรูปแบบ',
  },
  LEFT: {
    en: 'Returns the left part of a string.',
    ja: '文字列の先頭から指定文字数を返します。',
    de: 'Gibt den linken Teil einer Zeichenfolge zurück.',
    fr: "Renvoie la partie gauche d'une chaîne.",
    zh: '返回字符串的左侧部分。',
    ko: '문자열의 왼쪽 부분을 반환합니다.',
    th: 'ส่งกลับส่วนซ้ายของสตริง',
  },
  RIGHT: {
    en: 'Returns the right part of a string.',
    ja: '文字列の末尾から指定文字数を返します。',
    de: 'Gibt den rechten Teil einer Zeichenfolge zurück.',
    fr: "Renvoie la partie droite d'une chaîne.",
    zh: '返回字符串的右侧部分。',
    ko: '문자열의 오른쪽 부분을 반환합니다.',
    th: 'ส่งกลับส่วนขวาของสตริง',
  },
  MID: {
    en: 'Returns characters from the middle of a string.',
    ja: '文字列の中間から指定文字数を返します。',
    de: 'Gibt Zeichen aus der Mitte einer Zeichenfolge zurück.',
    fr: "Renvoie des caractères au milieu d'une chaîne.",
    zh: '返回字符串中间的字符。',
    ko: '문자열 중간의 문자를 반환합니다.',
    th: 'ส่งกลับอักขระจากตรงกลางของสตริง',
  },
  LEN: {
    en: 'Returns the length of a string.',
    ja: '文字列の文字数を返します。',
    de: 'Gibt die Länge einer Zeichenfolge zurück.',
    fr: "Renvoie la longueur d'une chaîne.",
    zh: '返回字符串的长度。',
    ko: '문자열의 길이를 반환합니다.',
    th: 'ส่งกลับความยาวของสตริง',
  },
  UPPER: {
    en: 'Converts a string to upper case.',
    ja: '文字列を大文字に変換します。',
    de: 'Wandelt eine Zeichenfolge in Großbuchstaben um.',
    fr: 'Convertit une chaîne en majuscules.',
    zh: '将字符串转换为大写。',
    ko: '문자열을 대문자로 변환합니다.',
    th: 'แปลงสตริงเป็นตัวพิมพ์ใหญ่',
  },
  LOWER: {
    en: 'Converts a string to lower case.',
    ja: '文字列を小文字に変換します。',
    de: 'Wandelt eine Zeichenfolge in Kleinbuchstaben um.',
    fr: 'Convertit une chaîne en minuscules.',
    zh: '将字符串转换为小写。',
    ko: '문자열을 소문자로 변환합니다.',
    th: 'แปลงสตริงเป็นตัวพิมพ์เล็ก',
  },
  AND: {
    en: 'TRUE only when every argument is TRUE.',
    ja: 'すべての引数が真のとき真を返します。',
    de: 'Gibt WAHR zurück, wenn alle Argumente WAHR sind.',
    fr: 'Renvoie VRAI uniquement si tous les arguments sont VRAI.',
    zh: '仅当所有参数均为 TRUE 时返回 TRUE。',
    ko: '모든 인수가 TRUE일 때만 TRUE를 반환합니다.',
    th: 'ส่งกลับ TRUE เมื่ออาร์กิวเมนต์ทั้งหมดเป็น TRUE เท่านั้น',
  },
  OR: {
    en: 'TRUE when any argument is TRUE.',
    ja: 'いずれかの引数が真のとき真を返します。',
    de: 'Gibt WAHR zurück, wenn mindestens ein Argument WAHR ist.',
    fr: 'Renvoie VRAI si au moins un argument est VRAI.',
    zh: '任一参数为 TRUE 时返回 TRUE。',
    ko: '하나 이상의 인수가 TRUE이면 TRUE를 반환합니다.',
    th: 'ส่งกลับ TRUE เมื่อมีอาร์กิวเมนต์ใดเป็น TRUE',
  },
  NOT: {
    en: 'Inverts a boolean.',
    ja: '論理値を反転します。',
    de: 'Kehrt einen Wahrheitswert um.',
    fr: 'Inverse une valeur logique.',
    zh: '反转逻辑值。',
    ko: '논리값을 반대로 바꿉니다.',
    th: 'กลับค่าบูลีน',
  },
  ISBLANK: {
    en: 'Tests whether a value is blank.',
    ja: '値が空白かどうかを返します。',
    de: 'Prüft, ob ein Wert leer ist.',
    fr: 'Teste si une valeur est vide.',
    zh: '检查值是否为空白。',
    ko: '값이 비어 있는지 확인합니다.',
    th: 'ตรวจสอบว่าค่าว่างหรือไม่',
  },
  ISNUMBER: {
    en: 'Tests whether a value is numeric.',
    ja: '値が数値かどうかを返します。',
    de: 'Prüft, ob ein Wert numerisch ist.',
    fr: 'Teste si une valeur est numérique.',
    zh: '检查值是否为数字。',
    ko: '값이 숫자인지 확인합니다.',
    th: 'ตรวจสอบว่าค่าเป็นตัวเลขหรือไม่',
  },
  NOW: {
    en: 'Returns the current date and time.',
    ja: '現在の日付と時刻を返します。',
    de: 'Gibt das aktuelle Datum und die aktuelle Uhrzeit zurück.',
    fr: "Renvoie la date et l'heure actuelles.",
    zh: '返回当前日期和时间。',
    ko: '현재 날짜와 시간을 반환합니다.',
    th: 'ส่งกลับวันที่และเวลาปัจจุบัน',
  },
  TODAY: {
    en: "Returns today's date.",
    ja: '今日の日付を返します。',
    de: 'Gibt das heutige Datum zurück.',
    fr: "Renvoie la date d'aujourd'hui.",
    zh: '返回今天的日期。',
    ko: '오늘 날짜를 반환합니다.',
    th: 'ส่งกลับวันที่วันนี้',
  },
  DATE: {
    en: 'Builds a date from year, month, day.',
    ja: '年・月・日からシリアル値を作成します。',
    de: 'Erstellt ein Datum aus Jahr, Monat und Tag.',
    fr: 'Construit une date à partir de l’année, du mois et du jour.',
    zh: '根据年、月、日生成日期。',
    ko: '연도, 월, 일로 날짜를 만듭니다.',
    th: 'สร้างวันที่จากปี เดือน และวัน',
  },
};

/** `Strings['fxDialog']` key holding each catalog category's title. */
export type CategoryLabelKey =
  | 'categoryLogical'
  | 'categoryLookup'
  | 'categoryText'
  | 'categoryDateTime'
  | 'categoryMath'
  | 'categoryFinancial'
  | 'categoryDynamicArray'
  | 'categoryStatistical'
  | 'categoryEngineering'
  | 'categoryInformation'
  | 'categoryDatabase'
  | 'categoryCompatibility'
  | 'categoryCube'
  | 'categoryWeb';

export const CATEGORY_LABEL_KEY: Readonly<Record<CatalogFunctionCategory, CategoryLabelKey>> = {
  logical: 'categoryLogical',
  lookup: 'categoryLookup',
  text: 'categoryText',
  datetime: 'categoryDateTime',
  math: 'categoryMath',
  financial: 'categoryFinancial',
  dynamicArray: 'categoryDynamicArray',
  statistical: 'categoryStatistical',
  engineering: 'categoryEngineering',
  information: 'categoryInformation',
  database: 'categoryDatabase',
  compatibility: 'categoryCompatibility',
  cube: 'categoryCube',
  web: 'categoryWeb',
};

/**
 * Legacy native metadata ordinal for a UI locale tag. German and French
 * function names can be localized by formulon; Chinese, Korean, and Thai use
 * canonical English function names while their provider descriptions remain
 * fully localized through the BCP-47 tag path below.
 */
export const catalogLocaleOrdinal = (locale: string): CatalogLocaleOrdinal => {
  const language = locale.trim().toLowerCase().replace('_', '-').split('-')[0];
  switch (language) {
    case 'ja':
      return 1;
    case 'de':
      return 2;
    case 'fr':
      return 3;
    default:
      return 0;
  }
};

/** Workbook surface the function pickers use for locale-dependent names. */
export interface FunctionNameReader {
  localeFacts?: () => { listSeparator?: string } | null;
  localizeFunctionName?: (name: string, locale: CatalogLocaleOrdinal) => string | null;
  canonicalizeFunctionName?: (name: string, locale: CatalogLocaleOrdinal) => string | null;
}

// Every reader call is guarded: a stale or disposing workbook throws, and the
// current catalog must stay usable until the host swaps it out.

/** Argument separator of the workbook's locale; `,` when unknown. */
export const functionListSeparator = (reader: FunctionNameReader | null): ',' | ';' => {
  try {
    return typeof reader?.localeFacts === 'function' && reader.localeFacts()?.listSeparator === ';'
      ? ';'
      : ',';
  } catch {
    return ',';
  }
};

/** Name a formula spells `canonicalName` with in `locale`; the canonical name when unmapped. */
export const nativeFunctionName = (
  reader: FunctionNameReader | null,
  canonicalName: string,
  locale: CatalogLocaleOrdinal,
): string => {
  try {
    const localized =
      typeof reader?.localizeFunctionName === 'function'
        ? reader.localizeFunctionName(canonicalName, locale)
        : null;
    return localized || canonicalName;
  } catch {
    return canonicalName;
  }
};

/**
 * Catalog entry name for a typed function name, or `null` when the catalog
 * has none. Same order as the engine write boundary: a canonical name wins,
 * then the engine's alias table, then a scan of each entry's native name.
 */
export const resolveCatalogFunctionName = (
  catalog: Pick<FunctionCatalogSnapshot, 'entries' | 'knownNames'>,
  reader: FunctionNameReader | null,
  rawName: string,
  locale: CatalogLocaleOrdinal,
): string | null => {
  const upper = rawName.trim().toLocaleUpperCase();
  if (catalog.knownNames.has(upper)) return upper;
  try {
    const canonical =
      typeof reader?.canonicalizeFunctionName === 'function'
        ? reader.canonicalizeFunctionName(rawName, locale)
        : null;
    if (canonical && catalog.knownNames.has(canonical)) return canonical;
  } catch {
    // Fall through to the native-name scan.
  }
  for (const entry of catalog.entries.values()) {
    if (nativeFunctionName(reader, entry.canonicalName, locale).toLocaleUpperCase() === upper) {
      return entry.canonicalName;
    }
  }
  return null;
};

/** Canonical BCP-47 tag used to select a host-provided function catalog. */
export const catalogLocaleTag = (locale: string): string => {
  const language = locale.trim().toLowerCase().replace(/_/g, '-').split('-')[0];
  switch (language) {
    case 'ja':
      return 'ja-JP';
    case 'de':
      return 'de-DE';
    case 'fr':
      return 'fr-FR';
    case 'zh':
      return 'zh-CN';
    case 'ko':
      return 'ko-KR';
    case 'th':
      return 'th-TH';
    case 'en':
      return 'en-US';
    default:
      return locale.trim() || 'en-US';
  }
};

type DescriptionLanguage = keyof FunctionDescription;

const descriptionLanguage = (locale: string | CatalogLocaleOrdinal): DescriptionLanguage => {
  if (typeof locale === 'number') {
    switch (locale) {
      case 1:
        return 'ja';
      case 2:
        return 'de';
      case 3:
        return 'fr';
      default:
        return 'en';
    }
  }
  switch (locale.trim().toLowerCase().replace(/_/g, '-').split('-')[0]) {
    case 'ja':
      return 'ja';
    case 'de':
      return 'de';
    case 'fr':
      return 'fr';
    case 'zh':
      return 'zh';
    case 'ko':
      return 'ko';
    case 'th':
      return 'th';
    default:
      return 'en';
  }
};

/** Engine-supplied description, else the built-in blurb, else empty. */
export const functionDescription = (
  entry: Pick<FunctionCatalogEntry, 'canonicalName' | 'description'>,
  locale: string | CatalogLocaleOrdinal,
): string => {
  const language = descriptionLanguage(locale);
  return entry.description ?? FUNCTION_DESCRIPTIONS[entry.canonicalName]?.[language] ?? '';
};
