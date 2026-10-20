import { inputLanguage } from './input-grammar.js';

/**
 * Text series the desktop spreadsheet's AutoFill recognises, per UI language:
 * its built-in custom lists, quarter labels that wrap after 4, ordinals, and
 * the digit run it steps inside other text. The lists are the application's
 * own, read through its object model, including its spelling quirks.
 */

type Lists = readonly (readonly string[])[];

const EN_DAYS = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
const EN_DAY_NAMES = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const EN_MONTHS = [
  'Jan',
  'Feb',
  'Mar',
  'Apr',
  'May',
  'Jun',
  'Jul',
  'Aug',
  'Sep',
  'Oct',
  'Nov',
  'Dec',
];
const EN_MONTH_NAMES = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];
const monthsWith = (suffix: string): string[] =>
  Array.from({ length: 12 }, (_, i) => `${i + 1}${suffix}`);
const ZODIAC = ['子', '丑', '寅', '卯', '辰', '巳', '午', '未', '申', '酉', '戌', '亥'];
const STEMS = ['甲', '乙', '丙', '丁', '戊', '己', '庚', '辛', '壬', '癸'];
const SEXAGENARY_KO = [
  ...'갑자 을축 병인 정묘 무진 기사 경오 신미 임신 계유 갑술 을해 병자 정축 무인 기묘 경진 신사 임오 계미'.split(
    ' ',
  ),
  ...'갑신 을유 병술 정해 무자 기축 경인 신묘 임진 계사 갑오 을미 병신년 정유 무술 기해 경자 신축 임인 계묘'.split(
    ' ',
  ),
  ...'갑진 을사 병오 정미 무신 기유 경술 신해 임자 계축 갑인 을묘 병진 정사 무오 기미 경신 신유 임술 계해'.split(
    ' ',
  ),
];

const LISTS: Readonly<Record<string, Lists>> = {
  en: [EN_DAYS, EN_DAY_NAMES, EN_MONTHS, EN_MONTH_NAMES],
  ja: [
    EN_DAYS,
    EN_DAY_NAMES,
    monthsWith(' 月'),
    EN_MONTH_NAMES,
    ['日', '月', '火', '水', '木', '金', '土'],
    ['日曜日', '月曜日', '火曜日', '水曜日', '木曜日', '金曜日', '土曜日'],
    monthsWith('月').map((m) => (m === '9月' ? '9 月' : m)),
    ['第 1 四半期', '第 2 四半期', '第 3 四半期', '第 4 四半期'],
    [
      '睦月',
      '如月',
      '弥生',
      '卯月',
      '皐月',
      '水無月',
      '文月',
      '葉月',
      '長月',
      '神無月',
      '霜月',
      '師走',
    ],
    ZODIAC,
    STEMS,
  ],
  de: [
    ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'],
    ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'],
    EN_MONTHS,
    EN_MONTH_NAMES,
    ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'],
    ['Montag', 'Dienstag', 'Mittwoch', 'Donnerstag', 'Freitag', 'Samstag', 'Sonntag'],
    ['Jan', 'Feb', 'Mär', 'Apr', 'Mai', 'Jun', 'Jul', 'Aug', 'Sep', 'Okt', 'Nov', 'Dez'],
    [
      'Januar',
      'Februar',
      'März',
      'April',
      'Mai',
      'Juni',
      'Juli',
      'August',
      'September',
      'Oktober',
      'November',
      'Dezember',
    ],
  ],
  fr: [
    ['Lun.', 'Mar.', 'Mer.', 'Jeu.', 'Ven.', 'Sam.', 'Dim.'],
    ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche'],
    ['janv', 'févr', 'mars', 'avr', 'mai', 'juin', 'juil', 'août', 'sept', 'oct', 'nov', 'déc'],
    [
      'janvier',
      'février',
      'mars',
      'avril',
      'mai',
      'juin',
      'juillet',
      'août',
      'septembre',
      'octobre',
      'novembre',
      'décembre',
    ],
  ],
  zh: [
    EN_DAYS,
    EN_DAY_NAMES,
    monthsWith('月'),
    EN_MONTH_NAMES,
    ['周日', '周一', '周二', '周三', '周四', '周五', '周六'],
    ['星期日', '星期一', '星期二', '星期三', '星期四', '星期五', '星期六'],
    [
      '一月',
      '二月',
      '三月',
      '四月',
      '五月',
      '六月',
      '七月',
      '八月',
      '九月',
      '十月',
      '十一月',
      '十二月',
    ],
    ['第一季', '第二季', '第三季', '第四季'],
    [
      '正月',
      '二月',
      '三月',
      '四月',
      '五月',
      '六月',
      '七月',
      '八月',
      '九月',
      '十月',
      '冬月',
      '腊月',
    ],
    ZODIAC,
    STEMS,
  ],
  ko: [
    EN_DAYS,
    EN_DAY_NAMES,
    EN_MONTHS,
    EN_MONTH_NAMES,
    ['일', '월', '화', '수', '목', '금', '토'],
    ['일요일', '월요일', '화요일', '수요일', '목요일', '금요일', '토요일'],
    monthsWith('월'),
    ['1사분기', '2사분기', '3사분기', '4사분기'],
    [
      '일월',
      '이월',
      '삼월',
      '사월',
      '오월',
      '유월',
      '칠월',
      '팔월',
      '구월',
      '시월',
      '십일월',
      '십이월',
    ],
    ['자', '축', '인', '묘', '진', '사', '오', '미', '신', '유', '술', '해'],
    ['갑', '을', '병', '정', '무', '기', '경', '신', '임', '계'],
    monthsWith('月'),
    ['1四分期', '2四分期', '3四分期', '4四分期'],
    SEXAGENARY_KO,
  ],
  th: [
    EN_DAYS,
    EN_DAY_NAMES,
    EN_MONTHS,
    EN_MONTH_NAMES,
    ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัส', 'ศุกร์', 'เสาร์'],
    ['วันอาทิตย์', 'วันจันทร์', 'วันอังคาร', 'วันพุธ', 'วันพฤหัสบดี', 'วันศุกร์', 'วันเสาร์'],
    [
      'ม.ค.',
      'ก.พ.',
      'มี.ค.',
      'เม.ย.',
      'พ.ค.',
      'มิ.ย.',
      'ก.ค.',
      'ส.ค.',
      'ก.ย.',
      'ต.ค.',
      'พ.ย.',
      'ธ.ค.',
    ],
    [
      'มกราคม',
      'กุมภาพันธ์',
      'มีนาคม',
      'เมษายน',
      'พฤษภาคม',
      'มิถุนายน',
      'กรกฎาคม',
      'สิงหาคม',
      'กันยายน',
      'ตุลาคม',
      'พฤศจิกายน',
      'ธันวาคม',
    ],
  ],
};

/** The built-in custom lists AutoFill and custom-list sorting use for `locale`. */
export const customListsFor = (locale: string | undefined): Lists =>
  LISTS[inputLanguage(locale)] ?? (LISTS.en as Lists);

type QuarterForm = 'q' | 'qtr' | 'quarter' | 'ordinalQtr' | 'cjk';

interface SeriesProfile {
  /** Ja/zh/ko: a list match keeps the list's spelling, and the last digit run
   *  anywhere in the text steps, full-width digits included. */
  cjk: boolean;
  foldWidth: boolean;
  /** Ko matches Hanja against the Hangul lists. */
  hanja: boolean;
  /** A trailing period is ignored when matching a list item. */
  trailingDot: boolean;
  ordinals: boolean;
  /** Quarter labels that wrap from 4 back to 1. */
  quarters: ReadonlySet<QuarterForm>;
}

const PROFILES: Readonly<Record<string, SeriesProfile>> = {
  en: {
    cjk: false,
    foldWidth: false,
    hanja: false,
    trailingDot: true,
    ordinals: true,
    quarters: new Set(['q', 'qtr', 'quarter', 'ordinalQtr']),
  },
  ja: {
    cjk: true,
    foldWidth: true,
    hanja: false,
    trailingDot: true,
    ordinals: true,
    quarters: new Set(['q', 'cjk']),
  },
  de: {
    cjk: false,
    foldWidth: false,
    hanja: false,
    trailingDot: true,
    ordinals: false,
    quarters: new Set(['q']),
  },
  fr: {
    cjk: false,
    foldWidth: false,
    hanja: false,
    trailingDot: true,
    ordinals: false,
    quarters: new Set(),
  },
  zh: {
    cjk: true,
    foldWidth: false,
    hanja: false,
    trailingDot: true,
    ordinals: true,
    quarters: new Set(['q']),
  },
  ko: {
    cjk: true,
    foldWidth: true,
    hanja: true,
    trailingDot: true,
    ordinals: true,
    quarters: new Set(['q', 'qtr', 'ordinalQtr']),
  },
  th: {
    cjk: false,
    foldWidth: false,
    hanja: false,
    trailingDot: false,
    ordinals: false,
    quarters: new Set(['q', 'qtr']),
  },
};

const profileFor = (locale: string | undefined): SeriesProfile =>
  PROFILES[inputLanguage(locale)] ?? (PROFILES.en as SeriesProfile);

const HANJA_TO_HANGUL: Readonly<Record<string, string>> = Object.fromEntries(
  [
    ['日月火水木金土曜', '일월화수목금토요'],
    ['子丑寅卯辰巳午未申酉戌亥', '자축인묘진사오미신유술해'],
    ['甲乙丙丁戊己庚辛壬癸', '갑을병정무기경신임계'],
    ['一二三四五六七八九十分期', '일이삼사오육칠팔구십분기'],
  ].flatMap(([hanja, hangul]) => [...(hanja ?? '')].map((ch, i) => [ch, hangul?.[i] ?? ch])),
);

const matchKey = (text: string, profile: SeriesProfile): string => {
  let key = profile.foldWidth ? text.normalize('NFKC') : text;
  if (profile.hanja) key = [...key].map((ch) => HANJA_TO_HANGUL[ch] ?? ch).join('');
  // Accents are ignored: `Dec` matches the French `déc`.
  return key.normalize('NFD').replace(/\p{M}/gu, '').normalize('NFC').toLocaleLowerCase();
};

/** Case the list item the way the source was cased: lower when the source
 *  starts lower, upper when its first two letters are upper, else capitalised. */
const caseLike = (item: string, source: string): string => {
  const letters = [...source].filter((ch) => ch.toLowerCase() !== ch.toUpperCase());
  const [first, second] = letters;
  if (!first) return item;
  if (first === first.toLowerCase()) return item.toLowerCase();
  if (second && second === second.toUpperCase()) return item.toUpperCase();
  return item.charAt(0).toUpperCase() + item.slice(1);
};

export type TextSeriesAtom =
  | { kind: 'list'; list: readonly string[]; listIndex: number; index: number; source: string }
  | {
      kind: 'number';
      /** Text before and after the stepped digits; equal templates form one series. */
      prefix: string;
      suffix: string;
      n: number;
      width: number;
      fullwidth: boolean;
      ordinal: boolean;
      /** Quarter labels step 1..4 and wrap. */
      wrap: boolean;
    };

const FULLWIDTH_ZERO = 0xff10;

const asciiDigits = (s: string): string =>
  s.replace(/[０-９]/g, (ch) => String((ch.codePointAt(0) ?? FULLWIDTH_ZERO) - FULLWIDTH_ZERO));

const ORDINAL_SUFFIX = /^(st|nd|rd|th)/i;

const ordinalSuffix = (n: number): string => {
  const abs = Math.abs(n);
  if (abs % 100 >= 11 && abs % 100 <= 13) return 'th';
  return ['th', 'st', 'nd', 'rd'][abs % 10] ?? 'th';
};

const QUARTER_PATTERNS: readonly [QuarterForm, RegExp][] = [
  ['q', /^(q\s?)(\d+)()$/i],
  ['qtr', /^(qtr\s?)(\d+)()$/i],
  ['quarter', /^(quarter\s?)(\d+)()$/i],
  ['ordinalQtr', /^()(\d+)((?:st|nd|rd|th) qtr)$/i],
  ['cjk', /^(第\s*)(\d+)(\s*四半期)$/],
];

const numberAtom = (
  prefix: string,
  digits: string,
  suffix: string,
  extra: { ordinal?: boolean; wrap?: boolean } = {},
): TextSeriesAtom => {
  const fullwidth = /[０-９]/.test(digits);
  return {
    kind: 'number',
    prefix,
    suffix: extra.ordinal ? suffix.replace(ORDINAL_SUFFIX, '') : suffix,
    n: Number(asciiDigits(digits)),
    width: digits.length,
    fullwidth,
    ordinal: extra.ordinal ?? false,
    wrap: extra.wrap ?? false,
  };
};

/** Classify a text cell as a step of a list or numbered series, or null when
 *  AutoFill copies it. */
export function textSeriesAtom(text: string, locale: string | undefined): TextSeriesAtom | null {
  const profile = profileFor(locale);
  const lists = customListsFor(locale);
  const key = matchKey(text, profile);
  const dotless = profile.trailingDot && key.endsWith('.') ? key.slice(0, -1) : null;
  // The last list containing the text wins, so a localized list overrides the English one.
  for (let listIndex = lists.length - 1; listIndex >= 0; listIndex -= 1) {
    const list = lists[listIndex] as readonly string[];
    const keys = list.map((item) => matchKey(item, profile));
    let index = keys.indexOf(key);
    if (index < 0 && dotless) index = keys.indexOf(dotless);
    if (index >= 0) return { kind: 'list', list, listIndex, index, source: text };
  }
  for (const [form, pattern] of QUARTER_PATTERNS) {
    if (!profile.quarters.has(form)) continue;
    const m = text.match(pattern);
    const n = Number(m?.[2]);
    if (m && n >= 1 && n <= 4) {
      return numberAtom(m[1] ?? '', m[2] ?? '', m[3] ?? '', {
        ordinal: form === 'ordinalQtr',
        wrap: true,
      });
    }
  }
  if (profile.ordinals) {
    const m = text.match(/^(\d+)((?:st|nd|rd|th)(?![A-Za-z]).*)$/is);
    if (m) return numberAtom('', m[1] ?? '', m[2] ?? '', { ordinal: true });
  }
  const digit = profile.cjk ? '[0-9０-９]' : '[0-9]';
  const trailing = text.match(new RegExp(`^(.*?)(${digit}+)$`, 's'));
  if (trailing) return numberAtom(trailing[1] ?? '', trailing[2] ?? '', '');
  const leading = text.match(/^([0-9]+)(\s.*)$/s);
  if (leading) return numberAtom('', leading[1] ?? '', leading[2] ?? '');
  if (profile.cjk) {
    const last = text.match(new RegExp(`^(.*?)(${digit}+)([^0-9０-９]*)$`, 's'));
    if (last) return numberAtom(last[1] ?? '', last[2] ?? '', last[3] ?? '');
  }
  return null;
}

/** Whether two atoms belong to one series, so AutoFill steps them together. */
export const sameTextSeries = (a: TextSeriesAtom, b: TextSeriesAtom): boolean => {
  if (a.kind === 'list' || b.kind === 'list') {
    return a.kind === 'list' && b.kind === 'list' && a.list === b.list;
  }
  return (
    a.prefix === b.prefix &&
    a.suffix === b.suffix &&
    a.ordinal === b.ordinal &&
    a.wrap === b.wrap &&
    a.fullwidth === b.fullwidth
  );
};

const toFullwidth = (s: string): string =>
  s.replace(/[0-9]/g, (ch) => String.fromCodePoint(FULLWIDTH_ZERO + Number(ch)));

/** The text at series position `n` (list index or number) shaped like `atom`.
 *  `casing` is the source text whose letter case a list item follows. */
export function textSeriesValue(
  atom: TextSeriesAtom,
  n: number,
  locale: string | undefined,
  casing: string,
): string {
  if (atom.kind === 'list') {
    const len = atom.list.length;
    const item = atom.list[((n % len) + len) % len] ?? '';
    return profileFor(locale).cjk ? item : caseLike(item, casing);
  }
  // Below zero the series continues upward through the absolute values.
  const value = atom.wrap ? ((((n - 1) % 4) + 4) % 4) + 1 : Math.abs(n);
  const digits = String(value).padStart(atom.width, '0');
  const shown = atom.fullwidth ? toFullwidth(digits) : digits;
  const ordinal = atom.ordinal ? ordinalSuffix(value) : '';
  return `${atom.prefix}${shown}${ordinal}${atom.suffix}`;
}
