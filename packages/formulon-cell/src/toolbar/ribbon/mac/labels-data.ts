import {
  type MacRibbonReasonStrings,
  macRibbonReasonsEn,
} from '../../../i18n/strings/mac/ribbon-reasons.js';
import type { Locale, Strings } from '../../../i18n/strings.js';
import { dictionaries } from '../../../i18n/strings.js';
import { MAC_BUNDLE_LABELS } from './labels-bundle.js';

export const MAC_LOCALES = ['ja', 'en', 'de', 'fr', 'zh', 'ko', 'th'] as const;

export type MacLocale = (typeof MAC_LOCALES)[number];
export type Label = Readonly<Record<Locale, string>>;
export type LegacyLabel = Pick<Label, 'ja' | 'en'>;
export type LabelLike = Label | LegacyLabel;
export type ExtraTranslations = Readonly<{
  de: string;
  fr: string;
  zh: string;
  ko: string;
  th: string;
}>;

type RibbonKey = Exclude<keyof Strings['ribbon'], 'tabs' | 'fontFamilies'>;

/**
 * Mac labels without a reusable central dictionary field. Values are kept by
 * stable id and supplied for every non-English locale explicitly. Product
 * names and standards such as SmartArt, Python, PDF, and 100% intentionally
 * remain unchanged where that is the Excel UI term.
 */
export const MAC_LABEL_TRANSLATIONS: Readonly<Record<string, ExtraTranslations>> = {
  chartBar: { de: 'Balken', fr: 'Barres', zh: '条形图', ko: '가로 막대형', th: 'แท่ง' },
  chartArea: { de: 'Fläche', fr: 'Aires', zh: '面积图', ko: '영역형', th: 'พื้นที่' },
  shapeLine: { de: 'Linie', fr: 'Ligne', zh: '直线', ko: '선', th: 'เส้น' },
  shapeRectangle: { de: 'Rechteck', fr: 'Rectangle', zh: '矩形', ko: '직사각형', th: 'สี่เหลี่ยมผืนผ้า' },
  shapeOval: { de: 'Oval', fr: 'Ovale', zh: '椭圆', ko: '타원', th: 'วงรี' },
  shapeDiamond: { de: 'Raute', fr: 'Losange', zh: '菱形', ko: '마름모', th: 'ข้าวหลามตัด' },
  shapeArrow: { de: 'Pfeil', fr: 'Flèche', zh: '箭头', ko: '화살표', th: 'ลูกศร' },
  headerFooter: {
    de: 'Kopf- und Fußzeile',
    fr: 'En-tête et pied de page',
    zh: '页眉和页脚',
    ko: '머리글 및 바닥글',
    th: 'หัวกระดาษและท้ายกระดาษ',
  },
  penBlack: {
    de: 'Schwarzer Stift',
    fr: 'Stylet noir',
    zh: '黑色笔',
    ko: '검은색 펜',
    th: 'ปากกาสีดำ',
  },
  penRed: { de: 'Roter Stift', fr: 'Stylet rouge', zh: '红色笔', ko: '빨간색 펜', th: 'ปากกาสีแดง' },
  trackpad: {
    de: 'Mit Trackpad zeichnen',
    fr: 'Dessiner avec le Trackpad',
    zh: '使用触控板绘图',
    ko: '트랙피드로 그리기',
    th: 'วาดด้วยแทร็คแพ็ด',
  },
  pageSetup: {
    de: 'Seiteneinrichtung',
    fr: 'Mise en page',
    zh: '页面设置',
    ko: '페이지 설정',
    th: 'ตั้งค่าหน้ากระดาษ',
  },
  headingsHeading: {
    de: 'Überschriften',
    fr: 'Titres',
    zh: '标题',
    ko: '머리글',
    th: 'หัวเรื่อง',
  },
  dynamicArray: {
    de: 'Dynamisches Array',
    fr: 'Tableau dynamique',
    zh: '动态数组',
    ko: '동적 배열',
    th: 'อาร์เรย์แบบไดนามิก',
  },
  reset: { de: 'Zurücksetzen', fr: 'Réinitialiser', zh: '重置', ko: '다시 설정', th: 'รีเซ็ต' },
  textToColumns: {
    de: 'Text in Spalten',
    fr: 'Convertir',
    zh: '分列',
    ko: '텍스트 나누기',
    th: 'ข้อความเป็นคอลัมน์',
  },
  analysis: {
    de: 'Analyse-Funktionen',
    fr: 'Utilitaire d’analyse',
    zh: '分析工具库',
    ko: '분석 도구',
    th: 'เครื่องมือวิเคราะห์',
  },
  macros: {
    de: 'Makros anzeigen',
    fr: 'Afficher les macros',
    zh: '查看宏',
    ko: '매크로 보기',
    th: 'ดูแมโคร',
  },
};

/** Group titles worded by a ribbon dictionary entry. */
export const MAC_GROUP_SOURCES: Readonly<Record<string, RibbonKey>> = {
  Options: 'options',
};

export const MAC_GROUP_TRANSLATIONS: Readonly<Record<string, ExtraTranslations>> = {
  'Data from Picture': {
    de: 'Daten aus Bild',
    fr: 'Données de l’image',
    zh: '来自图片的数据',
    ko: '그림의 데이터',
    th: 'ข้อมูลจากรูปภาพ',
  },
  'Protection Changes': {
    de: 'Änderungen',
    fr: 'Modifications',
    zh: '更改',
    ko: '변경 내용',
    th: 'การเปลี่ยนแปลง',
  },
  'Queries & Connections': {
    de: 'Abfragen und Verbindungen',
    fr: 'Requêtes et connexions',
    zh: '查询和连接',
    ko: '쿼리 및 연결',
    th: 'คิวรีและการเชื่อมต่อ',
  },
};

type MacRibbonReason = keyof MacRibbonReasonStrings;

/** Why a ribbon command or menu item is unavailable, in every locale, from the dictionaries. */
export const MAC_CONTEXT_LABELS = Object.fromEntries(
  (Object.keys(macRibbonReasonsEn) as MacRibbonReason[]).map((key) => [
    key,
    Object.fromEntries(
      MAC_LOCALES.map((locale) => [locale, dictionaries[locale].mac.ribbonReasons[key]]),
    ),
  ]),
) as Readonly<Record<MacRibbonReason, Label>>;

export const MAC_MENU_TRANSLATIONS: Readonly<Record<string, ExtraTranslations>> = {
  'mac.page.margins.normal': { de: 'Normal', fr: 'Normales', zh: '普通', ko: '보통', th: 'ปกติ' },
  'mac.page.size.letter': {
    de: 'Letter',
    fr: 'Lettre',
    zh: '信纸',
    ko: 'Letter',
    th: 'Letter',
  },
  'mac.page.size.legal': {
    de: 'Legal',
    fr: 'Juridique',
    zh: '法律',
    ko: 'Legal',
    th: 'Legal',
  },
  'mac.review.editNote': {
    de: 'Notiz bearbeiten',
    fr: 'Modifier la note',
    zh: '编辑备注',
    ko: '메모 편집',
    th: 'แก้ไขบันทึกย่อ',
  },
  'mac.view.freeze.off': {
    de: 'Fixierung aufheben',
    fr: 'Libérer les volets',
    zh: '取消冻结窗格',
    ko: '틀 고정 취소',
    th: 'ยกเลิกการตรึงแนว',
  },
};

const readCoreValue = (locale: Locale, key: RibbonKey): string => {
  const value = (dictionaries[locale].ribbon as unknown as Record<string, unknown>)[key];
  if (typeof value !== 'string')
    throw new Error(`Mac label source ribbon.${String(key)} is not a string.`);
  return value;
};

export const localizedFromCore = (key: RibbonKey): Label =>
  Object.fromEntries(MAC_LOCALES.map((locale) => [locale, readCoreValue(locale, key)])) as Label;

export const localizedFromTranslations = (
  seed: LegacyLabel,
  translations: ExtraTranslations,
): Label => ({ ...seed, ...translations });

export const localizedGroupLabel = (key: string, seed: LegacyLabel): Label => {
  const bundled = MAC_BUNDLE_LABELS[`groups:${key}`];
  if (bundled) return bundled;
  const source = MAC_GROUP_SOURCES[key];
  if (source) return localizedFromCore(source);
  const translations = MAC_GROUP_TRANSLATIONS[key];
  if (translations) return localizedFromTranslations(seed, translations);
  throw new Error(`Missing Mac ribbon translations for group: ${key}`);
};

export const localizedMenuLabel = (key: string, seed: LegacyLabel): Label => {
  const bundled = MAC_BUNDLE_LABELS[`menu:${key}`];
  if (bundled) return bundled;
  const translations = MAC_MENU_TRANSLATIONS[key];
  if (translations) return localizedFromTranslations(seed, translations);
  throw new Error(`Missing Mac ribbon translations for menu item: ${key}`);
};

/** Common function aliases used while a native live catalog is still being
 * assembled. A live engine's localized metadata supersedes these values. */
export const MAC_FUNCTION_TRANSLATIONS: Readonly<Record<string, ExtraTranslations>> = {
  SUM: { de: 'SUMME', fr: 'SOMME', zh: 'SUM', ko: 'SUM', th: 'SUM' },
  AVERAGE: { de: 'MITTELWERT', fr: 'MOYENNE', zh: 'AVERAGE', ko: 'AVERAGE', th: 'AVERAGE' },
  COUNT: { de: 'ANZAHL', fr: 'NB', zh: 'COUNT', ko: 'COUNT', th: 'COUNT' },
  COUNTA: { de: 'ANZAHL2', fr: 'NBVAL', zh: 'COUNTA', ko: 'COUNTA', th: 'COUNTA' },
  COUNTIF: { de: 'ZÄHLENWENN', fr: 'NB.SI', zh: 'COUNTIF', ko: 'COUNTIF', th: 'COUNTIF' },
  SUMIF: { de: 'SUMMEWENN', fr: 'SOMME.SI', zh: 'SUMIF', ko: 'SUMIF', th: 'SUMIF' },
  SUMIFS: { de: 'SUMMEWENNS', fr: 'SOMME.SI.ENS', zh: 'SUMIFS', ko: 'SUMIFS', th: 'SUMIFS' },
  IF: { de: 'WENN', fr: 'SI', zh: 'IF', ko: 'IF', th: 'IF' },
  AND: { de: 'UND', fr: 'ET', zh: 'AND', ko: 'AND', th: 'AND' },
  OR: { de: 'ODER', fr: 'OU', zh: 'OR', ko: 'OR', th: 'OR' },
  NOT: { de: 'NICHT', fr: 'NON', zh: 'NOT', ko: 'NOT', th: 'NOT' },
  XLOOKUP: { de: 'XVERWEIS', fr: 'RECHERCHEX', zh: 'XLOOKUP', ko: 'XLOOKUP', th: 'XLOOKUP' },
  VLOOKUP: { de: 'SVERWEIS', fr: 'RECHERCHEV', zh: 'VLOOKUP', ko: 'VLOOKUP', th: 'VLOOKUP' },
  HLOOKUP: { de: 'WVERWEIS', fr: 'RECHERCHEH', zh: 'HLOOKUP', ko: 'HLOOKUP', th: 'HLOOKUP' },
  MATCH: { de: 'VERGLEICH', fr: 'EQUIV', zh: 'MATCH', ko: 'MATCH', th: 'MATCH' },
  ROUND: { de: 'RUNDEN', fr: 'ARRONDI', zh: 'ROUND', ko: 'ROUND', th: 'ROUND' },
  TODAY: { de: 'HEUTE', fr: 'AUJOURDHUI', zh: 'TODAY', ko: 'TODAY', th: 'TODAY' },
  NOW: { de: 'JETZT', fr: 'MAINTENANT', zh: 'NOW', ko: 'NOW', th: 'NOW' },
  DATE: { de: 'DATUM', fr: 'DATE', zh: 'DATE', ko: 'DATE', th: 'DATE' },
};

export const localizedFunctionName = (name: string, locale: Locale): string => {
  if (locale === 'ja' || locale === 'en') return name;
  const translations = MAC_FUNCTION_TRANSLATIONS[name.toUpperCase()];
  return translations?.[locale] ?? name;
};

export const localizedLabel = (key: string, seed: LegacyLabel): Label => {
  const bundled = MAC_BUNDLE_LABELS[`labels:${key}`];
  if (bundled) return bundled;
  const translations = MAC_LABEL_TRANSLATIONS[key];
  if (translations) return localizedFromTranslations(seed, translations);
  throw new Error(`Missing Mac ribbon translations for label: ${key}`);
};
