import {
  type MacRibbonReasonStrings,
  macRibbonReasonsEn,
} from '../../../i18n/strings/mac/ribbon-reasons.js';
import type { Locale } from '../../../i18n/strings.js';
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

/**
 * Mac labels without a reusable central dictionary field. Values are kept by
 * stable id and supplied for every non-English locale explicitly. Product
 * names and standards such as SmartArt, Python, PDF, and 100% intentionally
 * remain unchanged where that is the Excel UI term.
 */
export const MAC_LABEL_TRANSLATIONS: Readonly<Record<string, ExtraTranslations>> = {
  shapeLine: { de: 'Linie', fr: 'Ligne', zh: '直线', ko: '선', th: 'เส้น' },
  shapeOval: { de: 'Oval', fr: 'Ovale', zh: '椭圆', ko: '타원', th: 'วงรี' },
  shapeDiamond: { de: 'Raute', fr: 'Losange', zh: '菱形', ko: '마름모', th: 'ข้าวหลามตัด' },
  penBlack: {
    de: 'Schwarzer Stift',
    fr: 'Stylet noir',
    zh: '黑色笔',
    ko: '검은색 펜',
    th: 'ปากกาสีดำ',
  },
  penRed: { de: 'Roter Stift', fr: 'Stylet rouge', zh: '红色笔', ko: '빨간색 펜', th: 'ปากกาสีแดง' },
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

export const localizedFromTranslations = (
  seed: LegacyLabel,
  translations: ExtraTranslations,
): Label => ({ ...seed, ...translations });

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
