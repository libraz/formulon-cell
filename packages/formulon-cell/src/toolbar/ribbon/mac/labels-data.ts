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

/** Why a Mac ribbon command or menu row is unavailable: a key of the dictionaries' reasons. */
export type MacRibbonReason = keyof MacRibbonReasonStrings;

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

export const localizedLabel = (key: string, seed: LegacyLabel): Label => {
  const bundled = MAC_BUNDLE_LABELS[`labels:${key}`];
  if (bundled) return bundled;
  const translations = MAC_LABEL_TRANSLATIONS[key];
  if (translations) return localizedFromTranslations(seed, translations);
  throw new Error(`Missing Mac ribbon translations for label: ${key}`);
};
