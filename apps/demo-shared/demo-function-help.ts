/**
 * Per-argument function help for the demo Mac formula palette.
 *
 * The engine ships no argument descriptions, so the demos supply a small
 * table through `getFunctionArgumentHelp`. Functions without an entry fall
 * back to the palette's catalog labels with no hint or reference link.
 */

import {
  dictionaryLocaleFor,
  type FunctionArgumentHelp,
  type FunctionArgumentHelpProvider,
  type Locale,
} from '@libraz/formulon-cell';

type DemoHelpLocale = Locale;

type DemoArgumentText = Record<DemoHelpLocale, string>;

const DEMO_ARGUMENT_HELP: Readonly<Record<string, readonly DemoArgumentText[]>> = {
  ABS: [
    {
      en: 'The real number to take the absolute value of.',
      de: 'Die reelle Zahl, deren Absolutwert berechnet werden soll.',
      fr: 'Le nombre réel dont vous souhaitez obtenir la valeur absolue.',
      zh: '要计算绝对值的实数。',
      ko: '절대값을 구할 실수입니다.',
      th: 'จำนวนจริงที่ต้องการหาค่าสัมบูรณ์',
      ja: '絶対値を求める実数を指定します。',
    },
  ],
  ACOS: [
    {
      en: 'Must be a number from -1 to 1.',
      ja: '-1 から 1 の範囲の数値を指定します。',
      de: 'Eine Zahl zwischen -1 und 1.',
      fr: 'Un nombre compris entre -1 et 1.',
      zh: '必须是 -1 到 1 之间的数字。',
      ko: '-1부터 1까지의 숫자여야 합니다.',
      th: 'ต้องเป็นจำนวนตั้งแต่ -1 ถึง 1',
    },
  ],
  ASIN: [
    {
      en: 'Must be a number from -1 to 1.',
      ja: '-1 から 1 の範囲の数値を指定します。',
      de: 'Eine Zahl zwischen -1 und 1.',
      fr: 'Un nombre compris entre -1 et 1.',
      zh: '必须是 -1 到 1 之间的数字。',
      ko: '-1부터 1까지의 숫자여야 합니다.',
      th: 'ต้องเป็นจำนวนตั้งแต่ -1 ถึง 1',
    },
  ],
  SQRT: [
    {
      en: 'The number to take the square root of. Must not be negative.',
      de: 'Die Zahl, deren Quadratwurzel berechnet werden soll. Darf nicht negativ sein.',
      fr: 'Le nombre dont vous souhaitez obtenir la racine carrée. Il ne doit pas être négatif.',
      zh: '要计算平方根的数字，不能为负数。',
      ko: '제곱근을 구할 숫자입니다. 음수일 수 없습니다.',
      th: 'จำนวนที่ต้องการหารากที่สอง ต้องไม่เป็นจำนวนลบ',
      ja: '平方根を求める数値を指定します。負の数は指定できません。',
    },
  ],
};

const REFERENCE_REGION: Record<DemoHelpLocale, string> = {
  en: 'en-us',
  ja: 'ja-jp',
  de: 'de-de',
  fr: 'fr-fr',
  zh: 'zh-cn',
  ko: 'ko-kr',
  th: 'th-th',
};

const helpLocale = (locale: string): DemoHelpLocale => dictionaryLocaleFor(locale);

export const demoFunctionArgumentHelp: FunctionArgumentHelpProvider = (
  functionName,
  argumentIndex,
  locale,
): FunctionArgumentHelp | null => {
  const text = DEMO_ARGUMENT_HELP[functionName]?.[argumentIndex];
  if (!text) return null;
  const lang = helpLocale(locale);
  return {
    description: text[lang],
    url: `https://support.microsoft.com/${REFERENCE_REGION[lang]}/excel/functions/${functionName.toLowerCase()}-function`,
  };
};
