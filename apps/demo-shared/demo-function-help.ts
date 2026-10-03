/**
 * Per-argument function help for the demo Mac formula palette.
 *
 * The engine ships no argument descriptions, so the demos supply a small
 * table through `getFunctionArgumentHelp`. Functions without an entry fall
 * back to the palette's catalog labels with no hint or reference link.
 */

import type { FunctionArgumentHelp, FunctionArgumentHelpProvider } from '@libraz/formulon-cell';

type DemoHelpLocale = 'en' | 'ja';

interface DemoArgumentText {
  en: string;
  ja: string;
}

const DEMO_ARGUMENT_HELP: Readonly<Record<string, readonly DemoArgumentText[]>> = {
  ABS: [
    {
      en: 'The real number to take the absolute value of.',
      ja: '絶対値を求める実数を指定します。',
    },
  ],
  ACOS: [{ en: 'Must be a number from -1 to 1.', ja: '-1 から 1 の範囲の数値を指定します。' }],
  ASIN: [{ en: 'Must be a number from -1 to 1.', ja: '-1 から 1 の範囲の数値を指定します。' }],
  SQRT: [
    {
      en: 'The number to take the square root of. Must not be negative.',
      ja: '平方根を求める数値を指定します。負の数は指定できません。',
    },
  ],
};

const REFERENCE_REGION: Record<DemoHelpLocale, string> = { en: 'en-us', ja: 'ja-jp' };

const helpLocale = (locale: string): DemoHelpLocale =>
  locale.trim().toLowerCase().startsWith('ja') ? 'ja' : 'en';

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
