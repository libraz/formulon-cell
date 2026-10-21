/** Kana forms a phonetic guide can be shown in. */
export type KanaForm = 'hiragana' | 'katakana' | 'halfKatakana' | 'asStored';

const HALF_BASE = 'ｧｱｨｲｩｳｪｴｫｵｶｷｸｹｺｻｼｽｾｿﾀﾁｯﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓｬﾔｭﾕｮﾖﾗﾘﾙﾚﾛﾜｦﾝ';
const FULL_BASE =
  'ァアィイゥウェエォオカキクケコサシスセソタチッツテトナニヌネノハヒフヘホマミムメモャヤュユョヨラリルレロワヲン';
const HALF_PUNCT = '｡｢｣､･ｰ';
const FULL_PUNCT = '。「」、・ー';
/** Full-width voiced/semi-voiced katakana and their base letter. */
const VOICED_BASE = 'カキクケコサシスセソタチツテトハヒフヘホウ';
const VOICED_FULL = 'ガギグゲゴザジズゼゾダヂヅデドバビブベボヴ';
const SEMI_BASE = 'ハヒフヘホ';
const SEMI_FULL = 'パピプペポ';

const halfToFull = new Map<string, string>();
const fullToHalf = new Map<string, string>();
for (const [i, ch] of [...HALF_BASE].entries()) {
  halfToFull.set(ch, FULL_BASE[i] as string);
  fullToHalf.set(FULL_BASE[i] as string, ch);
}
for (const [i, ch] of [...HALF_PUNCT].entries()) {
  halfToFull.set(ch, FULL_PUNCT[i] as string);
  fullToHalf.set(FULL_PUNCT[i] as string, ch);
}
const voiced = new Map<string, string>();
const semiVoiced = new Map<string, string>();
for (const [i, ch] of [...VOICED_FULL].entries()) {
  voiced.set(VOICED_BASE[i] as string, ch);
  fullToHalf.set(ch, `${fullToHalf.get(VOICED_BASE[i] as string)}ﾞ`);
}
for (const [i, ch] of [...SEMI_FULL].entries()) {
  semiVoiced.set(SEMI_BASE[i] as string, ch);
  fullToHalf.set(ch, `${fullToHalf.get(SEMI_BASE[i] as string)}ﾟ`);
}

const toFullKatakana = (text: string): string => {
  let out = '';
  const chars = [...text];
  for (let i = 0; i < chars.length; i++) {
    const ch = chars[i] as string;
    const full = halfToFull.get(ch);
    if (full === undefined) {
      // Hiragana becomes katakana; everything else is left alone.
      const code = ch.charCodeAt(0);
      out += code >= 0x3041 && code <= 0x3096 ? String.fromCharCode(code + 0x60) : ch;
      continue;
    }
    const mark = chars[i + 1];
    const combined =
      mark === 'ﾞ' ? voiced.get(full) : mark === 'ﾟ' ? semiVoiced.get(full) : undefined;
    if (combined !== undefined) i++;
    out += combined ?? full;
  }
  return out;
};

const toHiragana = (text: string): string =>
  [...toFullKatakana(text)]
    .map((ch) => {
      const code = ch.charCodeAt(0);
      return code >= 0x30a1 && code <= 0x30f6 ? String.fromCharCode(code - 0x60) : ch;
    })
    .join('');

const toHalfKatakana = (text: string): string =>
  [...toFullKatakana(text)].map((ch) => fullToHalf.get(ch) ?? ch).join('');

/** Rewrite `text`'s kana in `form`; characters that are not kana pass through. */
export function convertKana(text: string, form: KanaForm): string {
  switch (form) {
    case 'hiragana':
      return toHiragana(text);
    case 'katakana':
      return toFullKatakana(text);
    case 'halfKatakana':
      return toHalfKatakana(text);
    default:
      return text;
  }
}
