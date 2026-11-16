import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';
import { describe, expect, it } from 'vitest';

// UI text comes from the locale dictionaries, never from a branch on the locale code: a
// two-way branch silently leaves every other locale on one language.
const SRC = join(__dirname, '../../../src');
const DICTIONARY_DIR = 'i18n/strings/';

/** Non-text locale branches, each with the reason it is not UI wording. */
const ALLOWED: Readonly<Record<string, string>> = {
  'components/color-palette.ts:list-separator':
    'list punctuation joining colour names, not a translatable label',
  'interact/format-dialog-model.ts:format-code-lists':
    'number-format codes listed per locale, data rather than wording',
};

const sourceFiles = (dir: string): string[] =>
  readdirSync(dir).flatMap((name) => {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) return sourceFiles(path);
    return path.endsWith('.ts') ? [path] : [];
  });

const LOCALE_CODE = "'[a-z]{2}(?:-[A-Z]{2})?'";
const LOCALE_NAME = String.raw`\b\w*(?:lang|locale|Lang|Locale)\w*(?:\(\))?`;
const RULES: readonly { name: string; pattern: RegExp }[] = [
  {
    name: 'locale comparison choosing a string literal',
    pattern: new RegExp(String.raw`${LOCALE_NAME}\s*[!=]==\s*${LOCALE_CODE}\s*\?\s*['"\`]`, 'g'),
  },
  {
    name: 'locale comparison choosing a per-language field',
    pattern: new RegExp(
      String.raw`${LOCALE_NAME}\s*[!=]==\s*${LOCALE_CODE}\s*\?\s*[\w.\[\]]*\.(?:ja|en)\b`,
      'g',
    ),
  },
  {
    // A locale test in an `if` condition, with the wording set inside the guarded statement.
    name: 'locale-guarded statement setting a label literal',
    pattern: new RegExp(
      String.raw`\bif\s*\([^;{]*?${LOCALE_NAME}\s*[!=]==\s*${LOCALE_CODE}[^;{]*?\)\s*(?:\{[^}]*?|[^;{]*?)\b(?:label|title|text|caption|heading)\s*:\s*['"\`]`,
      'g',
    ),
  },
  {
    // A helper taking one Japanese and one English string fills every other locale from one of them.
    name: 'ja/en pair helper',
    pattern: /\(\s*(?:ja\s*:\s*string\s*,\s*en|en\s*:\s*string\s*,\s*ja)\s*:\s*string\b/g,
  },
  { name: 'per-language label field', pattern: /\b\w+Ja\b|\.ja\b|\b\w*_JA_\w*|\bJA_\w+/g },
];

const lineOf = (text: string, index: number): number => text.slice(0, index).split('\n').length;

const findings = (): string[] => {
  const found: string[] = [];
  for (const file of sourceFiles(SRC)) {
    const rel = relative(SRC, file).split('\\').join('/');
    if (rel.startsWith(DICTIONARY_DIR)) continue;
    const text = readFileSync(file, 'utf8');
    for (const rule of RULES) {
      for (const match of text.matchAll(rule.pattern)) {
        found.push(`${rel}:${lineOf(text, match.index ?? 0)} ${rule.name}: ${match[0]}`);
      }
    }
  }
  return found;
};

describe('locale-specific UI text', () => {
  it('comes from the dictionaries, not from a branch on the locale code', () => {
    const allowedFiles = Object.keys(ALLOWED).map((key) => key.split(':')[0]);
    const unexpected = findings().filter(
      (finding) => !allowedFiles.some((file) => finding.startsWith(`${file}:`)),
    );
    expect(unexpected).toEqual([]);
  });

  it('keeps each allowlisted branch present, so a stale entry is removed', () => {
    const all = findings();
    for (const key of Object.keys(ALLOWED)) {
      const file = key.split(':')[0];
      expect(all.some((finding) => finding.startsWith(`${file}:`))).toBe(true);
    }
  });
});

// Wording written per locale in code: an object keyed by locale code, or a label table read
// through its English field. Both belong in the dictionaries or the generated bundle labels.
const BUNDLE_LABELS_FILE = 'toolbar/ribbon/mac/labels-bundle.ts';
const LOCALE_TABLE_RULES: readonly { name: string; pattern: RegExp }[] = [
  { name: 'locale-keyed text', pattern: /\b(?:ja|en|de|fr|ko|th|zh)\s*:\s*['"`]/g },
  { name: 'English field of a label table', pattern: /\.en\b|\[\s*['"]en['"]\s*\]/g },
];

/**
 * Declarations still holding per-locale text, keyed `file:declaration` (`file:*` for the whole
 * file), each with the reason it is not converted yet. Entries leave as their tables move.
 */
const LOCALE_TABLES_PENDING: Readonly<Record<string, string>> = {
  'toolbar/ribbon/mac/labels-data.ts:MAC_LABEL_TRANSLATIONS':
    'ribbon command labels not yet on the bundle route',
  'toolbar/ribbon/mac/labels-data.ts:MAC_GROUP_TRANSLATIONS':
    'ribbon group titles not yet on the bundle route',
  'toolbar/ribbon/mac/labels-data.ts:MAC_MENU_TRANSLATIONS':
    'ribbon menu labels not yet on the bundle route',
  'toolbar/ribbon/mac/labels-data.ts:MAC_FUNCTION_TRANSLATIONS':
    'function names in the ribbon function menus, not yet on the bundle route',
  'toolbar/ribbon/mac/labels.ts:labelSeeds': 'ja/en command label seeds, not yet converted',
  'toolbar/ribbon/mac/labels.ts:MORE_FUNCTION_CATEGORY_LABELS':
    'function category names, not yet on the bundle route',
  'toolbar/ribbon/mac/labels.ts:groupSeeds': 'ja/en group title seeds, not yet converted',
  'toolbar/ribbon/mac/labels.ts:menuLabelSeeds': 'ja/en menu label seeds, not yet converted',
  'toolbar/ribbon/mac/labels.ts:text': 'English fallback for the ja/en seeds above',
  'toolbar/ribbon/mac/labels.ts:macRibbonDisabledReason':
    'menu items carry the English reason as its id, translated here at render',
  'toolbar/ribbon/mac/model.ts:*': 'ribbon model group titles and English-carried item labels',
  'toolbar/ribbon/mac/menu-manifest.ts:*':
    'menu items carry English labels and reasons, translated at render',
  'toolbar/ribbon/mac/shape-gallery.ts:galleryLabel':
    'gallery items carry the English label, translated at render',
  'toolbar/ribbon/mac/shape-gallery.ts:MAC_SHAPE_GALLERY':
    'gallery items carry the English reason, translated at render',
  'toolbar/ribbon/mac/automation-gallery.ts:automationScript':
    'gallery items carry the English label, translated at render',
  'toolbar/ribbon/menus/table.ts:LABELS': 'external-data menu labels, not yet on the bundle route',
  'interact/function-catalog-text.ts:FUNCTION_DESCRIPTIONS':
    'function descriptions, not yet in the dictionaries',
  'commands/input-grammar.ts:BACKSLASH_CURRENCY': 'currency symbol per input locale, data',
  'commands/input-grammar.ts:profileFor': 'English input grammar as the fallback, data',
  'commands/autofill-series.ts:customListsFor': 'English fill lists as the fallback, data',
  'commands/autofill-series.ts:profileFor': 'English fill grammar as the fallback, data',
  'commands/coerce-input.ts:errorLiteral': 'English error literals as the fallback, data',
  'engine/spreadsheet-profile.ts:PROFILE_SUFFIX_FOR_LOCALE': 'engine profile ids, data',
  'format/locale.ts:FORMAT_LOCALE_TAGS': 'BCP 47 tags, data',
};

/** Blanks comments, keeping offsets, so prose that names a field is not read as code. */
const withoutComments = (text: string): string =>
  text
    .replace(/\/\*[\s\S]*?\*\//g, (block) => block.replace(/[^\n]/g, ' '))
    .replace(/(^|[^:'"`\\])\/\/[^\n]*/g, (line, lead: string) => lead.padEnd(line.length, ' '));

const TOP_LEVEL_DECLARATION = /^(?:export\s+)?(?:const|let|function|class)\s+([A-Za-z_$][\w$]*)/gm;

const declarationAt = (text: string, index: number): string => {
  let name = '(module)';
  for (const match of text.slice(0, index).matchAll(TOP_LEVEL_DECLARATION)) name = match[1] ?? name;
  return name;
};

interface LocaleTableFinding {
  readonly site: string;
  readonly declaration: string;
}

const localeTableFindings = (): LocaleTableFinding[] => {
  const found: LocaleTableFinding[] = [];
  for (const file of sourceFiles(SRC)) {
    const rel = relative(SRC, file).split('\\').join('/');
    if (rel.startsWith(DICTIONARY_DIR) || rel === BUNDLE_LABELS_FILE) continue;
    const text = withoutComments(readFileSync(file, 'utf8'));
    for (const rule of LOCALE_TABLE_RULES) {
      for (const match of text.matchAll(rule.pattern)) {
        const index = match.index ?? 0;
        found.push({
          site: `${rel}:${lineOf(text, index)} ${rule.name}: ${match[0]}`,
          declaration: `${rel}:${declarationAt(text, index)}`,
        });
      }
    }
  }
  return found;
};

const pendingFor = (declaration: string): string | undefined =>
  LOCALE_TABLES_PENDING[declaration] ??
  LOCALE_TABLES_PENDING[`${declaration.slice(0, declaration.lastIndexOf(':'))}:*`];

describe('per-locale text tables in code', () => {
  it('appear only in the declarations still pending conversion', () => {
    const unexpected = localeTableFindings()
      .filter((finding) => pendingFor(finding.declaration) === undefined)
      .map((finding) => finding.site);
    expect(unexpected).toEqual([]);
  });

  it('keeps each pending declaration present, so a converted one leaves the list', () => {
    const declarations = new Set(localeTableFindings().map((finding) => finding.declaration));
    const files = new Set([...declarations].map((d) => `${d.slice(0, d.lastIndexOf(':'))}:*`));
    const stale = Object.keys(LOCALE_TABLES_PENDING).filter(
      (key) => !declarations.has(key) && !files.has(key),
    );
    expect(stale).toEqual([]);
  });
});
