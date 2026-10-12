import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const packageRoot = process.argv[2];
if (!packageRoot) throw new Error('Expected the staged npm package directory');
const pkg = JSON.parse(await readFile(resolve(packageRoot, 'package.json'), 'utf8'));
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';
const { default: factory } = await import(
  pathToFileURL(resolve(packageRoot, pkg.exports[entry].import)).href
);
const module = await factory();
assert.equal(module.versionString(), pkg.version, `WASM/package version (${entry})`);
const workbook = module.Workbook.createDefault();
let profiles = 0;
try {
  for (const host of ['win', 'mac']) {
    for (const locale of ['en_US', 'ja_JP', 'de_DE', 'fr_FR', 'zh_CN', 'ko_KR', 'th_TH']) {
      const profile = `${host}-365-${locale}`;
      const result = workbook.localeFacts(profile);
      assert.equal(result.status.ok, true, profile);
      assert.ok(result.facts, profile);
      assert.equal(result.facts.errorNames.length, 17, profile);
      assert.equal(result.facts.measured, host === 'mac' || locale === 'ja_JP', profile);
      const canonical = '=SUM(1.5,2)+IF(TRUE,0,1)';
      const localized = workbook.localizeFormula(canonical, profile);
      assert.equal(localized.status.ok, true, profile);
      if (locale === 'de_DE') assert.equal(localized.value, '=SUMME(1,5;2)+WENN(WAHR;0;1)');
      if (locale === 'fr_FR') assert.equal(localized.value, '=SOMME(1,5;2)+SI(VRAI;0;1)');
      const restored = workbook.canonicalizeFormula(localized.value, profile);
      assert.equal(restored.status.ok, true, profile);
      assert.equal(restored.value, canonical, profile);
      assert.equal(workbook.setFormula(0, 0, 0, restored.value).ok, true, profile);
      assert.equal(workbook.recalc().ok, true, profile);
      const value = workbook.getValue(0, 0, 0);
      assert.equal(value.status.ok, true, profile);
      assert.equal(value.value.number, 3.5, profile);
      const protectedText = '=SUM({1,2;3,4})+LEN("SUM(1,2)")';
      const converted = workbook.localizeFormula(protectedText, profile);
      assert.equal(converted.status.ok, true, profile);
      const inverse = workbook.canonicalizeFormula(converted.value, profile);
      assert.equal(inverse.status.ok, true, profile);
      assert.equal(inverse.value, protectedText, profile);
      profiles += 1;
    }
  }
  assert.equal(workbook.localeFacts('unknown').status.ok, false);
  console.log(`Locale API smoke: ${profiles} profiles (${entry})`);
} finally {
  workbook.delete();
}
