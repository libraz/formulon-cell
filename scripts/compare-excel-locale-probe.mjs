import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { buildRibbonModel, WorkbookHandle } from '../packages/formulon-cell/dist/index.js';

const capturePath = process.argv[2];
if (!capturePath)
  throw new Error('Usage: node scripts/compare-excel-locale-probe.mjs <capture.json>');
const capture = JSON.parse(fs.readFileSync(capturePath, 'utf8'));
assert.equal(capture.cleanup.preferences_restored_exactly, true);
assert.deepEqual(capture.locales.map((locale) => locale.tag).sort(), [
  'de_DE',
  'en_US',
  'fr_FR',
  'ja_JP',
  'ko_KR',
  'th_TH',
  'zh_CN',
]);
const report = { comparisons: 0, matches: 0, differences: [], locales: [] };
const wb = await WorkbookHandle.createDefault();
const addr = { sheet: 0, row: 0, col: 0 };
const compare = (locale, name, surface, excel, engine) => {
  report.comparisons++;
  if (excel === engine) report.matches++;
  else report.differences.push({ locale, name, surface, excel, engine });
};
try {
  for (const locale of capture.locales) {
    assert.equal(locale.status, 'ok');
    assert.equal(Object.keys(locale.cases).length, 18);
    assert.equal(wb.setSpreadsheetProfileId(`mac-${locale.tag}`), true);
    report.locales.push(locale.tag);
    const tabs = buildRibbonModel(locale.tag.slice(0, 2), { profile: 'excel365Mac' });
    const observedTabs = locale.ui.ribbon
      .filter((item) => item.role === 'AXRadioButton')
      .slice(0, 9);
    assert.equal(observedTabs.length, 9);
    assert.equal(tabs.length, 9);
    for (let index = 0; index < tabs.length; index++) {
      compare(
        locale.tag,
        tabs[index].id,
        'ribbon_tab',
        observedTabs[index].name,
        tabs[index].label,
      );
    }
    for (const [name, c] of Object.entries(locale.cases)) {
      assert.equal(c.status, 'ok');
      assert.equal(c.value_status, 'value');
      assert.equal(c.formula_local?.status, 'read');
      assert.equal(c.display_text?.status, 'read');
      wb.setBlank({ sheet: 0, row: 19, col: 0 });
      wb.setBlank({ sheet: 0, row: 20, col: 0 });
      if (name === 'countif_kana') {
        wb.setText({ sheet: 0, row: 19, col: 0 }, 'ガ');
        wb.setText({ sheet: 0, row: 20, col: 0 }, 'ｶﾞ');
      }
      wb.setFormula(addr, c.input);
      const v = wb.getValue(addr);
      const value = v.kind === 'number' || v.kind === 'text' || v.kind === 'bool' ? v.value : null;
      let excel = c.value;
      if (name === 'date' && typeof excel === 'string')
        excel = (Date.parse(`${excel}Z`) - Date.UTC(1899, 11, 30)) / 86400000;
      compare(locale.tag, name, 'value', excel, value);
      if (c.formula_local?.status === 'read' && c.formula_local.value)
        compare(
          locale.tag,
          name,
          'formula_local',
          c.formula_local.value,
          wb.getFormulaForDisplay(wb.cellFormula(addr)),
        );
      if (name !== 'date' && c.display_text?.status === 'read')
        compare(
          locale.tag,
          name,
          'display',
          c.display_text.value,
          wb.getDisplayText(addr.sheet, addr.row, addr.col),
        );
    }
  }
} finally {
  wb.dispose();
}
const output = path.join(
  path.dirname(capturePath),
  `excel-locale-parity-${capture.captured_at.slice(0, 10)}.json`,
);
assert.equal(report.comparisons, 434);
fs.writeFileSync(output, `${JSON.stringify(report, null, 2)}\n`);
console.log(JSON.stringify(report, null, 2));
if (report.differences.length) process.exitCode = 1;
