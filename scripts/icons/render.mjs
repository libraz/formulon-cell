/**
 * Icon review harness.
 *
 * Renders every ribbon icon to contact sheets and reports the geometry the
 * design tokens are supposed to hold constant. Icons are authored as path data,
 * so this is the only way to actually see what was written.
 *
 *   node scripts/icons/render.mjs            sheets + metrics
 *   node scripts/icons/render.mjs --metrics  metrics only, no screenshots
 *   node scripts/icons/render.mjs undo,redo  only the named icons
 */

import { mkdir, rm, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { chromium } from '@playwright/test';
import { rolldown } from 'rolldown';

const HERE = dirname(fileURLToPath(import.meta.url));
const ROOT = resolve(HERE, '../..');
const OUT = resolve(HERE, 'out');

const args = process.argv.slice(2);
const metricsOnly = args.includes('--metrics');
const sourceArg = args.find((a) => a.startsWith('--source='))?.slice('--source='.length);
const SOURCE = resolve(
  ROOT,
  sourceArg ?? 'packages/formulon-cell/src/toolbar/excel-ribbon-icons.ts',
);
const exportArg = args.find((a) => a.startsWith('--export='))?.slice('--export='.length);
const filter = args.find((a) => !a.startsWith('--'))?.split(',') ?? null;

/** Bundle the icon module so plain Node can import its TypeScript sources. */
const loadIcons = async () => {
  const bundle = await rolldown({
    input: SOURCE,
    platform: 'node',
    resolve: { extensionAlias: { '.js': ['.ts', '.js'] } },
  });
  const { output } = await bundle.generate({ format: 'esm' });
  await mkdir(OUT, { recursive: true });
  const temp = resolve(OUT, '.icons.bundle.mjs');
  await writeFile(temp, output[0].code);
  const module = await import(`${pathToFileURL(temp).href}?t=${process.hrtime.bigint()}`);
  await rm(temp, { force: true });
  return module[exportArg ?? 'EXCEL_RIBBON_ICON_PATHS'];
};

const svgFor = (segments, size) =>
  `<svg viewBox="0 0 24 24" width="${size}" height="${size}">${segments
    .map((s) =>
      [
        '<path',
        `d="${s.d}"`,
        `fill="${s.fill ?? 'none'}"`,
        s.stroke ? `stroke="${s.stroke}"` : '',
        s.strokeWidth ? `stroke-width="${s.strokeWidth}"` : '',
        s.strokeLinecap ? `stroke-linecap="${s.strokeLinecap}"` : '',
        s.strokeLinejoin ? `stroke-linejoin="${s.strokeLinejoin}"` : '',
        s.strokeDasharray ? `stroke-dasharray="${s.strokeDasharray}"` : '',
        s.fillRule ? `fill-rule="${s.fillRule}"` : '',
        s.transform ? `transform="${s.transform}"` : '',
        '/>',
      ]
        .filter(Boolean)
        .join(' '),
    )
    .join('')}</svg>`;

const SHEET_STYLE = `
body{margin:0;padding:16px;background:#fff;font:12px/1.3 -apple-system,'Helvetica Neue',sans-serif;color:#222}
.sheet{display:grid;grid-template-columns:repeat(6,1fr);gap:10px}
figure{margin:0;text-align:center}
.stage{position:relative;width:96px;height:96px;margin:0 auto;background:#fafafa;border:1px solid #e3e3e3}
.grid{position:absolute;inset:0;background-image:
 linear-gradient(to right,#e9eef5 1px,transparent 1px),linear-gradient(to bottom,#e9eef5 1px,transparent 1px),
 linear-gradient(to right,#c8d6e8 1px,transparent 1px),linear-gradient(to bottom,#c8d6e8 1px,transparent 1px);
 background-size:16px 16px,16px 16px,48px 48px,48px 48px}
.stage svg{position:relative}
figcaption{margin-top:4px;font-size:10px;word-break:break-all;display:flex;flex-direction:column;align-items:center;gap:2px}
.mini{display:flex;align-items:flex-end;gap:6px;height:22px}`;

const main = async () => {
  const icons = await loadIcons();
  const entries = Object.entries(icons).filter(([name]) => !filter || filter.includes(name));
  if (entries.length === 0) throw new Error('no icons matched');

  const browser = await chromium.launch();
  const page = await browser.newPage({ deviceScaleFactor: 2 });

  if (!metricsOnly) {
    const perSheet = 36;
    for (let i = 0, sheet = 1; i < entries.length; i += perSheet, sheet += 1) {
      const chunk = entries.slice(i, i + perSheet);
      const cells = chunk
        .map(
          ([name, segs]) =>
            `<figure><div class="stage"><div class="grid"></div>${svgFor(segs, 96)}</div>` +
            `<figcaption><span class="mini">${svgFor(segs, 16)}${svgFor(segs, 20)}</span>${name}</figcaption></figure>`,
        )
        .join('');
      const html = `<!doctype html><meta charset="utf-8"><style>${SHEET_STYLE}</style><div class="sheet">${cells}</div>`;
      await page.setContent(html);
      await page.screenshot({ path: resolve(OUT, `sheet-${sheet}.png`), fullPage: true });
      console.log(`scripts/icons/out/sheet-${sheet}.png  ${chunk.length} icons`);
    }
  }

  await page.setContent('<svg id="s" viewBox="0 0 24 24" width="240" height="240"></svg>');
  const rows = [];
  for (const [name, segs] of entries) {
    const box = await page.evaluate((segments) => {
      const NS = 'http://www.w3.org/2000/svg';
      const svg = document.getElementById('s');
      svg.textContent = '';
      for (const s of segments) {
        const p = document.createElementNS(NS, 'path');
        p.setAttribute('d', s.d);
        p.setAttribute('fill', s.fill ?? 'none');
        if (s.stroke) p.setAttribute('stroke', s.stroke);
        if (s.strokeWidth) p.setAttribute('stroke-width', s.strokeWidth);
        if (s.strokeLinecap) p.setAttribute('stroke-linecap', s.strokeLinecap);
        if (s.strokeLinejoin) p.setAttribute('stroke-linejoin', s.strokeLinejoin);
        if (s.strokeDasharray) p.setAttribute('stroke-dasharray', s.strokeDasharray);
        if (s.fillRule) p.setAttribute('fill-rule', s.fillRule);
        if (s.transform) p.setAttribute('transform', s.transform);
        svg.appendChild(p);
      }
      // Measured in screen space so transformed segments are accounted for,
      // then converted back to viewBox units.
      const frame = svg.getBoundingClientRect();
      const unit = frame.width / 24;
      let x0 = Infinity;
      let y0 = Infinity;
      let x1 = -Infinity;
      let y1 = -Infinity;
      for (const p of svg.querySelectorAll('path')) {
        const b = p.getBoundingClientRect();
        x0 = Math.min(x0, (b.x - frame.x) / unit);
        y0 = Math.min(y0, (b.y - frame.y) / unit);
        x1 = Math.max(x1, (b.right - frame.x) / unit);
        y1 = Math.max(y1, (b.bottom - frame.y) / unit);
      }
      return { x0, y0, x1, y1 };
    }, segs);
    rows.push({
      name,
      w: +(box.x1 - box.x0).toFixed(2),
      h: +(box.y1 - box.y0).toFixed(2),
      cx: +((box.x0 + box.x1) / 2).toFixed(2),
      cy: +((box.y0 + box.y1) / 2).toFixed(2),
      bleed: +Math.max(-box.x0, -box.y0, box.x1 - 24, box.y1 - 24).toFixed(2),
    });
  }
  await browser.close();

  await writeFile(resolve(OUT, 'metrics.json'), `${JSON.stringify(rows, null, 2)}\n`);

  const dims = rows.map((r) => Math.max(r.w, r.h)).sort((a, b) => a - b);
  const at = (q) => dims[Math.floor(dims.length * q)];
  console.log(`\nicons: ${rows.length}`);
  console.log(
    `longest axis: min ${dims[0]} / p25 ${at(0.25)} / p50 ${at(0.5)} / p75 ${at(0.75)} / max ${dims.at(-1)}`,
  );

  const report = (label, hits, format) => {
    console.log(`\n${label} (${hits.length})`);
    for (const r of hits.slice(0, 40)) console.log(`  ${r.name.padEnd(26)} ${format(r)}`);
    if (hits.length > 40) console.log(`  ... +${hits.length - 40}`);
  };

  report(
    'bleeds past the canvas',
    rows.filter((r) => r.bleed > 0.05).sort((a, b) => b.bleed - a.bleed),
    (r) => `by ${r.bleed}`,
  );
  report(
    'off-centre by more than 1px',
    rows
      .filter((r) => Math.abs(r.cx - 12) > 1 || Math.abs(r.cy - 12) > 1)
      .sort((a, b) => Math.hypot(b.cx - 12, b.cy - 12) - Math.hypot(a.cx - 12, a.cy - 12)),
    (r) => `centre (${r.cx}, ${r.cy})`,
  );
  report(
    'undersized: longest axis below 17',
    rows
      .filter((r) => Math.max(r.w, r.h) < 17)
      .sort((a, b) => Math.max(a.w, a.h) - Math.max(b.w, b.h)),
    (r) => `${r.w} x ${r.h}`,
  );
};

await main();
