import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { attachSessionIllustrations } from '../../../src/interact/session-illustrations.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');
const css = readFileSync(join(root, 'src/styles/core/app/overlays/drawings.css'), 'utf8');

// Excel for Mac (ja): a selected shape has eight 7 pt white square handles with
// a soft shadow, and a rotation handle above the top middle on an 11 pt stem.
describe('selected picture and shape handles', () => {
  it('adds a rotation handle above the top middle of an editable object', () => {
    const host = document.createElement('div');
    document.body.appendChild(host);
    const store = createSpreadsheetStore();
    mutators.upsertIllustration(store, {
      id: 'shape-1',
      kind: 'shape',
      shape: 'rectangle',
      sheet: 0,
    });
    const handle = attachSessionIllustrations({ host, store });
    const panel = host.querySelector<HTMLElement>('.fc-illustration');
    expect(panel?.querySelectorAll('.fc-drawing-handle')).toHaveLength(8);
    const rotate = panel?.querySelector('.fc-drawing-rotate');
    expect(rotate?.getAttribute('aria-hidden')).toBe('true');
    handle.detach();
    host.remove();
  });

  it('styles the handles as 7 px white squares with a shadow', () => {
    expect(css).toMatch(
      /\.fc-illustration--selected \.fc-drawing-handle\s*\{[\s\S]*?width: 7px;[\s\S]*?height: 7px;[\s\S]*?margin: -3.5px 0 0 -3.5px;[\s\S]*?border-radius: 0;[\s\S]*?background: #fff;[\s\S]*?box-shadow:/,
    );
  });

  it('sizes the rotation handle 12 px on an 11 px stem', () => {
    expect(css).toMatch(
      /\.fc-drawing-rotate\s*\{[\s\S]*?top: -26px;[\s\S]*?width: 12px;[\s\S]*?height: 12px;[\s\S]*?border-radius: 50%;/,
    );
    expect(css).toMatch(/\.fc-drawing-rotate::before\s*\{[\s\S]*?width: 1px;[\s\S]*?height: 11px;/);
  });
});
