import { readFileSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { afterEach, describe, expect, it } from 'vitest';
import { enhanceCustomSelect } from '../../../src/interact/custom-select.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../../..');

describe('enhanceCustomSelect', () => {
  afterEach(() => {
    document.body.innerHTML = '';
  });

  it('projects disabled option reasons through the shared a11y helper', () => {
    const select = document.createElement('select');
    select.setAttribute('aria-label', 'Calculation mode');
    select.innerHTML = `
      <option value="auto">Automatic</option>
      <option value="tables" disabled data-disabled-reason="Requires engine support">Data tables</option>
    `;
    document.body.appendChild(select);

    const handle = enhanceCustomSelect(select);
    document.querySelector<HTMLButtonElement>('.fc-select__button')?.click();

    const disabled = document.querySelector<HTMLButtonElement>(
      '.fc-select__option[data-value="tables"]',
    );
    const enabled = document.querySelector<HTMLButtonElement>(
      '.fc-select__option[data-value="auto"]',
    );

    expect(disabled?.disabled).toBe(true);
    expect(disabled?.dataset.disabledReason).toBe('Requires engine support');
    expect(disabled?.getAttribute('aria-description')).toBe('Requires engine support');
    expect(disabled?.title).toBe('Requires engine support');
    expect(enabled?.dataset.disabledReason).toBeUndefined();
    expect(enabled?.getAttribute('aria-description')).toBeNull();

    handle?.dispose();
  });

  it('keeps trigger and option row button DOM on the shared interaction primitive', () => {
    const source = readFileSync(join(root, 'src/interact/custom-select.ts'), 'utf8');

    expect(source).toContain("import { createInteractionButton } from './chip-button.js'");
    expect(source).toContain('function createCustomSelectButton');
    expect(source).toContain('function createCustomSelectOptionRow');
    expect(source).toContain('const button = createCustomSelectButton(select)');
    expect(source).toContain('const row = createCustomSelectOptionRow(opt)');
    expect(source).toContain('projectDisabledState(row, opt.disabled');
    expect(source).not.toContain("document.createElement('button')");
  });
});
