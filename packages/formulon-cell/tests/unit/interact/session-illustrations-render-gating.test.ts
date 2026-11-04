import { describe, expect, it } from 'vitest';
import { attachSessionIllustrations } from '../../../src/interact/session-illustrations.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

const setup = () => {
  const host = document.createElement('div');
  document.body.appendChild(host);
  const store = createSpreadsheetStore();
  mutators.upsertIllustration(store, { id: 'shape-gate', kind: 'shape', shape: 'oval', sheet: 0 });
  const handle = attachSessionIllustrations({ host, store });
  const panel = (): Element | null => host.querySelector('.fc-illustration');
  return { store, handle, panel };
};

describe('attachSessionIllustrations render gating', () => {
  it('keeps the panels when only the cursor selection changes', () => {
    const { store, handle, panel } = setup();
    const before = panel();
    expect(before).not.toBeNull();
    mutators.setActive(store, { sheet: 0, row: 4, col: 4 });
    expect(panel()).toBe(before);
    handle.detach();
  });

  it('rebuilds the panels when an illustration changes', () => {
    const { store, handle, panel } = setup();
    const before = panel();
    mutators.upsertIllustration(store, {
      id: 'shape-gate',
      kind: 'shape',
      shape: 'arrow',
      sheet: 0,
    });
    expect(panel()).not.toBe(before);
    handle.detach();
  });

  it('rebuilds the panels when the viewport scrolls', () => {
    const { store, handle, panel } = setup();
    const before = panel();
    mutators.scrollBy(store, 3, 0);
    expect(panel()).not.toBe(before);
    handle.detach();
  });
});
