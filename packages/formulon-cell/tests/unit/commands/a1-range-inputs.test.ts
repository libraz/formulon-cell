import { afterEach, describe, expect, it } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import { fixedFormPolicy } from '../../../src/commands/interaction-policy.js';
import type { Addr } from '../../../src/engine/types.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import {
  attachNavigationPolicy,
  isNavigationAddrAllowed,
  navigationBoundsFor,
} from '../../../src/interact/navigation-policy.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

describe('A1 text in host range options', () => {
  let dispose: (() => void) | undefined;
  afterEach(() => {
    dispose?.();
    dispose = undefined;
  });

  async function setup() {
    const store = createSpreadsheetStore();
    const wb = await WorkbookHandle.createDefault({ preferStub: true });
    wb.addSheet('Input');
    const controller = new InteractionController({
      store,
      getWb: () => wb,
      history: new History(),
    });
    const unregister = registerInteractionController(store, controller);
    dispose = () => {
      unregister();
      controller.dispose();
      wb.dispose();
    };
    return { store, wb, controller };
  }

  const edit = (addr: Addr) => ({
    type: 'cellBatch' as const,
    operation: 'valueEdit' as const,
    origin: 'editor' as const,
    changes: [{ addr, input: '1' }],
  });

  it('applies unqualified editable text on every sheet', async () => {
    const { controller } = await setup();
    controller.setPolicy(fixedFormPolicy(['B2:B11']));
    expect(controller.execute(edit({ sheet: 0, row: 1, col: 1 })).status).toBe('applied');
    expect(controller.execute(edit({ sheet: 1, row: 10, col: 1 })).status).toBe('applied');
    expect(controller.execute(edit({ sheet: 0, row: 1, col: 2 })).status).toBe('refused');
  });

  it('pins qualified editable text to its sheet', async () => {
    const { controller } = await setup();
    controller.setPolicy(fixedFormPolicy({ ranges: ['Input!B2'] }));
    expect(controller.execute(edit({ sheet: 1, row: 1, col: 1 })).status).toBe('applied');
    expect(controller.execute(edit({ sheet: 0, row: 1, col: 1 })).status).toBe('refused');
  });

  it('rejects bad editable text and unknown sheets when the policy is set', async () => {
    const { controller } = await setup();
    expect(() => fixedFormPolicy(['B2:'])).toThrow(/is not an A1 range/);
    expect(() => controller.setPolicy(fixedFormPolicy(['Nope!B2']))).toThrow(
      /sheet "Nope" does not exist/,
    );
    expect(controller.policy).toBeUndefined();
  });

  it('accepts A1 cells in applyChanges and rejects bad ones without writing', async () => {
    const { controller, wb, store } = await setup();
    mutators.setSheetIndex(store, 1);
    expect(controller.applyChanges([{ addr: 'C3', input: '7' }]).status).toBe('applied');
    expect(wb.getValue({ sheet: 1, row: 2, col: 2 })).toEqual({ kind: 'number', value: 7 });
    expect(controller.applyChanges([{ addr: 'Sheet1!A1', input: '5' }]).status).toBe('applied');
    expect(wb.getValue({ sheet: 0, row: 0, col: 0 })).toEqual({ kind: 'number', value: 5 });
    const bad = controller.applyChanges([
      { addr: 'D4', input: '1' },
      { addr: 'D4:E5', input: '1' },
    ]);
    expect(bad.status).toBe('refused');
    expect(bad.rejections[0]?.reason).toMatch(
      /applyChanges\[1\]\.addr: "D4:E5" is not a single cell/,
    );
    expect(wb.getValue({ sheet: 1, row: 3, col: 3 }).kind).toBe('blank');
  });

  it('resolves viewport range text on the active sheet and keeps the input', async () => {
    const { store, wb } = await setup();
    mutators.setSheetIndex(store, 1);
    const handle = attachNavigationPolicy(store, () => wb, { range: 'A1:D12' });
    expect(navigationBoundsFor(store)).toEqual({ sheet: 1, r0: 0, c0: 0, r1: 11, c1: 3 });
    expect(handle.options?.range).toBe('A1:D12');
    expect(() => handle.setOptions({ range: 'Sheet1!A1:B2' })).toThrow(/active sheet/);
    handle.dispose();
  });

  it('applies unqualified selectable text on every sheet', async () => {
    const { store, wb } = await setup();
    const handle = attachNavigationPolicy(store, () => wb, { selectable: ['B2:C3'] });
    expect(isNavigationAddrAllowed(store, { sheet: 0, row: 1, col: 1 })).toBe(true);
    expect(isNavigationAddrAllowed(store, { sheet: 1, row: 2, col: 2 })).toBe(true);
    expect(isNavigationAddrAllowed(store, { sheet: 1, row: 5, col: 5 })).toBe(false);
    handle.setOptions({ selectable: ['Input!B2'] });
    expect(isNavigationAddrAllowed(store, { sheet: 0, row: 5, col: 5 })).toBe(true);
    expect(isNavigationAddrAllowed(store, { sheet: 1, row: 5, col: 5 })).toBe(false);
    expect(handle.options?.selectable).toEqual(['Input!B2']);
    expect(() => handle.setOptions({ selectable: ['B2:'] })).toThrow(/is not an A1 range/);
    handle.dispose();
  });
});
