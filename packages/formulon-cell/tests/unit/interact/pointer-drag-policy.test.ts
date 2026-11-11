import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { History } from '../../../src/commands/history.js';
import {
  InteractionController,
  registerInteractionController,
} from '../../../src/commands/interaction-controller.js';
import {
  fixedFormPolicy,
  type InteractionPolicy,
  viewerPolicy,
} from '../../../src/commands/interaction-policy.js';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';
import { attachPointer } from '../../../src/interact/pointer.js';
import { gridHitState } from '../../../src/render/grid.js';
import { createSpreadsheetStore, mutators } from '../../../src/store/store.js';

// Header 52 x 30, cells 104 x 28: A1 spans x 52..156, y 30..58; its fill handle sits at 156, 58.
describe('pointer drags start as the host policy allows their operation', () => {
  let host: HTMLElement;
  let store: ReturnType<typeof createSpreadsheetStore>;
  let wb: WorkbookHandle;
  let controller: InteractionController;
  let detach: () => void;
  let unregister: () => void;

  const press = (type: string, x: number, y: number, button = 0): void => {
    host.dispatchEvent(
      new PointerEvent(type, {
        clientX: x,
        clientY: y,
        button,
        buttons: type === 'pointerup' ? 0 : button === 2 ? 2 : 1,
        bubbles: true,
        cancelable: true,
        pointerId: 1,
      }),
    );
  };
  const drag = (from: [number, number], to: [number, number], button = 0): void => {
    press('pointerdown', from[0], from[1], button);
    press('pointermove', to[0], to[1], button);
  };
  const withPolicy = (policy: InteractionPolicy | undefined): void => {
    controller.setPolicy(policy);
  };

  beforeEach(async () => {
    host = document.createElement('div');
    Object.assign(host, {
      setPointerCapture() {},
      releasePointerCapture() {},
      hasPointerCapture: () => false,
    });
    document.body.appendChild(host);
    store = createSpreadsheetStore();
    store.setState((s) => ({
      ...s,
      layout: {
        ...s.layout,
        defaultColWidth: 104,
        defaultRowHeight: 28,
        headerColWidth: 52,
        headerRowHeight: 30,
      },
    }));
    mutators.setActive(store, { sheet: 0, row: 0, col: 0 });
    wb = await WorkbookHandle.createDefault({ preferStub: true });
    controller = new InteractionController({ store, getWb: () => wb, history: new History() });
    unregister = registerInteractionController(store, controller);
    detach = attachPointer(host, store, wb);
  });

  afterEach(() => {
    detach();
    unregister();
    host.remove();
    wb.dispose();
  });

  it('moves by the selection border under a policy that restricts something else', () => {
    withPolicy({ defaultOperation: 'allow', operations: { insertRows: false } });
    drag([100, 58], [100, 120]);
    expect(store.getState().ui.movePreview).not.toBeNull();
  });

  it('keeps the border still while the policy limits cells or formats', () => {
    const policies: InteractionPolicy[] = [
      fixedFormPolicy({ ranges: ['A1:B5'] }),
      { defaultOperation: 'allow', operations: { format: false } },
      viewerPolicy(),
    ];
    for (const policy of policies) {
      withPolicy(policy);
      drag([100, 58], [100, 120]);
      expect(store.getState().ui.movePreview).toBeNull();
      press('pointerup', 100, 120);
    }
  });

  it('starts a right-button fill drag only when its menu fills are allowed', () => {
    const menu = vi.fn();
    host.addEventListener('fc:autofillmenu', menu);
    gridHitState(host).fillHandle = { x: 153, y: 55, w: 6, h: 6 };
    withPolicy({ defaultOperation: 'allow', operations: { insertRows: false } });
    drag([156, 58], [156, 120], 2);
    expect(store.getState().ui.fillPreview).not.toBeNull();
    press('pointerup', 156, 120, 2);
    expect(menu).toHaveBeenCalledTimes(1);

    withPolicy({ defaultOperation: 'allow', operations: { fill: false } });
    drag([156, 58], [156, 120], 2);
    expect(store.getState().ui.fillPreview).toBeNull();
    press('pointerup', 156, 120, 2);
    expect(menu).toHaveBeenCalledTimes(1);
  });

  it('starts a right-button border drag under the same rule as the left one', () => {
    const menu = vi.fn();
    host.addEventListener('fc:movemenu', menu);
    withPolicy({ defaultOperation: 'allow', operations: { insertRows: false } });
    drag([100, 58], [100, 120], 2);
    press('pointerup', 100, 120, 2);
    expect(menu).toHaveBeenCalledTimes(1);

    withPolicy(fixedFormPolicy({ ranges: ['A1:B5'] }));
    drag([100, 58], [100, 120], 2);
    press('pointerup', 100, 120, 2);
    expect(menu).toHaveBeenCalledTimes(1);
  });
});
