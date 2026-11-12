import { afterEach, describe, expect, it, vi } from 'vitest';
import { onOutsidePress } from '../../../src/interact/outside-press.js';

const press = (target: EventTarget, type = 'mousedown'): void => {
  target.dispatchEvent(new MouseEvent(type, { bubbles: true }));
};

describe('onOutsidePress', () => {
  afterEach(() => {
    vi.unstubAllGlobals();
    document.body.innerHTML = '';
  });

  it('reports a press outside every inside element, not one within them', () => {
    const menu = document.body.appendChild(document.createElement('div'));
    const anchor = document.body.appendChild(document.createElement('button'));
    const other = document.body.appendChild(document.createElement('span'));
    const onOutside = vi.fn();
    const detach = onOutsidePress(document, () => [menu, anchor, null], onOutside, {
      event: 'mousedown',
    });
    press(menu.appendChild(document.createElement('i')));
    press(anchor);
    expect(onOutside).not.toHaveBeenCalled();
    press(other);
    expect(onOutside).toHaveBeenCalledOnce();
    detach();
    press(other);
    expect(onOutside).toHaveBeenCalledOnce();
  });

  it('reads the inside elements at each press', () => {
    let popover: HTMLElement | null = null;
    const onOutside = vi.fn();
    const detach = onOutsidePress(document, () => [popover], onOutside, { event: 'pointerdown' });
    popover = document.body.appendChild(document.createElement('div'));
    press(popover, 'pointerdown');
    press(popover, 'mousedown');
    expect(onOutside).not.toHaveBeenCalled();
    press(document.body, 'pointerdown');
    expect(onOutside).toHaveBeenCalledOnce();
    detach();
  });

  it('treats nodes of another window as nodes', () => {
    // A document from another window has its own Node class.
    vi.stubGlobal('Node', class {});
    const panel = document.body.appendChild(document.createElement('div'));
    const onOutside = vi.fn();
    const detach = onOutsidePress(document, () => [panel], onOutside, { event: 'mousedown' });
    press(panel.appendChild(document.createElement('i')));
    expect(onOutside).not.toHaveBeenCalled();
    press(document.body);
    expect(onOutside).toHaveBeenCalledOnce();
    detach();
  });

  it('runs before the target handler in the capture phase, after it otherwise', () => {
    const order: string[] = [];
    const target = document.body.appendChild(document.createElement('div'));
    target.addEventListener('mousedown', () => order.push('target'));
    const captured = onOutsidePress(
      document,
      () => [],
      () => order.push('capture'),
      {
        event: 'mousedown',
      },
    );
    const bubbled = onOutsidePress(
      document,
      () => [],
      () => order.push('bubble'),
      {
        event: 'mousedown',
        capture: false,
      },
    );
    press(target);
    expect(order).toEqual(['capture', 'target', 'bubble']);
    captured();
    bubbled();
  });
});
