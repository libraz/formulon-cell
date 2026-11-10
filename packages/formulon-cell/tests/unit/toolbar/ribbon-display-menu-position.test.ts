import { afterEach, describe, expect, it } from 'vitest';
import { positionRibbonDisplayMenu } from '../../../src/toolbar/ribbon/render-ribbon.js';

const SCALE = 2;

/** A scaled host whose rendered geometry follows the inline `left`/`top`, plus a fixed bias. */
const build = (hostLeft: number, bias: number) => {
  const host = document.createElement('div');
  host.style.transform = `scale(${SCALE})`;
  Object.defineProperty(host, 'offsetWidth', { get: () => 80 });
  Object.defineProperty(host, 'offsetHeight', { get: () => 400 });
  host.getBoundingClientRect = () =>
    ({ left: hostLeft, top: 0, width: 80 * SCALE, height: 400 * SCALE }) as DOMRect;
  const toggle = document.createElement('button');
  toggle.getBoundingClientRect = () =>
    ({ left: 100, right: 146, top: 20, bottom: 44, width: 46, height: 24 }) as DOMRect;
  const menu = document.createElement('div');
  Object.defineProperty(menu, 'offsetWidth', { get: () => 168 });
  Object.defineProperty(menu, 'offsetHeight', { get: () => 120 });
  menu.getBoundingClientRect = () => {
    const left = hostLeft + (Number.parseFloat(menu.style.left) + bias) * SCALE;
    const top = Number.parseFloat(menu.style.top) * SCALE;
    const width = Number.parseFloat(menu.style.width || '168') * SCALE;
    return { left, top, width, height: 240, right: left + width, bottom: top + 240 } as DOMRect;
  };
  host.append(toggle, menu);
  document.body.appendChild(host);
  return { toggle, menu };
};

describe('positionRibbonDisplayMenu under a scaled host', () => {
  const previousWidth = window.innerWidth;
  const previousHeight = window.innerHeight;
  afterEach(() => {
    window.innerWidth = previousWidth;
    window.innerHeight = previousHeight;
    document.body.replaceChildren();
  });

  it('keeps the left edge inside the viewport', () => {
    window.innerWidth = 160;
    window.innerHeight = 900;
    const { toggle, menu } = build(0, 0);
    positionRibbonDisplayMenu(toggle, menu);
    expect(menu.getBoundingClientRect().left).toBeGreaterThanOrEqual(0);
    expect(menu.getBoundingClientRect().right).toBeLessThanOrEqual(160);
  });
});
