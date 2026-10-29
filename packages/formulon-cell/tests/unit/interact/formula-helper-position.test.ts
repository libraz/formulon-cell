import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  createFormulaCaretPositioner,
  positionFormulaHelper,
} from '../../../src/interact/formula-helper-position.js';

describe('formula helper positioning', () => {
  afterEach(() => {
    document.body.replaceChildren();
    vi.restoreAllMocks();
  });

  it('mirrors textarea typography, box model, wrapping, direction, and scroll offsets', () => {
    const input = document.createElement('textarea');
    input.value = 'formula text';
    input.style.cssText = [
      'font-family: monospace',
      'font-size: 16px',
      'line-height: 22px',
      'padding: 4px 7px',
      'border: 1px solid',
      'white-space: pre-wrap',
      'direction: rtl',
    ].join(';');
    document.body.append(input);
    const positioner = createFormulaCaretPositioner(input);
    positioner.caretRect();
    const mirror = document.querySelector<HTMLElement>('.fc-formula-helper-caret-mirror');

    expect(mirror?.style.fontFamily).toContain('monospace');
    expect(mirror?.style.fontSize).toBe('16px');
    expect(mirror?.style.lineHeight).toBe('22px');
    expect(mirror?.style.paddingLeft).toBe('7px');
    expect(mirror?.style.borderTopWidth).toBe('1px');
    expect(mirror?.style.whiteSpace).toBe('pre-wrap');
    expect(mirror?.style.direction).toBe('rtl');

    positioner.dispose();
    expect(document.querySelector('.fc-formula-helper-caret-mirror')).toBeNull();
  });

  it('keeps the complete formula suffix in the caret measurement mirror', () => {
    const input = document.createElement('textarea');
    input.value = '=SUM(A1)+B1';
    input.setSelectionRange(5, 5);
    document.body.append(input);
    const positioner = createFormulaCaretPositioner(input);

    positioner.caretRect();

    expect(
      document
        .querySelector('.fc-formula-helper-caret-mirror')
        ?.textContent?.replaceAll('\u200b', ''),
    ).toBe(input.value);
    positioner.dispose();
  });

  it('puts the panel below when it fits and flips above before clamping at viewport edges', () => {
    const panel = document.createElement('div');
    document.body.append(panel);
    vi.spyOn(panel, 'getBoundingClientRect').mockReturnValue({
      x: 0,
      y: 0,
      left: 0,
      top: 0,
      right: 220,
      bottom: 72,
      width: 220,
      height: 72,
      toJSON: () => ({}),
    } as DOMRect);
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(500);
    vi.spyOn(window, 'innerHeight', 'get').mockReturnValue(300);

    positionFormulaHelper(panel, { left: 100, right: 100, top: 100, bottom: 120 } as DOMRect);
    expect(Number.parseFloat(panel.style.top)).toBeGreaterThan(120);
    expect(Number.parseFloat(panel.style.left)).toBe(100);

    positionFormulaHelper(panel, { left: 420, right: 420, top: 245, bottom: 265 } as DOMRect);
    expect(Number.parseFloat(panel.style.top)).toBeLessThan(245);
    expect(Number.parseFloat(panel.style.left)).toBeLessThanOrEqual(276);
    expect(panel.style.maxHeight).not.toBe('');
    expect(panel.style.maxWidth).not.toBe('');
  });

  it('converts viewport constraints to CSS pixels inside a scaled portal', () => {
    const panel = document.createElement('div');
    document.body.append(panel);
    vi.spyOn(panel, 'getBoundingClientRect').mockReturnValue({
      x: 20,
      y: 20,
      left: 20,
      top: 20,
      right: 924,
      bottom: 420,
      width: 904,
      height: 400,
      toJSON: () => ({}),
    } as DOMRect);
    Object.defineProperty(panel, 'offsetWidth', { configurable: true, value: 452 });
    Object.defineProperty(panel, 'offsetHeight', { configurable: true, value: 200 });
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(840);
    vi.spyOn(window, 'innerHeight', 'get').mockReturnValue(620);

    positionFormulaHelper(panel, { left: 300, right: 300, top: 200, bottom: 220 } as DOMRect);

    expect(panel.style.maxWidth).toBe('414px');
    expect(panel.style.maxHeight).toBe('304px');
  });

  it('moves beside a visible helper when both vertical placements would overlap', () => {
    const panel = document.createElement('div');
    const avoid = document.createElement('div');
    document.body.append(panel, avoid);
    vi.spyOn(panel, 'getBoundingClientRect').mockImplementation(() => {
      const left = Number.parseFloat(panel.style.left) || 0;
      const top = Number.parseFloat(panel.style.top) || 0;
      return {
        x: left,
        y: top,
        left,
        top,
        right: left + 120,
        bottom: top + 60,
        width: 120,
        height: 60,
        toJSON: () => ({}),
      } as DOMRect;
    });
    vi.spyOn(avoid, 'getBoundingClientRect').mockReturnValue({
      x: 100,
      y: 60,
      left: 100,
      top: 60,
      right: 300,
      bottom: 260,
      width: 200,
      height: 200,
      toJSON: () => ({}),
    } as DOMRect);
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(500);
    vi.spyOn(window, 'innerHeight', 'get').mockReturnValue(300);

    positionFormulaHelper(panel, { left: 160, right: 160, top: 140, bottom: 160 } as DOMRect, {
      avoid,
    });

    expect(Number.parseFloat(panel.style.left)).toBeGreaterThanOrEqual(304);
  });

  it('checks overlap after viewport clamping moves a near-edge panel inward', () => {
    const panel = document.createElement('div');
    const avoid = document.createElement('div');
    document.body.append(panel, avoid);
    vi.spyOn(panel, 'getBoundingClientRect').mockImplementation(() => {
      const left = Number.parseFloat(panel.style.left) || 0;
      const top = Number.parseFloat(panel.style.top) || 0;
      return {
        x: left,
        y: top,
        left,
        top,
        right: left + 220,
        bottom: top + 72,
        width: 220,
        height: 72,
        toJSON: () => ({}),
      } as DOMRect;
    });
    vi.spyOn(avoid, 'getBoundingClientRect').mockReturnValue({
      x: 200,
      y: 164,
      left: 200,
      top: 164,
      right: 400,
      bottom: 236,
      width: 200,
      height: 72,
      toJSON: () => ({}),
    } as DOMRect);
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(500);
    vi.spyOn(window, 'innerHeight', 'get').mockReturnValue(300);

    positionFormulaHelper(panel, { left: 420, right: 420, top: 140, bottom: 160 } as DOMRect, {
      avoid,
    });

    const placed = panel.getBoundingClientRect();
    expect(placed.right).toBeLessThanOrEqual(494);
    expect(placed.top).toBeGreaterThanOrEqual(6);
    expect(placed.bottom).toBeLessThanOrEqual(294);
    expect(placed.left < 400 && placed.right > 200 && placed.top < 236 && placed.bottom > 164).toBe(
      false,
    );
  });

  it('notifies subscribers when input caret and ancestor viewport positions change', () => {
    const host = document.createElement('div');
    const input = document.createElement('textarea');
    host.append(input);
    document.body.append(host);
    const positioner = createFormulaCaretPositioner(input);
    const callback = vi.fn();
    const unsubscribe = positioner.subscribe(callback);

    input.dispatchEvent(new Event('scroll'));
    document.dispatchEvent(new Event('selectionchange'));
    window.dispatchEvent(new Event('resize'));
    expect(callback).toHaveBeenCalledTimes(3);

    unsubscribe();
    positioner.dispose();
    document.dispatchEvent(new Event('selectionchange'));
    expect(callback).toHaveBeenCalledTimes(3);
  });

  it('ignores scroll events from popup siblings while tracking input and ancestor scrolls', () => {
    const host = document.createElement('div');
    const input = document.createElement('textarea');
    host.append(input);
    const portal = document.createElement('div');
    const list = document.createElement('div');
    const option = document.createElement('button');
    list.append(option);
    portal.append(list);
    document.body.append(host, portal);
    const positioner = createFormulaCaretPositioner(input);
    const callback = vi.fn();
    const unsubscribe = positioner.subscribe(callback);

    input.dispatchEvent(new Event('scroll'));
    host.dispatchEvent(new Event('scroll'));
    option.dispatchEvent(new Event('scroll'));
    document.dispatchEvent(new Event('scroll'));
    window.dispatchEvent(new Event('scroll'));

    expect(callback).toHaveBeenCalledTimes(4);
    unsubscribe();
    positioner.dispose();
  });

  it('closes eligibility when the editor is parked outside the viewport', () => {
    const input = document.createElement('textarea');
    document.body.append(input);
    vi.spyOn(window, 'innerWidth', 'get').mockReturnValue(500);
    vi.spyOn(window, 'innerHeight', 'get').mockReturnValue(300);
    const bounds = vi.spyOn(input, 'getBoundingClientRect');
    const positioner = createFormulaCaretPositioner(input);
    bounds.mockReturnValue({
      x: -400,
      y: -60,
      left: -400,
      top: -60,
      right: -100,
      bottom: -40,
      width: 300,
      height: 20,
      toJSON: () => ({}),
    } as DOMRect);
    expect(positioner.isAnchorVisible()).toBe(false);

    bounds.mockReturnValue({
      x: 10,
      y: 10,
      left: 10,
      top: 10,
      right: 310,
      bottom: 30,
      width: 300,
      height: 20,
      toJSON: () => ({}),
    } as DOMRect);
    expect(positioner.isAnchorVisible()).toBe(true);
    bounds.mockReturnValue({
      x: 0,
      y: 0,
      left: 0,
      top: 0,
      right: 0,
      bottom: 0,
      width: 0,
      height: 0,
      toJSON: () => ({}),
    } as DOMRect);
    expect(positioner.isAnchorVisible()).toBe(true);
    positioner.dispose();
  });
});
