import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { Spreadsheet } from '../../../../src/mount.js';
import { type MountedStubSheet, mountStubSheet } from '../../../test-utils/mount.js';
import { scaledTimeout } from '../../../timeouts.js';
import { stubHelpers } from './fixtures.js';

vi.setConfig({ testTimeout: scaledTimeout(20_000) });

const HOST_LEFT = 300;
const HOST_TOP = 200;

const rect = (left: number, top: number, width: number, height: number): DOMRect =>
  ({
    left,
    top,
    width,
    height,
    right: left + width,
    bottom: top + height,
    x: left,
    y: top,
  }) as DOMRect;

/** Ribbon popups inside a host whose containment makes it the fixed containing block. */
describe('ribbon popups in a contained host', () => {
  let sheet: MountedStubSheet;
  let host: HTMLElement;

  beforeEach(async () => {
    sheet = await mountStubSheet({ locale: 'en' });
    host = document.createElement('div');
    host.style.contain = 'strict';
    host.getBoundingClientRect = () => rect(HOST_LEFT, HOST_TOP, 900, 500);
    document.body.appendChild(host);
  });
  afterEach(() => {
    sheet.dispose();
    host.remove();
  });

  const anchorAt = (button: HTMLElement, left: number, top: number): void => {
    button.getBoundingClientRect = () => rect(left, top, 60, 24);
  };

  it('places a dynamic dropdown relative to the host, not the viewport', () => {
    Spreadsheet.mountToolbar(host, sheet.instance, {
      dynamicDropdowns: true,
      helpers: stubHelpers(),
    });
    const button = host.querySelector<HTMLButtonElement>('[data-ribbon-command="shapesInsert"]');
    if (!button) throw new Error('no Shapes button');
    anchorAt(button, HOST_LEFT + 220, HOST_TOP + 15);
    button.click();
    const menu = host.querySelector<HTMLElement>('#menu-shapes-insert');
    expect(menu?.hidden).toBe(false);
    expect(menu?.style.left).toBe('220px');
    expect(menu?.style.top).toBe(`${15 + 24 + 3}px`);
  });
});
