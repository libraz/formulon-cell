import { appendFileSync } from 'node:fs';
import type { Page } from '@playwright/test';

/** Prefix of the console error a dialog layout fault raises. */
export const DIALOG_LAYOUT_PREFIX = '[fc-dialog-layout]';

/**
 * Watches every dialog the page shows and raises a console error, which the scenarios' console
 * checks then fail on, when two visible controls overlap, two text runs overlap, or a text run
 * overlaps a control it is not part of. A button or a caption may sit inside a text field, as an
 * icon, a clear button or a colour well's Automatic does; two buttons never share space. Rects are
 * clipped to their scrolling ancestors, so a list item scrolled out of view does not count, and
 * an open combobox popup is a layer above the dialog, so what it covers does not count either.
 */
const watchDialogLayout = (prefix: string): void => {
  const CONTROLS = [
    'button',
    'input:not([type="hidden"])',
    'select',
    'textarea',
    'a[href]',
    '[role="button"]',
    '[role="tab"]',
    '[role="option"]',
    '[role="checkbox"]',
    '[role="radio"]',
    '[role="menuitem"]',
    '[role="combobox"]',
    '[role="spinbutton"]',
  ].join(',');
  const MIN = 2;
  type Box = { x: number; y: number; r: number; b: number };
  const reported = new Set<string>();

  const visible = (el: Element): boolean => {
    const check = (el as Element & { checkVisibility?: (o: object) => boolean }).checkVisibility;
    if (check && !check.call(el, { opacityProperty: true, visibilityProperty: true })) return false;
    const rect = el.getBoundingClientRect();
    return rect.width >= MIN && rect.height >= MIN;
  };

  /** `rect` cut to the visible part of `el` and every clipping ancestor up to `root`; an
   *  ellipsised span clips its own text. */
  const clip = (el: Element, rect: DOMRect | Box, root: Element): Box | null => {
    let box: Box =
      'r' in rect ? rect : { x: rect.left, y: rect.top, r: rect.right, b: rect.bottom };
    for (
      let node: Element | null = el;
      node && node !== root.parentElement;
      node = node.parentElement
    ) {
      const style = getComputedStyle(node);
      if (style.overflowX === 'visible' && style.overflowY === 'visible') continue;
      const edge = node.getBoundingClientRect();
      box = {
        x: Math.max(box.x, edge.left),
        y: Math.max(box.y, edge.top),
        r: Math.min(box.r, edge.right),
        b: Math.min(box.b, edge.bottom),
      };
    }
    return box.r - box.x >= MIN && box.b - box.y >= MIN ? box : null;
  };

  const overlap = (a: Box, b: Box): Box | null => {
    const box = {
      x: Math.max(a.x, b.x),
      y: Math.max(a.y, b.y),
      r: Math.min(a.r, b.r),
      b: Math.min(a.b, b.b),
    };
    return box.r - box.x >= MIN && box.b - box.y >= MIN ? box : null;
  };
  const inside = (inner: Box, outer: Box): boolean =>
    inner.x >= outer.x - 1 &&
    inner.y >= outer.y - 1 &&
    inner.r <= outer.r + 1 &&
    inner.b <= outer.b + 1;

  const label = (el: Element): string => {
    const cls = [...el.classList].slice(0, 2).join('.');
    const text = (el.getAttribute('aria-label') ?? el.textContent ?? '').trim().slice(0, 24);
    return `${el.tagName.toLowerCase()}${cls ? `.${cls}` : ''}${text ? ` "${text}"` : ''}`;
  };
  const isField = (el: Element): boolean =>
    (el instanceof HTMLInputElement && !['checkbox', 'radio', 'button'].includes(el.type)) ||
    el instanceof HTMLTextAreaElement ||
    el instanceof HTMLSelectElement;

  const report = (root: Element, a: string, b: string, at: Box): void => {
    const name =
      root.getAttribute('aria-label') ??
      root.querySelector('h1,h2,h3,[class*="title"]')?.textContent ??
      '';
    const key = `${name}|${a}|${b}`;
    if (reported.has(key)) return;
    reported.add(key);
    const where = `${Math.round(at.x)},${Math.round(at.y)} ${Math.round(at.r - at.x)}x${Math.round(at.b - at.y)}`;
    console.error(`${prefix} ${name.trim().slice(0, 40)}: ${a} overlaps ${b} at ${where}`);
  };

  /** Whether `el` sits in the open popup of an expanded combobox: a listbox the combobox names in
   *  `aria-controls`, or one beside it in the same parent. The popup is a layer above the dialog,
   *  so what it covers is not a layout fault; the combobox itself is still checked. */
  const inOpenPopup = (el: Element): boolean => {
    const list = el.closest('[role="listbox"]');
    if (!list?.parentElement) return false;
    const opener = (combo: Element): boolean =>
      combo.getAttribute('aria-expanded') === 'true' &&
      combo.getAttribute('aria-haspopup') === 'listbox';
    if (
      list.id &&
      [...document.querySelectorAll('[role="combobox"]')].some(
        (combo) =>
          opener(combo) &&
          (combo.getAttribute('aria-controls') ?? '').split(/\s+/).includes(list.id),
      )
    )
      return true;
    return [...list.parentElement.children].some(
      (sibling) => sibling.getAttribute('role') === 'combobox' && opener(sibling),
    );
  };

  const checkDialog = (root: Element): void => {
    const controls: { el: Element; box: Box }[] = [];
    for (const el of root.querySelectorAll(CONTROLS)) {
      if (!visible(el) || inOpenPopup(el)) continue;
      const box = clip(el, el.getBoundingClientRect(), root);
      if (box) controls.push({ el, box });
    }
    // Controls nested in one another (a checkbox in a tab, an option in a list) are one unit.
    const top = controls.filter(
      ({ el }) => !controls.some((c) => c.el !== el && c.el.contains(el)),
    );
    for (let i = 0; i < top.length; i++) {
      for (let j = i + 1; j < top.length; j++) {
        const a = top[i];
        const b = top[j];
        if (!a || !b) continue;
        const at = overlap(a.box, b.box);
        if (!at) continue;
        if (isField(a.el) && !isField(b.el) && inside(b.box, a.box)) continue;
        if (isField(b.el) && !isField(a.el) && inside(a.box, b.box)) continue;
        report(root, label(a.el), label(b.el), at);
      }
    }

    const runs: { el: Element; box: Box }[] = [];
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let node = walker.nextNode(); node; node = walker.nextNode()) {
      const el = node.parentElement;
      if (!el || !node.textContent?.trim() || el.closest('select,option,textarea,script,style'))
        continue;
      if (!visible(el) || inOpenPopup(el)) continue;
      const range = document.createRange();
      range.selectNodeContents(node);
      for (const rect of range.getClientRects()) {
        if (rect.width < MIN || rect.height < MIN) continue;
        const box = clip(el, rect, root);
        if (box) runs.push({ el, box });
      }
    }
    for (let i = 0; i < runs.length; i++) {
      const a = runs[i];
      if (!a) continue;
      for (let j = i + 1; j < runs.length; j++) {
        const b = runs[j];
        if (!b || a.el === b.el) continue;
        const at = overlap(a.box, b.box);
        if (at) report(root, `text ${label(a.el)}`, `text ${label(b.el)}`, at);
      }
      for (const control of top) {
        if (control.el.contains(a.el) || a.el.contains(control.el)) continue;
        // A caption drawn inside a text field, as Automatic in a colour well is, is part of it.
        if (isField(control.el) && inside(a.box, control.box)) continue;
        const at = overlap(a.box, control.box);
        if (at) report(root, `text ${label(a.el)}`, label(control.el), at);
      }
    }
  };

  // Checks on the frame after any mutation, so a dialog a scenario passes through quickly is
  // checked whatever the machine load; nothing is checked before the fonts have loaded.
  let pending = 0;
  let fontsReady = false;
  const schedule = (): void => {
    if (pending || !fontsReady) return;
    pending = requestAnimationFrame(() => {
      pending = 0;
      for (const root of document.querySelectorAll('[role="dialog"],[role="alertdialog"]')) {
        if (visible(root)) checkDialog(root);
      }
    });
  };
  const start = (): void => {
    new MutationObserver(schedule).observe(document.documentElement, {
      subtree: true,
      childList: true,
      attributes: true,
      attributeFilter: ['hidden', 'class', 'style', 'aria-selected', 'aria-expanded', 'disabled'],
    });
    void document.fonts.ready.then(() => {
      fontsReady = true;
      schedule();
    });
  };
  if (document.documentElement) start();
  else document.addEventListener('DOMContentLoaded', start, { once: true });
};

/** Installs the dialog layout watch before the app loads. With `FC_DIALOG_LAYOUT_LOG` set, each
 *  fault is also appended to that file with the test's title. */
const installed = new WeakSet<Page>();

export const installDialogLayoutCheck = async (page: Page): Promise<void> => {
  // A scenario that mounts twice keeps one watch.
  if (installed.has(page)) return;
  installed.add(page);
  await page.addInitScript(watchDialogLayout, DIALOG_LAYOUT_PREFIX);
  const log = process.env.FC_DIALOG_LAYOUT_LOG;
  if (!log) return;
  page.on('console', (msg) => {
    if (msg.type() === 'error' && msg.text().startsWith(DIALOG_LAYOUT_PREFIX))
      appendFileSync(log, `${page.url()}\t${msg.text()}\n`);
  });
};
