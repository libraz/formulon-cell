import type { FeatureFlags } from '../extensions/index.js';
import type { Strings } from '../i18n/strings.js';

export type ChromeSlot = 'formulabar' | 'viewbar' | 'sheetbar' | 'statusbar' | 'watchDock';

export interface MountChrome {
  formulabar: HTMLDivElement;
  tag: HTMLInputElement;
  fx: HTMLButtonElement;
  fxCancel: HTMLButtonElement;
  fxAccept: HTMLButtonElement;
  fxInput: HTMLTextAreaElement;
  fxExpand: HTMLButtonElement;
  viewbar: HTMLDivElement;
  grid: HTMLDivElement;
  canvas: HTMLCanvasElement;
  a11y: HTMLDivElement;
  statusbar: HTMLDivElement;
  sheetbar: HTMLDivElement;
  firstSheet: HTMLButtonElement;
  lastSheet: HTMLButtonElement;
  sheetTabs: HTMLDivElement;
  addSheetBtn: HTMLButtonElement;
  sheetMenu: HTMLDivElement;
  watchDock: HTMLDivElement;
  refreshFormulaBarLabels(): void;
  setChromeAttached(slot: ChromeSlot, on: boolean): void;
}

interface CreateMountChromeOptions {
  host: HTMLElement;
  getStrings: () => Strings;
  flags: FeatureFlags;
  onSheetTabContextMenu(idx: number, tab: HTMLButtonElement, x: number, y: number): void;
}

export function createMountChrome({
  host,
  getStrings,
  flags,
  onSheetTabContextMenu,
}: CreateMountChromeOptions): MountChrome {
  const strings = getStrings();
  const formulabar = document.createElement('div');
  formulabar.className = 'fc-host__formulabar';

  const tag = document.createElement('input');
  tag.type = 'text';
  tag.className = 'fc-host__formulabar-tag';
  tag.spellcheck = false;
  tag.autocomplete = 'off';
  tag.setAttribute('aria-label', strings.a11y.nameBox);
  tag.value = 'A1';

  const fx = document.createElement('button');
  fx.type = 'button';
  fx.className = 'fc-host__formulabar-fx';
  fx.textContent = 'ƒx';
  fx.tabIndex = -1;
  fx.setAttribute('aria-label', strings.fxDialog?.fxButtonLabel ?? strings.a11y.formulaBar);

  const fxCancel = document.createElement('button');
  fxCancel.type = 'button';
  fxCancel.className = 'fc-host__formulabar-action fc-host__formulabar-action--cancel';
  fxCancel.textContent = '×';
  fxCancel.tabIndex = -1;
  fxCancel.disabled = true;
  fxCancel.setAttribute('aria-label', strings.a11y.cancelFormulaEdit);

  const fxAccept = document.createElement('button');
  fxAccept.type = 'button';
  fxAccept.className = 'fc-host__formulabar-action fc-host__formulabar-action--accept';
  fxAccept.textContent = '✓';
  fxAccept.tabIndex = -1;
  fxAccept.disabled = true;
  fxAccept.setAttribute('aria-label', strings.a11y.enterFormula);

  const fxInput = document.createElement('textarea');
  fxInput.className = 'fc-host__formulabar-input';
  fxInput.spellcheck = false;
  fxInput.autocomplete = 'off';
  fxInput.rows = 1;
  fxInput.wrap = 'soft';
  fxInput.setAttribute('aria-label', strings.a11y.formulaBar);

  const fxExpand = document.createElement('button');
  fxExpand.type = 'button';
  fxExpand.className = 'fc-host__formulabar-expand';
  fxExpand.setAttribute('aria-expanded', 'false');
  fxExpand.tabIndex = -1;
  fxExpand.textContent = '⌄';

  const refreshFormulaBarLabels = (): void => {
    const strings = getStrings();
    fx.setAttribute('aria-label', strings.fxDialog?.fxButtonLabel ?? strings.a11y.formulaBar);
    fxCancel.setAttribute('aria-label', strings.a11y.cancelFormulaEdit);
    fxAccept.setAttribute('aria-label', strings.a11y.enterFormula);
    fxInput.setAttribute('aria-label', strings.a11y.formulaBar);
    const expanded = fxExpand.getAttribute('aria-expanded') === 'true';
    fxExpand.setAttribute(
      'aria-label',
      expanded ? strings.a11y.collapseFormulaBar : strings.a11y.expandFormulaBar,
    );
  };
  refreshFormulaBarLabels();

  fxExpand.addEventListener('click', () => {
    const expanded = formulabar.dataset.fcExpanded === '1';
    if (expanded) {
      delete formulabar.dataset.fcExpanded;
      fxExpand.setAttribute('aria-expanded', 'false');
      fxExpand.textContent = '⌄';
      fxInput.rows = 1;
    } else {
      formulabar.dataset.fcExpanded = '1';
      fxExpand.setAttribute('aria-expanded', 'true');
      fxExpand.textContent = '⌃';
      fxInput.rows = 4;
    }
    refreshFormulaBarLabels();
  });
  formulabar.append(tag, fxCancel, fxAccept, fx, fxInput, fxExpand);

  const viewbar = document.createElement('div');
  viewbar.className = 'fc-viewbar';

  const grid = document.createElement('div');
  grid.className = 'fc-host__grid';
  const canvas = document.createElement('canvas');
  canvas.className = 'fc-host__canvas';
  grid.appendChild(canvas);

  const a11y = document.createElement('div');
  a11y.className = 'fc-host__a11y';
  a11y.setAttribute('aria-live', 'polite');
  grid.appendChild(a11y);

  const statusbar = document.createElement('div');
  statusbar.className = 'fc-host__statusbar';
  const sheetbar = document.createElement('div');
  sheetbar.className = 'fc-host__sheetbar';
  const sheetNav = document.createElement('div');
  sheetNav.className = 'fc-host__sheetbar-nav';
  const firstSheet = document.createElement('button');
  firstSheet.type = 'button';
  firstSheet.className = 'fc-host__sheetbar-navbtn';
  appendSheetbarIcon(firstSheet, ['M12.5 4.5 7 10l5.5 5.5']);
  const lastSheet = document.createElement('button');
  lastSheet.type = 'button';
  lastSheet.className = 'fc-host__sheetbar-navbtn';
  appendSheetbarIcon(lastSheet, ['M7.5 4.5 13 10l-5.5 5.5']);
  sheetNav.append(firstSheet, lastSheet);

  const sheetTabs = document.createElement('div');
  sheetTabs.className = 'fc-host__sheetbar-tabs';
  sheetTabs.setAttribute('role', 'tablist');
  const addSheetBtn = document.createElement('button');
  addSheetBtn.type = 'button';
  addSheetBtn.className = 'fc-host__sheetbar-add';
  appendSheetbarIcon(addSheetBtn, ['M10 4.5v11', 'M4.5 10h11']);
  sheetbar.append(sheetNav, sheetTabs, addSheetBtn);

  const sheetMenu = document.createElement('div');
  sheetMenu.className = 'fc-sheetmenu';
  sheetMenu.hidden = true;
  sheetMenu.setAttribute('role', 'menu');
  document.body.appendChild(sheetMenu);
  sheetbar.addEventListener(
    'contextmenu',
    (e) => {
      e.preventDefault();
      e.stopPropagation();
      const tab =
        e.target instanceof Element
          ? e.target.closest<HTMLButtonElement>('.fc-host__sheetbar-tab')
          : null;
      const idx = tab ? Number(tab.dataset.fcSheetIndex) : NaN;
      if (!tab || !Number.isInteger(idx)) return;
      onSheetTabContextMenu(idx, tab, e.clientX, e.clientY);
    },
    true,
  );
  sheetMenu.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    e.stopPropagation();
  });

  const watchDock = document.createElement('div');
  watchDock.dataset.fcWatch = 'dock';
  watchDock.className = 'fc-host__watchdock';

  const setChromeAttached = (slot: ChromeSlot, on: boolean): void => {
    const el =
      slot === 'formulabar'
        ? formulabar
        : slot === 'viewbar'
          ? viewbar
          : slot === 'sheetbar'
            ? sheetbar
            : slot === 'statusbar'
              ? statusbar
              : watchDock;
    if (on) {
      if (el.parentElement === host) return;
      if (slot === 'formulabar' || slot === 'viewbar') {
        host.insertBefore(el, grid);
      } else if (slot === 'sheetbar') {
        if (statusbar.parentElement === host) host.insertBefore(el, statusbar);
        else if (watchDock.parentElement === host) host.insertBefore(el, watchDock);
        else host.appendChild(el);
      } else if (slot === 'statusbar') {
        if (watchDock.parentElement === host) host.insertBefore(el, watchDock);
        else host.appendChild(el);
      } else {
        host.appendChild(el);
      }
    } else if (el.parentElement === host) {
      host.removeChild(el);
    }
  };

  host.appendChild(grid);
  setChromeAttached('formulabar', flags.formulaBar === true);
  setChromeAttached('viewbar', flags.viewToolbar === true);
  setChromeAttached('sheetbar', flags.sheetTabs === true);
  setChromeAttached('statusbar', flags.statusBar === true);
  setChromeAttached('watchDock', flags.watchWindow === true);

  return {
    formulabar,
    tag,
    fx,
    fxCancel,
    fxAccept,
    fxInput,
    fxExpand,
    viewbar,
    grid,
    canvas,
    a11y,
    statusbar,
    sheetbar,
    firstSheet,
    lastSheet,
    sheetTabs,
    addSheetBtn,
    sheetMenu,
    watchDock,
    refreshFormulaBarLabels,
    setChromeAttached,
  };
}

function appendSheetbarIcon(
  button: HTMLButtonElement,
  paths: readonly string[],
  viewBox = '0 0 20 20',
): void {
  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('class', 'fc-host__icon');
  svg.setAttribute('viewBox', viewBox);
  svg.setAttribute('fill', 'none');
  svg.setAttribute('stroke', 'currentColor');
  svg.setAttribute('stroke-width', '1.5');
  svg.setAttribute('stroke-linecap', 'round');
  svg.setAttribute('stroke-linejoin', 'round');
  svg.setAttribute('aria-hidden', 'true');
  for (const d of paths) {
    const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    path.setAttribute('d', d);
    svg.appendChild(path);
  }
  button.replaceChildren(svg);
}
