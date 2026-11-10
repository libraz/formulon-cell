import type { FeatureFlags } from '../extensions/index.js';
import { isMacHost } from '../extensions/ui-options.js';
import type { Strings } from '../i18n/strings.js';
import { overlayPortalFor } from '../interact/overlay-portal.js';
import { projectDisabledState } from '../toolbar/menu-a11y.js';
import { appendHostIcon, createHostButton } from './chrome-buttons.js';

export type ChromeSlot =
  | 'formulabar'
  | 'viewbar'
  | 'localeSwitcher'
  | 'sheetbar'
  | 'statusbar'
  | 'watchDock';

export interface MountChrome {
  formulabar: HTMLDivElement;
  tag: HTMLInputElement;
  fx: HTMLButtonElement;
  fxCancel: HTMLButtonElement;
  fxAccept: HTMLButtonElement;
  fxInput: HTMLTextAreaElement;
  fxExpand: HTMLButtonElement;
  viewbar: HTMLDivElement;
  localeSwitcher: HTMLDivElement;
  sheetWorkspace: HTMLDivElement;
  sheetSurface: HTMLDivElement;
  taskpaneDock: HTMLDivElement;
  grid: HTMLDivElement;
  canvas: HTMLCanvasElement;
  a11y: HTMLDivElement;
  a11yLive: HTMLDivElement;
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

  // The Mac formula bar names its controls with the desktop accessibility text.
  const barLabels = (current: typeof strings) => {
    const mac = host.dataset.fcPlatform === 'mac' ? current.mac.chrome : null;
    return {
      fx: mac?.fxButton ?? current.fxDialog?.fxButtonLabel ?? current.a11y.formulaBar,
      cancel: mac?.cancelFormulaEdit ?? current.a11y.cancelFormulaEdit,
      enter: mac?.enterFormula ?? current.a11y.enterFormula,
      expand: mac?.expandFormulaBar ?? current.a11y.expandFormulaBar,
    };
  };
  const fx = createHostButton({
    className: 'fc-host__formulabar-fx',
    text: 'ƒx',
    tabIndex: -1,
    ariaLabel: barLabels(strings).fx,
  });
  const setFxDialogAvailable = (available: boolean): void => {
    const strings = getStrings();
    const reason = available ? null : strings.fxDialog.fxButtonUnavailable;
    fx.style.cursor = available ? '' : 'default';
    projectDisabledState(fx, !available, reason, {
      datasetKey: 'disabledReason',
      titlePrefix: barLabels(strings).fx,
    });
  };
  setFxDialogAvailable(flags.fxDialog !== false);

  const fxCancel = createHostButton({
    className: 'fc-host__formulabar-action fc-host__formulabar-action--cancel',
    tabIndex: -1,
    ariaLabel: barLabels(strings).cancel,
    iconPaths: ['M6 6l8 8', 'M14 6l-8 8'],
  });
  projectDisabledState(fxCancel, true, strings.a11y.cancelFormulaEditUnavailable, {
    datasetKey: 'disabledReason',
    titlePrefix: barLabels(strings).cancel,
  });

  const fxAccept = createHostButton({
    className: 'fc-host__formulabar-action fc-host__formulabar-action--accept',
    tabIndex: -1,
    ariaLabel: barLabels(strings).enter,
    iconPaths: ['M5 10.5l3 3 7-7'],
  });
  projectDisabledState(fxAccept, true, strings.a11y.enterFormulaUnavailable, {
    datasetKey: 'disabledReason',
    titlePrefix: barLabels(strings).enter,
  });

  const fxInput = document.createElement('textarea');
  fxInput.className = 'fc-host__formulabar-input';
  fxInput.spellcheck = false;
  fxInput.autocomplete = 'off';
  fxInput.rows = 1;
  fxInput.wrap = 'soft';
  fxInput.setAttribute('aria-label', strings.a11y.formulaBar);

  const fxExpand = createHostButton({
    className: 'fc-host__formulabar-expand',
    ariaExpanded: false,
    tabIndex: -1,
    iconPaths: ['M5.5 7.5 10 12l4.5-4.5'],
  });

  let refreshGridLabels = (): void => {};
  const refreshFormulaBarLabels = (): void => {
    const strings = getStrings();
    const labels = barLabels(strings);
    fx.setAttribute('aria-label', labels.fx);
    setFxDialogAvailable(!fx.disabled);
    fxCancel.setAttribute('aria-label', labels.cancel);
    fxAccept.setAttribute('aria-label', labels.enter);
    fxInput.setAttribute('aria-label', strings.a11y.formulaBar);
    const expanded = fxExpand.getAttribute('aria-expanded') === 'true';
    fxExpand.setAttribute('aria-label', expanded ? strings.a11y.collapseFormulaBar : labels.expand);
    refreshGridLabels();
  };
  refreshFormulaBarLabels();

  fxExpand.addEventListener('click', () => {
    const expanded = formulabar.dataset.fcExpanded === '1';
    if (expanded) {
      delete formulabar.dataset.fcExpanded;
      fxExpand.setAttribute('aria-expanded', 'false');
      fxExpand.replaceChildren();
      appendHostIcon(fxExpand, ['M5.5 7.5 10 12l4.5-4.5']);
      fxInput.rows = 1;
    } else {
      formulabar.dataset.fcExpanded = '1';
      fxExpand.setAttribute('aria-expanded', 'true');
      fxExpand.replaceChildren();
      appendHostIcon(fxExpand, ['M5.5 12.5 10 8l4.5 4.5']);
      // 57 pt on the Mac desktop (three lines), four lines elsewhere.
      fxInput.rows = isMacHost(fxInput) ? 3 : 4;
    }
    refreshFormulaBarLabels();
  });
  formulabar.append(tag, fxCancel, fxAccept, fx, fxInput, fxExpand);

  const viewbar = document.createElement('div');
  viewbar.className = 'fc-viewbar';

  const localeSwitcher = document.createElement('div');
  localeSwitcher.className = 'fc-host__locale-switcher';

  const sheetWorkspace = document.createElement('div');
  sheetWorkspace.className = 'fc-host__sheet-workspace';
  const sheetSurface = document.createElement('div');
  sheetSurface.className = 'fc-host__sheet-surface';
  const taskpaneDock = document.createElement('div');
  taskpaneDock.className = 'fc-host__taskpane-dock';
  taskpaneDock.hidden = true;
  sheetWorkspace.append(sheetSurface, taskpaneDock);

  const grid = document.createElement('div');
  grid.className = 'fc-host__grid';
  grid.setAttribute('role', 'grid');
  grid.tabIndex = -1;
  const canvas = document.createElement('canvas');
  canvas.className = 'fc-host__canvas';
  canvas.setAttribute('aria-hidden', 'true');
  grid.appendChild(canvas);

  // Structured mirror of the painted cells: grid > rowgroup > row > gridcell.
  // chrome-sync fills in the rows.
  const a11y = document.createElement('div');
  a11y.className = 'fc-host__a11y';
  a11y.id = `${host.dataset.fcInstId ?? 'fc'}-a11y`;
  a11y.setAttribute('role', 'rowgroup');
  grid.appendChild(a11y);

  // The announcer is a separate node outside the grid: a live region is not a
  // valid child of role="grid", and the grid role must not sit on the region.
  const a11yLive = document.createElement('div');
  a11yLive.className = 'fc-host__a11y-live';
  a11yLive.id = `${a11y.id}-live`;
  a11yLive.setAttribute('aria-live', 'polite');
  a11yLive.setAttribute('aria-atomic', 'true');
  grid.setAttribute('aria-describedby', a11yLive.id);
  refreshGridLabels = (): void => {
    grid.setAttribute('aria-label', getStrings().a11y.grid);
  };
  refreshGridLabels();

  const statusbar = document.createElement('div');
  statusbar.className = 'fc-host__statusbar';
  statusbar.tabIndex = -1;
  const sheetbar = document.createElement('div');
  sheetbar.className = 'fc-host__sheetbar';
  const sheetNav = document.createElement('div');
  sheetNav.className = 'fc-host__sheetbar-nav';
  const firstSheet = createHostButton({
    className: 'fc-host__sheetbar-navbtn',
    iconPaths: ['M12.5 4.5 7 10l5.5 5.5'],
  });
  const lastSheet = createHostButton({
    className: 'fc-host__sheetbar-navbtn',
    iconPaths: ['M7.5 4.5 13 10l-5.5 5.5'],
  });
  sheetNav.append(firstSheet, lastSheet);

  const sheetTabs = document.createElement('div');
  sheetTabs.className = 'fc-host__sheetbar-tabs';
  sheetTabs.setAttribute('role', 'tablist');
  const addSheetBtn = createHostButton({
    className: 'fc-host__sheetbar-add',
    iconPaths: ['M10 4.5v11', 'M4.5 10h11'],
  });
  sheetbar.append(sheetNav, sheetTabs, addSheetBtn);

  const sheetMenu = document.createElement('div');
  sheetMenu.className = 'fc-sheetmenu';
  sheetMenu.hidden = true;
  sheetMenu.setAttribute('role', 'menu');
  overlayPortalFor(host).appendChild(sheetMenu);
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
          : slot === 'localeSwitcher'
            ? localeSwitcher
            : slot === 'sheetbar'
              ? sheetbar
              : slot === 'statusbar'
                ? statusbar
                : watchDock;
    if (on) {
      if (slot === 'formulabar') {
        const anchor =
          viewbar.parentElement === host
            ? viewbar
            : localeSwitcher.parentElement === host
              ? localeSwitcher
              : sheetWorkspace;
        host.insertBefore(el, anchor);
      } else if (slot === 'viewbar') {
        host.insertBefore(
          el,
          localeSwitcher.parentElement === host ? localeSwitcher : sheetWorkspace,
        );
      } else if (slot === 'localeSwitcher') {
        host.insertBefore(el, sheetWorkspace);
      } else if (slot === 'sheetbar') {
        sheetSurface.appendChild(el);
      } else if (slot === 'statusbar') {
        if (watchDock.parentElement === host) host.insertBefore(el, watchDock);
        else host.appendChild(el);
      } else {
        host.appendChild(el);
      }
    } else if (el.parentElement === host || el.parentElement === sheetSurface) {
      el.parentElement.removeChild(el);
    }
  };

  sheetSurface.appendChild(grid);
  host.appendChild(sheetWorkspace);
  host.appendChild(a11yLive);
  setChromeAttached('formulabar', flags.formulaBar === true);
  setChromeAttached('viewbar', flags.viewToolbar === true);
  setChromeAttached('localeSwitcher', flags.localeSwitcher === true);
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
    localeSwitcher,
    sheetWorkspace,
    sheetSurface,
    taskpaneDock,
    grid,
    canvas,
    a11y,
    a11yLive,
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
