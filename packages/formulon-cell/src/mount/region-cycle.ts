/**
 * F6 / Shift+F6 cycling of keyboard focus between the regions of a mounted sheet: ribbon, formula
 * bar, grid, sheet bar, status bar, then each open docked pane. A region that is absent or hidden
 * takes no stop. When the grid is the only region with a stop the key is left to the host page.
 */
import { focusableDialogItems, focusSheet } from '../interact/dialog-focus.js';
import { modalDialogOwnsKeys } from '../interact/dialog-shell.js';

export interface RegionCycleInput {
  host: HTMLElement;
  /** The ribbon's tab focus; absent while no ribbon is mounted. */
  focusRibbonTab: () => boolean;
}

/** A ribbon mounted outside the sheet's host element that takes part in the host's ring. */
interface ExternalRibbon {
  root: HTMLElement;
  /** Moves focus to the ribbon's active tab; false when it cannot take focus. */
  focusTab: () => boolean;
}

interface RingState {
  ribbons: Set<ExternalRibbon>;
  /** The ring's key handler once attached; external ribbons listen with it too. */
  onKeydown?: (event: KeyboardEvent) => void;
}

const rings = new WeakMap<HTMLElement, RingState>();

const ringOf = (host: HTMLElement): RingState => {
  let state = rings.get(host);
  if (!state) {
    state = { ribbons: new Set() };
    rings.set(host, state);
  }
  return state;
};

/**
 * Joins a ribbon mounted outside `host` to the host's ring as its first stop. Returns the function
 * that removes it. A ribbon inside `host` is already found by the ring and is not registered.
 */
export function registerExternalRibbon(host: HTMLElement, ribbon: ExternalRibbon): () => void {
  if (host.contains(ribbon.root)) return () => undefined;
  const state = ringOf(host);
  state.ribbons.add(ribbon);
  if (state.onKeydown) ribbon.root.addEventListener('keydown', state.onKeydown);
  return () => {
    state.ribbons.delete(ribbon);
    if (state.onKeydown) ribbon.root.removeEventListener('keydown', state.onKeydown);
  };
}

interface Region {
  /** The element that makes a focus inside it belong to this region. */
  root: HTMLElement;
  /** Moves focus to the region's entry stop; false when it cannot take focus. */
  enter: () => boolean;
  isGrid?: boolean;
}

const isShown = (el: HTMLElement, host: HTMLElement): boolean => {
  for (let node: HTMLElement | null = el; node && node !== host; node = node.parentElement) {
    if (node.hidden || node.getAttribute('aria-hidden') === 'true') return false;
    if (node.style.display === 'none') return false;
  }
  return true;
};

const focusElement = (el: HTMLElement | null): boolean => {
  if (!el) return false;
  el.focus({ preventScroll: true });
  return el.ownerDocument.activeElement === el;
};

const regionsOf = (host: HTMLElement, focusRibbonTab: () => boolean): Region[] => {
  const external = rings.get(host)?.ribbons ?? [];
  const find = (selector: string): HTMLElement | null => host.querySelector<HTMLElement>(selector);
  const regions: Region[] = [];
  const add = (root: HTMLElement | null, enter: (root: HTMLElement) => boolean, isGrid = false) => {
    if (root && isShown(root, host)) regions.push({ root, enter: () => enter(root), isGrid });
  };
  for (const ribbon of external) add(ribbon.root, ribbon.focusTab);
  add(find('.fc-host__ribbon'), focusRibbonTab);
  add(find('.fc-host__formulabar'), () => {
    const nameBox = find('.fc-host__formulabar-tag') as HTMLInputElement | null;
    if (!nameBox || !isShown(nameBox, host) || !focusElement(nameBox)) return false;
    nameBox.select();
    return true;
  });
  add(
    find('.fc-host__grid'),
    () => {
      focusSheet(host, { preventScroll: true });
      return host.ownerDocument.activeElement === find('.fc-host__grid');
    },
    true,
  );
  add(find('.fc-host__sheetbar'), (bar) =>
    focusElement(
      bar.querySelector<HTMLElement>('.fc-host__sheetbar-tabs [role="tab"][aria-selected="true"]'),
    ),
  );
  add(find('.fc-host__statusbar'), focusElement);
  for (const pane of host.querySelectorAll<HTMLElement>('[data-fc-surface="pane"]')) {
    add(pane, (root) => {
      const first = focusableDialogItems(root)[0];
      if (first) return focusElement(first);
      if (!root.hasAttribute('tabindex')) root.tabIndex = -1;
      return focusElement(root);
    });
  }
  return regions;
};

/** Registers the F6 ring on `host`; returns the function that removes it. */
export function attachRegionCycle(input: RegionCycleInput): () => void {
  const { host } = input;
  const onKeydown = (event: KeyboardEvent): void => {
    if (event.key !== 'F6' || event.ctrlKey || event.metaKey || event.altKey) return;
    if (modalDialogOwnsKeys(event)) return;
    const regions = regionsOf(host, input.focusRibbonTab);
    const active = host.ownerDocument.activeElement;
    let start = regions.findIndex((r) => active instanceof Node && r.root.contains(active));
    if (start < 0) start = regions.findIndex((r) => r.isGrid);
    if (start < 0) return;
    const direction = event.shiftKey ? -1 : 1;
    for (let step = 1; step < regions.length; step += 1) {
      const next = regions[(start + direction * step + regions.length) % regions.length];
      if (next?.enter()) {
        event.preventDefault();
        return;
      }
    }
  };
  const state = ringOf(host);
  state.onKeydown = onKeydown;
  host.addEventListener('keydown', onKeydown);
  for (const ribbon of state.ribbons) ribbon.root.addEventListener('keydown', onKeydown);
  return () => {
    host.removeEventListener('keydown', onKeydown);
    for (const ribbon of state.ribbons) ribbon.root.removeEventListener('keydown', onKeydown);
    state.onKeydown = undefined;
  };
}
