import {
  EMPTY_ACTIVE_STATE,
  mutators,
  projectActiveState,
  RIBBON_TAB_LABELS,
  type RibbonTab,
  toggleBold,
} from '@libraz/formulon-cell';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createApp, defineComponent, h, nextTick, type Ref, shallowRef } from 'vue';
import { useToolbarActive } from '../src/toolbar/active';
import { useToolbarDropdown } from '../src/toolbar/dropdown';
import { toolbarTabs } from '../src/toolbar/tabs';
import {
  installVueDomStubs,
  type MountedVueSpreadsheet,
  mountVueSpreadsheet,
  uninstallVueDomStubs,
} from './test-utils/mount';

/**
 * The Vue ribbon ships as a Single File Component (`SpreadsheetToolbar.vue`)
 * which our happy-dom vitest config can't parse without a `@vitejs/plugin-vue`
 * dependency the project deliberately avoids. Instead we test the building
 * blocks the SFC composes — `toolbarTabs`, `useToolbarActive`,
 * `useToolbarDropdown` — and we mount a minimal `<RibbonProbe>` Vue component
 * so all of it is exercised through Vue's reactivity, not in isolation.
 */

const flush = async (): Promise<void> => {
  for (let i = 0; i < 8; i += 1) await Promise.resolve();
  await nextTick();
};

interface RibbonProbeHandle {
  host: HTMLElement;
  /** Active state ref returned by `useToolbarActive`. */
  active: Ref<ReturnType<typeof projectActiveState>>;
  /** Open dropdown name. */
  openDropdown: Ref<string | null>;
  keydownDropdown: (event: KeyboardEvent) => void;
  toggleDropdown: (name: string) => void;
  pickDropdown: (name: string, value: string | number) => void;
  unmount: () => Promise<void>;
}

interface DropdownLogEntry {
  kind: string;
  value: unknown;
}

async function mountRibbonProbe(
  instance: MountedVueSpreadsheet['instance'],
  log: DropdownLogEntry[],
): Promise<RibbonProbeHandle> {
  installVueDomStubs();
  const host = document.createElement('div');
  document.body.appendChild(host);

  const instanceRef = shallowRef(instance);
  const activeRef = shallowRef<ReturnType<typeof projectActiveState>>(EMPTY_ACTIVE_STATE);
  const openRef = shallowRef<string | null>(null);
  let keydown: ((event: KeyboardEvent) => void) | null = null;
  let toggle: ((name: string) => void) | null = null;
  let pick: ((name: string, value: string | number) => void) | null = null;

  const Probe = defineComponent({
    setup() {
      const active = useToolbarActive(() => instanceRef.value);
      activeRef.value = active.value;

      const dd = useToolbarDropdown({
        onBorderPreset: (v) => log.push({ kind: 'borderPreset', value: v }),
        onFontFamily: (v) => log.push({ kind: 'fontFamily', value: v }),
        onFontSize: (v) => log.push({ kind: 'fontSize', value: v }),
        onMarginPreset: (v) => log.push({ kind: 'marginPreset', value: v }),
        onOpenPageSetup: () => log.push({ kind: 'openPageSetup', value: null }),
        onPageOrientation: (v) => log.push({ kind: 'pageOrientation', value: v }),
        onPaperSize: (v) => log.push({ kind: 'paperSize', value: v }),
      });

      keydown = dd.onDropdownKeydown;
      toggle = (name) => dd.toggleDropdown(name as never);
      pick = (name, value) => dd.onDropdownPick(name as never, value);

      return () => {
        // Mirror `active` and `openDropdown` into shallow refs the test reads.
        activeRef.value = active.value;
        openRef.value = dd.openDropdown.value;
        return h('div', { class: 'probe' }, [
          h('span', { 'data-testid': 'bold' }, String(active.value.bold)),
          h('span', { 'data-testid': 'fontSize' }, String(active.value.fontSize)),
          h('span', { 'data-testid': 'open' }, String(dd.openDropdown.value ?? '')),
        ]);
      };
    },
  });

  const app = createApp(Probe);
  app.mount(host);
  await flush();

  return {
    host,
    active: activeRef,
    openDropdown: openRef,
    keydownDropdown: (event) => {
      if (!keydown) throw new Error('keydown not yet bound');
      keydown(event);
    },
    toggleDropdown: (name) => {
      if (!toggle) throw new Error('toggle not yet bound');
      toggle(name);
    },
    pickDropdown: (name, value) => {
      if (!pick) throw new Error('pick not yet bound');
      pick(name, value);
    },
    async unmount() {
      app.unmount();
      await flush();
      host.remove();
      uninstallVueDomStubs();
    },
  };
}

describe('Vue toolbar — toolbarTabs builder', () => {
  it('returns one entry per non-file ribbon tab using RIBBON_TAB_LABELS', () => {
    const en = toolbarTabs('en');
    const expectedIds = (Object.keys(RIBBON_TAB_LABELS) as RibbonTab[]).filter(
      (id) => id !== 'file',
    );
    expect(en.map((t) => t.id)).toEqual(expectedIds);
    expect(en.map((t) => t.label)).toEqual(expectedIds.map((id) => RIBBON_TAB_LABELS[id].en));

    const ja = toolbarTabs('ja');
    expect(ja.map((t) => t.label)).toEqual(expectedIds.map((id) => RIBBON_TAB_LABELS[id].ja));
  });
});

describe('Vue <SpreadsheetToolbar> building blocks', () => {
  let mounted: MountedVueSpreadsheet | null = null;
  let probe: RibbonProbeHandle | null = null;

  afterEach(async () => {
    if (probe) {
      await probe.unmount();
      probe = null;
    }
    if (mounted) {
      await mounted.dispose();
      mounted = null;
    }
    document.body.replaceChildren();
  });

  it('useToolbarActive seeds from a live instance and reflects bold toggles', async () => {
    mounted = await mountVueSpreadsheet();
    const log: DropdownLogEntry[] = [];
    probe = await mountRibbonProbe(mounted.instance, log);

    expect(probe.active.value.bold).toBe(false);
    expect(probe.host.querySelector('[data-testid="bold"]')?.textContent).toBe('false');

    // Toggle bold via the same command the toolbar would invoke.
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 0, col: 0 });
    toggleBold(mounted.instance.store.getState(), mounted.instance.store);
    await flush();

    expect(projectActiveState(mounted.instance).bold).toBe(true);
    expect(probe.active.value.bold).toBe(true);
    expect(probe.host.querySelector('[data-testid="bold"]')?.textContent).toBe('true');
  });

  it('useToolbarDropdown opens, picks, and routes the right handler per dropdown name', async () => {
    mounted = await mountVueSpreadsheet();
    const log: DropdownLogEntry[] = [];
    probe = await mountRibbonProbe(mounted.instance, log);

    // Initially closed.
    expect(probe.openDropdown.value).toBeNull();

    probe.toggleDropdown('fontFamily');
    await flush();
    expect(probe.openDropdown.value).toBe('fontFamily');

    probe.pickDropdown('fontFamily', 'Calibri');
    await flush();
    expect(log).toContainEqual({ kind: 'fontFamily', value: 'Calibri' });
    // Picking auto-closes.
    expect(probe.openDropdown.value).toBeNull();

    // borderStyle is purely local — no handler should fire, but the
    // dropdown must still close.
    probe.toggleDropdown('borderStyle');
    await flush();
    expect(probe.openDropdown.value).toBe('borderStyle');
    probe.pickDropdown('borderStyle', 'thick');
    await flush();
    expect(log.find((e) => e.kind === 'borderStyle')).toBeUndefined();
    expect(probe.openDropdown.value).toBeNull();

    // margins=custom is a special-case redirect to onOpenPageSetup.
    probe.toggleDropdown('margins');
    probe.pickDropdown('margins', 'custom');
    await flush();
    expect(log).toContainEqual({ kind: 'openPageSetup', value: null });
    expect(log.find((e) => e.kind === 'marginPreset')).toBeUndefined();

    // margins=normal routes to onMarginPreset.
    probe.toggleDropdown('margins');
    probe.pickDropdown('margins', 'normal');
    await flush();
    expect(log).toContainEqual({ kind: 'marginPreset', value: 'normal' });
  });

  it('useToolbarDropdown closes the open dropdown on Escape and on outside click', async () => {
    mounted = await mountVueSpreadsheet();
    const log: DropdownLogEntry[] = [];
    probe = await mountRibbonProbe(mounted.instance, log);

    probe.toggleDropdown('fontSize');
    await flush();
    expect(probe.openDropdown.value).toBe('fontSize');

    // Escape from anywhere closes the dropdown.
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await flush();
    expect(probe.openDropdown.value).toBeNull();

    // Re-open and verify outside pointerdown closes too.
    probe.toggleDropdown('paperSize');
    await flush();
    expect(probe.openDropdown.value).toBe('paperSize');

    const outside = document.createElement('div');
    document.body.appendChild(outside);
    outside.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
    await flush();
    expect(probe.openDropdown.value).toBeNull();
    outside.remove();
  });

  it('useToolbarDropdown handles list keyboard navigation, pick, and focus return', async () => {
    mounted = await mountVueSpreadsheet();
    const log: DropdownLogEntry[] = [];
    probe = await mountRibbonProbe(mounted.instance, log);

    const root = document.createElement('div');
    root.className = 'demo__rb-dd';
    root.dataset.dropdownName = 'fontFamily';
    root.innerHTML = `
      <button class="demo__rb-dd__btn" type="button" aria-expanded="false">Font</button>
      <div class="demo__rb-dd__list" role="listbox">
        <button class="demo__rb-dd__opt" type="button" role="option" aria-selected="true">Aptos</button>
        <button class="demo__rb-dd__opt" type="button" role="option" aria-selected="false">Calibri</button>
        <button class="demo__rb-dd__opt" type="button" role="option" aria-selected="false">Consolas</button>
      </div>
    `;
    document.body.appendChild(root);
    root.addEventListener('keydown', probe.keydownDropdown);
    root.classList.add('demo__rb-dd--open');

    const button = root.querySelector<HTMLButtonElement>('.demo__rb-dd__btn');
    const options = root.querySelectorAll<HTMLButtonElement>('[role="option"]');
    expect(button).toBeTruthy();

    probe.toggleDropdown('fontFamily');
    await flush();
    options[0]?.focus();
    options[0]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
    await flush();
    expect(document.activeElement).toBe(options[1]);

    options[1]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'End', bubbles: true }));
    await flush();
    expect(document.activeElement).toBe(options[2]);

    options[2]?.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await flush();
    expect(probe.openDropdown.value).toBeNull();
    expect(document.activeElement).toBe(button);
    root.removeEventListener('keydown', probe.keydownDropdown);
    root.remove();
  });

  it('useToolbarActive cleans up its store subscription on unmount', async () => {
    mounted = await mountVueSpreadsheet();
    const log: DropdownLogEntry[] = [];
    probe = await mountRibbonProbe(mounted.instance, log);

    const beforeBold = probe.active.value.bold;
    expect(beforeBold).toBe(false);

    await probe.unmount();
    probe = null;

    // After unmount, store changes must not panic the (gone) subscriber.
    const errSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
    mutators.setActive(mounted.instance.store, { sheet: 0, row: 0, col: 0 });
    toggleBold(mounted.instance.store.getState(), mounted.instance.store);
    await flush();
    expect(errSpy).not.toHaveBeenCalled();
    errSpy.mockRestore();
  });

  it('useToolbarDropdown stops listening on document events after unmount', async () => {
    mounted = await mountVueSpreadsheet();
    const log: DropdownLogEntry[] = [];
    probe = await mountRibbonProbe(mounted.instance, log);

    probe.toggleDropdown('fontFamily');
    await flush();
    expect(probe.openDropdown.value).toBe('fontFamily');

    await probe.unmount();
    probe = null;

    // Escape after unmount should NOT mutate any captured state — and must
    // not throw. We exercise the path simply by dispatching the event; if
    // the listener wasn't removed we'd see a console error from Vue trying
    // to update an unmounted ref.
    const errSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await flush();
    expect(errSpy).not.toHaveBeenCalled();
    errSpy.mockRestore();
  });
});
