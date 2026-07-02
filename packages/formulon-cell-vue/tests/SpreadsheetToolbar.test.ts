// `SpreadsheetToolbar.vue` is now a ~100 LOC adapter on top of
// `Spreadsheet.mountToolbar`. The prior 3754 LOC of Vue-internal ribbon and
// the matching 1260 LOC test suite were retired in Phase 3-b.
//
// happy-dom + vitest can't compile `.vue` SFCs without `@vitejs/plugin-vue`,
// which the project deliberately avoids. We mirror the SFC's `<script setup>`
// inline as a `defineComponent` so the same `mountToolbar` wiring is
// exercised through Vue's reactivity. The smoke tests parallel the React
// adapter's coverage: DOM mount, tab forwarding both ways, ribbon-command
// hooks, and dispose on unmount.
import {
  type DynamicDropdownsCtx,
  Spreadsheet,
  type SpreadsheetInstance,
  type ToolbarInstance,
} from '@libraz/formulon-cell';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import {
  createApp,
  defineComponent,
  h,
  nextTick,
  onBeforeUnmount,
  onMounted,
  ref,
  watch,
} from 'vue';
import type { RibbonTab } from '../src/toolbar.js';
import {
  installVueDomStubs,
  type MountedVueSpreadsheet,
  mountVueSpreadsheet,
  uninstallVueDomStubs,
} from './test-utils/mount';

interface AdapterProps {
  instance: SpreadsheetInstance | null;
  activeTab: RibbonTab;
  locale: string;
  onSpellingReview?: () => void;
  onAccessibilityCheck?: () => void;
  onRunScript?: () => void;
  onDrawPen?: () => void;
  onDrawEraser?: () => void;
  onTranslate?: () => void;
  onAddIn?: () => void;
  onError?: (error: unknown) => void;
  onToolbarReady?: (toolbar: ToolbarInstance | null) => void;
  dropdownActions?: Partial<DynamicDropdownsCtx>;
}

// Verbatim copy of the SFC's `<script setup>` body, expressed as a
// `defineComponent` so happy-dom can mount it without the Vue compiler.
const AdapterComponent = defineComponent({
  name: 'SpreadsheetToolbarAdapter',
  props: {
    instance: { type: Object as () => SpreadsheetInstance | null, required: true },
    activeTab: { type: String as () => RibbonTab, required: true },
    locale: { type: String, required: true },
    onSpellingReview: { type: Function as unknown as () => () => void, default: undefined },
    onAccessibilityCheck: { type: Function as unknown as () => () => void, default: undefined },
    onRunScript: { type: Function as unknown as () => () => void, default: undefined },
    onDrawPen: { type: Function as unknown as () => () => void, default: undefined },
    onDrawEraser: { type: Function as unknown as () => () => void, default: undefined },
    onTranslate: { type: Function as unknown as () => () => void, default: undefined },
    onAddIn: { type: Function as unknown as () => () => void, default: undefined },
    onError: {
      type: Function as unknown as () => (error: unknown) => void,
      default: undefined,
    },
    onToolbarReady: {
      type: Function as unknown as () => (toolbar: ToolbarInstance | null) => void,
      default: undefined,
    },
    dropdownActions: {
      type: Object as () => Partial<DynamicDropdownsCtx> | undefined,
      default: undefined,
    },
  },
  emits: ['tabChange', 'error'],
  setup(props: AdapterProps, { emit }) {
    const hostEl = ref<HTMLDivElement | null>(null);
    let toolbar: ToolbarInstance | null = null;

    const mountToolbarFor = (instance: SpreadsheetInstance): void => {
      const host = hostEl.value;
      if (!host) return;
      if (toolbar) props.onToolbarReady?.(null);
      toolbar?.dispose();
      const dropdownOverrides: Partial<DynamicDropdownsCtx> = {
        applyScriptAction: (action) => {
          if (action === 'custom') props.onRunScript?.();
        },
        applyAddInAction: (action) => {
          if (action === 'manage') props.onAddIn?.();
        },
        ...props.dropdownActions,
      };
      try {
        toolbar = Spreadsheet.mountToolbar(host, instance, {
          lang: props.locale === 'en' ? 'en' : 'ja',
          activeTab: props.activeTab,
          onTabChange: (tab) => emit('tabChange', tab),
          dynamicDropdowns: dropdownOverrides,
          hooks: {
            review: {
              spelling: () => props.onSpellingReview?.(),
              accessibility: () => props.onAccessibilityCheck?.(),
              translate: () => props.onTranslate?.(),
            },
            drawing: {
              setInkMode: (mode) => {
                if (mode === 'pen') props.onDrawPen?.();
                else props.onDrawEraser?.();
              },
            },
          },
        });
      } catch (error) {
        props.onToolbarReady?.(null);
        props.onError?.(error);
        emit('error', error);
        return;
      }
      props.onToolbarReady?.(toolbar);
    };

    onMounted(() => {
      if (props.instance) mountToolbarFor(props.instance);
    });

    watch(
      () => [props.instance, props.locale] as const,
      ([instance], _prev, onCleanup) => {
        if (!hostEl.value) return;
        if (instance) mountToolbarFor(instance);
        else {
          props.onToolbarReady?.(null);
          toolbar?.dispose();
          toolbar = null;
        }
        onCleanup(() => {});
      },
      { flush: 'post' },
    );

    watch(
      () => props.activeTab,
      (next) => {
        if (toolbar && toolbar.getActiveTab() !== next) toolbar.setActiveTab(next);
      },
    );

    onBeforeUnmount(() => {
      props.onToolbarReady?.(null);
      toolbar?.dispose();
      toolbar = null;
    });

    return () => h('div', { ref: hostEl, style: { display: 'contents' } });
  },
});

const flush = async (): Promise<void> => {
  for (let i = 0; i < 8; i += 1) await Promise.resolve();
  await nextTick();
};

interface Harness {
  host: HTMLElement;
  tabChange: ReturnType<typeof vi.fn>;
  error: ReturnType<typeof vi.fn>;
  setActiveTab: (tab: RibbonTab) => Promise<void>;
  unmount: () => Promise<void>;
}

async function mountAdapter(
  instance: SpreadsheetInstance,
  callbacks: Omit<AdapterProps, 'instance' | 'activeTab' | 'locale'> = {},
  initialTab: RibbonTab = 'home',
): Promise<Harness> {
  const host = document.createElement('div');
  document.body.appendChild(host);
  const activeTabRef = ref<RibbonTab>(initialTab);
  const tabChange = vi.fn();
  const error = vi.fn();

  const root = defineComponent({
    setup() {
      return () =>
        h(AdapterComponent, {
          instance,
          activeTab: activeTabRef.value,
          locale: 'en',
          ...callbacks,
          onTabChange: (tab: RibbonTab) => tabChange(tab),
          onError: (err: unknown) => {
            callbacks.onError?.(err);
            error(err);
          },
        });
    },
  });

  const app = createApp(root);
  app.mount(host);
  await flush();

  return {
    host,
    tabChange,
    error,
    setActiveTab: async (tab) => {
      activeTabRef.value = tab;
      await flush();
    },
    unmount: async () => {
      app.unmount();
      await flush();
      host.remove();
    },
  };
}

describe('<SpreadsheetToolbar> Vue adapter (thin)', () => {
  let mounted: MountedVueSpreadsheet;

  beforeEach(async () => {
    installVueDomStubs();
    mounted = await mountVueSpreadsheet({ props: { locale: 'en' } });
  });

  afterEach(async () => {
    await mounted.dispose();
    document.body.querySelectorAll('.app__dlg').forEach((el) => {
      el.remove();
    });
    uninstallVueDomStubs();
  });

  it('mounts the core ribbon DOM into the wrapping host element', async () => {
    const onToolbarReady = vi.fn();
    const harness = await mountAdapter(mounted.instance, { onToolbarReady });
    expect(harness.host.querySelectorAll('[data-ribbon-tab]').length).toBeGreaterThan(0);
    expect(harness.host.querySelector('[data-ribbon-panel="home"]:not([hidden])')).toBeTruthy();
    expect(onToolbarReady).toHaveBeenCalledWith(
      expect.objectContaining({ applyCommand: expect.any(Function) }),
    );
    await harness.unmount();
    expect(onToolbarReady).toHaveBeenLastCalledWith(null);
  });

  it('forwards core toolbar mount failures to onError and error emit', async () => {
    const err = new Error('toolbar mount failed');
    const onError = vi.fn();
    const onToolbarReady = vi.fn();
    const spy = vi.spyOn(Spreadsheet, 'mountToolbar').mockImplementationOnce(() => {
      throw err;
    });
    const harness = await mountAdapter(mounted.instance, { onError, onToolbarReady });

    expect(onToolbarReady).toHaveBeenCalledWith(null);
    expect(onError).toHaveBeenCalledWith(err);
    expect(harness.error).toHaveBeenCalledWith(err);
    expect(harness.host.querySelector('[data-ribbon-tab]')).toBeNull();
    await harness.unmount();
    spy.mockRestore();
  });

  it('forwards tab-button clicks via tabChange emit', async () => {
    const harness = await mountAdapter(mounted.instance);
    const insertTab = harness.host.querySelector<HTMLButtonElement>('[data-ribbon-tab="insert"]');
    expect(insertTab).toBeTruthy();
    insertTab?.click();
    await flush();
    expect(harness.tabChange).toHaveBeenCalledWith('insert');
    await harness.unmount();
  });

  it('routes ribbon review/automation/drawing commands to the matching host callback', async () => {
    const onSpellingReview = vi.fn();
    const onAccessibilityCheck = vi.fn();
    const onTranslate = vi.fn();
    const onRunScript = vi.fn();
    const onAddIn = vi.fn();
    const onDrawPen = vi.fn();
    const onDrawEraser = vi.fn();
    const harness = await mountAdapter(
      mounted.instance,
      {
        onSpellingReview,
        onAccessibilityCheck,
        onTranslate,
        onRunScript,
        onAddIn,
        onDrawPen,
        onDrawEraser,
      },
      'review',
    );
    const clickCommand = (cmd: string): void => {
      harness.host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${cmd}"]`)?.click();
    };
    const clickAttr = (attr: string, value: string): void => {
      harness.host.querySelector<HTMLButtonElement>(`[data-${attr}="${value}"]`)?.click();
    };
    clickCommand('spellingReview');
    clickCommand('accessibility');
    clickCommand('translateReview');
    // Script / AddIn ribbon buttons open a menu on plain click; the host
    // callback fires only when the user picks the action wired to its prop.
    clickCommand('script');
    clickAttr('script-action', 'custom');
    clickCommand('addIn');
    clickAttr('add-in-action', 'manage');
    clickCommand('drawPen');
    clickCommand('drawErase');
    await flush();
    expect(onSpellingReview).toHaveBeenCalledTimes(1);
    expect(onAccessibilityCheck).toHaveBeenCalledTimes(1);
    expect(onTranslate).toHaveBeenCalledTimes(1);
    expect(onRunScript).toHaveBeenCalledTimes(1);
    expect(onAddIn).toHaveBeenCalledTimes(1);
    expect(onDrawPen).toHaveBeenCalledTimes(1);
    expect(onDrawEraser).toHaveBeenCalledTimes(1);
    await harness.unmount();
  });

  it('routes dropdownActions overrides through core dynamic-dropdowns dispatcher', async () => {
    const onProtect = vi.fn();
    const harness = await mountAdapter(
      mounted.instance,
      { dropdownActions: { applyProtectAction: onProtect } },
      'review',
    );
    harness.host.querySelector<HTMLButtonElement>('[data-protect-action="lock-cell"]')?.click();
    await flush();
    expect(onProtect).toHaveBeenCalledWith('lock-cell');
    await harness.unmount();
  });

  it('preserves core Insert activation parity for PivotTable, Table, and Pictures', async () => {
    mounted.instance.setFeatures({ pivotTableDialog: true, illustrations: true });
    await flush();
    const openPivotTableDialog = vi.spyOn(mounted.instance, 'openPivotTableDialog');
    const onToolbarReady = vi.fn();
    const harness = await mountAdapter(mounted.instance, { onToolbarReady }, 'insert');
    const toolbar = onToolbarReady.mock.calls
      .map((call) => call[0])
      .find((candidate) => candidate?.applyCommand);
    expect(toolbar).toBeTruthy();
    const pivotButton = harness.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pivotTableInsert"]',
    );
    const tableButton = harness.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="formatTableInsert"]',
    );
    const pictureButton = harness.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="pictureInsert"]',
    );
    expect(pivotButton?.dataset.ribbonActivation).toBe('splitPrimary');
    expect(tableButton?.dataset.ribbonActivation).toBe('dialog');
    expect(tableButton?.dataset.ribbonMenuId).toBeUndefined();
    expect(pictureButton?.dataset.ribbonActivation).toBe('gallery');

    expect(toolbar?.applyCommand('pivotTableInsert')).toBe(true);
    await flush();
    expect(openPivotTableDialog).toHaveBeenCalledTimes(1);

    expect(toolbar?.applyCommand('formatTableInsert')).toBe(true);
    await flush();
    expect(document.body.querySelector<HTMLElement>('.app__dlg')?.textContent).toContain(
      'Create Table',
    );
    document.body.querySelector<HTMLButtonElement>('.app__dlg .fc-fmtdlg__btn')?.click();
    await flush();

    pictureButton?.click();
    await flush();
    const pictureMenu = harness.host.querySelector<HTMLElement>('#menu-picture-insert');
    expect(pictureMenu?.hidden).toBe(false);
    expect(pictureMenu?.classList.contains('app__menu--visual')).toBe(true);
    expect(
      pictureMenu?.querySelector<HTMLButtonElement>('[data-picture-insert="stock"]'),
    ).toBeTruthy();

    await harness.unmount();
    openPivotTableDialog.mockRestore();
  });

  it('preserves core Home split/dropdown parity for Underline and Fill', async () => {
    const onToolbarReady = vi.fn();
    const harness = await mountAdapter(mounted.instance, { onToolbarReady });
    const toolbar = onToolbarReady.mock.calls
      .map((call) => call[0])
      .find((candidate): candidate is ToolbarInstance => !!candidate?.dropdownsApi);
    expect(toolbar).toBeTruthy();

    const underline = harness.host.querySelector<HTMLButtonElement>(
      '[data-ribbon-command="underline"]',
    );
    const underlineMenu = harness.host.querySelector<HTMLDivElement>('#menu-underline');
    expect(underline?.dataset.ribbonActivation).toBe('splitToggle');
    expect(underline?.dataset.ribbonMenuId).toBe('menu-underline');
    expect(underline?.getAttribute('aria-pressed')).toBe('false');
    underline?.click();
    await flush();
    expect(underline?.getAttribute('aria-pressed')).toBe('true');
    expect(underlineMenu?.hidden).toBe(true);
    toolbar.dropdownsApi?.openDynamicRibbonDropdown(
      { command: 'underline', menuId: 'menu-underline' },
      underline as HTMLButtonElement,
    );
    await flush();
    expect(underlineMenu?.hidden).toBe(false);
    expect(
      Array.from(
        harness.host.querySelectorAll<HTMLButtonElement>('#menu-underline .app__menu-item--iconic'),
      ).map((item) => item.textContent),
    ).toEqual(['Underline', 'Double Underline']);

    const fill = harness.host.querySelector<HTMLButtonElement>('[data-ribbon-command="fillHome"]');
    fill?.click();
    await flush();
    expect(
      Array.from(harness.host.querySelectorAll<HTMLButtonElement>('#menu-fill [data-fill]')).map(
        (item) => item.dataset.fill,
      ),
    ).toEqual(['down', 'right', 'up', 'left', 'series', 'days', 'weekdays', 'months', 'years']);

    await harness.unmount();
  });

  it('reacts to external activeTab prop changes without re-mounting the core toolbar', async () => {
    const harness = await mountAdapter(mounted.instance);
    expect(harness.host.querySelector('[data-ribbon-panel="home"]:not([hidden])')).toBeTruthy();
    await harness.setActiveTab('data');
    expect(harness.host.querySelector('[data-ribbon-panel="data"]:not([hidden])')).toBeTruthy();
    expect(harness.host.querySelector('[data-ribbon-panel="home"]:not([hidden])')).toBeFalsy();
    await harness.unmount();
  });

  it('cleans up the core toolbar when Vue unmounts the component', async () => {
    const harness = await mountAdapter(mounted.instance);
    expect(harness.host.querySelectorAll('[data-ribbon-tab]').length).toBeGreaterThan(0);
    await harness.unmount();
    expect(document.body.contains(harness.host)).toBe(false);
  });
});
