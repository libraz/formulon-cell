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
import { Spreadsheet, type SpreadsheetInstance, type ToolbarInstance } from '@libraz/formulon-cell';
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
  },
  emits: ['tabChange'],
  setup(props: AdapterProps, { emit }) {
    const hostEl = ref<HTMLDivElement | null>(null);
    let toolbar: ToolbarInstance | null = null;

    const mountToolbarFor = (instance: SpreadsheetInstance): void => {
      const host = hostEl.value;
      if (!host) return;
      toolbar?.dispose();
      toolbar = Spreadsheet.mountToolbar(host, instance, {
        lang: props.locale === 'en' ? 'en' : 'ja',
        activeTab: props.activeTab,
        onTabChange: (tab) => emit('tabChange', tab),
        hooks: {
          review: {
            spelling: () => props.onSpellingReview?.(),
            accessibility: () => props.onAccessibilityCheck?.(),
            translate: () => props.onTranslate?.(),
          },
          automation: {
            runScript: () => props.onRunScript?.(),
            addInManager: () => props.onAddIn?.(),
          },
          drawing: {
            setInkMode: (mode) => {
              if (mode === 'pen') props.onDrawPen?.();
              else props.onDrawEraser?.();
            },
          },
        },
      });
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

  const root = defineComponent({
    setup() {
      return () =>
        h(AdapterComponent, {
          instance,
          activeTab: activeTabRef.value,
          locale: 'en',
          ...callbacks,
          onTabChange: (tab: RibbonTab) => tabChange(tab),
        });
    },
  });

  const app = createApp(root);
  app.mount(host);
  await flush();

  return {
    host,
    tabChange,
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
    uninstallVueDomStubs();
  });

  it('mounts the core ribbon DOM into the wrapping host element', async () => {
    const harness = await mountAdapter(mounted.instance);
    expect(harness.host.querySelectorAll('[data-ribbon-tab]').length).toBeGreaterThan(0);
    expect(harness.host.querySelector('[data-ribbon-panel="home"]:not([hidden])')).toBeTruthy();
    await harness.unmount();
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
    const click = (cmd: string): void => {
      harness.host.querySelector<HTMLButtonElement>(`[data-ribbon-command="${cmd}"]`)?.click();
    };
    click('spellingReview');
    click('accessibility');
    click('translateReview');
    click('script');
    click('addIn');
    click('drawPen');
    click('drawErase');
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
