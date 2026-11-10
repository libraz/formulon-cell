import type { SpreadsheetInstance } from '@libraz/formulon-cell';
import { sameProp } from '@libraz/formulon-cell/adapter';
import {
  mountToolbar,
  type RibbonHooks,
  type RibbonTab,
  type Toolbar,
} from '@libraz/formulon-cell/toolbar';
import {
  type ComponentOptionsMixin,
  type DefineComponent,
  defineComponent,
  h,
  onBeforeUnmount,
  onMounted,
  type PropType,
  ref,
  watch,
} from 'vue';
import { forwardRibbonHooks, ribbonHooksShape } from './ribbon-hooks.js';

/** Props of {@link SpreadsheetToolbar}. */
export interface SpreadsheetToolbarProps {
  /** The spreadsheet the ribbon drives; nothing is mounted while it is `null`. */
  spreadsheet: SpreadsheetInstance | null;
  /** The selected tab, bound with `v-model:activeTab`. */
  activeTab?: RibbonTab;
  /** Host actions for the commands the ribbon cannot run on its own. Only adding or removing a
   *  hook remounts the ribbon; a new function identity is picked up in place. */
  hooks?: RibbonHooks;
  /** The tab surface; defaults to the platform's standard tabs. */
  ribbonTabs?: readonly RibbonTab[];
}

/** Events of {@link SpreadsheetToolbar}. `ready` carries the mounted toolbar, and `null` once it
 *  is disposed; `error` carries a failed toolbar mount. */
export type SpreadsheetToolbarEmits = {
  'update:activeTab': (tab: RibbonTab) => true;
  ready: (toolbar: Toolbar | null) => true;
  error: (error: unknown) => true;
};

const component = defineComponent({
  name: 'SpreadsheetToolbar',
  props: {
    spreadsheet: {
      type: Object as PropType<SpreadsheetInstance | null>,
      default: null,
    },
    activeTab: { type: String as PropType<RibbonTab>, default: undefined },
    hooks: { type: Object as PropType<RibbonHooks>, default: undefined },
    ribbonTabs: { type: Array as PropType<readonly RibbonTab[]>, default: undefined },
  },
  emits: {
    'update:activeTab': (_tab: RibbonTab) => true,
    ready: (_toolbar: Toolbar | null) => true,
    error: (_error: unknown) => true,
  },
  setup(props, { emit }) {
    const hostEl = ref<HTMLDivElement | null>(null);
    let toolbar: Toolbar | null = null;
    let reportedTab: RibbonTab | undefined;
    let generation = 0;

    const release = (): void => {
      generation += 1;
      if (!toolbar) return;
      const mounted = toolbar;
      toolbar = null;
      mounted.dispose();
      emit('ready', null);
    };

    const mountFor = (spreadsheet: SpreadsheetInstance | null): void => {
      release();
      const host = hostEl.value;
      if (!host || !spreadsheet) return;
      const current = generation;
      const { activeTab, ribbonTabs } = props;
      mountToolbar(host, spreadsheet, {
        hooks: forwardRibbonHooks(() => props.hooks),
        ...(ribbonTabs ? { ribbonTabs } : {}),
        ...(activeTab ? { activeTab } : {}),
      }).then(
        (mounted) => {
          if (current !== generation) {
            mounted.dispose();
            return;
          }
          toolbar = mounted;
          reportedTab = mounted.activeTab;
          mounted.on('change', (state) => {
            if (state.activeTab === reportedTab) return;
            reportedTab = state.activeTab;
            emit('update:activeTab', state.activeTab);
          });
          emit('ready', mounted);
        },
        (error: unknown) => {
          if (current === generation) emit('error', error);
        },
      );
    };

    onMounted(() => mountFor(props.spreadsheet));

    // The ribbon remounts when the spreadsheet, the tab surface or the set of hooks changes.
    watch(
      () => [props.spreadsheet, props.ribbonTabs, ribbonHooksShape(props.hooks)] as const,
      ([spreadsheet, tabs, shape], [previous, previousTabs, previousShape]) => {
        if (spreadsheet === previous && sameProp(tabs, previousTabs) && shape === previousShape) {
          return;
        }
        mountFor(spreadsheet);
      },
      { flush: 'post' },
    );

    watch(
      () => props.activeTab,
      (next) => {
        if (!toolbar || next === undefined || toolbar.activeTab === next) return;
        reportedTab = next;
        toolbar.update({ activeTab: next });
      },
    );

    onBeforeUnmount(release);

    // `display: contents` keeps the wrapper out of the flex chain the ribbon shell sits in.
    return () => h('div', { ref: hostEl, style: { display: 'contents' } });
  },
});

/** The ribbon of `spreadsheet`, placed apart from the grid. */
export const SpreadsheetToolbar: DefineComponent<
  SpreadsheetToolbarProps,
  Record<string, never>,
  Record<string, never>,
  Record<string, never>,
  Record<string, never>,
  ComponentOptionsMixin,
  ComponentOptionsMixin,
  SpreadsheetToolbarEmits
> = component as never;
