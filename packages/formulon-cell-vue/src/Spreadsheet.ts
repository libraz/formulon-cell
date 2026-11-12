import {
  type CellEvent,
  type ChangeEvent,
  type I18n,
  mount,
  type RecalcEvent,
  type SelectionState,
  type SpreadsheetInstance,
  type SpreadsheetOptions,
  type SpreadsheetPatch,
  type UpdateEvent,
  type WorkbookEvent,
} from '@libraz/formulon-cell';
import { optionsFromProps, patchFromOptions, reconcileProps } from '@libraz/formulon-cell/adapter';
import {
  type ComponentOptionsMixin,
  type CSSProperties,
  type DefineComponent,
  defineComponent,
  h,
  onBeforeUnmount,
  onMounted,
  type PropType,
  ref,
  shallowRef,
  type VNodeChild,
  watch,
} from 'vue';

type LocaleChangeEvent = Parameters<Parameters<I18n['on']>[1]>[0];
type Options = SpreadsheetOptions;

/**
 * Props of {@link Spreadsheet}: every `SpreadsheetOptions` key plus the Vue-only rendering props.
 * `seed` and `signal` are read at mount; every other option is applied to the live instance with
 * `update` when it changes structurally, so an inline object literal does not remount.
 */
export interface SpreadsheetProps extends SpreadsheetOptions {
  /** Rendered in place of the grid when the mount rejects. */
  errorFallback?: (error: unknown) => VNodeChild;
  class?: string | string[] | Record<string, boolean>;
  style?: CSSProperties;
}

/** Events of {@link Spreadsheet}. `error` carries a mount rejection, an option `update` refused,
 *  or an instance `error` event's `error`. */
export type SpreadsheetEmits = {
  ready: (spreadsheet: SpreadsheetInstance) => true;
  change: (event: ChangeEvent) => true;
  cell: (event: CellEvent) => true;
  workbook: (event: WorkbookEvent) => true;
  recalc: (event: RecalcEvent) => true;
  update: (event: UpdateEvent) => true;
  selectionChange: (selection: SelectionState) => true;
  localeChange: (event: LocaleChangeEvent) => true;
  error: (error: unknown) => true;
};

/** What a template ref to {@link Spreadsheet} exposes. */
export interface SpreadsheetExposed {
  /** The mounted spreadsheet; `null` before it is ready and after unmount. */
  readonly spreadsheet: SpreadsheetInstance | null;
}

const option = <K extends keyof Options>(type: PropType<Options[K]> | null) => ({
  type: type as PropType<Options[K]>,
  default: undefined,
});

/** One prop per `SpreadsheetOptions` key; the `satisfies` fails to compile on a missing or extra
 *  key, so a new option cannot reach the component unwired. */
const optionProps = {
  workbook: option<'workbook'>(Object as PropType<Options['workbook']>),
  seed: option<'seed'>(Function as PropType<Options['seed']>),
  ui: option<'ui'>(Object as PropType<Options['ui']>),
  policy: option<'policy'>(Object as PropType<Options['policy']>),
  viewport: option<'viewport'>(Object as PropType<Options['viewport']>),
  contextMenu: option<'contextMenu'>(Object as PropType<Options['contextMenu']>),
  overlays: option<'overlays'>(Object as PropType<Options['overlays']>),
  toolbar: option<'toolbar'>(Object as PropType<Options['toolbar']>),
  theme: option<'theme'>(String as PropType<Options['theme']>),
  locale: option<'locale'>(String as PropType<Options['locale']>),
  strings: option<'strings'>(Object as PropType<Options['strings']>),
  extensions: option<'extensions'>(Array as PropType<Options['extensions']>),
  formulaHints: option<'formulaHints'>(Array as PropType<Options['formulaHints']>),
  spelling: option<'spelling'>(Object as PropType<Options['spelling']>),
  autoCorrect: option<'autoCorrect'>(Object as PropType<Options['autoCorrect']>),
  host: option<'host'>(Object as PropType<Options['host']>),
  uploadStatus: option<'uploadStatus'>(String as PropType<Options['uploadStatus']>),
  macroRecording: option<'macroRecording'>(Boolean as PropType<Options['macroRecording']>),
  showAllNotes: option<'showAllNotes'>(Boolean as PropType<Options['showAllNotes']>),
  signal: option<'signal'>(null),
} satisfies { [K in keyof Options]-?: ReturnType<typeof option<K>> };

const component = defineComponent({
  name: 'Spreadsheet',
  props: {
    ...optionProps,
    errorFallback: {
      type: Function as PropType<(error: unknown) => VNodeChild>,
      default: undefined,
    },
    class: {
      type: [String, Array, Object] as PropType<SpreadsheetProps['class']>,
      default: undefined,
    },
    style: { type: Object as PropType<CSSProperties>, default: undefined },
  },
  emits: {
    ready: (_spreadsheet: SpreadsheetInstance) => true,
    change: (_event: ChangeEvent) => true,
    cell: (_event: CellEvent) => true,
    workbook: (_event: WorkbookEvent) => true,
    recalc: (_event: RecalcEvent) => true,
    update: (_event: UpdateEvent) => true,
    selectionChange: (_selection: SelectionState) => true,
    localeChange: (_event: LocaleChangeEvent) => true,
    error: (_error: unknown) => true,
  },
  setup(props, { emit, expose, slots }) {
    const hostEl = ref<HTMLDivElement | null>(null);
    // shallowRef so Vue does not walk the instance.
    const instance = shallowRef<SpreadsheetInstance | null>(null);
    const mountError = shallowRef<{ error: unknown } | null>(null);
    const offs: (() => void)[] = [];
    // The options the live instance was last brought to.
    let applied: Partial<SpreadsheetPatch> = {};
    let disposed = false;
    const report = (error: unknown): void => {
      if (!disposed) emit('error', error);
    };
    const currentOptions = (): SpreadsheetOptions => optionsFromProps(props as SpreadsheetOptions);

    onMounted(() => {
      const host = hostEl.value;
      if (!host) return;
      const options = currentOptions();
      mount(host, options).then(
        (inst) => {
          if (disposed) {
            inst.dispose();
            return;
          }
          offs.push(
            inst.on('change', (e) => emit('change', e)),
            inst.on('cell', (e) => emit('cell', e)),
            inst.on('workbook', (e) => emit('workbook', e)),
            inst.on('recalc', (e) => emit('recalc', e)),
            inst.on('update', (e) => emit('update', e)),
            inst.on('error', (e) => report(e.error)),
            inst.selection.on('change', (s) => emit('selectionChange', s)),
            inst.i18n.on('change', (e) => emit('localeChange', e)),
          );
          instance.value = inst;
          // Props that changed while the mount was pending.
          const latest = patchFromOptions(currentOptions());
          reconcileProps(inst, patchFromOptions(options), latest, report);
          applied = latest;
          emit('ready', inst);
        },
        (error: unknown) => {
          if (disposed) return;
          mountError.value = { error };
          report(error);
        },
      );
    });

    watch(
      () => patchFromOptions(currentOptions()),
      (next) => {
        const inst = instance.value;
        if (!inst) return;
        const previous = applied;
        applied = next;
        reconcileProps(inst, previous, next, report);
      },
    );

    onBeforeUnmount(() => {
      disposed = true;
      for (const off of offs) off();
      instance.value?.dispose();
      instance.value = null;
    });

    expose({
      get spreadsheet() {
        return instance.value;
      },
    } satisfies SpreadsheetExposed);

    return () => {
      const failed = mountError.value;
      const live = instance.value;
      return [
        h('div', { ref: hostEl, class: props.class, style: props.style }),
        failed && props.errorFallback ? props.errorFallback(failed.error) : null,
        live ? slots.default?.({ spreadsheet: live }) : null,
      ];
    };
  },
});

/**
 * The spreadsheet as a Vue component. It mounts once; option props are applied to the live
 * instance with `update`, and the instance is disposed on unmount. The default slot receives
 * `{ spreadsheet }` once it is mounted.
 */
export const Spreadsheet: DefineComponent<
  SpreadsheetProps,
  SpreadsheetExposed,
  Record<string, never>,
  Record<string, never>,
  Record<string, never>,
  ComponentOptionsMixin,
  ComponentOptionsMixin,
  SpreadsheetEmits
> = component as never;
