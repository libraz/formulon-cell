import {
  type CellChangeEvent,
  type ExtensionInput,
  type FeatureFlags,
  type LocaleChangeEvent,
  type MountOptions,
  type RecalcEvent,
  type SelectionChangeEvent,
  Spreadsheet as SpreadsheetCore,
  type SpreadsheetInstance,
  type ThemeChangeEvent,
  type WorkbookChangeEvent,
  type WorkbookHandle,
} from '@libraz/formulon-cell';
import {
  type CSSProperties,
  computed,
  defineComponent,
  h,
  onBeforeUnmount,
  onMounted,
  type PropType,
  type Ref,
  ref,
  shallowRef,
  type VNodeChild,
  watch,
} from 'vue';

export type SpreadsheetExposed = {
  readonly instance: Ref<SpreadsheetInstance | null>;
};

const applyRuntimeProps = async (
  inst: SpreadsheetInstance,
  props: {
    workbook?: WorkbookHandle;
    theme?: MountOptions['theme'];
    locale?: MountOptions['locale'];
    strings?: MountOptions['strings'];
    features?: FeatureFlags;
    extensions?: ExtensionInput[];
  },
  baseline: {
    workbook?: WorkbookHandle;
    theme?: MountOptions['theme'];
    locale?: MountOptions['locale'];
    strings?: MountOptions['strings'];
    features?: FeatureFlags;
    extensions?: ExtensionInput[];
  } = {},
): Promise<void> => {
  if (props.workbook && props.workbook !== baseline.workbook && props.workbook !== inst.workbook) {
    await inst.setWorkbook(props.workbook);
  }
  if (props.theme && props.theme !== baseline.theme) inst.setTheme(props.theme);
  if (props.locale && props.locale !== baseline.locale) inst.i18n.setLocale(props.locale);
  if (props.strings && props.strings !== baseline.strings) {
    inst.i18n.extend(inst.i18n.locale, props.strings);
  }
  if (props.features !== baseline.features) inst.setFeatures(props.features ?? {});
  if (props.extensions !== baseline.extensions) inst.setExtensions(props.extensions);
};

export const Spreadsheet = defineComponent({
  name: 'Spreadsheet',
  props: {
    workbook: { type: Object as PropType<WorkbookHandle>, default: undefined },
    theme: { type: String as PropType<MountOptions['theme']>, default: undefined },
    locale: { type: String as PropType<MountOptions['locale']>, default: undefined },
    strings: { type: Object as PropType<MountOptions['strings']>, default: undefined },
    features: { type: Object as PropType<FeatureFlags>, default: undefined },
    extensions: { type: Array as PropType<ExtensionInput[]>, default: undefined },
    functions: { type: Array as PropType<MountOptions['functions']>, default: undefined },
    seed: { type: Function as PropType<MountOptions['seed']>, default: undefined },
    errorFallback: {
      type: Function as PropType<(error: unknown) => VNodeChild>,
      default: undefined,
    },
    class: { type: [String, Array, Object] as PropType<string | string[] | object>, default: '' },
    style: { type: Object as PropType<CSSProperties>, default: undefined },
  },
  emits: {
    ready: (_inst: SpreadsheetInstance) => true,
    cellChange: (_e: CellChangeEvent) => true,
    selectionChange: (_e: SelectionChangeEvent) => true,
    workbookChange: (_e: WorkbookChangeEvent) => true,
    localeChange: (_e: LocaleChangeEvent) => true,
    themeChange: (_e: ThemeChangeEvent) => true,
    recalc: (_e: RecalcEvent) => true,
    error: (_e: unknown) => true,
  },
  setup(props, { emit, expose }) {
    const hostEl = ref<HTMLDivElement | null>(null);
    // shallowRef so Vue doesn't deep-walk the spreadsheet's internal state.
    const instance = shallowRef<SpreadsheetInstance | null>(null);
    const mountError = shallowRef<unknown>(null);
    const eventDisposers: (() => void)[] = [];
    let disposed = false;

    onMounted(async () => {
      const host = hostEl.value;
      if (!host) return;
      const opts: MountOptions = {};
      if (props.workbook) opts.workbook = props.workbook;
      if (props.theme) opts.theme = props.theme;
      if (props.locale) opts.locale = props.locale;
      if (props.strings) opts.strings = props.strings;
      if (props.features) opts.features = props.features;
      if (props.extensions) opts.extensions = props.extensions;
      if (props.functions) opts.functions = props.functions;
      if (props.seed) opts.seed = props.seed;
      opts.renderError = !props.errorFallback;
      opts.onError = (error) => {
        mountError.value = error;
        emit('error', error);
      };
      const mountedWith = {
        workbook: opts.workbook,
        theme: opts.theme,
        locale: opts.locale,
        strings: opts.strings,
        features: opts.features,
        extensions: opts.extensions,
      };
      let inst: SpreadsheetInstance;
      try {
        inst = await SpreadsheetCore.mount(host, opts);
      } catch (error) {
        mountError.value = error;
        return;
      }
      if (disposed) {
        inst.dispose();
        return;
      }
      mountError.value = null;
      instance.value = inst;
      eventDisposers.push(
        inst.on('cellChange', (e) => emit('cellChange', e)),
        inst.on('selectionChange', (e) => emit('selectionChange', e)),
        inst.on('workbookChange', (e) => emit('workbookChange', e)),
        inst.on('localeChange', (e) => emit('localeChange', e)),
        inst.on('themeChange', (e) => emit('themeChange', e)),
        inst.on('recalc', (e) => emit('recalc', e)),
      );
      await applyRuntimeProps(inst, props, mountedWith);
      if (disposed) return;
      emit('ready', inst);
    });

    // Theme / locale / workbook are all cheap to swap via the imperative
    // API — react to prop changes without re-mounting.
    watch(
      () => props.theme,
      (next) => {
        if (next && instance.value) instance.value.setTheme(next);
      },
    );
    watch(
      () => props.locale,
      (next) => {
        if (next && instance.value) instance.value.i18n.setLocale(next);
      },
    );
    watch(
      () => props.strings,
      (next) => {
        if (next && instance.value) instance.value.i18n.extend(instance.value.i18n.locale, next);
      },
      { deep: true },
    );
    watch(
      () => props.workbook,
      (next) => {
        if (next && instance.value && next !== instance.value.workbook) {
          void instance.value.setWorkbook(next);
        }
      },
    );
    watch(
      () => props.features,
      (next) => {
        if (instance.value) instance.value.setFeatures(next ?? {});
      },
      { deep: true },
    );
    watch(
      () => props.extensions,
      (next) => {
        if (instance.value) instance.value.setExtensions(next);
      },
    );

    onBeforeUnmount(() => {
      disposed = true;
      for (const d of eventDisposers) d();
      instance.value?.dispose();
      instance.value = null;
    });

    expose({ instance });

    const renderHost = computed(() =>
      h('div', { ref: hostEl, class: props.class, style: props.style }),
    );
    return () => {
      const fallback =
        mountError.value && props.errorFallback ? props.errorFallback(mountError.value) : null;
      return fallback ? [renderHost.value, fallback] : renderHost.value;
    };
  },
});
