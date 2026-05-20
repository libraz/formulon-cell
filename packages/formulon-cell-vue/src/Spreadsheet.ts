import {
  type CellChangeEvent,
  type ExtensionInput,
  type FeatureFlags,
  type LocaleChangeEvent,
  type MountOptions,
  type PrinterProfile,
  type RecalcEvent,
  resolveSpreadsheetUiOptions,
  type SelectionChangeEvent,
  Spreadsheet as SpreadsheetCore,
  type SpreadsheetInstance,
  type SpreadsheetUiOptions,
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
    ui?: SpreadsheetUiOptions;
    workbook?: WorkbookHandle;
    theme?: MountOptions['theme'];
    locale?: MountOptions['locale'];
    strings?: MountOptions['strings'];
    features?: FeatureFlags;
    extensions?: ExtensionInput[];
    printerProfiles?: readonly PrinterProfile[];
    printerProfileId?: string;
    refreshPrinterProfiles?: MountOptions['refreshPrinterProfiles'];
    captureScreenClip?: MountOptions['captureScreenClip'];
    uploadStatus?: MountOptions['uploadStatus'];
    macroRecording?: MountOptions['macroRecording'];
  },
  baseline: {
    ui?: SpreadsheetUiOptions;
    workbook?: WorkbookHandle;
    theme?: MountOptions['theme'];
    locale?: MountOptions['locale'];
    strings?: MountOptions['strings'];
    features?: FeatureFlags;
    extensions?: ExtensionInput[];
    printerProfiles?: readonly PrinterProfile[];
    printerProfileId?: string;
    refreshPrinterProfiles?: MountOptions['refreshPrinterProfiles'];
    captureScreenClip?: MountOptions['captureScreenClip'];
    uploadStatus?: MountOptions['uploadStatus'];
    macroRecording?: MountOptions['macroRecording'];
  } = {},
): Promise<void> => {
  const ui = props.ui ? resolveSpreadsheetUiOptions(props.ui) : null;
  if (props.workbook && props.workbook !== baseline.workbook && props.workbook !== inst.workbook) {
    await inst.setWorkbook(props.workbook);
  }
  const nextTheme = props.theme ?? ui?.theme;
  const baselineTheme =
    baseline.theme ?? (baseline.ui ? resolveSpreadsheetUiOptions(baseline.ui).theme : undefined);
  if (nextTheme && nextTheme !== baselineTheme) inst.setTheme(nextTheme);
  if (props.locale && props.locale !== baseline.locale) inst.i18n.setLocale(props.locale);
  if (props.strings && props.strings !== baseline.strings) {
    inst.i18n.extend(inst.i18n.locale, props.strings);
  }
  if (props.features !== baseline.features || props.ui !== baseline.ui) {
    inst.setFeatures({ ...(ui?.features ?? {}), ...(props.features ?? {}) });
  }
  if (props.extensions !== baseline.extensions) inst.setExtensions(props.extensions);
  if (props.printerProfiles !== baseline.printerProfiles) {
    inst.setPrinterProfiles(props.printerProfiles);
  }
  if (props.printerProfileId !== baseline.printerProfileId) {
    inst.setPrinterProfileId(props.printerProfileId);
  }
  if (props.uploadStatus !== baseline.uploadStatus) {
    inst.setUploadStatus(props.uploadStatus ?? null);
  }
  if (props.macroRecording !== baseline.macroRecording) {
    inst.setMacroRecording(props.macroRecording ?? null);
  }
};

export const Spreadsheet: ReturnType<typeof defineComponent> = defineComponent({
  name: 'Spreadsheet',
  props: {
    ui: { type: Object as PropType<SpreadsheetUiOptions>, default: undefined },
    workbook: { type: Object as PropType<WorkbookHandle>, default: undefined },
    theme: { type: String as PropType<MountOptions['theme']>, default: undefined },
    locale: { type: String as PropType<MountOptions['locale']>, default: undefined },
    strings: { type: Object as PropType<MountOptions['strings']>, default: undefined },
    features: { type: Object as PropType<FeatureFlags>, default: undefined },
    extensions: { type: Array as PropType<ExtensionInput[]>, default: undefined },
    printerProfiles: { type: Array as PropType<readonly PrinterProfile[]>, default: undefined },
    printerProfileId: { type: String, default: undefined },
    refreshPrinterProfiles: {
      type: Function as PropType<MountOptions['refreshPrinterProfiles']>,
      default: undefined,
    },
    captureScreenClip: {
      type: Function as PropType<MountOptions['captureScreenClip']>,
      default: undefined,
    },
    uploadStatus: { type: String as PropType<MountOptions['uploadStatus']>, default: undefined },
    macroRecording: {
      type: Boolean as PropType<MountOptions['macroRecording']>,
      default: undefined,
    },
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
      if (props.ui) opts.ui = props.ui;
      if (props.workbook) opts.workbook = props.workbook;
      if (props.theme) opts.theme = props.theme;
      if (props.locale) opts.locale = props.locale;
      if (props.strings) opts.strings = props.strings;
      if (props.features) opts.features = props.features;
      if (props.extensions) opts.extensions = props.extensions;
      if (props.printerProfiles) opts.printerProfiles = props.printerProfiles;
      if (props.printerProfileId) opts.printerProfileId = props.printerProfileId;
      if (props.refreshPrinterProfiles) opts.refreshPrinterProfiles = props.refreshPrinterProfiles;
      if (props.captureScreenClip) {
        opts.captureScreenClip = () => props.captureScreenClip?.();
      }
      if (props.uploadStatus !== undefined) opts.uploadStatus = props.uploadStatus;
      if (props.macroRecording !== undefined) opts.macroRecording = props.macroRecording;
      if (props.functions) opts.functions = props.functions;
      if (props.seed) opts.seed = props.seed;
      opts.renderError = !props.errorFallback;
      opts.onError = (error) => {
        mountError.value = error;
        emit('error', error);
      };
      const mountedWith = {
        workbook: opts.workbook,
        ui: opts.ui,
        theme: opts.theme,
        locale: opts.locale,
        strings: opts.strings,
        features: opts.features,
        extensions: opts.extensions,
        printerProfiles: opts.printerProfiles,
        printerProfileId: opts.printerProfileId,
        captureScreenClip: opts.captureScreenClip,
        uploadStatus: opts.uploadStatus,
        macroRecording: opts.macroRecording,
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
      () => [props.theme, props.ui] as const,
      ([nextTheme, nextUi]) => {
        const resolved = nextUi ? resolveSpreadsheetUiOptions(nextUi) : null;
        const theme = nextTheme ?? resolved?.theme;
        if (theme && instance.value) instance.value.setTheme(theme);
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
      () => [props.features, props.ui] as const,
      ([nextFeatures, nextUi]) => {
        const resolved = nextUi ? resolveSpreadsheetUiOptions(nextUi) : null;
        if (instance.value) {
          instance.value.setFeatures({ ...(resolved?.features ?? {}), ...(nextFeatures ?? {}) });
        }
      },
      { deep: true },
    );
    watch(
      () => props.extensions,
      (next) => {
        if (instance.value) instance.value.setExtensions(next);
      },
    );
    watch(
      () => props.printerProfiles,
      (next) => {
        if (instance.value) instance.value.setPrinterProfiles(next);
      },
    );
    watch(
      () => props.printerProfileId,
      (next) => {
        if (instance.value) instance.value.setPrinterProfileId(next);
      },
    );
    watch(
      () => props.uploadStatus,
      (next) => {
        if (instance.value) instance.value.setUploadStatus(next ?? null);
      },
    );
    watch(
      () => props.macroRecording,
      (next) => {
        if (instance.value) instance.value.setMacroRecording(next ?? null);
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
