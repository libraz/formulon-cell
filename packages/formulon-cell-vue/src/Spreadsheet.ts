import {
  type CellChangeEvent,
  type ChangeBatchResult,
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
    policy?: MountOptions['policy'];
    viewport?: MountOptions['viewport'];
    contextMenu?: MountOptions['contextMenu'];
    overlays?: MountOptions['overlays'];
    toolbar?: MountOptions['toolbar'];
    theme?: MountOptions['theme'];
    locale?: MountOptions['locale'];
    strings?: MountOptions['strings'];
    features?: FeatureFlags;
    extensions?: ExtensionInput[];
    printerProfiles?: readonly PrinterProfile[];
    printerProfileId?: string;
    refreshPrinterProfiles?: MountOptions['refreshPrinterProfiles'];
    captureScreenClip?: MountOptions['captureScreenClip'];
    getFunctionArgumentHelp?: MountOptions['getFunctionArgumentHelp'];
    uploadStatus?: MountOptions['uploadStatus'];
    macroRecording?: MountOptions['macroRecording'];
  },
  baseline: {
    ui?: SpreadsheetUiOptions;
    workbook?: WorkbookHandle;
    policy?: MountOptions['policy'];
    viewport?: MountOptions['viewport'];
    contextMenu?: MountOptions['contextMenu'];
    overlays?: MountOptions['overlays'];
    toolbar?: MountOptions['toolbar'];
    theme?: MountOptions['theme'];
    locale?: MountOptions['locale'];
    strings?: MountOptions['strings'];
    features?: FeatureFlags;
    extensions?: ExtensionInput[];
    printerProfiles?: readonly PrinterProfile[];
    printerProfileId?: string;
    refreshPrinterProfiles?: MountOptions['refreshPrinterProfiles'];
    captureScreenClip?: MountOptions['captureScreenClip'];
    getFunctionArgumentHelp?: MountOptions['getFunctionArgumentHelp'];
    uploadStatus?: MountOptions['uploadStatus'];
    macroRecording?: MountOptions['macroRecording'];
  } = {},
): Promise<void> => {
  const resolvedUi = resolveSpreadsheetUiOptions(props.ui);
  const ui = props.ui ? resolvedUi : null;
  if (props.workbook && props.workbook !== baseline.workbook && props.workbook !== inst.workbook) {
    await inst.setWorkbook(props.workbook);
  }
  if (props.policy !== baseline.policy) inst.setPolicy(props.policy);
  if (props.viewport !== baseline.viewport) inst.setViewportOptions(props.viewport);
  if (props.contextMenu !== baseline.contextMenu) inst.setContextMenu(props.contextMenu);
  if (props.overlays !== baseline.overlays) inst.setOverlayOptions(props.overlays);
  if (props.ui !== baseline.ui) inst.setUi(props.ui);
  const nextTheme = props.theme ?? ui?.theme;
  const baselineTheme =
    baseline.theme ?? (baseline.ui ? resolveSpreadsheetUiOptions(baseline.ui).theme : undefined);
  if (nextTheme && nextTheme !== baselineTheme) inst.setTheme(nextTheme);
  if (props.locale && props.locale !== baseline.locale) inst.i18n.setLocale(props.locale);
  if (props.strings && props.strings !== baseline.strings) {
    inst.i18n.extend(inst.i18n.locale, props.strings);
  }
  if (props.features !== baseline.features || props.ui !== baseline.ui) {
    inst.setFeatures({ ...resolvedUi.features, ...(props.features ?? {}) });
  }
  if (props.toolbar !== baseline.toolbar || props.ui !== baseline.ui) {
    inst.setToolbar(props.toolbar ?? (props.ui ? resolvedUi.ribbon : false));
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
    policy: { type: Object as PropType<MountOptions['policy']>, default: undefined },
    viewport: { type: Object as PropType<MountOptions['viewport']>, default: undefined },
    contextMenu: { type: Object as PropType<MountOptions['contextMenu']>, default: undefined },
    overlays: { type: Object as PropType<MountOptions['overlays']>, default: undefined },
    toolbar: {
      type: [Boolean, Object] as PropType<MountOptions['toolbar']>,
      default: undefined,
    },
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
    getFunctionArgumentHelp: {
      type: Function as PropType<MountOptions['getFunctionArgumentHelp']>,
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
    'change-batch': (_e: ChangeBatchResult) => true,
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
      if (props.policy !== undefined) opts.policy = props.policy;
      if (props.viewport !== undefined) opts.viewport = props.viewport;
      if (props.contextMenu !== undefined) opts.contextMenu = props.contextMenu;
      if (props.overlays !== undefined) opts.overlays = props.overlays;
      if (props.toolbar !== undefined) opts.toolbar = props.toolbar;
      if (props.theme) opts.theme = props.theme;
      if (props.locale) opts.locale = props.locale;
      if (props.strings) opts.strings = props.strings;
      if (props.features) opts.features = props.features;
      if (props.extensions) opts.extensions = props.extensions;
      if (props.printerProfiles) opts.printerProfiles = props.printerProfiles;
      if (props.printerProfileId) opts.printerProfileId = props.printerProfileId;
      opts.refreshPrinterProfiles = () => props.refreshPrinterProfiles?.();
      opts.captureScreenClip = () => props.captureScreenClip?.();
      opts.getFunctionArgumentHelp = (name, index, locale) =>
        props.getFunctionArgumentHelp?.(name, index, locale);
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
        policy: opts.policy,
        viewport: opts.viewport,
        contextMenu: opts.contextMenu,
        overlays: opts.overlays,
        toolbar: opts.toolbar,
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
        inst.on('changeBatch', (e) => emit('change-batch', e)),
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
      (nextTheme) => {
        const resolved = props.ui ? resolveSpreadsheetUiOptions(props.ui) : null;
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
      ([nextFeatures, nextUi], previous) => {
        const resolved = resolveSpreadsheetUiOptions(nextUi);
        if (instance.value) {
          const uiChanged = nextUi !== previous?.[1];
          if (uiChanged) instance.value.setUi(nextUi);
          instance.value.setFeatures({ ...resolved.features, ...(nextFeatures ?? {}) });
          instance.value.setToolbar(props.toolbar ?? (nextUi ? resolved.ribbon : false));
          if (uiChanged) {
            const theme = props.theme ?? (nextUi ? resolved.theme : undefined);
            if (theme) instance.value.setTheme(theme);
          }
        }
      },
      { deep: true },
    );
    watch(
      () => props.policy,
      (next) => {
        if (instance.value) instance.value.setPolicy(next);
      },
      { deep: true },
    );
    watch(
      () => props.viewport,
      (next) => {
        if (instance.value) instance.value.setViewportOptions(next);
      },
      { deep: true },
    );
    watch(
      () => props.contextMenu,
      (next) => {
        if (instance.value) instance.value.setContextMenu(next);
      },
      { deep: true },
    );
    watch(
      () => props.overlays,
      (next) => {
        if (instance.value) instance.value.setOverlayOptions(next);
      },
      { deep: true },
    );
    watch(
      () => props.toolbar,
      (next) => {
        if (instance.value) {
          const resolved = resolveSpreadsheetUiOptions(props.ui);
          instance.value.setToolbar(next ?? (props.ui ? resolved.ribbon : false));
        }
      },
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
