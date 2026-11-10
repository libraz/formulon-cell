import {
  applyControlledLocale,
  applyControlledProps,
  applyUiFeaturesToolbar,
  type CellChangeEvent,
  type ChangeBatchResult,
  type ExtensionInput,
  type FeatureFlags,
  type LocaleChangeEvent,
  type MountOptions,
  type PrinterProfile,
  type LegacyRecalcEvent as RecalcEvent,
  resolveAdapterTheme,
  type SelectionChangeEvent,
  Spreadsheet as SpreadsheetCore,
  type LegacySpreadsheetInstance as SpreadsheetInstance,
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

export const Spreadsheet: ReturnType<typeof defineComponent> = defineComponent({
  name: 'Spreadsheet',
  props: {
    ui: { type: Object as PropType<SpreadsheetUiOptions>, default: undefined },
    workbook: { type: Object as PropType<WorkbookHandle>, default: undefined },
    getWorkbookName: {
      type: Function as PropType<MountOptions['getWorkbookName']>,
      default: undefined,
    },
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
    let localeEventGeneration = 0;
    const notifyLocaleChange = (event: LocaleChangeEvent): void => {
      localeEventGeneration += 1;
      emit('localeChange', event);
    };
    const getLocaleEventGeneration = (): number => localeEventGeneration;

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
      opts.getWorkbookName = () => props.getWorkbookName?.();
      opts.captureScreenClip = () => props.captureScreenClip?.();
      opts.getFunctionArgumentHelp = (name, index, locale) =>
        props.getFunctionArgumentHelp?.(name, index, locale);
      if (props.uploadStatus !== undefined) opts.uploadStatus = props.uploadStatus;
      if (props.macroRecording !== undefined) opts.macroRecording = props.macroRecording;
      if (props.functions) opts.functions = props.functions;
      if (props.seed) opts.seed = props.seed;
      opts.renderError = !props.errorFallback;
      opts.onError = (error) => {
        if (disposed) return;
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
        if (!disposed) mountError.value = error;
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
        inst.on('localeChange', notifyLocaleChange),
        inst.on('themeChange', (e) => emit('themeChange', e)),
        inst.on('recalc', (e) => emit('recalc', e)),
      );
      try {
        await applyControlledProps(inst, props, mountedWith, {
          onLocaleChange: notifyLocaleChange,
          onError: (error) => emit('error', error),
          getLocaleEventGeneration,
        });
      } catch (error) {
        // Runtime prop reconciliation can fail while the already-mounted
        // instance remains usable. Keep it live and let the specific
        // reconciliation path report its cause through `error`.
        if (!disposed) emit('error', error);
      }
      if (disposed) return;
      emit('ready', inst);
    });

    // Theme / locale / workbook are all cheap to swap via the imperative
    // API — react to prop changes without re-mounting.
    watch(
      () => props.theme,
      () => {
        const theme = resolveAdapterTheme(props);
        if (theme && instance.value) instance.value.setTheme(theme);
      },
    );
    watch(
      () => props.locale,
      (next) => {
        if (next && instance.value) {
          const accepted = applyControlledLocale(instance.value, next, {
            onLocaleChange: notifyLocaleChange,
            onError: (error) => emit('error', error),
            getLocaleEventGeneration,
          });
          if (accepted && props.strings) {
            instance.value.i18n.extend(instance.value.i18n.locale, props.strings);
          }
        }
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
        const inst = instance.value;
        if (next && inst && next !== inst.workbook) {
          void inst.setWorkbook(next).catch((error: unknown) => {
            if (!disposed && instance.value === inst) emit('error', error);
          });
        }
      },
    );
    watch(
      () => [props.features, props.ui, props.toolbar] as const,
      ([nextFeatures, nextUi, nextToolbar], previous) => {
        if (instance.value) {
          const uiChanged = nextUi !== previous?.[1];
          applyUiFeaturesToolbar(
            instance.value,
            { ui: nextUi, features: nextFeatures, toolbar: nextToolbar },
            uiChanged,
          );
          if (uiChanged) {
            const theme = resolveAdapterTheme({ theme: props.theme, ui: nextUi });
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
