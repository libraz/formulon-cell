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
import type { CSSProperties, ForwardedRef, ReactNode } from 'react';
import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from 'react';

export interface SpreadsheetProps {
  /** Optional pre-loaded workbook (e.g. from xlsx bytes). When omitted, a
   *  fresh default workbook is created on mount. */
  workbook?: WorkbookHandle;
  /** Host interaction restrictions for embedded forms and viewers. */
  policy?: MountOptions['policy'];
  /** Visible/selectable bounds and keyboard-navigation behavior. */
  viewport?: MountOptions['viewport'];
  /** Built-in, transformed, or host-owned context menu. */
  contextMenu?: MountOptions['contextMenu'];
  /** Overlay root inside a host modal or fullscreen boundary. */
  overlays?: MountOptions['overlays'];
  /** Simplified Excel-365-style UI preset and feature switches. `theme` and
   *  `features` props override the matching values when both are supplied. */
  ui?: SpreadsheetUiOptions;
  /** Mount or hide the core ribbon. Omitted follows the selected `ui` profile;
   *  with no `ui`, the wrapper keeps its legacy ribbonless default. */
  toolbar?: MountOptions['toolbar'];
  /** Theme. When the prop changes after mount, the component calls
   *  `instance.setTheme()` to keep the spreadsheet in sync. */
  theme?: MountOptions['theme'];
  /** UI locale. When the prop changes after mount, the component calls
   *  `instance.i18n.setLocale()`. */
  locale?: MountOptions['locale'];
  /** Per-string overrides applied on top of the chosen locale. */
  strings?: MountOptions['strings'];
  /** Toggle individual built-in features. */
  features?: FeatureFlags;
  /** Custom extensions added on top of built-ins. */
  extensions?: ExtensionInput[];
  /** Initial set of host-side custom functions registered against
   *  `instance.formula`. */
  functions?: MountOptions['functions'];
  /** Optional cell-seeding callback (mostly useful for demos). */
  seed?: MountOptions['seed'];
  /** Host-provided printer profiles for the built-in print/PDF flow. */
  printerProfiles?: readonly PrinterProfile[];
  /** Active host printer profile id used by the built-in print/PDF flow. */
  printerProfileId?: string;
  /** Host refresh hook for native/Electron printer discovery. */
  refreshPrinterProfiles?: MountOptions['refreshPrinterProfiles'];
  /** Host capture hook for Insert > Screenshot > Screen Clipping. */
  captureScreenClip?: MountOptions['captureScreenClip'];
  /** Host lookup for per-argument help in the Mac formula palette. */
  getFunctionArgumentHelp?: MountOptions['getFunctionArgumentHelp'];
  /** Host-driven status bar Upload Status indicator. */
  uploadStatus?: MountOptions['uploadStatus'];
  /** Host-driven status bar Macro Recording indicator. */
  macroRecording?: MountOptions['macroRecording'];
  /** Fires once after mount with the live instance. Use this to wire toolbars
   *  / menus that talk to the spreadsheet's imperative API. */
  onReady?: (instance: SpreadsheetInstance) => void;
  /** Fires when the component cannot mount a spreadsheet instance. */
  onError?: (error: unknown) => void;
  /** Optional framework-native fallback shown after mount failure. */
  errorFallback?: ReactNode | ((error: unknown) => ReactNode);
  /** Fires every time a cell value changes (engine-side). Use this to
   *  mirror the spreadsheet into outer state (Redux, Zustand, server). */
  onCellChange?: (e: CellChangeEvent) => void;
  /** Fires after a policy-aware cell batch is applied or rejected. */
  onChangeBatch?: (e: ChangeBatchResult) => void;
  /** Fires when the active cell / range moves. */
  onSelectionChange?: (e: SelectionChangeEvent) => void;
  /** Fires after `instance.setWorkbook(next)` swaps the engine. */
  onWorkbookChange?: (e: WorkbookChangeEvent) => void;
  /** Fires when `instance.i18n.setLocale` switches the active dictionary. */
  onLocaleChange?: (e: LocaleChangeEvent) => void;
  /** Fires when the host theme attribute changes. */
  onThemeChange?: (e: ThemeChangeEvent) => void;
  /** Fires after the engine reports a recalc batch. */
  onRecalc?: (e: RecalcEvent) => void;
  className?: string;
  style?: CSSProperties;
  /** Optional render prop — receives the live instance after mount. Useful
   *  for embedding a toolbar component that needs ref access. */
  children?: ReactNode | ((instance: SpreadsheetInstance) => ReactNode);
}

export interface SpreadsheetRef {
  readonly instance: SpreadsheetInstance | null;
}

const applyRuntimeProps = async (
  inst: SpreadsheetInstance,
  props: SpreadsheetProps,
  baseline: SpreadsheetProps = {},
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

const SpreadsheetComponent = (
  props: SpreadsheetProps,
  ref: ForwardedRef<SpreadsheetRef>,
): ReactNode => {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const instanceRef = useRef<SpreadsheetInstance | null>(null);
  const [liveInstance, setLiveInstance] = useState<SpreadsheetInstance | null>(null);
  const [mountError, setMountError] = useState<unknown>(null);
  // Keep the latest props in a ref so the mount effect doesn't have to
  // re-run when callbacks change. Mounting is expensive (it creates the
  // wb + canvas + listeners) so we only re-mount on workbook identity.
  const propsRef = useRef(props);
  propsRef.current = props;
  const previousUiRef = useRef(props.ui);

  useImperativeHandle(
    ref,
    () => ({
      get instance() {
        return instanceRef.current;
      },
    }),
    [],
  );

  useEffect(() => {
    let disposed = false;
    const host = hostRef.current;
    if (!host) return undefined;
    const eventDisposers: (() => void)[] = [];
    void (async () => {
      const cur = propsRef.current;
      const inst = await SpreadsheetCore.mount(host, {
        ...(cur.workbook ? { workbook: cur.workbook } : {}),
        ...(cur.policy !== undefined ? { policy: cur.policy } : {}),
        ...(cur.viewport !== undefined ? { viewport: cur.viewport } : {}),
        ...(cur.contextMenu !== undefined ? { contextMenu: cur.contextMenu } : {}),
        ...(cur.overlays !== undefined ? { overlays: cur.overlays } : {}),
        ...(cur.ui ? { ui: cur.ui } : {}),
        ...(cur.toolbar !== undefined ? { toolbar: cur.toolbar } : {}),
        ...(cur.theme ? { theme: cur.theme } : {}),
        ...(cur.locale ? { locale: cur.locale } : {}),
        ...(cur.strings ? { strings: cur.strings } : {}),
        ...(cur.features ? { features: cur.features } : {}),
        ...(cur.extensions ? { extensions: cur.extensions } : {}),
        ...(cur.functions ? { functions: cur.functions } : {}),
        ...(cur.seed ? { seed: cur.seed } : {}),
        ...(cur.printerProfiles ? { printerProfiles: cur.printerProfiles } : {}),
        ...(cur.printerProfileId ? { printerProfileId: cur.printerProfileId } : {}),
        refreshPrinterProfiles: () => propsRef.current.refreshPrinterProfiles?.(),
        captureScreenClip: () => propsRef.current.captureScreenClip?.(),
        getFunctionArgumentHelp: (name, index, locale) =>
          propsRef.current.getFunctionArgumentHelp?.(name, index, locale),
        ...(cur.uploadStatus !== undefined ? { uploadStatus: cur.uploadStatus } : {}),
        ...(cur.macroRecording !== undefined ? { macroRecording: cur.macroRecording } : {}),
        renderError: !cur.errorFallback,
        onError: (error) => {
          if (disposed) return;
          setMountError(error);
          propsRef.current.onError?.(error);
        },
      });
      if (disposed) {
        inst.dispose();
        return;
      }
      setMountError(null);
      instanceRef.current = inst;
      // Wire event props through `inst.on(...)`. Each handler reads from
      // propsRef, so callers can swap callbacks without re-mounting.
      eventDisposers.push(
        inst.on('changeBatch', (e) => propsRef.current.onChangeBatch?.(e)),
        inst.on('cellChange', (e) => propsRef.current.onCellChange?.(e)),
        inst.on('selectionChange', (e) => propsRef.current.onSelectionChange?.(e)),
        inst.on('workbookChange', (e) => propsRef.current.onWorkbookChange?.(e)),
        inst.on('localeChange', (e) => propsRef.current.onLocaleChange?.(e)),
        inst.on('themeChange', (e) => propsRef.current.onThemeChange?.(e)),
        inst.on('recalc', (e) => propsRef.current.onRecalc?.(e)),
      );
      if (propsRef.current !== cur) {
        await applyRuntimeProps(inst, propsRef.current, cur);
      }
      if (disposed) return;
      setLiveInstance(inst);
      propsRef.current.onReady?.(inst);
    })().catch((error: unknown) => {
      if (disposed) return;
      setMountError(error);
    });
    return () => {
      disposed = true;
      for (const d of eventDisposers) d();
      const inst = instanceRef.current;
      instanceRef.current = null;
      setLiveInstance(null);
      inst?.dispose();
    };
    // Mount once; prop mutations land via imperative methods on `instance`
    // and event handlers re-read from propsRef on each fire.
  }, []);

  // Forward reactive theme / locale / features / extensions prop changes
  // to the running instance via imperative APIs — avoids re-mounting.
  useEffect(() => {
    const inst = instanceRef.current;
    if (!inst || !props.workbook || props.workbook === inst.workbook) return;
    void inst.setWorkbook(props.workbook);
  }, [props.workbook]);

  useEffect(() => {
    const inst = instanceRef.current;
    if (!inst || !props.locale) return;
    inst.i18n.setLocale(props.locale);
  }, [props.locale]);

  useEffect(() => {
    const inst = instanceRef.current;
    if (!inst || !props.strings) return;
    inst.i18n.extend(inst.i18n.locale, props.strings);
  }, [props.strings]);

  useEffect(() => {
    const inst = instanceRef.current;
    const uiChanged = previousUiRef.current !== props.ui;
    previousUiRef.current = props.ui;
    if (!inst) return;
    const resolvedUi = resolveSpreadsheetUiOptions(props.ui);
    if (uiChanged) inst.setUi(props.ui);
    inst.setFeatures({ ...resolvedUi.features, ...(props.features ?? {}) });
    inst.setToolbar(props.toolbar ?? (props.ui ? resolvedUi.ribbon : false));
  }, [props.features, props.ui, props.toolbar]);

  useEffect(() => {
    const inst = instanceRef.current;
    if (!inst) return;
    const ui = props.ui ? resolveSpreadsheetUiOptions(props.ui) : null;
    const nextTheme = props.theme ?? ui?.theme;
    if (nextTheme) inst.setTheme(nextTheme);
  }, [props.theme, props.ui]);

  useEffect(() => {
    const inst = instanceRef.current;
    if (inst) inst.setPolicy(props.policy);
  }, [props.policy]);

  useEffect(() => {
    const inst = instanceRef.current;
    if (inst) inst.setViewportOptions(props.viewport);
  }, [props.viewport]);

  useEffect(() => {
    const inst = instanceRef.current;
    if (inst) inst.setContextMenu(props.contextMenu);
  }, [props.contextMenu]);

  useEffect(() => {
    const inst = instanceRef.current;
    if (inst) inst.setOverlayOptions(props.overlays);
  }, [props.overlays]);

  useEffect(() => {
    const inst = instanceRef.current;
    if (!inst) return;
    inst.setExtensions(props.extensions);
  }, [props.extensions]);

  useEffect(() => {
    const inst = instanceRef.current;
    if (!inst) return;
    inst.setPrinterProfiles(props.printerProfiles);
  }, [props.printerProfiles]);

  useEffect(() => {
    const inst = instanceRef.current;
    if (!inst) return;
    inst.setPrinterProfileId(props.printerProfileId);
  }, [props.printerProfileId]);

  useEffect(() => {
    const inst = instanceRef.current;
    if (!inst) return;
    inst.setUploadStatus(props.uploadStatus ?? null);
  }, [props.uploadStatus]);

  useEffect(() => {
    const inst = instanceRef.current;
    if (!inst) return;
    inst.setMacroRecording(props.macroRecording ?? null);
  }, [props.macroRecording]);

  // Children are rendered outside the host element since the spreadsheet
  // owns the host's children (`replaceChildren` on mount). When children is
  // a render prop, defer execution until the instance is available.
  const children =
    typeof props.children === 'function'
      ? liveInstance
        ? props.children(liveInstance)
        : null
      : props.children;
  const errorFallback =
    mountError && props.errorFallback
      ? typeof props.errorFallback === 'function'
        ? props.errorFallback(mountError)
        : props.errorFallback
      : null;

  return (
    <>
      <div ref={hostRef} className={props.className} style={props.style} />
      {errorFallback}
      {children}
    </>
  );
};

export const Spreadsheet = forwardRef<SpreadsheetRef, SpreadsheetProps>(SpreadsheetComponent);
Spreadsheet.displayName = 'Spreadsheet';
