// Controlled-prop reconcile policy shared by the React and Vue adapters, so
// the framework wrappers stay thin and cannot drift apart.

import type { PrinterProfile } from '../commands/printer-profile.js';
import type { WorkbookHandle } from '../engine/workbook-handle.js';
import type { LocaleChangeEvent } from '../events.js';
import type { FeatureFlags } from '../extensions/features.js';
import type { ExtensionInput } from '../extensions/types.js';
import {
  resolveSpreadsheetUiOptions,
  type SpreadsheetUiOptions,
} from '../extensions/ui-options.js';
import { DEFAULT_LOCALE } from '../i18n/locales.js';
import type { SelectionSlice } from '../store/types.js';
import type { RibbonHooks } from '../toolbar/ribbon/apply-ribbon-command.js';
import type { DynamicDropdownsCtx } from '../toolbar/ribbon/dynamic-dropdowns.js';
import { createDefaultDynamicDropdownsCtx } from './dynamic-dropdowns-defaults.js';
import type { MountOptions, SpreadsheetInstance } from './types.js';

/** Selection an adapter reports while no instance is mounted. */
export const ADAPTER_FALLBACK_SELECTION: SelectionSlice = {
  active: { sheet: 0, row: 0, col: 0 },
  anchor: { sheet: 0, row: 0, col: 0 },
  range: { sheet: 0, r0: 0, c0: 0, r1: 0, c1: 0 },
};

/** Locale an adapter reports while no instance is mounted. */
export const ADAPTER_FALLBACK_LOCALE = DEFAULT_LOCALE;

/** Props the adapters reconcile against a running instance. */
export interface AdapterControlledProps {
  workbook?: WorkbookHandle;
  policy?: MountOptions['policy'];
  viewport?: MountOptions['viewport'];
  contextMenu?: MountOptions['contextMenu'];
  overlays?: MountOptions['overlays'];
  ui?: SpreadsheetUiOptions;
  toolbar?: MountOptions['toolbar'];
  theme?: MountOptions['theme'];
  locale?: MountOptions['locale'];
  strings?: MountOptions['strings'];
  features?: FeatureFlags;
  extensions?: ExtensionInput[];
  printerProfiles?: readonly PrinterProfile[];
  printerProfileId?: string;
  uploadStatus?: MountOptions['uploadStatus'];
  macroRecording?: MountOptions['macroRecording'];
}

export interface AdapterReconcileHandlers {
  onLocaleChange?: (event: LocaleChangeEvent) => void;
  onError?: (error: unknown) => void;
  getLocaleEventGeneration?: () => number;
}

/** Apply a controlled locale. Returns false (after reporting the rejection
 *  through the handlers) when the instance refuses it. */
export const applyControlledLocale = (
  inst: SpreadsheetInstance,
  requested: string,
  handlers: AdapterReconcileHandlers = {},
): boolean => {
  const { onLocaleChange, onError, getLocaleEventGeneration } = handlers;
  const generationBefore = getLocaleEventGeneration?.();
  let accepted = false;
  let failure: unknown;
  try {
    accepted = inst.setLocale(requested);
  } catch (error) {
    failure = error;
  }
  if (accepted) return true;

  // Core normally emits this during its rollback; the explicit fallback covers
  // a host instance that rejects without emitting.
  if (generationBefore === undefined || getLocaleEventGeneration?.() === generationBefore) {
    onLocaleChange?.({ locale: inst.locale, strings: inst.i18n.strings });
  }
  onError?.(
    failure ??
      new Error(
        `formulon-cell: locale change to "${requested}" was rejected; active locale is "${inst.locale}"`,
      ),
  );
  return false;
};

/** Effective theme: the explicit `theme` prop, else the `ui` preset's theme. */
export const resolveAdapterTheme = (
  props: Pick<AdapterControlledProps, 'theme' | 'ui'>,
): MountOptions['theme'] =>
  props.theme ?? (props.ui ? resolveSpreadsheetUiOptions(props.ui).theme : undefined);

/** Reconcile ui / features / toolbar; `uiChanged` also re-applies the ui preset. */
export const applyUiFeaturesToolbar = (
  inst: SpreadsheetInstance,
  props: Pick<AdapterControlledProps, 'ui' | 'features' | 'toolbar'>,
  uiChanged: boolean,
): void => {
  const resolved = resolveSpreadsheetUiOptions(props.ui);
  if (uiChanged) inst.setUi(props.ui);
  inst.setFeatures({ ...resolved.features, ...(props.features ?? {}) });
  inst.setToolbar(props.toolbar ?? (props.ui ? resolved.ribbon : false));
};

/** Apply every prop that differs from `baseline` (props captured at mount
 *  time while the async mount was pending). */
export const applyControlledProps = async (
  inst: SpreadsheetInstance,
  props: AdapterControlledProps,
  baseline: AdapterControlledProps = {},
  handlers: AdapterReconcileHandlers = {},
): Promise<void> => {
  const resolvedUi = resolveSpreadsheetUiOptions(props.ui);
  const ui = props.ui ? resolvedUi : null;
  if (props.workbook && props.workbook !== baseline.workbook && props.workbook !== inst.workbook) {
    try {
      await inst.setWorkbook(props.workbook);
    } catch (error) {
      // Core keeps the previous workbook active on failure; report it and keep
      // reconciling the independent props.
      handlers.onError?.(error);
    }
  }
  if (props.policy !== baseline.policy) inst.setPolicy(props.policy);
  if (props.viewport !== baseline.viewport) inst.setViewportOptions(props.viewport);
  if (props.contextMenu !== baseline.contextMenu) inst.setContextMenu(props.contextMenu);
  if (props.overlays !== baseline.overlays) inst.setOverlayOptions(props.overlays);
  if (props.ui !== baseline.ui) inst.setUi(props.ui);
  const nextTheme = props.theme ?? ui?.theme;
  const baselineTheme = resolveAdapterTheme(baseline);
  if (nextTheme && nextTheme !== baselineTheme) inst.setTheme(nextTheme);
  const localeAccepted =
    props.locale && props.locale !== baseline.locale
      ? applyControlledLocale(inst, props.locale, handlers)
      : undefined;
  if (props.strings && (props.strings !== baseline.strings || localeAccepted === true)) {
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

/** Host callbacks the ribbon adapters forward. Read at call time so a
 *  late-swapped callback still fires. */
export interface AdapterRibbonCallbacks {
  onSpellingReview?: () => void;
  onAccessibilityCheck?: () => void;
  onRunScript?: () => void;
  onDrawPen?: () => void;
  onDrawEraser?: () => void;
  onTranslate?: () => void;
  onAddIn?: () => void;
}

/** Build the dropdown overrides and hooks for `mountToolbar`. A provided
 *  callback handles only its own item (`custom` script / `manage` add-in /
 *  its review hook); everything else reaches core's defaults. Host
 *  `dropdownActions` win over the wiring. */
export const buildAdapterRibbonWiring = (
  inst: SpreadsheetInstance,
  getCallbacks: () => AdapterRibbonCallbacks,
  dropdownActions?: Partial<DynamicDropdownsCtx>,
): { dynamicDropdowns: Partial<DynamicDropdownsCtx>; hooks: RibbonHooks } => {
  const present = getCallbacks();
  let coreDefaults: DynamicDropdownsCtx | null = null;
  const getCoreDefaults = (): DynamicDropdownsCtx => {
    coreDefaults ??= createDefaultDynamicDropdownsCtx(inst);
    return coreDefaults;
  };
  const dynamicDropdowns: Partial<DynamicDropdownsCtx> = {};
  if (present.onRunScript) {
    dynamicDropdowns.applyScriptAction = (action) =>
      action === 'custom'
        ? getCallbacks().onRunScript?.()
        : getCoreDefaults().applyScriptAction(action);
  }
  if (present.onAddIn) {
    dynamicDropdowns.applyAddInAction = (action) =>
      action === 'manage' ? getCallbacks().onAddIn?.() : getCoreDefaults().applyAddInAction(action);
  }
  Object.assign(dynamicDropdowns, dropdownActions);
  const review: NonNullable<RibbonHooks['review']> = {};
  if (present.onSpellingReview) review.spelling = () => getCallbacks().onSpellingReview?.();
  if (present.onAccessibilityCheck) {
    review.accessibility = () => getCallbacks().onAccessibilityCheck?.();
  }
  if (present.onTranslate) review.translate = () => getCallbacks().onTranslate?.();
  return {
    dynamicDropdowns,
    hooks: {
      review,
      drawing: {
        setInkMode: (mode) => {
          if (mode === 'pen') getCallbacks().onDrawPen?.();
          else getCallbacks().onDrawEraser?.();
        },
      },
    },
  };
};
