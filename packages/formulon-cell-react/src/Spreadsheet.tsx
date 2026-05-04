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
import { useEffect, useImperativeHandle, useRef } from 'react';
import type { CSSProperties, ForwardedRef, ReactNode } from 'react';
import { forwardRef } from 'react';

export interface SpreadsheetProps {
  /** Optional pre-loaded workbook (e.g. from xlsx bytes). When omitted, a
   *  fresh default workbook is created on mount. */
  workbook?: WorkbookHandle;
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
  /** Fires once after mount with the live instance. Use this to wire toolbars
   *  / menus that talk to the spreadsheet's imperative API. */
  onReady?: (instance: SpreadsheetInstance) => void;
  /** Fires every time a cell value changes (engine-side). Use this to
   *  mirror the spreadsheet into outer state (Redux, Zustand, server). */
  onCellChange?: (e: CellChangeEvent) => void;
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

const SpreadsheetComponent = (
  props: SpreadsheetProps,
  ref: ForwardedRef<SpreadsheetRef>,
): ReactNode => {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const instanceRef = useRef<SpreadsheetInstance | null>(null);
  // Keep the latest props in a ref so the mount effect doesn't have to
  // re-run when callbacks change. Mounting is expensive (it creates the
  // wb + canvas + listeners) so we only re-mount on workbook identity.
  const propsRef = useRef(props);
  propsRef.current = props;

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
        ...(cur.theme ? { theme: cur.theme } : {}),
        ...(cur.locale ? { locale: cur.locale } : {}),
        ...(cur.strings ? { strings: cur.strings } : {}),
        ...(cur.features ? { features: cur.features } : {}),
        ...(cur.extensions ? { extensions: cur.extensions } : {}),
        ...(cur.functions ? { functions: cur.functions } : {}),
        ...(cur.seed ? { seed: cur.seed } : {}),
      });
      if (disposed) {
        inst.dispose();
        return;
      }
      instanceRef.current = inst;
      // Wire event props through `inst.on(...)`. Each handler reads from
      // propsRef, so callers can swap callbacks without re-mounting.
      eventDisposers.push(
        inst.on('cellChange', (e) => propsRef.current.onCellChange?.(e)),
        inst.on('selectionChange', (e) => propsRef.current.onSelectionChange?.(e)),
        inst.on('workbookChange', (e) => propsRef.current.onWorkbookChange?.(e)),
        inst.on('localeChange', (e) => propsRef.current.onLocaleChange?.(e)),
        inst.on('themeChange', (e) => propsRef.current.onThemeChange?.(e)),
        inst.on('recalc', (e) => propsRef.current.onRecalc?.(e)),
      );
      cur.onReady?.(inst);
    })();
    return () => {
      disposed = true;
      for (const d of eventDisposers) d();
      instanceRef.current?.dispose();
      instanceRef.current = null;
    };
    // Only re-mount on workbook identity change — props mutations land via
    // imperative methods on `instance` (setTheme, i18n.setLocale, …) and
    // event handlers re-read from propsRef on each fire.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.workbook]);

  // Forward reactive theme / locale / features / extensions prop changes
  // to the running instance via imperative APIs — avoids re-mounting.
  useEffect(() => {
    const inst = instanceRef.current;
    if (!inst || !props.theme) return;
    inst.setTheme(props.theme);
  }, [props.theme]);

  useEffect(() => {
    const inst = instanceRef.current;
    if (!inst || !props.locale) return;
    inst.i18n.setLocale(props.locale);
  }, [props.locale]);

  useEffect(() => {
    const inst = instanceRef.current;
    if (!inst) return;
    inst.setFeatures(props.features ?? {});
  }, [props.features]);

  useEffect(() => {
    const inst = instanceRef.current;
    if (!inst) return;
    inst.setExtensions(props.extensions);
  }, [props.extensions]);

  // Children are rendered outside the host element since the spreadsheet
  // owns the host's children (`replaceChildren` on mount). When children is
  // a render prop, defer execution until the instance is available.
  const children =
    typeof props.children === 'function' ? props.children(instanceRef.current!) : props.children;

  return (
    <>
      <div ref={hostRef} className={props.className} style={props.style} />
      {children}
    </>
  );
};

export const Spreadsheet = forwardRef<SpreadsheetRef, SpreadsheetProps>(SpreadsheetComponent);
Spreadsheet.displayName = 'Spreadsheet';
