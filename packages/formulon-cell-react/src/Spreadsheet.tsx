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
import type { CSSProperties, ForwardedRef, ReactNode } from 'react';
import { forwardRef, useEffect, useImperativeHandle, useRef, useState } from 'react';

type LocaleChangeEvent = Parameters<Parameters<I18n['on']>[1]>[0];

/**
 * Props of {@link Spreadsheet}: every `SpreadsheetOptions` key, the instance events as `on*`
 * callbacks, and the React-only rendering props. `seed` and `signal` are read at mount; every other
 * option is applied to the live instance with `update` when it changes structurally, so an inline
 * object literal does not remount.
 */
export interface SpreadsheetProps extends SpreadsheetOptions {
  /** Fires once the spreadsheet is mounted. */
  onReady?: (spreadsheet: SpreadsheetInstance) => void;
  onChange?: (event: ChangeEvent) => void;
  onCell?: (event: CellEvent) => void;
  onWorkbook?: (event: WorkbookEvent) => void;
  onRecalc?: (event: RecalcEvent) => void;
  onUpdate?: (event: UpdateEvent) => void;
  onSelectionChange?: (selection: SelectionState) => void;
  onLocaleChange?: (event: LocaleChangeEvent) => void;
  /** A mount rejection, an option `update` refused, or an instance `error` event's `error`. */
  onError?: (error: unknown) => void;
  /** Rendered in place of the grid when the mount rejects. */
  errorFallback?: ReactNode | ((error: unknown) => ReactNode);
  className?: string;
  style?: CSSProperties;
  /** Rendered after the host; a function receives the spreadsheet once it is mounted. */
  children?: ReactNode | ((spreadsheet: SpreadsheetInstance) => ReactNode);
}

/** What a ref to {@link Spreadsheet} holds. */
export interface SpreadsheetRef {
  /** The mounted spreadsheet; `null` before it is ready and after unmount. */
  readonly spreadsheet: SpreadsheetInstance | null;
}

const SpreadsheetComponent = (
  props: SpreadsheetProps,
  ref: ForwardedRef<SpreadsheetRef>,
): ReactNode => {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const instanceRef = useRef<SpreadsheetInstance | null>(null);
  // The options the live instance was last brought to.
  const appliedRef = useRef<Partial<SpreadsheetPatch>>({});
  const [live, setLive] = useState<SpreadsheetInstance | null>(null);
  const [mountError, setMountError] = useState<{ error: unknown } | null>(null);
  // Callbacks and options are read through this ref, so neither remounts the spreadsheet.
  const propsRef = useRef(props);
  propsRef.current = props;

  useImperativeHandle(
    ref,
    () => ({
      get spreadsheet() {
        return instanceRef.current;
      },
    }),
    [],
  );

  useEffect(() => {
    const host = hostRef.current;
    if (!host) return undefined;
    let disposed = false;
    const report = (error: unknown): void => {
      if (!disposed) propsRef.current.onError?.(error);
    };
    const options = optionsFromProps(propsRef.current);
    const offs: (() => void)[] = [];
    mount(host, options).then(
      (inst) => {
        if (disposed) {
          inst.dispose();
          return;
        }
        const p = (): SpreadsheetProps => propsRef.current;
        offs.push(
          inst.on('change', (e) => p().onChange?.(e)),
          inst.on('cell', (e) => p().onCell?.(e)),
          inst.on('workbook', (e) => p().onWorkbook?.(e)),
          inst.on('recalc', (e) => p().onRecalc?.(e)),
          inst.on('update', (e) => p().onUpdate?.(e)),
          inst.on('error', (e) => report(e.error)),
          inst.selection.on('change', (s) => p().onSelectionChange?.(s)),
          inst.i18n.on('change', (e) => p().onLocaleChange?.(e)),
        );
        instanceRef.current = inst;
        // Props that changed while the mount was pending.
        const latest = patchFromOptions(optionsFromProps(p()));
        reconcileProps(inst, patchFromOptions(options), latest, report);
        appliedRef.current = latest;
        setLive(inst);
        p().onReady?.(inst);
      },
      (error: unknown) => {
        if (disposed) return;
        setMountError({ error });
        report(error);
      },
    );
    return () => {
      disposed = true;
      for (const off of offs) off();
      const inst = instanceRef.current;
      instanceRef.current = null;
      setLive(null);
      inst?.dispose();
    };
  }, []);

  useEffect(() => {
    const inst = instanceRef.current;
    if (!inst) return;
    const next = patchFromOptions(optionsFromProps(props));
    const previous = appliedRef.current;
    appliedRef.current = next;
    reconcileProps(inst, previous, next, (error) => propsRef.current.onError?.(error));
  });

  const { children, errorFallback } = props;
  const fallback =
    mountError && errorFallback
      ? typeof errorFallback === 'function'
        ? errorFallback(mountError.error)
        : errorFallback
      : null;
  return (
    <>
      <div ref={hostRef} className={props.className} style={props.style} />
      {fallback}
      {typeof children === 'function' ? (live ? children(live) : null) : children}
    </>
  );
};

/**
 * The spreadsheet as a React component. It mounts once; option props are applied to the live
 * instance with `update`, and the instance is disposed on unmount.
 */
export const Spreadsheet = forwardRef<SpreadsheetRef, SpreadsheetProps>(SpreadsheetComponent);
Spreadsheet.displayName = 'Spreadsheet';
