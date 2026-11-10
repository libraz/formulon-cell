import type { SpreadsheetInstance } from '@libraz/formulon-cell';
import { sameProp } from '@libraz/formulon-cell/adapter';
import {
  mountToolbar,
  type RibbonHooks,
  type RibbonTab,
  type Toolbar,
} from '@libraz/formulon-cell/toolbar';
import { type ReactElement, useEffect, useRef } from 'react';
import { forwardRibbonHooks, ribbonHooksShape } from './ribbon-hooks.js';

/** Props of {@link SpreadsheetToolbar}. */
export interface SpreadsheetToolbarProps {
  /** The spreadsheet the ribbon drives; nothing is mounted while it is `null`. */
  spreadsheet: SpreadsheetInstance | null;
  /** The selected tab, controlled; follow `onTabChange` to keep it in step. */
  activeTab?: RibbonTab;
  /** The tab selected at mount when `activeTab` is omitted. */
  defaultActiveTab?: RibbonTab;
  onTabChange?: (tab: RibbonTab) => void;
  /** Host actions for the commands the ribbon cannot run on its own. Only adding or removing a
   *  hook remounts the ribbon; a new function identity is picked up in place. */
  hooks?: RibbonHooks;
  /** The tab surface; defaults to the platform's standard tabs. */
  ribbonTabs?: readonly RibbonTab[];
  /** Receives the mounted toolbar, and `null` once it is disposed. */
  onReady?: (toolbar: Toolbar | null) => void;
  /** Receives a failed toolbar mount. */
  onError?: (error: unknown) => void;
}

/** Returns the previous value while `value` is structurally the same. */
const useSameProp = <T,>(value: T): T => {
  const ref = useRef(value);
  if (!sameProp(ref.current, value)) ref.current = value;
  return ref.current;
};

/** The ribbon of `spreadsheet`, placed apart from the grid. */
export const SpreadsheetToolbar = (props: SpreadsheetToolbarProps): ReactElement => {
  const { spreadsheet, activeTab } = props;
  const hostRef = useRef<HTMLDivElement | null>(null);
  const toolbarRef = useRef<Toolbar | null>(null);
  const reportedTabRef = useRef<RibbonTab | undefined>(undefined);
  const propsRef = useRef(props);
  propsRef.current = props;
  const ribbonTabs = useSameProp(props.ribbonTabs);
  const hooksShape = ribbonHooksShape(props.hooks);

  // biome-ignore lint/correctness/useExhaustiveDependencies: hooksShape stands for the hooks present; the hooks themselves are read at call time
  useEffect(() => {
    const host = hostRef.current;
    if (!host || !spreadsheet) return undefined;
    let disposed = false;
    let mounted: Toolbar | null = null;
    const initialTab = propsRef.current.activeTab ?? propsRef.current.defaultActiveTab;
    mountToolbar(host, spreadsheet, {
      hooks: forwardRibbonHooks(() => propsRef.current.hooks),
      ...(ribbonTabs ? { ribbonTabs } : {}),
      ...(initialTab ? { activeTab: initialTab } : {}),
    }).then(
      (toolbar) => {
        if (disposed) {
          toolbar.dispose();
          return;
        }
        mounted = toolbar;
        toolbarRef.current = toolbar;
        reportedTabRef.current = toolbar.activeTab;
        toolbar.on('change', (state) => {
          if (state.activeTab === reportedTabRef.current) return;
          reportedTabRef.current = state.activeTab;
          propsRef.current.onTabChange?.(state.activeTab);
        });
        propsRef.current.onReady?.(toolbar);
      },
      (error: unknown) => {
        if (!disposed) propsRef.current.onError?.(error);
      },
    );
    return () => {
      disposed = true;
      if (!mounted) return;
      toolbarRef.current = null;
      mounted.dispose();
      propsRef.current.onReady?.(null);
    };
  }, [spreadsheet, ribbonTabs, hooksShape]);

  useEffect(() => {
    const toolbar = toolbarRef.current;
    if (!toolbar || activeTab === undefined || toolbar.activeTab === activeTab) return;
    reportedTabRef.current = activeTab;
    toolbar.update({ activeTab });
  }, [activeTab]);

  // `display: contents` keeps the wrapper out of the flex chain the ribbon shell sits in.
  return <div ref={hostRef} style={{ display: 'contents' }} />;
};
