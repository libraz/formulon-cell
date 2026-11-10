// React adapter on top of `Spreadsheet.mountToolbar`. The wrapper owns just
// three concerns: mounting / disposing the core toolbar against a host div,
// keeping `activeTab` reactive in both directions, and forwarding the
// optional review/automation/drawing/file callbacks as hooks.
//
// Everything else — DOM, helpers, menus, hook defaults — is provided by core
// so this file stays small. The ribbon is the core's DOM; select it through
// the core's `data-ribbon-*` attributes rather than internal class names.
import {
  buildAdapterRibbonWiring,
  dictionaryLocaleFor,
  openHostMenuFirstDropdown,
  Spreadsheet,
  type LegacySpreadsheetInstance as SpreadsheetInstance,
  type ToolbarInstance,
} from '@libraz/formulon-cell';
import { type ReactElement, useEffect, useRef } from 'react';

import type { RibbonTab, SpreadsheetToolbarProps } from './toolbar/model.js';

export type { RibbonTab, SpreadsheetToolbarProps } from './toolbar/model.js';

type CallbackBag = {
  onSpellingReview?: () => void;
  onAccessibilityCheck?: () => void;
  onRunScript?: () => void;
  onDrawPen?: () => void;
  onDrawEraser?: () => void;
  onTranslate?: () => void;
  onAddIn?: () => void;
  onError?: (error: unknown) => void;
  onToolbarReady?: (toolbar: ToolbarInstance | null) => void;
  onTabChange: (tab: RibbonTab) => void;
};

export const SpreadsheetToolbar = ({
  instance,
  activeTab,
  onTabChange,
  locale,
  onSpellingReview,
  onAccessibilityCheck,
  onRunScript,
  onDrawPen,
  onDrawEraser,
  onTranslate,
  onAddIn,
  onError,
  onToolbarReady,
  dropdownActions,
  ribbonTabs,
}: SpreadsheetToolbarProps): ReactElement => {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const toolbarRef = useRef<ToolbarInstance | null>(null);

  // Keep the latest callbacks in a ref so the mount effect only re-runs on
  // instance / locale change. Hook implementations close over the ref so a
  // late-mutated callback still fires on the next click.
  const callbacks: CallbackBag = {
    onTabChange,
    onSpellingReview,
    onAccessibilityCheck,
    onRunScript,
    onDrawPen,
    onDrawEraser,
    onTranslate,
    onAddIn,
    onError,
    onToolbarReady,
  };
  const callbacksRef = useRef<CallbackBag>(callbacks);
  callbacksRef.current = callbacks;

  // Presence (not identity) of these props decides whether core defaults are
  // overridden, so only a change in presence re-mounts the toolbar.
  const hasSpelling = Boolean(onSpellingReview);
  const hasAccessibility = Boolean(onAccessibilityCheck);
  const hasTranslate = Boolean(onTranslate);
  const hasRunScript = Boolean(onRunScript);
  const hasAddIn = Boolean(onAddIn);

  // biome-ignore lint/correctness/useExhaustiveDependencies: activeTab handled by the tab-switch effect below; including it would re-mount the toolbar on every tab click
  useEffect(() => {
    const host = hostRef.current;
    if (!host || !instance) return undefined;
    // Provided callbacks override only their own items; host `dropdownActions`
    // win. Pass a memoized object to avoid re-mounting on every render.
    const wiring = buildAdapterRibbonWiring(
      instance as SpreadsheetInstance,
      () => callbacksRef.current,
      dropdownActions,
    );
    let tb: ToolbarInstance;
    try {
      tb = Spreadsheet.mountToolbar(host as HTMLElement, instance as SpreadsheetInstance, {
        lang: dictionaryLocaleFor(locale),
        activeTab,
        ribbonTabs,
        onTabChange: (tab) => callbacksRef.current.onTabChange(tab),
        interceptCommand: (id, button) => openHostMenuFirstDropdown(toolbarRef.current, id, button),
        // Opt into core's default dropdown-menu click delegator so Fill / Clear
        // / AutoSum / etc. work without each consumer reimplementing the
        // playground's `createDynamicDropdowns` wiring.
        dynamicDropdowns: wiring.dynamicDropdowns,
        hooks: wiring.hooks,
      });
    } catch (error) {
      callbacksRef.current.onToolbarReady?.(null);
      callbacksRef.current.onError?.(error);
      return undefined;
    }
    toolbarRef.current = tb;
    callbacksRef.current.onToolbarReady?.(tb);
    return () => {
      callbacksRef.current.onToolbarReady?.(null);
      tb.dispose();
      toolbarRef.current = null;
    };
  }, [
    instance,
    locale,
    dropdownActions,
    ribbonTabs,
    hasSpelling,
    hasAccessibility,
    hasTranslate,
    hasRunScript,
    hasAddIn,
  ]);

  // Forward external tab changes into the toolbar without re-mounting.
  useEffect(() => {
    const tb = toolbarRef.current;
    if (!tb) return;
    if (tb.getActiveTab() !== activeTab) tb.setActiveTab(activeTab);
  }, [activeTab]);

  // `display: contents` keeps the wrapper out of the layout tree so the
  // core ribbon-shell (`flex: 0 0 auto`) sees the parent flex column
  // directly. Without this the extra div breaks the flex chain and the
  // sibling sheet element collapses to zero height.
  return <div ref={hostRef} style={{ display: 'contents' }} />;
};

export const Toolbar = SpreadsheetToolbar;
