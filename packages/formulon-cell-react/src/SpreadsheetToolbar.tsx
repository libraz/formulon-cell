// React adapter on top of `Spreadsheet.mountToolbar`. The wrapper owns just
// three concerns: mounting / disposing the core toolbar against a host div,
// keeping `activeTab` reactive in both directions, and forwarding the
// optional review/automation/drawing/file callbacks as hooks.
//
// Everything else — DOM, helpers, menus, hook defaults — is provided by core
// so this file stays small. The legacy SpreadsheetToolbar (~3.9k LOC of
// React-native ribbon UI) was retired in Phase 3-b; consumers that depended
// on internal class names or sub-component exports should migrate to the
// core's `data-ribbon-*` attributes for selection.
import {
  type DynamicDropdownsCtx,
  Spreadsheet,
  type SpreadsheetInstance,
  type ToolbarInstance,
} from '@libraz/formulon-cell';
import { type ReactElement, useEffect, useRef } from 'react';

import type { SpreadsheetToolbarProps } from './toolbar/model.js';

export type { RibbonTab, SpreadsheetToolbarProps } from './toolbar/model.js';

type CallbackBag = {
  onSpellingReview?: () => void;
  onAccessibilityCheck?: () => void;
  onRunScript?: () => void;
  onDrawPen?: () => void;
  onDrawEraser?: () => void;
  onTranslate?: () => void;
  onAddIn?: () => void;
  onToolbarReady?: (toolbar: ToolbarInstance | null) => void;
  onTabChange: (tab: import('./toolbar/model.js').RibbonTab) => void;
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
  onToolbarReady,
  dropdownActions,
  ribbonTabs,
}: SpreadsheetToolbarProps): ReactElement => {
  const hostRef = useRef<HTMLDivElement | null>(null);
  const toolbarRef = useRef<ToolbarInstance | null>(null);

  // Keep the latest callbacks in a ref so the mount effect only re-runs on
  // instance / locale change. Hook implementations close over the ref so a
  // late-mutated callback still fires on the next click.
  const callbacksRef = useRef<CallbackBag>({
    onTabChange,
    onSpellingReview,
    onAccessibilityCheck,
    onRunScript,
    onDrawPen,
    onDrawEraser,
    onTranslate,
    onAddIn,
    onToolbarReady,
  });
  callbacksRef.current = {
    onTabChange,
    onSpellingReview,
    onAccessibilityCheck,
    onRunScript,
    onDrawPen,
    onDrawEraser,
    onTranslate,
    onAddIn,
    onToolbarReady,
  };

  // biome-ignore lint/correctness/useExhaustiveDependencies: activeTab handled by the tab-switch effect below; including it would re-mount the toolbar on every tab click
  useEffect(() => {
    const host = hostRef.current;
    if (!host || !instance) return undefined;
    // Forward host-supplied callbacks onto the matching `scriptAction` /
    // `addInAction` menu items so a plain click on the Script / AddIn ribbon
    // button opens its menu (built-in UX) and the host hook fires only when
    // the user picks the action that maps to its prop. Other menu items keep
    // the default no-op until a future PR extends the prop surface.
    // Host-supplied `dropdownActions` win over the wrapper's built-in
    // script/addIn wiring — consumers can fully replace those if they want.
    // Pass a memoized object to avoid re-mounting the toolbar on every render.
    const dropdownOverrides: Partial<DynamicDropdownsCtx> = {
      applyScriptAction: (action) => {
        if (action === 'custom') callbacksRef.current.onRunScript?.();
      },
      applyAddInAction: (action) => {
        if (action === 'manage') callbacksRef.current.onAddIn?.();
      },
      ...dropdownActions,
    };
    const tb = Spreadsheet.mountToolbar(host as HTMLElement, instance as SpreadsheetInstance, {
      lang: locale === 'en' ? 'en' : 'ja',
      activeTab,
      ribbonTabs,
      onTabChange: (tab) => callbacksRef.current.onTabChange(tab),
      // Opt into core's default dropdown-menu click delegator so Fill / Clear
      // / AutoSum / etc. work without each consumer reimplementing the
      // playground's `createDynamicDropdowns` wiring.
      dynamicDropdowns: dropdownOverrides,
      hooks: {
        review: {
          spelling: () => callbacksRef.current.onSpellingReview?.(),
          accessibility: () => callbacksRef.current.onAccessibilityCheck?.(),
          translate: () => callbacksRef.current.onTranslate?.(),
        },
        drawing: {
          setInkMode: (mode) => {
            if (mode === 'pen') callbacksRef.current.onDrawPen?.();
            else callbacksRef.current.onDrawEraser?.();
          },
        },
      },
    });
    toolbarRef.current = tb;
    callbacksRef.current.onToolbarReady?.(tb);
    return () => {
      callbacksRef.current.onToolbarReady?.(null);
      tb.dispose();
      toolbarRef.current = null;
    };
  }, [instance, locale, dropdownActions, ribbonTabs]);

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
