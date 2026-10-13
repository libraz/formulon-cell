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
  createDefaultDynamicDropdownsCtx,
  type DynamicDropdownsCtx,
  dictionaryLocaleFor,
  openHostMenuFirstDropdown,
  type RibbonHooks,
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
  onError?: (error: unknown) => void;
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
    onError,
    onToolbarReady,
  };

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
    // A provided prop handles only its own menu item (`custom` for scripts,
    // `manage` for add-ins); every other action, and every action when the
    // prop is omitted, reaches core's default handler. Likewise review hooks
    // are registered only for provided props so core's built-in reports stay.
    // Host-supplied `dropdownActions` win over the wrapper's wiring.
    // Pass a memoized object to avoid re-mounting the toolbar on every render.
    let coreDefaults: DynamicDropdownsCtx | null = null;
    const getCoreDefaults = (): DynamicDropdownsCtx => {
      coreDefaults ??= createDefaultDynamicDropdownsCtx(instance as SpreadsheetInstance);
      return coreDefaults;
    };
    const dropdownOverrides: Partial<DynamicDropdownsCtx> = {};
    if (hasRunScript) {
      dropdownOverrides.applyScriptAction = (action) =>
        action === 'custom'
          ? callbacksRef.current.onRunScript?.()
          : getCoreDefaults().applyScriptAction(action);
    }
    if (hasAddIn) {
      dropdownOverrides.applyAddInAction = (action) =>
        action === 'manage'
          ? callbacksRef.current.onAddIn?.()
          : getCoreDefaults().applyAddInAction(action);
    }
    Object.assign(dropdownOverrides, dropdownActions);
    const review: NonNullable<RibbonHooks['review']> = {};
    if (hasSpelling) review.spelling = () => callbacksRef.current.onSpellingReview?.();
    if (hasAccessibility)
      review.accessibility = () => callbacksRef.current.onAccessibilityCheck?.();
    if (hasTranslate) review.translate = () => callbacksRef.current.onTranslate?.();
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
        dynamicDropdowns: dropdownOverrides,
        hooks: {
          review,
          drawing: {
            setInkMode: (mode) => {
              if (mode === 'pen') callbacksRef.current.onDrawPen?.();
              else callbacksRef.current.onDrawEraser?.();
            },
          },
        },
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
