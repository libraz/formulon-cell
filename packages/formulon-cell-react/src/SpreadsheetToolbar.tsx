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
import { Spreadsheet, type SpreadsheetInstance, type ToolbarInstance } from '@libraz/formulon-cell';
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
  };

  useEffect(() => {
    const host = hostRef.current;
    if (!host || !instance) return undefined;
    const tb = Spreadsheet.mountToolbar(host as HTMLElement, instance as SpreadsheetInstance, {
      lang: locale === 'en' ? 'en' : 'ja',
      activeTab,
      onTabChange: (tab) => callbacksRef.current.onTabChange(tab),
      // Opt into core's default dropdown-menu click delegator so Fill / Clear
      // / AutoSum / etc. work without each consumer reimplementing the
      // playground's `createDynamicDropdowns` wiring.
      dynamicDropdowns: true,
      hooks: {
        review: {
          spelling: () => callbacksRef.current.onSpellingReview?.(),
          accessibility: () => callbacksRef.current.onAccessibilityCheck?.(),
          translate: () => callbacksRef.current.onTranslate?.(),
        },
        automation: {
          runScript: () => callbacksRef.current.onRunScript?.(),
          addInManager: () => callbacksRef.current.onAddIn?.(),
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
    return () => {
      tb.dispose();
      toolbarRef.current = null;
    };
    // activeTab is intentionally NOT a dep: changing it should drive the
    // toolbar via the effect below, not re-mount the whole thing.
    // biome-ignore lint/correctness/useExhaustiveDependencies: see comment above
  }, [instance, locale]);

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
