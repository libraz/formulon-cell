import type { SpreadsheetInstance } from '@libraz/formulon-cell';
import {
  type ReactElement,
  useCallback,
  useEffect,
  useRef,
  useState,
  useSyncExternalStore,
} from 'react';
import {
  DEFAULT_DEMO_QUICK_ACCESS,
  DEMO_QUICK_ACCESS_GLYPHS,
  type DemoQuickAccessDeps,
  type DemoQuickAccessId,
  type DemoUiStrings,
  demoQuickAccessEnabled,
  demoQuickAccessMenu,
  runDemoQuickAccess,
  toggleDemoQuickAccess,
} from '../../demo-shared/index.js';
import { DemoGlyph } from './DemoGlyph.js';

interface DemoQuickAccessProps {
  ui: DemoUiStrings;
  instance: SpreadsheetInstance | null;
  deps: DemoQuickAccessDeps;
}

const noHistory = (): (() => void) => () => {};

/** The Mac title bar's Quick Access commands and their customize menu. */
export const DemoQuickAccess = ({ ui, instance, deps }: DemoQuickAccessProps): ReactElement => {
  const [shown, setShown] = useState<readonly DemoQuickAccessId[]>(DEFAULT_DEMO_QUICK_ACCESS);
  const [autosave, setAutosave] = useState(false);
  const [menuOpen, setMenuOpen] = useState(false);
  const menuRef = useRef<HTMLDivElement | null>(null);
  const chevronRef = useRef<HTMLButtonElement | null>(null);
  // Re-render when the history changes so Undo enables and disables with it.
  useSyncExternalStore(
    instance ? (notify) => instance.history.on('change', notify) : noHistory,
    () => instance?.history.canUndo ?? false,
  );

  const closeMenu = useCallback(() => setMenuOpen(false), []);
  useEffect(() => {
    if (!menuOpen) return;
    const onPointer = (event: PointerEvent): void => {
      const target = event.target as Node;
      if (!menuRef.current?.contains(target) && !chevronRef.current?.contains(target)) closeMenu();
    };
    const onKey = (event: KeyboardEvent): void => {
      if (event.key !== 'Escape') return;
      closeMenu();
      chevronRef.current?.focus();
    };
    document.addEventListener('pointerdown', onPointer, true);
    document.addEventListener('keydown', onKey, true);
    return () => {
      document.removeEventListener('pointerdown', onPointer, true);
      document.removeEventListener('keydown', onKey, true);
    };
  }, [closeMenu, menuOpen]);

  return (
    <>
      {shown.map((id) =>
        id === 'autosave' ? (
          <label key={id} className="demo__autosave">
            <span>{ui.quickAccess.autosave}</span>
            <input
              type="checkbox"
              role="switch"
              checked={autosave}
              onChange={(e) => setAutosave(e.currentTarget.checked)}
            />
          </label>
        ) : (
          <button
            key={id}
            type="button"
            className={`demo__title-icon demo__qat-${id}`}
            aria-label={ui.quickAccess.buttons[id]}
            title={ui.quickAccess.buttons[id]}
            disabled={!demoQuickAccessEnabled(id, instance)}
            onClick={() => runDemoQuickAccess(id, deps)}
          >
            {DEMO_QUICK_ACCESS_GLYPHS[id] ? (
              <DemoGlyph name={DEMO_QUICK_ACCESS_GLYPHS[id]} />
            ) : null}
            {id === 'undo' ? <span className="demo__qat-caret" aria-hidden="true" /> : null}
          </button>
        ),
      )}
      <div className="demo__qat-customize">
        <button
          ref={chevronRef}
          type="button"
          className="demo__title-icon demo__qat-more"
          aria-label={ui.quickAccess.customize}
          title={ui.quickAccess.customize}
          aria-haspopup="menu"
          aria-expanded={menuOpen}
          onClick={() => setMenuOpen((open) => !open)}
        >
          <DemoGlyph name="moreCommands" />
        </button>
        {menuOpen ? (
          <div
            ref={menuRef}
            className="demo__qat-menu"
            role="menu"
            aria-label={ui.quickAccess.customize}
          >
            <div className="demo__qat-menu-heading" role="presentation">
              {ui.quickAccess.customize}
            </div>
            {demoQuickAccessMenu(ui, shown).map((item) => (
              <button
                key={item.id}
                type="button"
                role="menuitemcheckbox"
                aria-checked={item.checked}
                className="demo__qat-menu-item"
                onClick={() => {
                  setShown((current) => toggleDemoQuickAccess(current, item.id));
                  closeMenu();
                }}
              >
                {item.label}
              </button>
            ))}
            <div className="demo__qat-menu-separator" role="separator" />
            <button type="button" role="menuitem" className="demo__qat-menu-item" disabled>
              {ui.quickAccess.moreCommands}
            </button>
          </div>
        ) : null}
      </div>
    </>
  );
};
