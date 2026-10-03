import type { RibbonTab, SpreadsheetInstance } from '@libraz/formulon-cell';
import {
  type ReactElement,
  type RefObject,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import {
  type DemoPlatform,
  type DemoSearchItem,
  type DemoSearchUsagePrior,
  type DemoUiStrings,
  demoSearchOptionId,
  installDemoSearchShortcut,
  loadDemoSearchUsagePrior,
  queryDemoSearchItems,
  recordDemoSearchUsage,
  resolveDemoSearchKey,
  saveDemoSearchUsagePrior,
} from '../../demo-shared/index.js';
import { DemoIcon } from './DemoIcon.js';

interface DemoTitleBarProps {
  ui: DemoUiStrings;
  platform: DemoPlatform;
  locale: string;
  bookName: string;
  instance: SpreadsheetInstance | null;
  showPanel: boolean;
  searchItems: readonly DemoSearchItem[];
  /** Attached to the Quick Access toolbar so the host can route F6 into it. */
  quickAccessRef: RefObject<HTMLDivElement | null>;
  onSave: () => void;
  onOpenBackstage: () => void;
  onTogglePanel: () => void;
  onTabChange: (tab: RibbonTab) => void;
}

export const DemoTitleBar = ({
  ui,
  platform,
  locale,
  bookName,
  instance,
  showPanel,
  searchItems,
  quickAccessRef,
  onSave,
  onOpenBackstage,
  onTogglePanel,
  onTabChange,
}: DemoTitleBarProps): ReactElement => {
  const [searchQuery, setSearchQuery] = useState('');
  const [searchOpen, setSearchOpen] = useState(false);
  const [searchActiveIndex, setSearchActiveIndex] = useState(-1);
  const [searchUsagePrior, setSearchUsagePrior] = useState<DemoSearchUsagePrior>(() =>
    loadDemoSearchUsagePrior(),
  );
  const searchInputRef = useRef<HTMLInputElement | null>(null);

  useEffect(() => installDemoSearchShortcut(() => searchInputRef.current, platform), [platform]);
  useEffect(() => saveDemoSearchUsagePrior(searchUsagePrior), [searchUsagePrior]);

  const filteredCommands = useMemo(() => {
    return queryDemoSearchItems(searchItems, searchQuery, 8, searchUsagePrior);
  }, [searchItems, searchQuery, searchUsagePrior]);

  const runCommand = useCallback(
    (cmd: DemoSearchItem) => {
      setSearchUsagePrior((prior) => recordDemoSearchUsage(prior, cmd));
      if (cmd.tab) onTabChange(cmd.tab);
      cmd.run();
      setSearchQuery('');
      setSearchOpen(false);
      setSearchActiveIndex(-1);
    },
    [onTabChange],
  );

  return (
    <header className="demo__head">
      <div className="fc-tb__titlebar">
        <div
          ref={quickAccessRef}
          className="demo__quick"
          role="toolbar"
          aria-label={ui.quickAccessToolbar}
        >
          {platform === 'mac' ? (
            <button
              type="button"
              className="demo__brand-mark"
              aria-label={ui.file}
              title={ui.file}
              onClick={onOpenBackstage}
            >
              <DemoIcon name="app" />
            </button>
          ) : (
            <span className="demo__brand-mark" aria-hidden="true">
              <DemoIcon name="app" />
            </span>
          )}
          <button type="button" className="demo__title-icon" aria-label={ui.save} onClick={onSave}>
            <DemoIcon name="save" />
          </button>
          <button
            type="button"
            className="demo__title-icon"
            aria-label={ui.undo}
            onClick={() => instance?.undo()}
          >
            <DemoIcon name="undo" />
          </button>
          <button
            type="button"
            className="demo__title-icon"
            aria-label={ui.redo}
            onClick={() => instance?.redo()}
          >
            <DemoIcon name="redo" />
          </button>
        </div>
        <div className="fc-tb__title">
          <strong>{bookName}</strong>
          <span>{ui.saved}</span>
        </div>
        <div className="fc-tb__search">
          <DemoIcon name="search" />
          <input
            ref={searchInputRef}
            type="search"
            role="combobox"
            placeholder={ui.search}
            aria-label={
              platform === 'mac' ? (locale === 'ja' ? '検索' : 'Search') : ui.searchCommands
            }
            aria-controls="demo-search-results"
            aria-expanded={searchOpen}
            aria-activedescendant={
              searchOpen && searchActiveIndex >= 0
                ? demoSearchOptionId(searchActiveIndex)
                : undefined
            }
            value={searchQuery}
            onFocus={() => {
              setSearchOpen(true);
              setSearchActiveIndex(-1);
            }}
            onChange={(e) => {
              const input = e.currentTarget;
              setSearchQuery(input.value);
              // Escape clears a search input natively, and that clear lands
              // as a change on an input we just blurred. Only a change the
              // user typed reopens the list.
              setSearchOpen(document.activeElement === input);
              setSearchActiveIndex(-1);
            }}
            onKeyDown={(e) => {
              const action = resolveDemoSearchKey(
                e.key,
                searchActiveIndex,
                filteredCommands.length,
              );
              if (action?.kind === 'close') {
                setSearchOpen(false);
                setSearchActiveIndex(-1);
                e.currentTarget.blur();
              } else if (action?.kind === 'move') {
                e.preventDefault();
                setSearchOpen(true);
                setSearchActiveIndex(action.index);
              } else if (action?.kind === 'run') {
                e.preventDefault();
                const command = filteredCommands[action.index];
                if (command) runCommand(command);
              }
            }}
            onBlur={() => setSearchOpen(false)}
          />
          {searchOpen ? (
            <div id="demo-search-results" className="fc-tb__command-menu" role="listbox">
              {filteredCommands.length === 0 ? (
                <div className="fc-tb__command-empty">{ui.noCommands}</div>
              ) : (
                filteredCommands.map((cmd, index) => (
                  <button
                    key={cmd.id}
                    id={demoSearchOptionId(index)}
                    type="button"
                    role="option"
                    aria-selected={index === searchActiveIndex}
                    aria-disabled={cmd.disabled ? 'true' : undefined}
                    data-disabled-reason={cmd.disabledReason}
                    className={`fc-tb__command-item${
                      index === searchActiveIndex ? ' fc-tb__command-item--active' : ''
                    }${cmd.disabled ? ' fc-tb__command-item--disabled' : ''}`}
                    onMouseDown={(e) => e.preventDefault()}
                    onMouseEnter={() => setSearchActiveIndex(index)}
                    onClick={() => runCommand(cmd)}
                  >
                    <strong>{cmd.label}</strong>
                    <span>{cmd.hint}</span>
                  </button>
                ))
              )}
            </div>
          ) : null}
        </div>
        <div className="demo__account">
          <button type="button" className="demo__share">
            {ui.share}
          </button>
          <button
            type="button"
            className={`demo__share${showPanel ? ' demo__share--active' : ''}`}
            onClick={onTogglePanel}
            aria-pressed={showPanel}
          >
            {ui.demoPane}
          </button>
          <span className="demo__avatar" role="img" aria-label={ui.signedInUser}>
            FC
          </span>
        </div>
      </div>
    </header>
  );
};
