import type {
  RibbonTab,
  LegacySpreadsheetInstance as SpreadsheetInstance,
} from '@libraz/formulon-cell';
import {
  type ReactElement,
  type RefObject,
  useCallback,
  useEffect,
  useMemo,
  useReducer,
  useRef,
  useState,
} from 'react';
import {
  buildDemoSearchSuggestions,
  type DemoPlatform,
  type DemoSearchItem,
  type DemoSearchUsagePrior,
  type DemoUiStrings,
  demoSearchOptionId,
  INITIAL_DEMO_SEARCH_STATE,
  installDemoSearchShortcut,
  loadDemoSearchUsagePrior,
  queryDemoSearchItems,
  recordDemoSearchUsage,
  reduceDemoSearch,
  resolveDemoSearchKey,
  saveDemoSearchUsagePrior,
} from '../../demo-shared/index.js';
import { DemoDocumentMenu } from './DemoDocumentMenu.js';
import { DemoGlyph } from './DemoGlyph.js';
import { DemoIcon } from './DemoIcon.js';
import { DemoQuickAccess } from './DemoQuickAccess.js';

interface DemoTitleBarProps {
  ui: DemoUiStrings;
  platform: DemoPlatform;
  bookName: string;
  instance: SpreadsheetInstance | null;
  showPanel: boolean;
  searchItems: readonly DemoSearchItem[];
  /** Attached to the Quick Access toolbar so the host can route F6 into it. */
  quickAccessRef: RefObject<HTMLDivElement | null>;
  onSave: () => void;
  onSaveAs: () => void;
  onRename: (name: string) => void;
  onNewWorkbook: () => void;
  /** Runs a ribbon command id through the toolbar; false when none matched. */
  applyCommand: (commandId: string) => boolean;
  onOpenBackstage: () => void;
  onTogglePanel: () => void;
  onTabChange: (tab: RibbonTab) => void;
}

export const DemoTitleBar = ({
  ui,
  platform,
  bookName,
  instance,
  showPanel,
  searchItems,
  quickAccessRef,
  onSave,
  onSaveAs,
  onRename,
  onNewWorkbook,
  applyCommand,
  onOpenBackstage,
  onTogglePanel,
  onTabChange,
}: DemoTitleBarProps): ReactElement => {
  const [search, dispatchSearch] = useReducer(reduceDemoSearch, INITIAL_DEMO_SEARCH_STATE);
  const { query: searchQuery, open: searchOpen, activeIndex: searchActiveIndex } = search;
  const [searchUsagePrior, setSearchUsagePrior] = useState<DemoSearchUsagePrior>(() =>
    loadDemoSearchUsagePrior(),
  );
  const searchInputRef = useRef<HTMLInputElement | null>(null);
  const mac = platform === 'mac';
  // The Mac title bar shows a magnifier button until the search is opened.
  const [searchActive, setSearchActive] = useState(false);
  const openSearch = useCallback(() => setSearchActive(true), []);
  useEffect(() => {
    if (searchActive) searchInputRef.current?.focus();
  }, [searchActive]);

  useEffect(
    () =>
      installDemoSearchShortcut(
        () => searchInputRef.current,
        platform,
        mac ? openSearch : undefined,
      ),
    [mac, openSearch, platform],
  );
  useEffect(() => saveDemoSearchUsagePrior(searchUsagePrior), [searchUsagePrior]);

  const filteredCommands = useMemo(() => {
    return queryDemoSearchItems(searchItems, searchQuery, 8, searchUsagePrior);
  }, [searchItems, searchQuery, searchUsagePrior]);

  const suggestionRows = useMemo(
    () =>
      mac
        ? buildDemoSearchSuggestions(ui, { instance, setRibbonTab: onTabChange, applyCommand })
        : [],
    [applyCommand, instance, mac, onTabChange, ui],
  );
  const suggestions = useMemo(
    () => suggestionRows.flatMap((section) => section.items),
    [suggestionRows],
  );
  // An empty Mac query lists the suggestions; a typed one lists the matches.
  const showSuggestions = mac && searchQuery.trim() === '';
  const rowCount = showSuggestions ? suggestions.length : filteredCommands.length;

  const closeSearch = useCallback(() => {
    dispatchSearch({ type: 'reset' });
    setSearchActive(false);
  }, []);

  const runCommand = useCallback(
    (cmd: DemoSearchItem) => {
      setSearchUsagePrior((prior) => recordDemoSearchUsage(prior, cmd));
      if (cmd.tab) onTabChange(cmd.tab);
      cmd.run();
      dispatchSearch({ type: 'reset' });
    },
    [onTabChange],
  );

  const runRow = useCallback(
    (index: number) => {
      if (!showSuggestions) {
        const command = filteredCommands[index];
        if (command) runCommand(command);
        return;
      }
      const suggestion = suggestions[index];
      if (!suggestion || suggestion.disabled) return;
      closeSearch();
      suggestion.run();
    },
    [closeSearch, filteredCommands, runCommand, showSuggestions, suggestions],
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
          {mac ? (
            <DemoQuickAccess
              ui={ui}
              instance={instance}
              deps={{
                instance,
                newWorkbook: onNewWorkbook,
                saveWorkbook: onSave,
                saveWorkbookAs: onSaveAs,
                setRibbonTab: onTabChange,
                applyCommand,
              }}
            />
          ) : (
            <>
              <button
                type="button"
                className="demo__title-icon"
                aria-label={ui.save}
                title={ui.save}
                onClick={onSave}
              >
                <DemoIcon name="save" />
              </button>
              <button
                type="button"
                className="demo__title-icon"
                aria-label={ui.undo}
                title={ui.undo}
                onClick={() => instance?.undo()}
              >
                <DemoIcon name="undo" />
              </button>
              <button
                type="button"
                className="demo__title-icon"
                aria-label={ui.redo}
                title={ui.redo}
                onClick={() => instance?.redo()}
              >
                <DemoIcon name="redo" />
              </button>
            </>
          )}
        </div>
        {mac ? (
          <DemoDocumentMenu ui={ui} bookName={bookName} onRename={onRename} />
        ) : (
          <div className="fc-tb__title">
            <strong>{bookName}</strong>
            <span>{ui.saved}</span>
          </div>
        )}
        {mac && !searchActive ? (
          <button
            type="button"
            className="demo__search-button"
            aria-label={ui.searchSuggestions.button}
            title={ui.searchSuggestions.button}
            onClick={openSearch}
          >
            <DemoGlyph name="find" />
          </button>
        ) : (
          <div className="fc-tb__search">
            {mac ? <DemoGlyph name="find" /> : <DemoIcon name="search" />}
            <input
              ref={searchInputRef}
              type="search"
              role="combobox"
              placeholder={mac ? ui.searchSuggestions.button : ui.search}
              aria-label={mac ? ui.search : ui.searchCommands}
              aria-controls="demo-search-results"
              aria-expanded={searchOpen}
              aria-activedescendant={
                searchOpen && searchActiveIndex >= 0
                  ? demoSearchOptionId(searchActiveIndex)
                  : undefined
              }
              value={searchQuery}
              onFocus={() => dispatchSearch({ type: 'focus' })}
              onChange={(e) => {
                const input = e.currentTarget;
                dispatchSearch({
                  type: 'input',
                  value: input.value,
                  focused: document.activeElement === input,
                });
              }}
              onKeyDown={(e) => {
                const action = resolveDemoSearchKey(e.key, searchActiveIndex, rowCount);
                if (action?.kind === 'close') {
                  dispatchSearch({ type: 'key', action });
                  e.currentTarget.blur();
                } else if (action?.kind === 'move') {
                  e.preventDefault();
                  dispatchSearch({ type: 'key', action });
                } else if (action?.kind === 'run') {
                  e.preventDefault();
                  runRow(action.index);
                }
              }}
              onBlur={() => {
                dispatchSearch({ type: 'blur' });
                if (mac && searchQuery.trim() === '') setSearchActive(false);
              }}
            />
            {searchOpen && showSuggestions ? (
              <div
                id="demo-search-results"
                className="fc-tb__command-menu demo__search-suggestions"
                role="listbox"
                aria-label={ui.searchSuggestions.popup}
              >
                {suggestionRows.map((section) => (
                  <div
                    key={section.id}
                    className={`demo__search-section demo__search-section--${section.id}`}
                    role="group"
                    aria-label={section.heading ?? undefined}
                  >
                    {section.heading ? (
                      <div className="demo__search-heading" role="presentation">
                        {section.heading}
                      </div>
                    ) : null}
                    {section.items.map((item) => {
                      const index = suggestions.indexOf(item);
                      return (
                        <button
                          key={item.id}
                          id={demoSearchOptionId(index)}
                          type="button"
                          role="option"
                          aria-selected={index === searchActiveIndex}
                          aria-disabled={item.disabled ? 'true' : undefined}
                          className={`demo__search-suggestion${
                            index === searchActiveIndex ? ' demo__search-suggestion--active' : ''
                          }`}
                          onMouseDown={(e) => e.preventDefault()}
                          onMouseEnter={() => dispatchSearch({ type: 'hover', index })}
                          onClick={() => runRow(index)}
                        >
                          <DemoGlyph name={item.glyph} />
                          <span className="demo__search-suggestion-label">{item.label}</span>
                          {item.submenu ? (
                            <span className="demo__search-submenu" aria-hidden="true" />
                          ) : null}
                        </button>
                      );
                    })}
                  </div>
                ))}
              </div>
            ) : searchOpen ? (
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
                      onMouseEnter={() => dispatchSearch({ type: 'hover', index })}
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
        )}
        <div className="demo__account">
          {mac ? null : (
            <button type="button" className="demo__share">
              {ui.share}
            </button>
          )}
          <button
            type="button"
            className={`demo__share demo__pane-toggle${showPanel ? ' demo__share--active' : ''}`}
            onClick={onTogglePanel}
            aria-pressed={showPanel}
          >
            {ui.demoPane}
          </button>
          {mac ? (
            <>
              <button
                type="button"
                className="demo__title-icon demo__comments"
                aria-label={ui.comments}
                title={ui.comments}
                onClick={() => applyCommand('mac.review.showComments')}
              >
                <DemoGlyph name="comments" />
              </button>
              <button type="button" className="demo__share-pill">
                <DemoGlyph name="share" />
                <span>{ui.share}</span>
                <span className="demo__share-caret" aria-hidden="true" />
              </button>
            </>
          ) : (
            <span className="demo__avatar" role="img" aria-label={ui.signedInUser}>
              FC
            </span>
          )}
        </div>
      </div>
    </header>
  );
};
