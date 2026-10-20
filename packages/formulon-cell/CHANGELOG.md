# Changelog

All notable changes to `@libraz/formulon-cell` are documented here. The
format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/);
versioning is [SemVer](https://semver.org/).

## Unreleased

### Added

- A selected chart outlines its data on the grid in the spreadsheet's colours: series names in red, categories (or X values) in purple, values in blue. The outline is published as `state.ui.chartSourceRefs` and set, with the selected object, by `mutators.setSelectedObject`.
- Chart data points show a tooltip with the series, point and value; pie and doughnut slices add their share of the total. `sessionCharts` strings gain `pointTooltip` and `pointValue`.
- Double-clicking a chart opens its settings, and Escape or a press outside the chart deselects it.
- A selected chart, picture or shape takes the place of the cell selection: the name box shows its name the way the spreadsheet numbers objects (`Chart 1`, `Picture 1`, `Oval 2`), the formula bar is empty, and the cell cursor, range shading, fill handle and header highlight are hidden. The selection is published as `state.ui.selectedObject`. `sessionCharts` strings gain `objectName`.
- Pictures and shapes have their own `sessionIllustrations` strings, with the spreadsheet's object names in all seven locales. `attachSessionIllustrations` takes them as `labels` and its handle gains `setLabels`.
- Right-clicking a chart, picture or shape opens its own menu instead of the cell menu, also from Shift+F10 or the context-menu key. A chart's menu offers Change Chart Type, Select Data and Format Chart Area, which open the matching settings tab, and every object can be brought to the front or sent to the back. `arrangeSessionChart` reorders a chart among its sheet's charts as one undo step. `sessionCharts` strings gain `changeChartType`, `selectData`, `formatChartArea`, `bringToFront` and `sendToBack`; `sessionIllustrations` strings gain `bringToFront` and `sendToBack`.
- Holding Alt while moving or sizing a chart, picture or shape docks the dragged edges to the nearest gridlines.
- Typed error literals such as `#N/A` and `#DIV/0!` are stored as error values, with the German and French spellings in those locales. `CoercedInput` gains an `error` kind.

### Changed

- The chart frame follows the desktop spreadsheet: the title is drawn inside the chart, the frame is flat with a hairline border, a selected chart has eight sizing handles, and the elements, styles and filters buttons sit outside its right edge as icons. A chart moves from a press anywhere on it, and a Shift-dragged corner keeps its shape. `sessionCharts` strings gain `elementsButton`, `stylesButton` and `filtersButton`.
- Automatic value axes are scaled the way the spreadsheet scales them: a twentieth of the data span is added as headroom and the bound is rounded out to a major unit of 1, 2 or 5 times a power of ten that leaves at least five intervals. A positive axis starts at zero unless the data spans less than a sixth of its maximum. The default column chart of values up to 15 now runs from 0 to 16 in steps of 2.
- Chart titles are 14 px, and a horizontal legend is centred under or over the chart.
- A data label that would run off the chart is moved back inside it instead of being shortened.
- Pictures and shapes are selected and sized like charts: eight sizing handles, a press anywhere moves the object once it travels a few pixels, the frame follows the pointer and the drag is recorded as one undo step, and Escape or a press outside deselects. A picture keeps its shape from a corner; a shape keeps it with Shift.
- Typed input is parsed the way the desktop spreadsheet parses it in each of the seven locales, checked against a capture from the desktop application. A General cell now picks up the number format the desktop spreadsheet attaches: `#,##0` for grouped numbers, `0.00E+00` for exponents, the system short date for dates, `# ?/?` for mixed fractions, and the locale's currency format with no decimals for whole amounts. Month/day (`1/5`), year/month, CJK and Japanese-era dates, date + time pairs, hour-only AM/PM times and fractional seconds are recognised, and month names follow the locale. Booleans and errors must be typed without surrounding spaces, and German and French accept only their own spellings.

### Removed

- The chart close button, the `close` string of `sessionCharts`, and the `closeLabel` option of `attachSessionCharts`. Delete removes a selected chart.
- The `closeLabel`, `pictureLabel`, `shapeLabel` and `resizeLabel` options of `attachSessionIllustrations`, and the `resize` string of `sessionCharts`.

### Fixed

- Charts and anchored pictures scroll off with their anchor cell instead of sticking to the first visible column or row, and they pass under the headers and frozen panes.
- The long date format offered for Japanese and Chinese was missing its closing quote and displayed the date serial.
- Loading a number format no longer drops the backslash-escaped currency symbol or the parentheses of a negative section, and date, time and custom formats keep their quoted literals and locale tags. `h:mm` is read as a time rather than a date-time.
- Number and currency formats whose negatives are parenthesized pad positive values with `_)`, as the desktop spreadsheet writes them, so positives line up with negatives.

## 0.10.0 — 2026-10-20

### Added

- Charts module, exported as `@libraz/formulon-cell/charts`. `layoutChart` lays out column, bar, line, area, pie, doughnut and scatter charts independently of the renderer. The module also provides chart and series types, scales with logarithmic axes, number formatting, data-label options, scatter line fitting and style resolution.
- Session charts render through the charts module and have a settings panel. It covers chart type and grouping, series direction, empty-cell handling, axes (including log scale and label position), legend, data labels, and series and point formatting. Chart edits respect sheet protection and the interaction policy, and support undo/redo.
- Charts in loaded workbooks are read and edited through the engine's chart API, and they save back to XLSX. Charts the module cannot draw show a placeholder.
- Chart series can be added and removed from the settings panel or with `addSessionChartSeries` / `removeSessionChartSeries`. Both support undo/redo, and the last series cannot be removed. The series name must be a cell reference.
- Pivot value-field settings list every show-values-as kind, including the new % Of, % Running Total In, Rank Smallest to Largest and Rank Largest to Smallest. The base field and base item selects appear only for the kinds that use them. `PivotShowValuesAs` gains `PercentOf`, `PercentOfRunningTotal`, `RankAscending` and `RankDescending`.
- Pictures can be reordered in the engine (bring forward / send backward).

### Changed

- Update `@libraz/formulon` to 0.16.0. The engine adds a chart API, stores formula text the way the spreadsheet stores entered text (`1.50` reads back as `1.5`), and resolves built-in number format ids to the active profile's locale.
- `WorkbookHandle.removeImage` is renamed to `removeDrawing` and now acts on charts as well as pictures. It is joined by `setDrawingAnchor`, `setDrawingZOrder`, `snapshotDrawing` and `restoreDrawing`. `EngineCapabilities` gains `drawings` and `charts`.
- `SessionChartKind` is a deprecated alias of `ChartKind`.

### Removed

- The root exports `SessionChartSeriesPoint`, `sessionChartSeries`, `applyChartsSnapshot`, `captureChartsSnapshot` and `recordChartsChange`. Chart data comes from the charts module, and chart history goes through the operation history.

## 0.9.1 — 2026-10-14

### Changed

- Update `@libraz/formulon` to 0.15.2, including Windows Excel arithmetic rounding and collation, XLSB differential formats, and pivot tables that open in Excel.
- `WorkbookHandle.setPivotFieldSort` takes a `PivotSortOrder` (`Descending`, `Ascending`, `Manual`) in place of the `ascending` flag. `PivotSortOrder` is exported from the package root.
- Defined names read back without a leading `=`, matching names loaded from a file. `upsertDefinedName` and `createDefinedNamesFromSelection` return entries in that stored form. The Name Manager and the Name Box list still show references with a leading `=`.
- Icon-set rules with fewer than N-1 thresholds for an N-icon set are no longer written to the workbook; they still render locally. An empty threshold list writes the default thresholds.
- Conditional-format font differences that do not set a size no longer write one, so the rule leaves the cell's font size unchanged.
- `mountToolbar` now defaults to the standard ribbon tab set (`EXCEL365_STANDARD_RIBBON_TABS`, or the Mac set on macOS) when `ribbonTabs` is omitted. Pass `RIBBON_TABS` or `[...EXCEL365_STANDARD_RIBBON_TABS, ...OPTIONAL_RIBBON_TABS]` explicitly to keep the Draw, Automate and Acrobat tabs.

### Removed

- `CellRegistry.registerEditor`, `resolveEditor` and the `CellEditorEntry` / `CellEditorHandle` types. Registered editors were never mounted.
- The `ui.lockTheme` option, which had no effect.
- The `@libraz/formulon-cell/extensions/*` subpath export. Import extension factories from `@libraz/formulon-cell/extensions`.

### Fixed

- React and Vue `SpreadsheetToolbar` fall back to the core Spelling, Accessibility and Translate reports and the core script and add-in actions when the matching callback prop is omitted. A provided `onRunScript` / `onAddIn` handles only its own menu item.
- Autocomplete and the argument helper show a custom function's `meta.args` signature.
- The stylesheet header no longer claims that `mount()` injects styles at runtime.
- The File tab only switches the active tab and fires `onTabChange('file')`; hosts render their own backstage from that callback. The docs for `createBackstageView`, `setBackstageOpen` and `onTabChange` now say so.
- Back and Escape now close the backstage opened through `setBackstageOpen(true)` and restore the previously active tab. Other backstage actions remain host-handled.

## 0.9.0 — 2026-10-12

### Added

- Localized ribbon screen tips on hover and keyboard focus, with platform shortcuts, viewport clamping and Escape dismissal.
- Seven built-in locales (English, Japanese, German, French, Chinese, Korean and Thai), synchronized calculation profiles, and `instance.setLocale()` for runtime switching.
- Explicit locale and comma/semicolon delimiter options for CSV import/export.
- An opt-in language selector controlled by `features.localeSwitcher`, including live visibility changes.
- Locale-aware scalar input, numeric editing and clipboard text. German and French function aliases and boolean tokens follow the active profile.
- Native formula conversion for localized input and display, including decimal, argument and array separators, boolean and error tokens. Stored formulas retain canonical syntax.
- Typed AutoFilter hydration and native filter evaluation, with the existing evaluator retained for color criteria.
- Native validation of proposed values for supported rules and paged enumeration for Circle Invalid Data. Custom formulas retain the local input validator because native evaluation reads saved workbook values.
- Threaded comments with replies, editing, resolution, navigation and undo/redo. Authors and threads persist in XLSX; legacy notes remain separate.
- Native image insertion, deletion and geometry edits with undo/redo, original image bytes and XLSX save/load. Loaded images are read-only to preserve drawing metadata; native picture reordering remains unavailable.
- Workbook theme color and font presets in Page Layout, including the Mac ribbon. Theme edits support undo/redo independently of the app's paper/ink/contrast theme.
- Native detailed pagination for print, including page dimensions, repeated titles and placement. Interactive page views use native geometry for layouts representable by the grid; multiple or overlapping print areas retain the local preview paginator.
- AutoFit row height and column width across disjoint multi-range selections. Each range keeps its own cross-axis span, rows or columns selected more than once take the largest fit, and one undo restores all of them.

### Changed

- Update `@libraz/formulon` to 0.15.1, including measured Mac locale collation, locale tags in number formats and native row and column style records, and migrate named-style writes to its record API.
- Require Node.js 22 or later for all three packages.
- Narrow ribbon groups in stages when space runs short: labels wrap, then large buttons become stacked rows, then icon-only rows. Command order stays the same.
- Give function-library, calculation, view, help, arrange and selection-pane commands dedicated icons on the default and Mac ribbons.
- Use native range enumeration for populated-cell searches and native formula reads for A1/R1C1 display.

### Fixed

- Keep ribbon commands reachable on narrow screens, prevent translated labels from collapsing icons, and use the measured German and French font-control glyphs.
- Convert font sizes between native points and public CSS pixels, and synchronize point-based ribbon controls with cell painting, editing, dialogs and autofit.
- Preserve open dialog, find, comment and filter drafts when a locale switch is rejected.
- Parse localized conditional-format thresholds and canonicalize custom conditional-format and validation formulas when committed.
- Evaluate conditional-format formulas through the workbook profile and invalidate their display cache after locale changes.
- Show Watch Window values and formulas in the current profile and accept a live `getWorkbookName` callback in core, React and Vue.
- Use the current locale for number-format shortcuts and F4 repeat after a language switch, and keep workbook shortcuts out of native select controls.
- Associate language-switch failures with the selector for assistive technology, refresh repeated failure notices, and allow long notices to wrap.
- Preserve the active calculation profile for standalone formula evaluation while keeping workbook cells and names isolated.
- Localize scientific notation, Show Formulas and HTML/TSV clipboard output using the active language profile.
- Roll back newly registered named styles when a later catalog write fails, and retain built-in style ids 48–53.
- Convert engine column widths in characters and row heights in points to CSS pixels. Read sheet defaults and clear removed row-height overrides back to auto.
- Preserve XF apply flags, quote prefixes and protection metadata across format edits.
- Resolve theme, indexed and automatic colours for cell-format projection while retaining the original selectors for save/load.
- Render inherited row and column styles without writing them back as direct cell formatting, and use native display formatting where applicable.
- Preserve absolute and mixed references when editing formulas displayed in R1C1 notation.
- Restore native style records, whole-sheet font and row/column formats when undoing or redoing cell, row and column insertion and deletion, including inserted copied cells.
- Widen a styled row or column span when a row or column is inserted inside it, without saving the new gap as an explicit default style.
- Stop Goal Seek from reporting a stale workbook when the saved file differs only in archive timestamps.
- Wrap cell text at locale-aware word boundaries, so text without spaces such as Japanese or Chinese wraps instead of overflowing, and break oversized words without splitting grapheme clusters. Wrapped lines, padding and indent stay consistent across zoom levels, and AutoFit row height counts the same lines the grid draws.
- Show the function name in the active language in the cell editor and formula-bar argument helpers.
- Show a hyperlink's screen tip on hover when the cell has no comment.
- Use the full font name, including its theme font role, as the tooltip and accessible name of font dropdown options.
- Stop dialogs and the filter dropdown from moving focus to their initial field after the user has already started typing in another field, which sent the typed text to the wrong input.

## 0.8.0 — 2026-10-04

### Added

- Opt-in Mac platform profile. `ui.platform` accepts `'default'` (the default), `'mac'`, or `'auto'` to detect it from the browser; `resolveSpreadsheetPlatform` exposes the same resolution. The Mac profile mounts its own ribbon tab set, Mac keyboard shortcuts, and a Draw tab with an ink layer. `RibbonProfile` and `BuildRibbonModelOptions` select the tab set for hosts that build the ribbon model themselves.
- A nonmodal formula argument palette for the Mac profile. It suspends an in-progress cell or formula-bar edit and restores it on Cancel, edits the function call under the caret of a compound formula without rewriting the text around it, and records insertions in Recent.
- `MountOptions.getFunctionArgumentHelp` supplies per-argument hints and reference links to the palette. The React and Vue components accept it as a prop and export `FunctionArgumentHelp` and `FunctionArgumentHelpProvider`.
- Goal Seek, Consolidate, Subtotal, Insert Sparkline, Insert Slicer, and Workbook Statistics dialogs. Subtotal's structural changes go through the interaction policy and are undoable.
- Function Arguments and the function picker read the live workbook catalog, group functions into family categories (`FunctionCategory`), and keep a Recent list. `FxDialogOpenOptions` opens the dialog on a category.
- Multi-range selections: modifier-drag adds ranges, navigation moves within them, Ctrl/Cmd+Enter fills every selected cell, and ribbon formatting and Clear apply to every area after authorizing each one.
- Named styles apply across the selected areas using their included format groups. Format Cells gains an underline style selector.
- `AutofitOptions` for `autofitColWidth` / `autofitRowHeight`, and `formatA1Cell`.

### Changed

- Dialog format changes undo and redo against the engine as well as the store, and restore both if replay fails.
- Ribbon state updates index buttons by command once per update instead of querying each command.

### Fixed

- Input on a merged cell is validated and written at the merge anchor, and protection is checked across every cell of the merge.
- The inline editor stays open with the typed text when a workbook write fails, and reports the failure.
- Editor, formula bar, Backspace, and ribbon Clear commit through the registered interaction controller, so the policy and validation rules apply to each of them. Selection fill validates every target before writing.
- Inverted selection ranges are normalized instead of rejected, and whole-row and whole-column checks agree across selection mutators.
- Ctrl+Enter in the formula bar fills the selection when no controller is registered.
- Composite undo and redo are re-authorized against the current policy.
- Double-click fill unmerges the destination, and text scripts applied to a range recalculate once.
- Autofit measures in the rendered cell font and reserves room for the table header button.
- Manual calculation mode is read through the engine's `calcMode()`.
- Function syntax renders from one formatter, autocomplete accepts dotted function names, and engine and built-in descriptions use the same catalog locale.
- The conditional-format formula splitter tracks nested parentheses, braces, and brackets.
- The context menu keeps the format intent and clears range unions.
- Host features resolve the workbook lazily after `setWorkbook`.
- The ribbon display menu stays above a peeking ribbon, and Custom Sort takes focus synchronously.

## 0.7.0 — 2026-10-01

### Added

- Restricted embedding. `MountOptions.policy` authorizes every user mutation (edit, clear, paste, fill, structure, sheet, export) through one command service shared by the editor, formula bar, clipboard, keyboard, pointer, context menu, and ribbon; undo and redo re-check the current policy. `fixedFormPolicy(ranges)` and `viewerPolicy()` cover the common cases, and `policy.restrict(context)` adds host-specific denials. Omitting `policy` keeps the previous behaviour.
- The `embedded` UI profile mounts only the grid, with features, ribbon, and print off until enabled. `viewport` keeps navigation, selection, and rendering inside a fixed range; `contextMenu` selects built-in items, transforms them, or hands the menu to the host.
- Runtime setters `setPolicy`, `setViewportOptions`, `setContextMenu`, `setOverlayOptions`, `setUi`, and `setToolbar`, plus `applyChanges` for trusted host updates and a `changeBatch` event for policy-aware batches.
- `OverlayOptions.root` places menus and dialogs inside a host modal or fullscreen element. Without it, overlays follow the nearest open dialog or fullscreen element that contains the sheet.
- The React and Vue components accept `ui`, `policy`, `viewport`, `contextMenu`, `overlays`, and `toolbar` props and apply changes without re-mounting. React reports batches through `onChangeBatch`; Vue emits `change-batch`.
- `WorkbookHandle.applyCellPatchAtomic` writes a batch of cell patches with one recalc and restores the touched cells if the engine fails.
- Whole-row and whole-column copy and cut can be inserted as bands (`insertCopiedBand`), keeping formulas, formats, merges, notes, and row/column sizes, including across sheets, as a single undo step. `HistoryTransaction` and `captureSnapshotFromCopyResult` are exported for hosts that drive the same path.
- Data bars gain min/max scale points, axis position and colour, negative fill, borders, and solid fill; icon sets gain a per-threshold comparison operator.

### Changed

- **Breaking:** `PivotDateGrouping.Week` is replaced by `Days`. Use `Days` with `intervalDays: 7` for weekly buckets.
- Update the engine dependency to `@libraz/formulon` 0.12.0 and adapt status-bearing accessor results. The default single-threaded WASM loads without COOP/COEP headers.
- Register pivot number formats as engine format IDs and align pivot data fields and filters with the engine API.
- Formula references follow structural edits: inserting or deleting rows, columns, and cell bands, sorting, filling, cut-and-paste, and band moves rewrite references (including cross-sheet, whole-row/column, and structured references) instead of leaving them pointing at the old cells. Edits that would push content past the sheet edge are rejected before anything is written.
- Fill uses a least-squares linear trend for numeric series and tiles mixed formula/value sources with shifted references.
- Merging validates the range, carries visual formatting and perimeter borders, and refuses ranges that overlap a table.
- Repeated and transposed pastes fill the selected destination range, and oversized pastes are refused.
- Split conditional-format formula types, parsing, and evaluation from rule rendering.

### Fixed

- Paste falls back to the internal clipboard payload when the system clipboard cannot be read, and ignores a payload from an older copy.
- Cell format changes made through a bound store are mirrored to the engine, and loading a workbook no longer writes back what it just read.
- The format dialog previews the active cell's actual value with the real number format.
- Split buttons open their menu from the chevron, and disabled ribbon buttons ignore clicks.
- AutoSum treats the non-anchor cells of a merge as occupied.
- Sorting and context-menu sorts are undoable and carry comments and error values with their rows.
- Retry WASM initialization after a failed load.
- Render React children after asynchronous spreadsheet mount and support object or array selectors in `useSpreadsheet`.
- Reset Vue selection and locale state and release subscriptions when the spreadsheet instance is cleared.
- Preserve conditional-format icon floors, data-bar direction, and gradient settings across workbook load and save. Align bar geometry and gradients with the sheet and rule direction.
- Keep engine result status properties out of public array results and suppress late mount error callbacks after React unmounts.
- Preserve strict icon comparisons and scale signed data bars within each side of the zero axis.
- Preserve pivot decimal formats during hydration, update row and column field order during creation and editing, and name all supported aggregations consistently.

## 0.6.0 — 2026-08-23

### Added

- Worksheet print settings round-trip through xlsx. The Page Setup record was
  a session-only UI model because the engine offered getters and no setters;
  now every control that maps to an OOXML attribute is hydrated from a loaded
  workbook and written back on save — orientation, paper size, scale,
  fit-to-page, margins, print options, header/footer, print area, print titles,
  manual page breaks, and the attributes beside the engine's typed setter
  (black and white, draft quality, cell comments, cell errors, page order,
  first page number, print quality). A sheet whose paper is outside the sizes
  the UI names keeps its OOXML code in the new `PageSetup.paperSizeCode` and
  saves back on that paper rather than being silently re-papered as A4.

  `PageSetup.printableBounds` is the one field that stays session state, and
  for a reason rather than a gap: it records a physical printer's minimum
  margins, which OOXML has no concept of. `page-setup` joins the workbook
  compatibility summary so a host can say which of the two a workbook gets.

- Sheet tabs carry the three-state visibility a workbook can express, not just
  hidden-or-not. `LayoutSlice.veryHiddenSheets` records the subset a workbook
  marked very hidden, and the tab menu's Unhide list leaves those out — which
  is the whole point of the state, and what a spreadsheet's own Unhide dialog
  does. Hiding and unhiding no longer quietly demotes a very-hidden sheet to
  plain hidden on save. The tab menu gains an entry that states it, and
  `setSheetTabVisibility` is the command behind it; `setSheetHidden` stays the
  two-state form and now refuses to demote a very-hidden tab, matching the
  engine's own flag setter.

- A PivotTable filter can select the blank member of a field. The filter item
  list was built from labels alone, and the blank member has none, so it was
  dropped from the checklist entirely and a filter could never name it. Items
  now carry the cache index `<item x="N">` addresses them by, the blank member
  is listed under a `(blank)` placeholder, and the writeback states each item
  by index through `addPivotFieldItemAt`. A field whose members were inferred
  from the projected layout rather than read out of the cache has no index, and
  still goes by label.

- The iterative-calculation dialog opens on the settings the workbook actually
  carries. It had no way to read them back, so a file that arrived with
  iteration switched on still showed the dialog's own defaults, and clicking OK
  wrote those defaults over the file's.

- Phonetic guides are addressed span by span. `WorkbookHandle` gains
  `getCellPhoneticRuns` / `setCellPhoneticRuns`, and `commands/phonetic`
  exports `phoneticReading`, `phoneticReadingAt` and `setPhoneticReading` —
  the last being what a single-field editor should call, since it declines to
  overwrite a guide's spans with a reading the user did not change.

- `WorkbookHandle` gains `getIterative`, `setSheetVisibility`,
  `setWorkbookDefaultFont`, the print accessors (`getSheetPageSetup` /
  `setSheetPageSetup`, `getSheetPageMargins` /
  `setSheetPageMargins`, `getSheetPrintOptions` / `setSheetPrintOptions`,
  `getSheetHeaderFooter` / `setSheetHeaderFooter`, `getSheetPrintArea` /
  `setSheetPrintArea`, `getSheetPrintTitles` / `setSheetPrintTitles`,
  `getSheetPageBreaks` / `setSheetPageBreaks`, `getSheetPageSetupXml` /
  `setSheetPageSetupXml`, `getSheetPageSetupExtras` / `setSheetPageSetupExtras`)
  and the `printSettings` / `printSettingsXml` / `pageBreaks` /
  `sheetVisibility` / `iterativeSettings` / `pivotItemByCacheIndex` /
  `phoneticRuns` / `defaultFont` capability flags. `getSheetView` now reports
  `visibility` beside `tabHidden`, and `getPivotTables` reports
  `fieldItemIndexes` beside `fieldItems`.

  `setSheetPageSetupExtras` merges into the raw `<pageSetup>` fragment the
  engine keeps as its writer's source of truth, so an attribute neither side
  models — `copies`, or an `r:id` naming the sheet's printerSettings part —
  survives the write untouched. `mergePageSetupFragment` is exported for a host
  that needs to reach an attribute this package does not name.

### Changed

- The engine dependency moves to `@libraz/formulon` 0.11.1.

- `WorkbookHandle.withAutoFilterSyncMuted` is now `withEngineSyncMuted`: the
  bound store mirrors page setup as well as AutoFilter into the engine, and one
  scope suppresses both.

- A new workbook states its locale default font as font 0, through the new
  `WorkbookHandle.setWorkbookDefaultFont`, rather than appending it to the font
  table. The engine now seeds a fresh workbook's style table with the records a
  spreadsheet writes, so an appended font would have left index 0 — and with it
  every cell carrying no format at all — on the seeded Calibri. An engine that
  cannot restate its default falls back to naming the locale font on each
  formatted cell.

- `CellFormat.phonetic` carries the guide's runs (`readonly PhoneticRun[]`)
  rather than one string. A phonetic guide annotates spans of the cell text,
  and OOXML stores one `<rPh>` block per span, so the single-string model could
  only ever hold the readings run together — see the round-trip fix below. Read
  the editor's view of a guide with the exported `phoneticReading`.

### Fixed

- A partially annotated phonetic guide survives an edit. Every format sync
  rewrote the cell's guide as one reading for the whole cell, so touching any
  format on a cell whose furigana annotated its kanji separately collapsed the
  spans into a single whole-string annotation. Guides are now read and written
  span by span, and the phonetic editor — which offers one field, and so can
  only state a whole-cell reading — leaves the guide alone when the reading
  comes back unchanged. Each run is drawn centred over the span it annotates
  instead of all of them centred over the cell.

  A guide travels with the format it belongs to, so the writeback trims it to
  the spans that cover the target cell's own text — pasting formats onto a
  shorter cell would otherwise state runs reaching past the end of the string,
  which the engine takes verbatim and writes into the file.

- A font's `<scheme>` theme link survives a format edit. The link is what keeps
  a font tracking the workbook theme; records authored from UI state carried
  none, so bolding a cell rewrote its font as a literal name. The link now
  carries while the cell stays on the workbook's body font, and is dropped only
  when a family is named explicitly — which is the edit that genuinely picks a
  literal typeface.

- Data-validation writeback keeps rejecting blank cells only where the rule
  says so. The engine flipped `addValidation`'s `allowBlank` default from true
  to false; the writeback already spelled the field explicitly on every rule,
  so no stored rule changes meaning.

## 0.5.1 — 2026-08-20

### Added

- `WorkbookHandle.withBatchedRecalc(fn)` collapses a multi-cell write into a
  single recalc. Scopes nest, and the pass runs when the outermost one exits —
  including when the body throws, so a partially applied write is never left
  stale. Every bulk writer goes through it: paste, paste special, insert copied
  cells, fill, sort, remove duplicates, filter, text to columns, structure
  edits, find and replace, import, and undo / redo replay.
- `WorkbookHandle.recalcAuto()` is the edit-driven recalc, kept apart from
  `recalc()` so Calculate Now stays unconditional while the automatic passes
  can be suppressed.

### Changed

- Every cell write recalcs through the handle. `setNumber`, `setText`,
  `setBool`, `setError`, and `setBlank` schedule a pass, so dependents settle
  after an edit without the host calling `recalc()` itself.
- Manual calc mode is enforced here instead of being round-trip metadata the
  engine ignores: it suppresses the automatic pass after a write, and leaving
  Manual runs the recalc those edits were denied. `recalc()` stays
  unconditional in every mode.
- `RecalcEvent.dirty` carries the cells written since the previous pass — the
  edits the recalc was answering — rather than a set reported by the engine.
  Dependents recomputed by the pass are not listed, and the set is empty for a
  recalc no edit triggered.
- Ribbon icons are composed from a shared icon kit rather than written as
  literal path data: one canvas and optical size, a five-step stroke scale, a
  closed palette, and shared primitives for arrow heads, sheet frames, corner
  badges, and letterforms, with the artwork split into per-category glyph
  modules. Individual icons shift slightly as their colours and stroke widths
  move onto the shared scale. An icon segment gained optional `fillRule` and
  `transform`, so a host painting the exported ribbon icon paths through its
  own SVG serializer has to honour both — otherwise the icons carrying a
  letterform land at the wrong size and position.

### Fixed

- The cell-format dialog announced the format title while its heading read
  Data Validation. `role="dialog"` sits on the overlay, but the mode-specific
  accessible name was written to the panel inside it, so assistive technology
  never saw the name the dialog was opened under.
- The ribbon display-options menu ignored Escape when it had been opened by
  click. Opening the menu re-renders the ribbon, which drops focus back to the
  document, and Escape was only handled from the menu's own options.

### Removed

- `WorkbookHandle.setViewportHint()` and `clearViewportHint()`, along with the
  partial-recalc-on-write path they fed. `partialRecalc()` is still available
  but opt-in: dirty cells outside its rectangle keep their previous value until
  a later pass reaches them, which is not a trade a cell write should make on
  the host's behalf.

## 0.5.0 — 2026-08-19

### Added

- Page Layout and Page Break Preview are now rendered views rather than a mode
  stamped on the host. Page gutters and ruler bands fold into the grid geometry
  through `PagedLayout`, so cell rects, hit-testing, and the axis layouts shift
  without every painter learning about pages. The renderer paints the desk and
  paper, page frames, header and footer slots, rulers, break lines, the
  out-of-print wash, and a page-number watermark. Page boundaries come from the
  same band splitter the print path uses, so an on-screen break lands where the
  printed document splits. Break lines and the print-area frame are draggable,
  the Page Layout rulers set page margins, and header and footer slots are
  editable in place. Entering Page Break Preview zooms out to fit a page and
  restores the previous zoom on the way out. New `pageView` strings (`en` / `ja`)
  and page-view color tokens across the `paper`, `ink`, and `contrast` themes.
- Right-to-left sheets mirror the whole grid axis. `<sheetView rightToLeft>` is
  hydrated on load and writable through `setSheetRightToLeft`; arrow keys,
  `Alt+PageUp` / `Alt+PageDown`, horizontal wheel scrolling, and column-resize
  drags follow the visual direction, while Tab stays in index order. The header
  wedge, AutoFilter chevron, row-number strip, freeze divider, and the in-cell
  error / validation / comment / lock affordances anchor to the mirrored side,
  and overflow, centre-across, general alignment, and indent follow reading
  order. `colLeftEdge` is renamed `colLeadingEdge`, and geometry functions that
  emit or consume a screen x now take the `ViewLayout` / `ViewState` pair
  projected by `layoutForView`.
- Spreadsheet keyboard and interaction parity: End mode with status-bar state
  and current-region `Ctrl+A` expansion, Backspace clearing only the active
  cell, IME-composition-guarded shortcuts, `Alt+Down` opening a validation
  dropdown, repeat-last-action (`F4`) for format commands, an Insert / Delete
  Cells direction dialog, an Open Hyperlink context-menu entry with scheme
  validation, and direct shortcuts for number formats, insert/delete cells,
  hide/unhide rows and columns, filter, table, flash fill, and new sheet. Host
  shortcuts are derived from the ribbon's own ARIA shortcut declarations.
- The copy marquee has a full lifecycle. A `copyMode` (`'copy' | 'cut'`) field
  records which command raised it, so a copy marquee stands across repeat pastes
  while a cut marquee is consumed by its first paste. Escape and entering edit
  mode cancel copy mode, Enter is a one-shot paste that ends it, and row/column
  inserts and deletes shift or drop the marquee so it keeps outlining the copied
  band. Row and column header menus gain Insert Copied Cells, which opens as
  many whole rows or columns as the copied band is deep or wide and pastes into
  them.
- OOXML cell-format fidelity: the full fill-pattern repertoire with a pattern
  gallery in the Format Cells dialog, underline variants (single / double /
  accounting) with double-underline painting, the full border-style repertoire
  with SVG line samples, extended horizontal and vertical alignment, negative-
  number sections and locale-tagged currency tokens in generated format codes,
  a Fraction number category replacing Date & Time, superscript / subscript,
  justify-last-line for distributed alignment, and cell phonetic guides painted
  above the base text. The dialog also has a `dxf` mode so conditional
  formatting can edit a custom format, plus 2/3-color-scale style selection.
- Named cell styles are written as real OOXML styles: every referenced style
  becomes a `cellStyleXfs` + `cellStyle` pair carrying its `builtinId`, a
  style's own formatting stays on its style-xf row, and each styled cell points
  at it through `xfId` — so editing a style reaches every cell using it instead
  of collapsing into direct formatting. Format as Table persists a real
  `ListObject`, and AutoFilter definitions survive a save and are cleared on
  removal. Named cell styles, loaded tables, Format as Table, phonetic guides,
  and AutoFilter are reported as writable in the compatibility summary.
- `sheetView` display flags (`showGridLines`, `showHeaders`, `showZeros`,
  `rightToLeft`) are hydrated on load and written back, with `showZeros`
  plumbed into the grid paint context so zero values can be blanked.
- `DEFAULT_OFF_FEATURE_IDS` and `isFeatureDefaultOff(id)` are exported, so a
  host rendering a feature-toggle UI reads a flag the same way `resolveFlags`
  resolves it instead of hard-coding the default-off list.
- `openHostMenuFirstDropdown` and `RIBBON_HOST_MENU_FIRST_COMMANDS` are
  exported for split buttons whose menu is the entry point when the host owns
  the actions behind it (`script`, `addIn`). Pass the helper straight through
  as `mountToolbar`'s `interceptCommand`: it opens the dropdown for those
  commands and returns `false` for everything else, so the default dispatch
  still runs. Both framework adapters now use it instead of carrying their own
  copy.
- `formatSheetAbsoluteRange` is exported for rendering a range the way the
  desktop dialogs present one — sheet-qualified and absolute.

### Changed

- The flat View toolbar (`viewToolbar`) now ships default-off, alongside
  `watchWindow` and `slicer`, and no longer appears in `allBuiltIns()`. The
  ribbon's View tab already carries every control it offered — Zero Values and
  the right-to-left toggle were added there to complete it — and a second
  always-visible strip for the same commands has no desktop counterpart. Hosts
  that want the strip can still switch the flag on or compose the
  `viewToolbar()` factory explicitly.
- The engine dependency moves to `@libraz/formulon` 0.10.0 and its style-record
  model. `getCellXf`, `getFontRecord`, `getFillRecord`, and `getBorderRecord`
  return whole records, so indent, text rotation, shrink-to-fit, reading order,
  `xfId`, and colour provenance survive a round trip. The adapter-local
  `TableInput` and `CellXf` shims are gone now that the engine ships them.
- The Format Cells dialog moves the font and border color palettes into
  flyouts opened from a chevron beside each color input, placed as fixed panels
  clamped to the viewport and closed on tab switch, outside mousedown, dialog
  close, or a first Escape that no longer closes the dialog itself.
- The ribbon icon artwork is redrawn with heavier strokes and larger glyph
  counters so a glyph stays legible at ribbon size. The set of icon keys is
  unchanged, so no control changes which icon it resolves to.
- The READMEs lead with what the library is and when to reach for it, state the
  current compatibility scope (which workbook workflows are ready and which UI
  areas are not yet on par), add a Non-goals section, and normalize the badge
  sets across the root and package READMEs.
- The accessibility mirror is a valid ARIA grid. Cells were being emitted
  straight into the live region, which made every `gridcell` parentless and
  turned each selection change into an announcement of the whole mirror. The
  mirror is now `grid` → `rowgroup` → `row` → `gridcell`, and the polite
  announcer is a separate node outside the grid. Mirrored cell ids change from
  `<id>-a11y-viewport-cell-<r>-<c>` to `<id>-a11y-cell-<r>-<c>`, rows are
  addressable as `<id>-a11y-row-<r>`, `aria-owns` is gone from the grid, and
  `ExtensionContext.a11y` now points at the live region rather than the mirror
  container.
- The Create Table and Create PivotTable range fields show sheet-qualified
  absolute references (`Sheet1!$A$1:$B$3`), matching how the desktop dialogs
  present them. Create PivotTable also defaults to New Worksheet instead of
  Existing Worksheet; the destination picker stays disabled until Existing
  Worksheet is selected.
- Ribbon chrome meets WCAG 2.2 AA: group labels carry enough contrast at 10px,
  the ribbon display toggle and the stacked group buttons are at least 24x24,
  the Home tab's dense groups fit the panel without horizontal overflow, and
  the table-style gallery is a capped scrollable panel instead of growing past
  the viewport.
- `@libraz/formulon-cell-vue` re-exports the `Toolbar` alias instead of
  importing its own SFC to alias it. The self-import was a module cycle that
  read the default binding during evaluation, which threw before the component
  existed.
- The toolchain is pinned through `mise.toml` (Node 22, Yarn 4.18) instead of a
  `volta` block.

### Fixed

- A superseded instance can no longer tear down the live overlay portal. The
  portal is stamped with the claiming host's instance id, and a disposing
  instance skips removal when the portal no longer belongs to it — so a remount
  of the same host cannot detach the container the live instance renders its
  menus, dialogs, and dropdowns into.
- The inline editor stays welded to its cell. It repositions and resizes from a
  store subscription, so viewport, layout, and format changes during an edit
  follow the cell; a cell scrolled under the frozen band or the headers parks
  the editor outside the grid clip rect, keeping DOM focus, caret, and IME
  composition alive. The editor also takes the cell's own fill and text color
  and drops its radius, ring, and shadow transition so an edit reads as in
  place.
- Ribbon commands take focus when clicked. WebKit follows the macOS convention
  of not focusing a `<button>` on click, so ribbon keyboard navigation and any
  dialog that restores focus to the command that opened it both lost their
  anchor there. Whatever the command focuses next — a menu item, a dialog
  field, the sheet — still wins.

## 0.4.0 — 2026-07-05

### Added

- `Spreadsheet.mount(host, { toolbar: true })` builds the ribbon toolbar inside
  the host in a single call — no separate `mountToolbar` wiring — and exposes it
  as `instance.toolbar` (`null` when not requested). The ribbon shell is
  inserted at the top of `.fc-host` through a `display: contents` host, so it
  becomes the first item in the host's flex column with the grid filling the
  rest, and it shares the host's `data-fc-theme`. `instance.dispose()` tears the
  toolbar down with everything else.
- `mount`'s `toolbar` option now also accepts a `MountToolbarOptions` object, not
  just `true`. Pass one to add backstage content (`createBackstageView`), custom
  hooks, submenu factories, a ribbon-tab profile, or lifecycle callbacks while
  keeping the single-call mount — the fields you set are merged over (and
  override) the built-in defaults (locale-derived language, dynamic dropdowns),
  and the toolbar is still wired against the freshly mounted instance for you.
  Embeds that need a fully separate toolbar host can still call
  `Spreadsheet.mountToolbar` directly.

### Changed

- All floating UI (context menus, dialogs, tooltips, dropdowns, popovers) now
  mounts into a single per-instance overlay portal appended to `<body>` and
  tagged with the host's `data-fc-theme`, instead of being teleported directly
  onto `<body>` with theme tokens hand-copied onto each overlay. Overlays now
  inherit paper / ink / contrast tokens through normal CSS cascade, so a new
  theme token reaches every overlay automatically — no per-overlay forwarding
  list to keep in sync. Removed the `inheritHostTokens` helper and its
  hard-coded token allow-list. This also fixes the dark-theme context-menu
  glyph regression structurally: `--fc-menu-icon-filter` now flows to the menu
  like any other token.
- Single CSS entry: `@libraz/formulon-cell/styles.css` now includes the ribbon
  toolbar styles, so it is the only stylesheet an embedder needs (grid +
  toolbar + all three palettes). The individual `styles/paper.css`,
  `styles/ink.css`, `styles/contrast.css`, and `styles/toolbar.css` exports
  remain for granular setups.
- The ribbon toolbar's CSS class prefixes are unified from `demo__*` / `app__*`
  to a single `fc-tb__*` (e.g. `demo__ribbon-shell` / `app__menu-item` →
  `fc-tb__ribbon-shell` / `fc-tb__menu-item`), matching the grid's `fc-*` /
  `fc-host__*` convention. Hosts that styled or queried ribbon markup by class
  should update the prefix or, preferably, target the stable `data-ribbon-*`
  attributes instead. The published toolbar stylesheet also no longer ships the
  reference demo app's own page-shell styling (brand bar, options panel, preset
  picker, feature toggles, event log); that CSS moved to the demo apps, so the
  bundle now contains only genuine toolbar chrome.
- The ribbon toolbar's overridable theme tokens are renamed from `--demo-*` to
  `--fc-tb-*` (e.g. `--demo-ribbon-bg` → `--fc-tb-ribbon-bg`), so the public
  override surface no longer ships under a `demo` name and cannot collide with
  the grid's `--fc-*` tokens. Embedders that overrode `--demo-*` must switch to
  the `--fc-tb-*` names. Internal ribbon class names are unchanged; hosts should
  target the `data-ribbon-*` attributes rather than class names.
- One brand accent across grid and toolbar. When the ribbon is mounted inside
  `.fc-host` (single-call `mount({ toolbar })`, or any `SpreadsheetToolbar` whose
  host descends from the grid host), `--fc-tb-accent` / `--fc-tb-accent-strong`
  now fall back to the grid's `--fc-accent` / `--fc-accent-strong`, and
  `--fc-tb-accent-soft` derives from the resolved accent. A single
  `.fc-host { --fc-accent: … }` override now themes both surfaces — no more
  setting the grid and toolbar accents separately. Standalone toolbars (host not
  under `.fc-host`) keep the `--fc-tb-*` literals, which now match the grid's
  per-theme accent so both mount modes look identical (the ink ribbon accent
  moves from its former independent value to the shared theme accent).
- One theme vocabulary across grid and toolbar. The ribbon toolbar now reads
  the same `paper` / `ink` / `contrast` values off the shared `data-fc-theme`
  attribute, instead of a separate `light` / `dark` / `contrast` vocabulary on
  its own `data-theme` attribute. Put `data-fc-theme` on a common ancestor of
  the grid and toolbar — or just drive `instance.setTheme(name)` — and both
  surfaces theme together through CSS cascade. `MountToolbarOptions.theme`,
  `ToolbarInstance.getTheme()` / `setTheme()`, and the `onThemeChange` callback
  now speak `ThemeName` (`paper` / `ink` / `contrast`).

### Fixed

- Picking a palette from the ribbon's Page Layout → Themes gallery now actually
  re-themes the grid. The tiles emit the grid's `ThemeName` directly, so a
  selection reaches `setTheme()` instead of stamping an unmatched
  `data-fc-theme="dark"` that left the grid unstyled.
- The distributed toolbar stylesheet no longer ships demo-page globals
  (`* { box-sizing }`, `html, body, #root` sizing, and a `body { background }`
  rule). Embedding the toolbar previously restyled the host page's `<body>`;
  those rules now live in the demo apps where they belong.
- Context-menu item padding is still re-asserted unlayered, so an aggressive
  host `button { padding: 0 }` reset (as shipped by VitePress, Tailwind
  preflight, or normalize.css) cannot collapse it and clip the keyboard-shortcut
  hints against the menu's right edge.

## 0.3.1 — 2026-07-04

### Added

- Upgraded to the formulon 0.9.5 calc engine.
- Array-aware F9 formula preview: an array- or spill-returning selection now
  renders as a spreadsheet array constant (`{a,b;c,d}`) instead of collapsing
  to its top-left value, backed by the engine's ad-hoc `evaluateFormulaArray`.
  Exposed as `WorkbookHandle.evaluateFormulaArray` behind the new
  `arrayFormulaEvaluation` capability, with a scalar fallback when the engine
  does not provide it.
- Host-injectable function metadata: `WorkbookHandle.setFunctionMetadataProvider`
  merges localized signature / description / alias overrides over the engine's
  structural function catalog, resolved by locale-override → entry-default →
  engine-value precedence. Ships the pure `mergeFunctionMetadata` helper and
  re-exports the `EvalArrayResult` and `FunctionMetadata*` types.

## 0.3.0 — 2026-07-04

### Added

- Ribbon toolbar, dialogs, and menu chrome are now built once in
  `@libraz/formulon-cell` and shared by every host. `Spreadsheet.mountToolbar`
  gained the full ribbon activation model, dynamic-dropdown dispatcher, and
  dialog set (Sort, Text to Columns, Remove Duplicates, Advanced Filter,
  Conditional Formatting, PivotTable, sheet/view, and protection dialogs).
- Full desktop-spreadsheet-style chrome: backstage/file menu, printer profile
  picker, command search, key tips, and drawing/illustration tools (shapes
  with corner-radius editing, duplicate, and line/opacity controls).
- PivotTable creation dialog with per-field value settings (summarize-by,
  number format, show-values-as) and pivot cache refresh.
- Conditional formatting gains standard-deviation-based rules and
  formula-driven ranges, backed by a statistical- and lookup-aware formula
  evaluator, with rule edits now written back to the workbook.
- Hyperlink display text and full-fidelity cell snapshots (formatting,
  comments, hyperlinks) now survive clipboard copy/paste and xlsx
  export/round-trip.
- Named ranges can be scoped to the active sheet instead of always being
  workbook-global.
- Comments are hydrated via engine-wide enumeration where the underlying
  engine supports it, instead of per-cell lookups only.
- Print/export: pagination tiling, repeating title columns, "fit to N pages"
  scaling, and PDF export are exposed as first-class commands.
- Data tools: color-based sort/filter and improved filter/slicer semantics.
- Format-as-Table now routes through a dedicated Create Table dialog from
  the Home tab.
- Formula preview (F9) is evaluated through the real engine instead of a
  static stub, and gains keyboard-navigable results.
- Spreadsheet-style border drawing UI, paste of previously copied cells, and
  a number-format dropdown were added to the toolbar.
- The canvas grid now exposes ARIA grid semantics, and dialogs, popovers,
  and menus gained keyboard navigation and focus handling.
- New `ja`/`en` strings for pivot layout, named-range scope, and
  fill-preview UI.
- Core now publishes building blocks for hosts that extend or audit the
  ribbon: `ribbonActivationEntries`, `ribbonSurfaceCommandIds`,
  `DYNAMIC_RIBBON_DROPDOWN_HANDLER_ATTRS`, `attachRangePickerButton`,
  `appendConditionalApplyFormatControls`, `conditionalStyleOptions`,
  `showReport`, `reportDialogLabels`, `projectDisabledReason`, and
  `projectDisabledState`.
- `SpreadsheetToolbar` (React/Vue) gained `onError` and `onToolbarReady`
  props to surface toolbar mount failures and access the mounted toolbar
  instance, plus expanded ribbon type re-exports.

### Changed

- **React and Vue `SpreadsheetToolbar` components were rewritten from
  framework-native ribbon implementations into thin adapters over
  `Spreadsheet.mountToolbar`.** The documented prop surface (`instance`,
  `activeTab`, `onTabChange`, `locale`, the review/drawing/script hook
  callbacks) is unchanged and additive-only, but the internal DOM
  structure, CSS class names, and any previously-importable sub-components
  are not preserved. Hosts that styled or queried internal ribbon markup
  directly should switch to the `data-ribbon-*` attributes exposed by core.
- Floating UI (dialogs, dropdowns, popovers) now uses a consistent z-index
  tier so it reliably layers above host-provided modals.
- Grid header, ribbon, and dialog styling were realigned to a consistent
  desktop-spreadsheet baseline.

### Fixed

- Toolbar instances and dialogs no longer leak listeners/DOM on dispose.
- Dialogs are portaled to `document.body` so they no longer clip inside
  scrollable/overflow-hidden hosts; clipboard shortcuts route correctly and
  the cell editor keeps focus more reliably during interaction.
- Range-scanning commands (formula preview, reference rewriting, and
  similar bulk operations) now guard against unbounded selections to avoid
  slowdowns on very large ranges.
- The format submenu is registered in the dynamic-dropdown dispatch keys so
  its ribbon dropdown opens correctly.
- Self-package imports resolve to relative paths, producing clean library
  builds without unresolved import warnings.

## 0.2.0 — 2026-05-11

### Added

- Viewport zoom is now applied uniformly to geometry and hit-testing.
  `colWidth`, `rowHeight`, `frozenColsWidth`, `frozenRowsHeight`,
  `colX`, `rowY`, `cellRect`, `hitTest`, `buildColLayout`, and
  `buildRowLayout` accept an optional `ViewportSlice` argument and
  multiply visible dimensions by `viewport.zoom`. Default zoom is `1`,
  so existing callers continue to work unchanged.
- React and Vue companion packages now publish a `SpreadsheetToolbar`
  ribbon component sharing the same tab model and command surface.

### Fixed

- General-format numbers that overflow their column shrink to fit
  before falling back to `####` (released in 0.1.1, documented here).

## 0.1.1 — 2026-05-11

### Changed

- Reinstated `publishConfig.provenance: true` after configuring npm
  trusted-publisher (OIDC) bindings for the three packages.

### Fixed

- Render: shrink overflowing General-format numbers before falling
  back to `####`.

## 0.1.0 — 2026-05-11

Initial public release.

### Added

- `Spreadsheet.mount()` with extension-based composition. Built-in
  extensions: formula bar, status bar, context menu, find/replace, format
  dialog, format painter, conditional formatting, named ranges, hyperlink
  dialog, paste-special, validation, autocomplete, hover comments,
  clipboard, wheel scroll, keymap.
- `presets.minimal() / .standard() / .full()` for one-line setups.
- Runtime i18n via `instance.i18n.setLocale / extend / register`. `ja` and
  `en` ship in the box; new locales can be registered at runtime.
- `paper` / `ink` themes wired through `data-fc-theme` attribute and CSS
  custom properties — paint canvas reads the same tokens.
- WASM loaded via the portable `new URL(asset, import.meta.url)` pattern,
  so the package works under any modern bundler. Falls back to an
  in-memory stub when `crossOriginIsolated` is unavailable.

[0.10.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.10.0
[0.9.1]: https://github.com/libraz/formulon-cell/releases/tag/v0.9.1
[0.9.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.9.0
[0.8.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.8.0
[0.7.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.7.0
[0.6.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.6.0
[0.5.1]: https://github.com/libraz/formulon-cell/releases/tag/v0.5.1
[0.5.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.5.0
[0.4.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.4.0
[0.3.1]: https://github.com/libraz/formulon-cell/releases/tag/v0.3.1
[0.3.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.3.0
[0.2.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.2.0
[0.1.1]: https://github.com/libraz/formulon-cell/releases/tag/v0.1.1
[0.1.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.1.0
