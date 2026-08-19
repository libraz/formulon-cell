# Changelog

All notable changes to `@libraz/formulon-cell` are documented here. The
format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/);
versioning is [SemVer](https://semver.org/).

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

[0.5.1]: https://github.com/libraz/formulon-cell/releases/tag/v0.5.1
[0.5.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.5.0
[0.4.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.4.0
[0.3.1]: https://github.com/libraz/formulon-cell/releases/tag/v0.3.1
[0.3.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.3.0
[0.2.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.2.0
[0.1.1]: https://github.com/libraz/formulon-cell/releases/tag/v0.1.1
[0.1.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.1.0
