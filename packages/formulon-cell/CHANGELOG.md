# Changelog

All notable changes to `@libraz/formulon-cell` are documented here. The
format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/);
versioning is [SemVer](https://semver.org/).

## Unreleased

### Added

- `Spreadsheet.mount(host, { toolbar: true })` builds the ribbon toolbar inside
  the host in a single call — no separate `mountToolbar` wiring — and exposes it
  as `instance.toolbar` (`null` when not requested). The ribbon shell is
  inserted at the top of `.fc-host` through a `display: contents` host, so it
  becomes the first item in the host's flex column with the grid filling the
  rest, and it shares the host's `data-fc-theme`. `instance.dispose()` tears the
  toolbar down with everything else. Embeds that need custom ribbon hooks, tabs,
  or dropdown overrides can still call `Spreadsheet.mountToolbar` against their
  own element.

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

[0.3.1]: https://github.com/libraz/formulon-cell/releases/tag/v0.3.1
[0.3.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.3.0
[0.2.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.2.0
[0.1.1]: https://github.com/libraz/formulon-cell/releases/tag/v0.1.1
[0.1.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.1.0
