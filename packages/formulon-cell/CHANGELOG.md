# Changelog

All notable changes to `@libraz/formulon-cell` are documented here. The
format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/);
versioning is [SemVer](https://semver.org/).

## 0.1.0 — Unreleased

Initial public release.

### Added

- `Spreadsheet.mount()` with extension-based composition. Built-in
  extensions: formula bar, status bar, context menu, find/replace, format
  dialog, format painter, conditional formatting, named ranges, hyperlink
  dialog, paste-special, validation, autocomplete, hover comments,
  clipboard, wheel scroll, keymap.
- `presets.minimal() / .standard() / .excel()` for one-line setups.
- Runtime i18n via `instance.i18n.setLocale / extend / register`. `ja` and
  `en` ship in the box; new locales can be registered at runtime.
- `paper` / `ink` themes wired through `data-fc-theme` attribute and CSS
  custom properties — paint canvas reads the same tokens.
- WASM loaded via the portable `new URL(asset, import.meta.url)` pattern,
  so the package works under any modern bundler. Falls back to an
  in-memory stub when `crossOriginIsolated` is unavailable.

[0.1.0]: https://github.com/libraz/formulon-cell/releases/tag/v0.1.0
