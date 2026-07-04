# formulon-cell stylesheets

This directory ships **two independent stylesheets** plus a token vocabulary:

| Entry                 | Public name                                  | Loaded by                                                      |
| --------------------- | -------------------------------------------- | -------------------------------------------------------------- |
| `index.css`           | `@libraz/formulon-cell/styles.css`           | host page mounting `Spreadsheet`                               |
| `toolbar.css`         | `@libraz/formulon-cell/styles/toolbar.css`   | host page mounting `SpreadsheetToolbar` (react / vue wrappers) |
| `tokens.css`          | (declarative only — no styling effect)       | imported by `index.css` for editor IntelliSense                |

## Cascade layers

`layers.css` declares this order (lowest to highest precedence — later wins):

```css
@layer fc.reset, fc.base, fc.theme, fc.tokens, fc.surface, fc.app;
```

Every rule **must** live inside one of these layers so that consumer-side styles
(which are unlayered, and therefore beat any layered rule) can override our
defaults without `!important`. The single source of truth is `index.css`; never
declare layers in a leaf stylesheet.

## Theming surface

`tokens.css` declares the **public** custom-properties consumers may set on
`.fc-host` (or any ancestor). They are intentionally declared with `initial` so
that omitting a token in a custom theme falls through to the cascade rather
than the variable's last-set value.

| Theme file            | Selector                              | Notes                                                  |
| --------------------- | ------------------------------------- | ------------------------------------------------------ |
| `theme-paper.css`     | `.fc-host:not([data-fc-theme])` + `[data-fc-theme="paper"]` | default; light spreadsheet look                        |
| `theme-ink.css`       | `[data-fc-theme="ink"]`               | dark mode                                              |
| `theme-contrast.css`  | `[data-fc-theme="contrast"]`          | WCAG-AAA hard-edge contrast                            |

To add a brand theme: copy one of these files, change the data-attribute, and
override only the tokens you need.

The host instance API (`inst.setTheme('paper' | 'ink' | 'contrast' | …)`)
writes `data-fc-theme` on `.fc-host` for you. The ribbon toolbar bundle reads
the same `data-fc-theme` attribute and the same `paper` / `ink` / `contrast`
vocabulary, so putting `data-fc-theme` on a common ancestor of the grid and the
toolbar themes both surfaces together through the cascade — one attribute, one
set of theme names.

## Token namespaces

| Prefix          | Where defined                           | Purpose                                                          |
| --------------- | --------------------------------------- | ---------------------------------------------------------------- |
| `--fc-*`        | `tokens.css` + `theme-*.css`            | the **public** override surface for the spreadsheet itself.       |
| `--fc-tb-*`      | `toolbar/base/tokens.css`               | the **public** override surface for the ribbon toolbar bundle.    |
| `--fc-z-*`      | `tokens.css`                            | stacking floors (set on `:where(html)`, not on `.fc-host`).       |

When in doubt:
- adding a color used by the spreadsheet body / dialogs → `--fc-*`
- adding a color used by the ribbon toolbar             → `--fc-tb-*`
- nothing about the spreadsheet engine changes between themes → use literals
  inside the leaf stylesheet, not a new token.

## When to add a new token vs. a literal hex

Add a new token when **at least one** of these is true:

- the value differs between paper / ink / contrast themes,
- consumers reasonably want to override it (brand color, focus ring, …), or
- the same value appears in three or more places.

For one-off illustrations (data-bar legends, swatch grids, icon glyphs in
preset thumbnails) literals are fine — these are not themable surfaces.

## File layout

```
index.css            — entry: imports layers, reset, base, themes, surface, app, print
toolbar.css          — entry: imports the ribbon-toolbar bundle (separate consumer surface)
tokens.css           — declares the public --fc-* variable vocabulary (no styling)
theme-{paper,ink,contrast}.css — assigns values to those tokens per theme
core/
  reset.css          — minimal box-sizing / margin reset scoped to .fc-host
  base.css           — typography & root vars common to all themes
  surface.css        — sheet, header rail, formula bar, status bar (re-imports)
  app.css            — overlays, dialogs, panels, popups, editor helpers (re-imports)
  print.css          — print stylesheet (hides chrome, expands grid)
  surface/           — surface sub-files (formulabar, statusbar, sheetbar, …)
  app/               — app sub-files (dialogs, overlays, panels, popups, …)
toolbar/
  base/              — ribbon document chrome (titlebar, command bar, backstage, tokens)
  ribbon/            — ribbon controls (buttons, dropdowns, color flyout, cf icons)
  panels/            — side panel + modal dialogs + responsive overrides
```

Each `*.css` is intentionally narrow (≤ ~300 lines as a rule of thumb). When a
file passes that threshold, split it by the component its rules describe and
re-export via `@import`.

## Shadow / hairline tokens

Reuse `--fc-shadow-{2,4,8,16}` and `--fc-hairline` rather than re-inventing
rgba shadow stacks. The shadow scale follows Fluent depth tiers; both increase
in opacity for the ink theme and switch to solid black outlines for the
contrast theme.
