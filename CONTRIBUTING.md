# Contributing to formulon-cell

Contributions are welcome, including small and unfinished ones. A maintainer reviews every pull request, finishes what is missing, and merges it, so you do not need to get everything in this document right, or even read all of it.

## What a pull request needs

- A description of what the change does and why. If the intent is clear, the maintainer can verify the behaviour.
- If you can, a build and the tests for the part you changed.

That is all. A red check, a missing translation, or a test you did not write is fine: say what you know in the pull request, and the maintainer will finish it or tell you what is left. Target `develop` if you can; a pull request against `main` is retargeted.

For a larger change, such as a new dialog or a change to the mount options, opening an issue first to describe the idea can save you work. It is not required.

Reporting a bug you do not intend to fix is a useful contribution too, and needs none of the setup below.

### What the maintainer can verify

The maintainer tests on macOS in every shipped language and on Windows in Japanese only. If your change depends on Windows in another language (a keyboard layout, an IME, a system font, a locale-specific shortcut), describe what you see there, ideally with a screenshot, because it cannot be reproduced on the maintainer's side.

## What you are contributing to

A TypeScript UI layer over the [formulon](https://github.com/libraz/formulon) WebAssembly calc engine:

- `packages/formulon-cell` is the framework-free core: canvas grid, formula bar, ribbon, dialogs, i18n and themes.
- `packages/formulon-cell-react` and `packages/formulon-cell-vue` are thin adapters over the core's mount call.
- `apps/react-demo` and `apps/vue-demo` are the demo apps the end-to-end tests run against; `apps/demo-shared` holds the source both demos share.

Formula evaluation, recalculation and xlsx parsing belong to the engine. A formula that computes the wrong value is an engine issue; report it in the [formulon repository](https://github.com/libraz/formulon/issues).

The rest of this document has two parts. [Checklist](#checklist) describes what a finished change looks like. [Reference](#reference) explains how the checks behave.

## Development setup

Node 22 and Yarn 4 (the versions are pinned in `mise.toml` and `package.json`). Use Yarn; npm, pnpm and bun are not supported in this repository.

```bash
yarn install
yarn dev:react        # or yarn dev:vue
yarn test             # core unit tests (vitest)
yarn lint             # Biome; yarn lint:fix applies fixes
yarn typecheck
```

End-to-end tests run Playwright against both demos:

```bash
yarn test:e2e:install # once, to download the browsers
yarn build:all
yarn test:e2e
```

## What CI requires

A red check on your pull request is not a reason to hold it back; the maintainer can look at it with you.

| Workflow | Runs on | What it checks |
| --- | --- | --- |
| `ci.yml` (`build`) | pushes to `main` and pull requests to `main`, on Node 22 and 24 | `yarn lint`, `yarn build:all`, `yarn typecheck`, and the core unit tests with coverage |
| `visual-baselines.yml` | started by hand | regenerates the visual regression snapshots |
| `publish.yml` | version tags | publishes the three packages to npm |

## Checklist

This is what a change looks like by the time it is merged. The maintainer checks it and fills in what is missing, so you are not expected to know all of it.

- Keep the React and Vue packages thin: menus, commands, ribbon definitions and dialogs live in the core, and the adapters import them.
- Write cell values through `WorkbookHandle`, and wrap a write that touches several cells in `withBatchedRecalc`.
- Use the in-memory stub engine (`preferStub: true`) only in tests and explicit demos, never as a fallback when the WebAssembly engine fails to load.
- Build ribbon icons from `src/toolbar/icon-kit`, and run `yarn icons:review` to look at the contact sheets before and after an icon change.
- Add a new UI string to every locale table under `src/i18n/strings/`. English text in a language you do not know is fine; the maintainer translates it.
- Add unit tests under `packages/formulon-cell/tests` for the code you changed. A change visible in both demos gets a scenario under `tests/e2e-shared/scenarios`.
- Where a change fixes a defect, include a test that fails without the fix.
- Write comments in English and give exported symbols a TSDoc comment.

### Commit messages

Conventional Commits, in English: `type(scope): summary`, lowercase after the colon, no trailing period. A body, when there is one, is a blank line followed by `- ` bullets. Commit messages can be rewritten at merge, so do not let them hold up a pull request.

## Reference

Nothing in this section asks anything of you; read it when a checklist item or a check puzzles you.

- The engine marks cells dirty on write but never recalculates on its own. `WorkbookHandle` recalculates after every write outside Manual calculation mode, so a loop of single-cell writes recalculates once per cell unless it is batched.
- The stub engine leaves formula evaluation, recalculation and xlsx round-trip incomplete on purpose. Falling back to it silently would show a sheet that looks loaded and computes nothing.
- Each locale table is typed as the full string set, so a string added to one locale and not the others fails `yarn typecheck`.
- Ribbon icons are composed from shared tokens and primitives rather than literal path data, which keeps stroke widths and optical sizes consistent across the set.
- `ci.yml` does not run the end-to-end tests; they run locally.
- A pull request to `develop` does not start `ci.yml`; the maintainer runs the same checks before merging.

## Reporting issues

Search the existing issues first, then include what you did, what happened, what you expected, the package version, your OS and its display language, and the browser. If a desktop spreadsheet application behaves differently, describe what it does. For a problem with a loaded file, the file matters more than anything else in the report.

Security vulnerabilities go through the process in [SECURITY.md](SECURITY.md), not the public issue tracker.

## Licence

By contributing, you agree that your contributions are licensed under the Apache License 2.0.
