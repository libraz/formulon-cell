# Changesets

Automated versioning for `@libraz/formulon-cell` and its framework adapters
(`@libraz/formulon-cell-react`, `@libraz/formulon-cell-vue`).

## Add a changeset

```sh
yarn changeset
```

The CLI walks you through choosing affected packages and a bump (patch /
minor / major) and writes a markdown file to this directory.

## Release flow

1. Open a PR with code + a changeset markdown.
2. After merging to `main`, run `yarn version` locally — it consumes the
   changeset markdowns, bumps versions, and updates each package's
   `CHANGELOG.md`.
3. Tag the release commit (`git tag v0.x.y && git push --tags`).
4. The `Publish and Create Release` GitHub Action picks up the tag and
   publishes all three packages to npm with `--access public`.

The three published packages are kept on a single fixed version (see
`config.json`'s `fixed` array) so the adapter peer-deps never diverge.
