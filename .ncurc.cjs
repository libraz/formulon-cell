/** npm-check-updates configuration. */
module.exports = {
  reject: [
    // Pinned to 6.x. tests/unit/toolbar/ribbon-menu-primitives/fixtures.ts imports the
    // TypeScript compiler API (`import ts from 'typescript'`) to parse the ribbon
    // menu sources; 7.x no longer exposes that API from the package entry point.
    'typescript',
  ],
};
