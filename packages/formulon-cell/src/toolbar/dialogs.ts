// Re-export shim. The real implementations now live in ./dialogs/*.ts.
// Kept here so existing `import { ... } from './dialogs.js'` callers keep
// working without churning every import site at once.
export * from './dialogs/index.js';
