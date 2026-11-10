import type { BuiltInExtension, Extension } from './types.js';

const builtIns = new WeakSet<object>();

/** Marks `ext` as a built-in feature so the registry hands it the engine-level context. The mark
 *  is on the object, not the id: a user extension that replaces a built-in by id stays a user one. */
export const builtIn = (ext: BuiltInExtension): Extension => {
  const out = ext as unknown as Extension;
  builtIns.add(out);
  return out;
};

/** Whether `ext` was created by a built-in factory. */
export const isBuiltInExtension = (ext: Extension): boolean => builtIns.has(ext);
