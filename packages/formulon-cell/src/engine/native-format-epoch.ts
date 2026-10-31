const workbookEpochs = new WeakMap<object, number>();

/** Cheap cache token for native workbook format state. */
export const nativeFormatEpoch = (workbook: object): number => workbookEpochs.get(workbook) ?? 0;

/** Advance when StoreMirror publishes native format state or a muted section completes. */
export const advanceNativeFormatEpoch = (workbook: object): number => {
  const next = nativeFormatEpoch(workbook) + 1;
  workbookEpochs.set(workbook, next);
  return next;
};
