import { showPrompt } from './prompt.js';

export interface RenameSheetDialogOptions {
  title: string;
  label: string;
  initial: string;
  requiredMessage: string;
  okLabel: string;
  cancelLabel: string;
}

export const showRenameSheetDialog = async (
  opts: RenameSheetDialogOptions,
): Promise<string | null> => {
  const value = await showPrompt({
    title: opts.title,
    label: opts.label,
    initial: opts.initial,
    okLabel: opts.okLabel,
    cancelLabel: opts.cancelLabel,
    validate: (v) => (v.trim() ? null : opts.requiredMessage),
  });
  return value === null ? null : value.trim();
};
