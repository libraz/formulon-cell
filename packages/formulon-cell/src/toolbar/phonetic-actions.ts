import {
  phoneticReadingAt,
  phoneticSettingsAt,
  setPhoneticReading,
  setPhoneticSettings,
} from '../commands/phonetic.js';
import type { SpreadsheetInstance } from '../mount/types.js';
import { showPhoneticSettingsDialog } from './dialogs/phonetic-settings.js';
import { showPrompt } from './dialogs/prompt.js';
import type { ToolbarText } from './ribbon-model.js';

/** What the phonetic dialogs need from their host: the strings and the glue
 *  that repaints and returns focus once a dialog commits. */
export interface PhoneticActionEnv {
  inst: SpreadsheetInstance;
  text: ToolbarText;
  refreshCells: () => void;
  focusSheet: () => void;
}

/** Open the reading editor on the active cell. */
export function openPhoneticEditor(env: PhoneticActionEnv): void {
  const { inst, text } = env;
  if (!inst.workbook?.capabilities.phonetic) return;
  const addr = inst.store.getState().selection.active;
  const initial = phoneticReadingAt(inst.store, addr);
  void showPrompt({
    title: text.phoneticDialogTitle,
    label: text.phoneticDialogLabel,
    initial,
    okLabel: text.ok,
    cancelLabel: text.cancel,
  }).then((phonetic) => {
    const workbook = inst.workbook;
    if (phonetic === null || !workbook) return;
    if (!setPhoneticReading(inst.store, workbook, addr, phonetic, initial)) return;
    env.refreshCells();
    env.focusSheet();
  });
}

/** Open the phonetic settings dialog over the selection. */
export function openPhoneticSettings(env: PhoneticActionEnv): void {
  const { inst, text } = env;
  const workbook = inst.workbook;
  if (!workbook?.capabilities.phoneticProperties) return;
  const state = inst.store.getState();
  const range = state.selection.range;
  void showPhoneticSettingsDialog({
    title: text.phoneticSettings,
    typeLabel: text.phoneticType,
    alignmentLabel: text.phoneticAlignment,
    typeOptions: [
      { value: 'hiragana', label: text.phoneticHiragana },
      { value: 'katakana', label: text.phoneticKatakana },
      { value: 'halfKatakana', label: text.phoneticHalfKatakana },
    ],
    alignmentOptions: [
      { value: 'normal', label: text.phoneticAlignNormal },
      { value: 'left', label: text.phoneticAlignLeft },
      { value: 'center', label: text.phoneticAlignCenter },
      { value: 'distributed', label: text.phoneticAlignDistributed },
    ],
    initial: phoneticSettingsAt(inst.store, state.selection.active),
    okLabel: text.ok,
    cancelLabel: text.cancel,
  }).then((settings) => {
    if (settings === null) return;
    setPhoneticSettings(inst.store, workbook, range, settings);
    env.refreshCells();
    env.focusSheet();
  });
}
