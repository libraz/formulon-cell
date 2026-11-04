import type { CellFormatStrings } from '../cell-format/types.js';

type CfRulesDialogStrings = CellFormatStrings['cfRulesDialog'];

/** Labels only the Mac Manage Rules dialog shows. */
export interface MacCfRulesDialogLabels {
  /** Caption beside the rule order arrows. */
  orderLabel: string;
  /** Caption above the selected rule's range field. */
  appliesToLabel: string;
}

/** Overrides of `cfRulesDialog`, plus the Mac-only labels, that apply on the Mac host only. */
export type MacCfRulesDialogStrings = Partial<CfRulesDialogStrings> & MacCfRulesDialogLabels;

/** The `cfRulesDialog` keys the Mac dialog can override. A locale whose Mac
 *  wording is not measured repeats its own shared text. */
const MAC_CF_RULES_DIALOG_KEYS = [
  'title',
  'headerRule',
  'duplicate',
] as const satisfies readonly (keyof CfRulesDialogStrings)[];

export const macCfRulesDialogFrom = (
  dialog: CfRulesDialogStrings,
  labels: MacCfRulesDialogLabels,
): MacCfRulesDialogStrings => ({
  ...Object.fromEntries(MAC_CF_RULES_DIALOG_KEYS.map((key) => [key, dialog[key]])),
  ...labels,
});

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macCfRulesDialogJa: MacCfRulesDialogStrings = {
  title: 'ルールの管理',
  headerRule: 'ルール (表示順で適用)',
  duplicate: 'ルールの重複',
  orderLabel: 'ルールの順序を変更:',
  appliesToLabel: '適用対象:',
};
