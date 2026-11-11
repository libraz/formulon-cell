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

/** The `cfRulesDialog` keys the Mac dialog can override. Each locale starts from its own shared text and overrides the fields the Excel bundle answers. */
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

/** Mac app wording from the Excel bundle; not answered by it: title, headerRule. */
export const macCfRulesDialogDe: Partial<MacCfRulesDialogStrings> = {
  duplicate: 'Regel duplizieren',
  orderLabel: 'Reihenfolge der Regeln ändern:',
  appliesToLabel: 'Gilt für:',
};

/** Mac app wording from the Excel bundle; not answered by it: title, headerRule. */
export const macCfRulesDialogFr: Partial<MacCfRulesDialogStrings> = {
  duplicate: 'Dupliquer la règle',
  orderLabel: 'Modifier l’ordre de règle :',
  appliesToLabel: 'S’applique à :',
};

/** Mac app wording from the Excel bundle; not answered by it: title, headerRule. */
export const macCfRulesDialogKo: Partial<MacCfRulesDialogStrings> = {
  duplicate: '중복 규칙',
  orderLabel: '규칙 순서 변경:',
  appliesToLabel: '적용 대상:',
};

/** Mac app wording from the Excel bundle; not answered by it: title, headerRule. */
export const macCfRulesDialogTh: Partial<MacCfRulesDialogStrings> = {
  duplicate: 'ทำซ้ำกฎ',
  orderLabel: 'เปลี่ยนลำดับของกฎ:',
  appliesToLabel: 'นำไปใช้กับ:',
};

/** Mac app wording from the Excel bundle; not answered by it: title, headerRule. */
export const macCfRulesDialogZh: Partial<MacCfRulesDialogStrings> = {
  duplicate: '复制规则',
  orderLabel: '更改规则次序:',
  appliesToLabel: '适用于:',
};
