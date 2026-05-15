import type { ReactElement } from 'react';
import type { BuildRibbonGroupsOptions } from './group-types.js';
import { RIBBON_TAB_LABELS } from './model.js';

type AddInGroupOptions = Pick<
  BuildRibbonGroupsOptions,
  'group' | 'iconLabel' | 'instance' | 'lang' | 'onAddIn' | 'onRunScript' | 'tool' | 'tr'
>;

export const buildAddInRibbonGroups = ({
  group,
  iconLabel,
  instance,
  lang,
  onAddIn,
  onRunScript,
  tool,
  tr,
}: AddInGroupOptions): { automate: ReactElement[]; acrobat: ReactElement[] } => ({
  automate: [
    group(
      RIBBON_TAB_LABELS.automate[lang],
      [
        tool(
          'script',
          tr.script,
          iconLabel('script', tr.script),
          () => onRunScript?.(),
          false,
          ' demo__rb--wide',
          !onRunScript,
        ),
      ],
      'tiles',
    ),
  ],
  acrobat: [
    group(
      tr.addIn,
      [
        tool(
          'addIn',
          tr.addIn,
          iconLabel('addIn', tr.addIn),
          () => onAddIn?.(),
          false,
          ' demo__rb--wide',
          !onAddIn,
        ),
      ],
      'tiles',
    ),
    group(
      tr.pdf,
      [
        tool(
          'pdf',
          tr.pdf,
          iconLabel('pdf', tr.pdf),
          () => instance?.print(),
          false,
          ' demo__rb--wide',
        ),
      ],
      'tiles',
    ),
  ],
});
