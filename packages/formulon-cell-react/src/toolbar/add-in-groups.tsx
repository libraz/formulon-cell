import type { ReactElement } from 'react';
import type { BuildRibbonGroupsOptions } from './group-types.js';

type AddInGroupOptions = Pick<
  BuildRibbonGroupsOptions,
  | 'addInMenu'
  | 'pdfMenu'
  | 'group'
  | 'iconLabel'
  | 'strings'
  | 'onRunScript'
  | 'onRecordActions'
  | 'onAllScripts'
  | 'tool'
  | 'tr'
>;

export const buildAddInRibbonGroups = ({
  addInMenu,
  pdfMenu,
  group,
  iconLabel,
  strings,
  onRunScript,
  onRecordActions,
  onAllScripts,
  tool,
  tr,
}: AddInGroupOptions): { automate: ReactElement[]; acrobat: ReactElement[] } => {
  return {
    automate: [
      group(
        strings.ribbon.tabs.automate,
        [
          tool(
            'script',
            tr.script,
            iconLabel('script', tr.script),
            () => onRunScript?.(),
            false,
            ' demo__rb--wide',
            !onRunScript,
            !!onRunScript,
          ),
          tool(
            'recordActions',
            tr.recordActions,
            iconLabel('script', tr.recordActions),
            () => onRecordActions?.(),
            false,
            ' demo__rb--wide',
            !onRecordActions,
            !!onRecordActions,
          ),
          tool(
            'allScripts',
            tr.allScripts,
            iconLabel('script', tr.allScripts),
            () => onAllScripts?.(),
            false,
            ' demo__rb--wide',
            !onAllScripts,
            !!onAllScripts,
          ),
        ],
        'tiles',
      ),
    ],
    acrobat: [
      group(tr.addIn, [addInMenu], 'tiles'),
      group(tr.pdf, [pdfMenu], 'tiles'),
    ],
  };
};
