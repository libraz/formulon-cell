import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { describe, expect, it } from 'vitest';
import { buildRibbonModel, fluentIconPaths } from '../../../src/index.js';

type ReactRibbonControlKind = 'tool' | 'select' | 'color' | 'break';

interface RibbonControl {
  id: string;
  kind: ReactRibbonControlKind;
}

const reactToolbarSource = (name: string): string =>
  readFileSync(resolve(process.cwd(), `../formulon-cell-react/src/toolbar/${name}`), 'utf8');

const reactRibbonControls = (): RibbonControl[] => {
  const source = `${reactToolbarSource('groups.tsx')}\n${reactToolbarSource('add-in-groups.tsx')}`;
  const controls: RibbonControl[] = [];
  const re = /\b(tool|select|optionSelect|color|rowBreak)(?:<[^>]+>)?\(\s*['"]([^'"]+)/g;
  for (const match of source.matchAll(re)) {
    const [, kind, id] = match;
    if (!kind || !id) continue;
    controls.push({
      id,
      kind:
        kind === 'rowBreak'
          ? 'break'
          : kind === 'optionSelect'
            ? 'select'
            : (kind as ReactRibbonControlKind),
    });
  }
  if (source.includes('mergeMenu,')) {
    const wrapIndex = controls.findIndex((control) => control.id === 'wrap');
    controls.splice(wrapIndex + 1, 0, { id: 'merge', kind: 'select' });
  }
  return controls;
};

const modelControls = (): RibbonControl[] =>
  buildRibbonModel('en').flatMap((tab) =>
    tab.groups.flatMap((group) =>
      group.commands.map((command) => ({
        id: command.id,
        kind:
          command.kind === 'select' || command.kind === 'color'
            ? command.kind
            : command.kind === 'break'
              ? 'break'
              : 'tool',
      })),
    ),
  );

describe('toolbar/ribbon-model', () => {
  it('keeps the shared ribbon command surface aligned with the React toolbar', () => {
    expect(modelControls()).toEqual(reactRibbonControls());
  });

  it('has Fluent SVG paths for every icon used by the ribbon model', () => {
    const missing = buildRibbonModel('en')
      .flatMap((tab) => tab.groups)
      .flatMap((group) => group.commands)
      .filter((command) => command.icon && !fluentIconPaths(command.icon))
      .map((command) => `${command.id}:${command.icon}`);

    expect(missing).toEqual([]);
  });
});
