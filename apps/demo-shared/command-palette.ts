// Lightweight command palette that filters ribbon commands as the user types
// into a search input. Framework-agnostic: hosts pass a plain
// `HTMLInputElement` plus a container the menu can be appended to. React/Vue
// demos call it from `useEffect` / `onMounted`; the playground wires it
// straight from its boot sequence.

import type { ToolbarText } from '@libraz/formulon-cell';

export interface CommandPaletteItem {
  id: string;
  label: string;
  hint: string;
}

export interface CommandPaletteOptions {
  input: HTMLInputElement;
  container: HTMLElement;
  ribbonText: ToolbarText;
  ribbonLang: 'ja' | 'en';
  applyCommand: (id: string) => boolean;
}

const noCommandsLabel = (lang: 'ja' | 'en'): string =>
  lang === 'ja' ? '一致するコマンドはありません' : 'No matching commands';

const buildCommands = (ribbonText: ToolbarText, lang: 'ja' | 'en'): CommandPaletteItem[] => {
  const ja = lang === 'ja';
  return [
    {
      id: 'formatCells',
      label: ribbonText.formatCells,
      hint: ja ? 'セルの書式を変更します' : 'Change cell formatting',
    },
    {
      id: 'conditional',
      label: ribbonText.conditional,
      hint: ja ? '条件付き書式を編集' : 'Edit conditional formatting',
    },
    {
      id: 'findHome',
      label: ribbonText.find,
      hint: ja ? 'シート内を検索 / 置換' : 'Find or replace on the sheet',
    },
    {
      id: 'pageSetup',
      label: ribbonText.pageSetup,
      hint: ja ? 'ページ設定を開く' : 'Open page setup',
    },
    {
      id: 'evaluateFormula',
      label: ribbonText.evaluateFormula,
      hint: ja ? '数式を 1 ステップずつ評価' : 'Step through a formula',
    },
    {
      id: 'recalcNow',
      label: ribbonText.recalc,
      hint: ja ? 'ブックを再計算' : 'Recalculate the workbook',
    },
    {
      id: 'sortAscHome',
      label: ribbonText.sortAscending,
      hint: ja ? '昇順に並べ替え' : 'Sort ascending',
    },
    {
      id: 'sortDesc',
      label: ribbonText.sortDescending,
      hint: ja ? '降順に並べ替え' : 'Sort descending',
    },
    {
      id: 'freeze',
      label: ribbonText.freeze,
      hint: ja ? 'ウィンドウ枠を固定' : 'Freeze panes',
    },
  ];
};

const matches = (cmd: CommandPaletteItem, query: string): boolean => {
  const haystack = `${cmd.label} ${cmd.hint}`.toLowerCase();
  return haystack.includes(query);
};

export const createCommandPalette = (opts: CommandPaletteOptions): { dispose: () => void } => {
  const { input, container, ribbonText, ribbonLang, applyCommand } = opts;
  const commands = buildCommands(ribbonText, ribbonLang);

  let menu: HTMLDivElement | null = null;

  const closeMenu = (): void => {
    menu?.remove();
    menu = null;
  };

  const openMenu = (): void => {
    const query = input.value.trim().toLowerCase();
    const visible = query
      ? commands.filter((cmd) => matches(cmd, query)).slice(0, 8)
      : commands.slice(0, 8);
    if (!menu) {
      menu = document.createElement('div');
      menu.className = 'demo__command-menu';
      container.appendChild(menu);
    } else {
      menu.replaceChildren();
    }
    if (visible.length === 0) {
      const empty = document.createElement('div');
      empty.className = 'demo__command-empty';
      empty.textContent = noCommandsLabel(ribbonLang);
      menu.appendChild(empty);
      return;
    }
    for (const cmd of visible) {
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'demo__command-item';
      const label = document.createElement('strong');
      label.textContent = cmd.label;
      const hint = document.createElement('span');
      hint.textContent = cmd.hint;
      btn.append(label, hint);
      // mousedown fires before blur; preventDefault keeps focus on the input
      // so the click handler runs before the menu disappears.
      btn.addEventListener('mousedown', (e) => e.preventDefault());
      btn.addEventListener('click', () => {
        applyCommand(cmd.id);
        input.value = '';
        closeMenu();
        input.blur();
      });
      menu.appendChild(btn);
    }
  };

  const onInput = (): void => openMenu();
  const onFocus = (): void => openMenu();
  const onBlur = (): void => closeMenu();
  const onKeyDown = (event: KeyboardEvent): void => {
    if (event.key === 'Escape') {
      closeMenu();
      input.blur();
      return;
    }
    if (event.key === 'Enter') {
      const query = input.value.trim().toLowerCase();
      const first = query ? commands.find((cmd) => matches(cmd, query)) : null;
      if (first) {
        event.preventDefault();
        applyCommand(first.id);
        input.value = '';
        closeMenu();
        input.blur();
      }
    }
  };

  input.addEventListener('input', onInput);
  input.addEventListener('focus', onFocus);
  input.addEventListener('blur', onBlur);
  input.addEventListener('keydown', onKeyDown);

  return {
    dispose: () => {
      closeMenu();
      input.removeEventListener('input', onInput);
      input.removeEventListener('focus', onFocus);
      input.removeEventListener('blur', onBlur);
      input.removeEventListener('keydown', onKeyDown);
    },
  };
};
