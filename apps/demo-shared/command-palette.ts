// Lightweight command palette that filters ribbon commands as the user types
// into a search input. Framework-agnostic: hosts pass a plain
// `HTMLInputElement` plus a container the menu can be appended to. React/Vue
// demos call it from `useEffect` / `onMounted`.

import {
  buildRibbonSearchIndex,
  projectDisabledReason,
  queryRibbonSearchIndex,
  type RibbonSearchItem,
  type RibbonTab,
} from '@libraz/formulon-cell';

export interface CommandPaletteItem {
  id: string;
  label: string;
  hint: string;
}

export interface CommandPaletteOptions {
  input: HTMLInputElement;
  container: HTMLElement;
  ribbonLang: 'ja' | 'en';
  applyCommand: (id: string) => boolean;
  selectTab?: (tab: RibbonTab) => void;
}

const noCommandsLabel = (lang: 'ja' | 'en'): string =>
  lang === 'ja' ? '一致するコマンドはありません' : 'No matching commands';

const toPaletteItem = (item: RibbonSearchItem): CommandPaletteItem => ({
  id: item.commandId ?? item.id,
  label: item.label,
  hint: [
    item.kind === 'tab' || item.kind === 'help' ? item.hint : `${item.hint} · ${item.tab}`,
    item.disabledReason,
  ]
    .filter(Boolean)
    .join(' · '),
});

export const createCommandPalette = (opts: CommandPaletteOptions): { dispose: () => void } => {
  const { input, container, ribbonLang, applyCommand, selectTab } = opts;
  const commands = buildRibbonSearchIndex(ribbonLang, { includeDisabled: true });

  let menu: HTMLDivElement | null = null;
  let usagePrior: { commandBoosts?: Record<string, number> } = {};

  const recordUsage = (commandId: string | undefined): void => {
    if (!commandId) return;
    usagePrior = {
      commandBoosts: {
        ...(usagePrior.commandBoosts ?? {}),
        [commandId]: Math.min(100, (usagePrior.commandBoosts?.[commandId] ?? 0) + 12),
      },
    };
  };

  const closeMenu = (): void => {
    menu?.remove();
    menu = null;
  };

  const openMenu = (): void => {
    const query = input.value.trim().toLowerCase();
    const visible = queryRibbonSearchIndex(commands, query, 8, { usagePrior });
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
      const paletteItem = toPaletteItem(cmd);
      const btn = document.createElement('button');
      btn.type = 'button';
      btn.className = 'demo__command-item';
      if (cmd.disabled) btn.setAttribute('aria-disabled', 'true');
      projectDisabledReason(btn, cmd.disabledReason ?? null, {
        datasetKey: 'disabledReason',
        title: false,
      });
      const label = document.createElement('strong');
      label.textContent = paletteItem.label;
      const hint = document.createElement('span');
      hint.textContent = paletteItem.hint;
      btn.append(label, hint);
      // mousedown fires before blur; preventDefault keeps focus on the input
      // so the click handler runs before the menu disappears.
      btn.addEventListener('mousedown', (e) => e.preventDefault());
      btn.addEventListener('click', () => {
        if (cmd.commandId) {
          recordUsage(cmd.commandId);
          applyCommand(cmd.commandId);
        } else selectTab?.(cmd.tab);
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
      const first = queryRibbonSearchIndex(commands, query, 1, { usagePrior })[0] ?? null;
      if (first) {
        event.preventDefault();
        if (first.commandId) {
          recordUsage(first.commandId);
          applyCommand(first.commandId);
        } else selectTab?.(first.tab);
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
