import type { resolveSpreadsheetUiOptions, ThemeName } from '../extensions/index.js';
import {
  type ResolvedSpreadsheetPlatform,
  resolveSpreadsheetPlatform,
} from '../extensions/ui-options.js';
import type { Strings } from '../i18n/strings.js';
import type { createSpreadsheetStore } from '../store/store.js';

let mountCounter = 0;

export function prepareMountHost(
  host: HTMLElement,
  strings: Strings,
  theme: ThemeName | undefined,
  platform?: ResolvedSpreadsheetPlatform,
): string {
  host.classList.add('fc-host');
  host.setAttribute('tabindex', '0');
  host.setAttribute('role', 'region');
  host.setAttribute('aria-roledescription', 'spreadsheet');
  host.setAttribute('aria-label', strings.a11y.spreadsheet);
  host.dataset.fcTheme = theme ?? 'paper';
  host.dataset.fcPlatform = platform ?? resolveSpreadsheetPlatform();
  host.replaceChildren();

  const instanceId = `fc-${++mountCounter}`;
  host.dataset.fcInstId = instanceId;
  return instanceId;
}

export function releaseMountHost(host: HTMLElement, instanceId: string): void {
  if (host.dataset.fcInstId !== instanceId) return;
  host.replaceChildren();
  host.classList.remove('fc-host');
  host.removeAttribute('tabindex');
  host.removeAttribute('role');
  host.removeAttribute('aria-roledescription');
  host.removeAttribute('aria-label');
  delete host.dataset.fcInstId;
  delete host.dataset.fcEngineState;
  delete host.dataset.fcTheme;
  delete host.dataset.fcPlatform;
}

function mountErrorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export function renderMountError(
  host: HTMLElement,
  error: unknown,
  strings: Strings['mountError'],
): void {
  const panel = document.createElement('div');
  panel.className = 'fc-mount-error';
  panel.setAttribute('role', 'alert');

  const title = document.createElement('strong');
  title.textContent = strings.title;

  const help = document.createElement('p');
  help.textContent = strings.engineHelp;

  const detail = document.createElement('code');
  detail.textContent = mountErrorMessage(error);

  panel.append(title, help, detail);
  host.replaceChildren(panel);
}

export function applyPlatformLayoutDefaults(
  store: ReturnType<typeof createSpreadsheetStore>,
  platform: ReturnType<typeof resolveSpreadsheetUiOptions>['platform'],
): void {
  const defaults =
    platform === 'mac'
      ? { defaultColWidth: 75, headerColWidth: 26 }
      : {
          defaultColWidth: 64,
          headerColWidth: 32,
        };
  store.setState((state) => ({
    ...state,
    layout: { ...state.layout, ...defaults },
  }));
}
