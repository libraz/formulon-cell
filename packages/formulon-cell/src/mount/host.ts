import type { ThemeName } from '../extensions/index.js';
import type { Strings } from '../i18n/strings.js';

let mountCounter = 0;

export function prepareMountHost(
  host: HTMLElement,
  strings: Strings,
  theme: ThemeName | undefined,
): string {
  host.classList.add('fc-host');
  host.setAttribute('tabindex', '0');
  host.setAttribute('role', 'region');
  host.setAttribute('aria-roledescription', 'spreadsheet');
  host.setAttribute('aria-label', strings.a11y.spreadsheet);
  host.dataset.fcTheme = theme ?? 'paper';
  host.replaceChildren();

  const instanceId = `fc-${++mountCounter}`;
  host.dataset.fcInstId = instanceId;
  return instanceId;
}

export function releaseMountHost(host: HTMLElement, instanceId: string): void {
  if (host.dataset.fcInstId !== instanceId) return;
  host.replaceChildren();
  host.classList.remove('fc-host');
  delete host.dataset.fcInstId;
  delete host.dataset.fcEngineState;
}
