import { findActiveSignature } from '../commands/refs.js';

export interface ArgHelperHandle {
  /** Re-evaluate the tooltip against the current input value/caret. */
  refresh(): void;
  /** Hide the tooltip. */
  close(): void;
  detach(): void;
}

export interface ArgHelperDeps {
  /** The textarea/input being edited. */
  input: HTMLInputElement | HTMLTextAreaElement;
}

/**
 * Floating tooltip that mirrors Excel's "ScreenTip" — when the caret sits
 * inside a known function call, it shows `NAME(arg1, **arg2**, [arg3])` with
 * the active argument bolded. Hangs off `document.body`; the caller calls
 * `refresh()` from input/keyup handlers.
 */
export function attachArgHelper(deps: ArgHelperDeps): ArgHelperHandle {
  const { input } = deps;
  let root: HTMLDivElement | null = null;

  const close = (): void => {
    if (!root) return;
    root.remove();
    root = null;
  };

  const positionAboveCaret = (el: HTMLDivElement): void => {
    const rect = input.getBoundingClientRect();
    el.style.left = `${rect.left}px`;
    el.style.top = `${Math.max(0, rect.top - 26)}px`;
  };

  const render = (name: string, args: readonly string[], active: number): void => {
    let el = root;
    if (!el) {
      el = document.createElement('div');
      el.className = 'fc-arghelper';
      el.setAttribute('role', 'tooltip');
      document.body.appendChild(el);
      root = el;
    }
    el.innerHTML = '';
    const head = document.createElement('span');
    head.className = 'fc-arghelper__name';
    head.textContent = `${name}(`;
    el.appendChild(head);
    args.forEach((arg, i) => {
      if (i > 0) {
        const sep = document.createElement('span');
        sep.className = 'fc-arghelper__sep';
        sep.textContent = ', ';
        el.appendChild(sep);
      }
      const span = document.createElement('span');
      span.className = 'fc-arghelper__arg';
      // Anchor the highlight on the last arg when the caret has run past the
      // declared count — Excel does the same so variadic tails stay visible.
      const isActive = i === active || (i === args.length - 1 && active >= args.length);
      if (isActive) span.classList.add('fc-arghelper__arg--active');
      span.textContent = arg;
      el.appendChild(span);
    });
    const tail = document.createElement('span');
    tail.className = 'fc-arghelper__name';
    tail.textContent = ')';
    el.appendChild(tail);
    positionAboveCaret(el);
  };

  const refresh = (): void => {
    const text = input.value;
    const caret = input.selectionStart ?? text.length;
    const sig = findActiveSignature(text, caret);
    if (!sig || sig.args.length === 0) {
      close();
      return;
    }
    render(sig.name, sig.args, sig.activeArgIndex);
  };

  return {
    refresh,
    close,
    detach() {
      close();
    },
  };
}
