import { caretInsideImplicitIntersection, findActiveSignature } from '../commands/refs.js';
import { createFormulaCaretPositioner } from './formula-helper-position.js';
import { overlayPortalFor } from './overlay-portal.js';

let argHelperSeq = 0;

export interface ArgHelperHandle {
  /** Re-evaluate the tooltip against the current input value/caret. */
  refresh(): void;
  /** Hide the tooltip. */
  close(): void;
  /** Hide while an IME composition is in progress. */
  setComposing?(composing: boolean): void;
  setLabels(labels: Partial<ArgHelperLabels>): void;
  detach(): void;
}

export interface ArgHelperLabels {
  implicitIntersection: string;
}

export interface ArgHelperDeps {
  /** The textarea/input being edited. */
  input: HTMLInputElement | HTMLTextAreaElement;
  /** Read the workbook's live formula list separator. */
  getListSeparator?: () => string;
  /** Resolve a localized function alias to its canonical catalog name. */
  resolveFunctionName?: (rawName: string) => string | null;
  /** Resolve a canonical catalog name to the workbook's live display name. */
  getFunctionDisplayName?: (canonicalName: string) => string | null;
  /** Argument labels for a custom function (`CustomFunctionMeta.args`). */
  getCustomFunctionArgs?: (name: string) => readonly string[] | undefined;
  labels?: Partial<ArgHelperLabels>;
}

/**
 * Floating tooltip that mirrors the "ScreenTip" tooltip — when the caret sits
 * inside a known function call, it shows `NAME(arg1, **arg2**, [arg3])` with
 * the active argument bolded. Hangs off the overlay portal; the caller calls
 * `refresh()` from input/keyup handlers.
 */
export function attachArgHelper(deps: ArgHelperDeps): ArgHelperHandle {
  const { input } = deps;
  let labels: ArgHelperLabels = {
    implicitIntersection: 'Implicit intersection',
    ...deps.labels,
  };
  let root: HTMLDivElement | null = null;
  let rootId = '';
  let isComposing = false;
  const positioner = createFormulaCaretPositioner(input);

  const close = (): void => {
    if (!root) return;
    root.remove();
    root = null;
    input.removeAttribute('aria-describedby');
  };

  const positionAboveCaret = (el: HTMLDivElement): void => {
    positioner.position(
      el,
      'above',
      overlayPortalFor(input).querySelector<HTMLElement>('.fc-autocomplete'),
    );
  };

  const render = (
    name: string,
    args: readonly string[],
    active: number,
    implicit: boolean,
  ): void => {
    let el = root;
    if (!el) {
      el = document.createElement('div');
      argHelperSeq += 1;
      rootId = `fc-arghelper-${argHelperSeq}`;
      el.id = rootId;
      el.className = 'fc-arghelper';
      el.setAttribute('role', 'tooltip');
      overlayPortalFor(input).appendChild(el);
      root = el;
    }
    input.setAttribute('aria-describedby', el.id);
    el.replaceChildren();
    if (implicit) {
      const chip = document.createElement('span');
      chip.className = 'fc-arghelper__chip';
      chip.dataset.fcKind = 'implicit-intersection';
      chip.textContent = '@';
      chip.title = labels.implicitIntersection;
      el.appendChild(chip);
    }
    const head = document.createElement('span');
    head.className = 'fc-arghelper__name';
    head.textContent = `${name}(`;
    el.appendChild(head);
    args.forEach((arg, i) => {
      if (i > 0) {
        const sep = document.createElement('span');
        sep.className = 'fc-arghelper__sep';
        sep.textContent = `${deps.getListSeparator?.() === ';' ? ';' : ','} `;
        el.appendChild(sep);
      }
      const span = document.createElement('span');
      span.className = 'fc-arghelper__arg';
      // Anchor the highlight on the last arg when the caret has run past the
      // declared count — desktop spreadsheets do the same so variadic tails stay visible.
      const isActive = i === active || (i === args.length - 1 && active >= args.length);
      if (isActive) {
        span.classList.add('fc-arghelper__arg--active');
        span.setAttribute('aria-current', 'true');
      }
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
    if (isComposing || !positioner.isAnchorVisible()) {
      close();
      return;
    }
    const text = input.value;
    const caret = input.selectionStart ?? text.length;
    const sig = findActiveSignature(text, caret, {
      listSeparator: deps.getListSeparator?.(),
      resolveKnownName: deps.resolveFunctionName,
      getCustomFunctionArgs: deps.getCustomFunctionArgs,
    });
    if (!sig || sig.args.length === 0) {
      close();
      return;
    }
    const implicit = caretInsideImplicitIntersection(text, caret);
    const displayName = deps.getFunctionDisplayName?.(sig.name) ?? sig.name;
    render(displayName, sig.args, sig.activeArgIndex, implicit);
  };

  const setComposing = (composing: boolean): void => {
    if (isComposing === composing) return;
    isComposing = composing;
    if (composing) close();
    else refresh();
  };

  const onCompositionStart = (): void => setComposing(true);
  const onCompositionEnd = (): void => setComposing(false);
  input.addEventListener('compositionstart', onCompositionStart);
  input.addEventListener('compositionend', onCompositionEnd);
  positioner.subscribe(() => {
    if (root) refresh();
  });

  return {
    refresh,
    close,
    setComposing,
    setLabels(next) {
      labels = { ...labels, ...next };
      refresh();
    },
    detach() {
      close();
      input.removeEventListener('compositionstart', onCompositionStart);
      input.removeEventListener('compositionend', onCompositionEnd);
      positioner.dispose();
    },
  };
}
