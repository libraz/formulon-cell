import { vi } from 'vitest';

/** Runs `scenario` while tracking every add/removeEventListener, then reports
 *  how many listeners were registered and how many are still attached to a
 *  target that outlives a detached dialog (document, window, host, or any
 *  still-connected node). Listeners on detached dialog nodes die with them. */
export function trackConnectedListenerLeaks(scenario: () => void): {
  registered: number;
  leaked: number;
} {
  const live = new Map<EventTarget, Set<string>>();
  const ids = new WeakMap<object, number>();
  let registered = 0;
  const key = (type: string, listener: object, options: unknown): string => {
    let id = ids.get(listener);
    if (id === undefined) {
      id = registered++;
      ids.set(listener, id);
    }
    const capture =
      typeof options === 'boolean' ? options : !!(options as { capture?: boolean })?.capture;
    return `${type}:${id}:${capture}`;
  };
  // The DOM environment's EventTarget is not the Node global one; patch the
  // prototype element listeners actually resolve through.
  let proto: EventTarget | null = Object.getPrototypeOf(document.createElement('div'));
  while (proto && !Object.hasOwn(proto, 'addEventListener')) proto = Object.getPrototypeOf(proto);
  if (!proto) throw new Error('EventTarget prototype not found');
  const realAdd = proto.addEventListener;
  const realRemove = proto.removeEventListener;
  const addSpy = vi.spyOn(proto, 'addEventListener').mockImplementation(function (
    this: EventTarget,
    type,
    listener,
    options,
  ) {
    if (listener) {
      const set = live.get(this) ?? new Set<string>();
      set.add(key(type, listener, options));
      live.set(this, set);
    }
    realAdd.call(this, type, listener, options);
  });
  const removeSpy = vi.spyOn(proto, 'removeEventListener').mockImplementation(function (
    this: EventTarget,
    type,
    listener,
    options,
  ) {
    if (listener) live.get(this)?.delete(key(type, listener, options));
    realRemove.call(this, type, listener, options);
  });
  try {
    scenario();
  } finally {
    addSpy.mockRestore();
    removeSpy.mockRestore();
  }
  const leaked = [...live]
    .filter(([target]) => !(target instanceof Node) || target.isConnected)
    .reduce((n, [, set]) => n + set.size, 0);
  return { registered, leaked };
}
