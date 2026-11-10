/** Stops a subscription. Calling it again does nothing. */
export type Disposer = () => void;

/** Reports a failure nobody can catch (a listener or teardown step) through `reportError`, else a microtask rethrow. */
export function reportUncaught(error: unknown): void {
  const report = (globalThis as { reportError?: (error: unknown) => void }).reportError;
  if (typeof report === 'function') {
    report(error);
    return;
  }
  queueMicrotask(() => {
    throw error;
  });
}

/**
 * Typed event emitter behind every public `on()`. Listeners run in subscription order; a throwing
 * listener is reported through {@link reportUncaught} and the rest still run. A listener added
 * during an emit does not see that emit.
 */
export class Emitter<Events extends object> {
  private readonly listeners = new Map<keyof Events, Set<(event: never) => void>>();

  /** Subscribes `listener` to `name`; the returned disposer is idempotent. */
  on<K extends keyof Events>(name: K, listener: (event: Events[K]) => void): Disposer {
    let set = this.listeners.get(name);
    if (!set) {
      set = new Set();
      this.listeners.set(name, set);
    }
    // Wrap so the same function subscribed twice gets two independent subscriptions.
    const entry = (event: never): void => listener(event);
    set.add(entry);
    return () => {
      set.delete(entry);
    };
  }

  /** Calls every listener of `name` with `event`. */
  emit<K extends keyof Events>(name: K, event: Events[K]): void {
    const set = this.listeners.get(name);
    if (!set || set.size === 0) return;
    for (const listener of [...set]) {
      if (!set.has(listener)) continue;
      try {
        (listener as (event: Events[K]) => void)(event);
      } catch (error) {
        reportUncaught(error);
      }
    }
  }

  /** Drops every listener. */
  clear(): void {
    this.listeners.clear();
  }
}
