import { afterEach } from 'vitest';

// Mounting the spreadsheet schedules deferred callbacks (requestAnimationFrame,
// setTimeout) that capture ribbon/grid DOM. Under happy-dom these sit in the
// environment's task queue — a GC root — and only fire when the event loop
// turns. Tests run synchronously back to back, so without an explicit turn the
// queue grows unbounded across a file and exhausts the heap. Yielding once per
// test drains it, keeping heap flat. In a real browser these callbacks fire
// every frame, so this is a test-environment concern only.
afterEach(async () => {
  await new Promise((resolve) => setTimeout(resolve));
});
