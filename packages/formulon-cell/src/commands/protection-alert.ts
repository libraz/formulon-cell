import type { WorkbookHandle } from '../engine/workbook-handle.js';

/** Stop-level alert shown when sheet protection refuses a gesture the user could make. */
export interface ProtectionAlert {
  severity: 'stop';
  message: string;
}

export interface ProtectionAlertMessages {
  /** A slicer item was clicked but the sheet refuses the filter it would apply. */
  filterRefused: string;
}

const sinks = new WeakMap<
  WorkbookHandle,
  { sink: (alert: ProtectionAlert) => void; messages: ProtectionAlertMessages }
>();

/** Route protection refusal alerts for `wb` to the host's alert surface, worded in the UI
 *  language. */
export function setProtectionAlertSink(
  wb: WorkbookHandle,
  sink: ((alert: ProtectionAlert) => void) | null,
  messages?: ProtectionAlertMessages,
): void {
  if (sink && messages) sinks.set(wb, { sink, messages });
  else sinks.delete(wb);
}

/** Tell the user the protected sheet refuses the filter a slicer click would apply. */
export function alertFilterRefused(wb: WorkbookHandle): void {
  const entry = sinks.get(wb);
  entry?.sink({ severity: 'stop', message: entry.messages.filterRefused });
}
