const raw = process.env.FC_TEST_TIMEOUT_SCALE ?? '1';
const scale = Number(raw);
if (!Number.isFinite(scale) || scale <= 0) {
  throw new Error(`FC_TEST_TIMEOUT_SCALE must be a positive number, got "${raw}"`);
}

/** A test or hook timeout multiplied by FC_TEST_TIMEOUT_SCALE (default 1), so a host with
 *  slower cores runs the suite with every limit stretched from one place. */
export const scaledTimeout = (ms: number): number => Math.round(ms * scale);
