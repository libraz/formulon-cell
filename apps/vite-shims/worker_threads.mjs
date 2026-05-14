export const workerData = undefined;
export const parentPort = undefined;

export class Worker {
  constructor() {
    throw new Error('Node worker_threads is unavailable in browser builds.');
  }
}
