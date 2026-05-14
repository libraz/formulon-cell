export function createRequire() {
  return () => {
    throw new Error('Node createRequire is unavailable in browser builds.');
  };
}

export default { createRequire };
