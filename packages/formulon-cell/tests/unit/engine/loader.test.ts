import { afterEach, describe, expect, it, vi } from 'vitest';

describe('engine/loader', () => {
  afterEach(() => {
    vi.doUnmock('@libraz/formulon');
    vi.unstubAllGlobals();
    vi.resetModules();
  });

  it('fails loudly when SharedArrayBuffer is unavailable', async () => {
    const createFormulon = vi.fn();
    vi.doMock('@libraz/formulon', () => ({ default: createFormulon }));
    vi.stubGlobal('SharedArrayBuffer', undefined);
    vi.resetModules();

    const { loadFormulon } = await import('../../../src/engine/loader.js');

    await expect(loadFormulon()).rejects.toThrow(/SharedArrayBuffer is missing/);
    expect(createFormulon).not.toHaveBeenCalled();
  });

  it('propagates WASM initialization failures instead of falling back to stub', async () => {
    const createFormulon = vi.fn(() => Promise.reject(new Error('boom')));
    vi.doMock('@libraz/formulon', () => ({ default: createFormulon }));
    vi.stubGlobal('SharedArrayBuffer', class FakeSharedArrayBuffer {});
    vi.resetModules();

    const { isUsingStub, loadFormulon } = await import('../../../src/engine/loader.js');

    await expect(loadFormulon()).rejects.toThrow(/formulon WASM init failed: Error: boom/);
    expect(isUsingStub()).toBe(false);
  });

  it('still allows the explicit test/demo stub path', async () => {
    const createFormulon = vi.fn();
    vi.doMock('@libraz/formulon', () => ({ default: createFormulon }));
    vi.stubGlobal('SharedArrayBuffer', undefined);
    vi.resetModules();

    const { isUsingStub, loadFormulon } = await import('../../../src/engine/loader.js');
    const onFallback = vi.fn();

    const module = await loadFormulon({ preferStub: true, onFallback });

    expect(module.versionString()).toBe('stub');
    expect(isUsingStub()).toBe(true);
    expect(onFallback).toHaveBeenCalledWith('preferStub set — using stub engine');
    expect(createFormulon).not.toHaveBeenCalled();
  });
});
