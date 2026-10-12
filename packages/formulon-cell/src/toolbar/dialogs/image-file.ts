export interface PickImageFileResult {
  src: string;
  alt: string;
  /** Original bytes for native drawing insertion. Keeping this alongside the
   * data URL avoids a lossy decode/re-encode before save. */
  bytes: Uint8Array;
}

export interface PickImageFileOptions {
  accept?: string;
  document?: Document;
}

export function pickImageFileDataUrl(
  options: PickImageFileOptions = {},
): Promise<PickImageFileResult | null> {
  const doc = options.document ?? globalThis.document;
  if (!doc) return Promise.resolve(null);
  return new Promise((resolve) => {
    const input = doc.createElement('input');
    let settled = false;
    input.type = 'file';
    input.accept = options.accept ?? 'image/*';
    input.style.position = 'fixed';
    input.style.left = '-9999px';
    input.style.top = '0';

    const cleanup = (): void => {
      input.removeEventListener('change', onChange);
      input.removeEventListener('cancel', onCancel);
      input.remove();
    };
    const finish = (result: PickImageFileResult | null): void => {
      if (settled) return;
      settled = true;
      cleanup();
      resolve(result);
    };
    const onCancel = (): void => finish(null);
    const onChange = (): void => {
      const file = input.files?.[0];
      if (!file) {
        finish(null);
        return;
      }
      const dataReader = new FileReader();
      const bytesReader = new FileReader();
      let src = '';
      let bytes: Uint8Array | null = null;
      let failed = false;
      const canReadBytes =
        typeof (bytesReader as FileReader & { readAsArrayBuffer?: unknown }).readAsArrayBuffer ===
        'function';
      const complete = (): void => {
        if (failed) return;
        if (src && (bytes || !canReadBytes)) {
          finish({ src, alt: file.name, bytes: bytes ?? new Uint8Array() });
        }
      };
      dataReader.addEventListener('load', () => {
        src = typeof dataReader.result === 'string' ? dataReader.result : '';
        complete();
      });
      if (canReadBytes) {
        bytesReader.addEventListener('load', () => {
          bytes =
            bytesReader.result instanceof ArrayBuffer ? new Uint8Array(bytesReader.result) : null;
          complete();
        });
      }
      const onError = (): void => {
        failed = true;
        finish(null);
      };
      dataReader.addEventListener('error', onError);
      if (canReadBytes) bytesReader.addEventListener('error', onError);
      dataReader.readAsDataURL(file);
      if (canReadBytes) bytesReader.readAsArrayBuffer(file);
    };

    input.addEventListener('change', onChange);
    input.addEventListener('cancel', onCancel);
    doc.body?.appendChild(input);
    input.click();
  });
}
