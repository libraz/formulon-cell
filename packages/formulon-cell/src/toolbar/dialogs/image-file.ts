export interface PickImageFileResult {
  src: string;
  alt: string;
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
      const reader = new FileReader();
      reader.addEventListener('load', () => {
        const src = typeof reader.result === 'string' ? reader.result : '';
        finish(src ? { src, alt: file.name } : null);
      });
      reader.addEventListener('error', () => finish(null));
      reader.readAsDataURL(file);
    };

    input.addEventListener('change', onChange);
    input.addEventListener('cancel', onCancel);
    doc.body?.appendChild(input);
    input.click();
  });
}
