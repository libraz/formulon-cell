import type { ExcelRibbonIconName } from '@libraz/formulon-cell';
import { type ReactElement, useLayoutEffect, useRef } from 'react';
import { createDemoGlyph } from '../../demo-shared/index.js';

/** Mounts a core icon-kit glyph; the element is built by the core renderer. */
export const DemoGlyph = ({ name }: { name: ExcelRibbonIconName }): ReactElement => {
  const ref = useRef<HTMLSpanElement | null>(null);
  useLayoutEffect(() => {
    const svg = createDemoGlyph(name);
    ref.current?.replaceChildren(...(svg ? [svg] : []));
  }, [name]);
  return <span ref={ref} className="demo__glyph" aria-hidden="true" />;
};
