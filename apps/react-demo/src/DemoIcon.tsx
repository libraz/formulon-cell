import type { ReactElement } from 'react';
import { DEMO_ICONS, type DemoIconName } from '../../demo-shared/index.js';

export const DemoIcon = ({ name }: { name: DemoIconName }): ReactElement => (
  <svg
    className="fc-tb__rb-icon"
    viewBox="0 0 20 20"
    strokeWidth="1.45"
    strokeLinecap="round"
    strokeLinejoin="round"
    aria-hidden="true"
  >
    {DEMO_ICONS[name].map((segment) => (
      <path
        key={segment.d}
        d={segment.d}
        fill={segment.fill ?? 'none'}
        stroke={segment.stroke ?? 'currentColor'}
      />
    ))}
  </svg>
);
