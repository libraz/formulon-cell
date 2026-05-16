import { ICON_PATHS, type IconName } from '@libraz/formulon-cell';
import { defineComponent, h, type PropType } from 'vue';

export type { IconName };
export { ICON_PATHS };

export const RibbonIcon = defineComponent({
  name: 'RibbonIcon',
  props: {
    name: { type: String as PropType<IconName>, required: true },
  },
  setup(iconProps) {
    return () =>
      h(
        'svg',
        {
          class: 'demo__rb-icon',
          viewBox: '0 0 24 24',
          fill: 'currentColor',
          'aria-hidden': 'true',
        },
        ICON_PATHS[iconProps.name].map((d) => h('path', { d })),
      );
  },
});
