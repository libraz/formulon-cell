export type ExcelRibbonIconSegment = {
  d: string;
  fill?: string;
  stroke?: string;
  strokeWidth?: string;
  strokeLinecap?: 'butt' | 'round' | 'square';
  strokeLinejoin?: 'arcs' | 'bevel' | 'miter' | 'miter-clip' | 'round';
  strokeDasharray?: string;
};

export type ExcelRibbonIconDefinition = readonly ExcelRibbonIconSegment[];

const ink = '#1f1f1f';
const grid = '#8a8f98';
const gridLight = '#d9d9d9';
const excelGreen = '#107c41';
const excelBlue = '#2f75b5';
const excelRed = '#c00000';
const excelYellow = '#ffd966';
const excelOrange = '#ed7d31';
const clipboardTan = '#f4c27a';

export const EXCEL_RIBBON_ICON_PATHS = {
  undo: [
    {
      d: 'M5.5 9h8a5 5 0 0 1 0 10H9.5',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.7',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
    {
      d: 'M9.75 4.75 5.5 9l4.25 4.25',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.9',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  redo: [
    {
      d: 'M18.5 9h-8a5 5 0 0 0 0 10h4.5',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.7',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
    {
      d: 'M14.25 4.75 18.5 9l-4.25 4.25',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.9',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  paste: [
    { d: 'M6.5 5.5h11v14h-11z', fill: '#ffffff', stroke: ink, strokeWidth: '1.35' },
    { d: 'M8.5 4.5h7v3h-7z', fill: clipboardTan, stroke: '#9a6a2f', strokeWidth: '1.15' },
    { d: 'M10.25 3.7h3.5v1.5h-3.5z', fill: '#f7d99b', stroke: '#9a6a2f', strokeWidth: '1.1' },
    {
      d: 'M9 10.75h6M9 13.25h6M9 15.75h4',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.2',
      strokeLinecap: 'round',
    },
  ],
  pasteFormulas: [
    { d: 'M6.5 5.5h11v14h-11z', fill: '#ffffff', stroke: ink, strokeWidth: '1.35' },
    { d: 'M8.5 4.5h7v3h-7z', fill: clipboardTan, stroke: '#9a6a2f', strokeWidth: '1.15' },
    {
      d: 'M9.75 16.8v-5.2c0-1 .55-1.5 1.45-1.4M8.55 12.9h2.9M12.55 12.4l2.9 4.4M15.45 12.4l-2.9 4.4',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.3',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  pasteValues: [
    { d: 'M6.5 5.5h11v14h-11z', fill: '#ffffff', stroke: ink, strokeWidth: '1.35' },
    { d: 'M8.5 4.5h7v3h-7z', fill: clipboardTan, stroke: '#9a6a2f', strokeWidth: '1.15' },
    {
      d: 'M8.15 12.6 9.05 11.8v5M10.75 12.7a1.05 1.05 0 0 1 2 .5c0 1.15-2 1.85-2 3.5h2.1M14.15 11.8h1.7l-1.05 1.7a1.2 1.2 0 1 1-1 1.85',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.25',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  pasteTranspose: [
    { d: 'M6.5 5.5h11v14h-11z', fill: '#ffffff', stroke: ink, strokeWidth: '1.35' },
    { d: 'M8.5 4.5h7v3h-7z', fill: clipboardTan, stroke: '#9a6a2f', strokeWidth: '1.15' },
    {
      d: 'M9.95 16.6v-4.3a1.3 1.3 0 0 1 1.3-1.3h4.2M14.05 10.7l1.9 1.6-1.9 1.6M8.05 14.7l1.9 1.9 1.9-1.9',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.35',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  pasteSpecial: [
    { d: 'M6.5 5.5h11v14h-11z', fill: '#ffffff', stroke: ink, strokeWidth: '1.35' },
    { d: 'M8.5 4.5h7v3h-7z', fill: clipboardTan, stroke: '#9a6a2f', strokeWidth: '1.15' },
    {
      d: 'M9 10.5h6M9 12.75h4',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.2',
      strokeLinecap: 'round',
    },
    { d: 'M13.5 14.75h5v5h-5z', fill: '#ffffff', stroke: ink, strokeWidth: '1.25' },
    {
      d: 'M14.9 16.4h2.2M14.9 18.1h2.2',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.2',
      strokeLinecap: 'round',
    },
  ],
  cut: [
    {
      d: 'M6.1 17.1m-2.6 0a2.6 2.6 0 1 0 5.2 0a2.6 2.6 0 1 0 -5.2 0',
      fill: excelBlue,
    },
    {
      d: 'M17.9 17.1m-2.6 0a2.6 2.6 0 1 0 5.2 0a2.6 2.6 0 1 0 -5.2 0',
      fill: excelBlue,
    },
    {
      d: 'M6.1 17.1m-1.25 0a1.25 1.25 0 1 0 2.5 0a1.25 1.25 0 1 0 -2.5 0',
      fill: '#ffffff',
    },
    {
      d: 'M17.9 17.1m-1.25 0a1.25 1.25 0 1 0 2.5 0a1.25 1.25 0 1 0 -2.5 0',
      fill: '#ffffff',
    },
    {
      d: 'M8 15.6 11.1 12.5M16 15.6 12.9 12.5',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.55',
      strokeLinecap: 'round',
    },
    {
      d: 'M12.3 11.5 19.2 5.1M11.7 11.5 4.8 5.1',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.75',
      strokeLinecap: 'round',
    },
    {
      d: 'M12 12m-1.3 0a1.3 1.3 0 1 0 2.6 0a1.3 1.3 0 1 0 -2.6 0',
      fill: '#ffffff',
      stroke: ink,
      strokeWidth: '1.1',
    },
  ],
  copy: [
    { d: 'M5.6 4.6h9.4v10.8H5.6z', fill: '#f4f5f7', stroke: grid, strokeWidth: '1.25' },
    { d: 'M8.7 7.1h9.7v11.8H8.7z', fill: '#ffffff', stroke: ink, strokeWidth: '1.35' },
    {
      d: 'M10.6 10.2h5.9M10.6 12.8h5.9M10.6 15.4h4.2',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.15',
      strokeLinecap: 'round',
    },
  ],
  paint: [
    {
      d: 'M6 19.6c3.4-1.6 8.6-1.6 12 0',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.8',
      strokeLinecap: 'round',
    },
    { d: 'M10.5 4.6h3v5.9h-3z', fill: clipboardTan, stroke: '#9a6a2f', strokeWidth: '1.2' },
    { d: 'M9.4 10.5h5.2v2.1H9.4z', fill: '#f3f2f1', stroke: grid, strokeWidth: '1.2' },
    {
      d: 'M8.9 12.6h6.2v3.4a1.2 1.2 0 0 1-1.2 1.2h-3.8a1.2 1.2 0 0 1-1.2-1.2z',
      fill: excelBlue,
      stroke: ink,
      strokeWidth: '1.2',
      strokeLinejoin: 'round',
    },
  ],
  borders: [
    { d: 'M9.5 5.5v13M14.5 5.5v13M5.5 9.5h13M5.5 14.5h13', fill: 'none', stroke: grid },
    { d: 'M5.5 5.5h13v13h-13z', fill: 'none', stroke: ink, strokeWidth: '1.5' },
  ],
  fillColor: [
    {
      d: 'M9 7.9a3 3 0 0 1 6 0',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.3',
      strokeLinecap: 'round',
    },
    {
      d: 'M6.5 7.9h11l-1 7.6a1.6 1.6 0 0 1-1.6 1.4h-5.8a1.6 1.6 0 0 1-1.6-1.4z',
      fill: '#ffffff',
      stroke: ink,
      strokeWidth: '1.35',
      strokeLinejoin: 'round',
    },
    {
      d: 'M7.4 11.6h9.2l-.5 3.9a1 1 0 0 1-1 .9h-6.2a1 1 0 0 1-1-.9z',
      fill: excelYellow,
    },
    { d: 'M5 18.4h14v2H5z', fill: excelYellow },
    { d: 'M5 18.4h14v2H5z', fill: 'none', stroke: '#b79500' },
  ],
  fontColor: [
    { d: 'M5.5 16.5 10.4 4h3.2L18.5 16.5h-2.2l-1.1-3h-6.4l-1.1 3H5.5Z', fill: ink },
    { d: 'M9.5 11.5h5L12 4.6 9.5 11.5Z', fill: '#ffffff' },
    { d: 'M5 18.4h14v2H5z', fill: excelRed },
  ],
  fontGrow: [
    { d: 'M3.75 17.5 7.45 6h2.2l3.6 11.5h-1.9l-.8-2.5H6.45l-.8 2.5H3.75Z', fill: ink },
    { d: 'M6.95 13.4h3.9L8.9 7.5 6.95 13.4Z', fill: '#ffffff' },
    {
      d: 'M17.35 15.5V7.5M15.25 9.8l2.1-2.3 2.1 2.3',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.6',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  fontShrink: [
    { d: 'M3.75 17.5 7.45 6h2.2l3.6 11.5h-1.9l-.8-2.5H6.45l-.8 2.5H3.75Z', fill: ink },
    { d: 'M6.95 13.4h3.9L8.9 7.5 6.95 13.4Z', fill: '#ffffff' },
    {
      d: 'M17.35 7.5v8M15.25 13.2l2.1 2.3 2.1-2.3',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.6',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  bold: [
    {
      d: 'M7 5h6.1c2.4 0 3.9 1.2 3.9 3.2 0 1.3-.7 2.3-1.8 2.8 1.5.4 2.5 1.5 2.5 3.2 0 2.3-1.8 3.8-4.5 3.8H7Z',
      fill: ink,
    },
    { d: 'M10 7.4v2.7h2.7c.9 0 1.4-.5 1.4-1.3s-.6-1.4-1.5-1.4Z', fill: '#ffffff' },
    { d: 'M10 12.5v3.1h3c1 0 1.6-.6 1.6-1.5s-.6-1.6-1.7-1.6Z', fill: '#ffffff' },
  ],
  italic: [{ d: 'M10.1 5h8.1v2.2h-2.6l-3.5 8.6h2.5V18H6.5v-2.2h2.7l3.5-8.6h-2.6Z', fill: ink }],
  underline: [
    {
      d: 'M7.1 4.3h3v7c0 1.7.7 2.6 2 2.6s2-.9 2-2.6v-7h3v7.1c0 3.3-1.9 5.1-5 5.1s-5-1.8-5-5.1Z',
      fill: ink,
    },
    { d: 'M5.7 18.4h12.6v2H5.7z', fill: excelGreen },
  ],
  underlineSingle: [
    {
      d: 'M7.1 4.3h3v7c0 1.7.7 2.6 2 2.6s2-.9 2-2.6v-7h3v7.1c0 3.3-1.9 5.1-5 5.1s-5-1.8-5-5.1Z',
      fill: ink,
    },
    { d: 'M5.7 18.4h12.6v2H5.7z', fill: excelGreen },
  ],
  underlineDouble: [
    {
      d: 'M7.1 4.3h3v7c0 1.7.7 2.6 2 2.6s2-.9 2-2.6v-7h3v7.1c0 3.3-1.9 5.1-5 5.1s-5-1.8-5-5.1Z',
      fill: ink,
    },
    { d: 'M5.7 17.4h12.6v1.4H5.7z', fill: excelGreen },
    { d: 'M5.7 19.6h12.6v1.4H5.7z', fill: excelGreen },
  ],
  strike: [
    {
      d: 'M7.2 8.7c0-2 1.9-3.2 4.8-3.2 2.2 0 3.8.7 5 2l-1.7 1.7c-.8-.8-1.9-1.2-3.4-1.2-1.1 0-1.8.3-1.8.8 0 .6.8.9 2.9 1.4 2.8.7 4.4 1.8 4.4 4.1 0 2.4-2.2 4-5.3 4-2.6 0-4.5-.8-5.8-2.5l1.9-1.6c1 1.1 2.2 1.6 3.8 1.6 1.4 0 2.3-.5 2.3-1.3 0-.7-.7-1.1-2.7-1.6-2.8-.7-4.4-1.7-4.4-4.2Z',
      fill: ink,
    },
    { d: 'M5.6 10.7h12.8v2.4H5.6z', fill: excelGreen },
  ],
  currency: [
    { d: 'M4.5 6.5h15v11h-15z', fill: '#ffffff', stroke: ink, strokeWidth: '1.3' },
    { d: 'M6.5 8.5h11v7h-11z', fill: '#eef6ee', stroke: gridLight, strokeWidth: '1' },
    {
      d: 'M9.4 9.4 12 12.2l2.6-2.8M12 12.2v3M10.1 12.3h3.8M10.1 13.9h3.8',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.45',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  percent: [
    {
      d: 'M7.4 6.8a2.5 2.5 0 1 0 0 5 2.5 2.5 0 1 0 0-5',
      fill: '#ffffff',
      stroke: ink,
      strokeWidth: '1.7',
    },
    {
      d: 'M16.6 12.2a2.5 2.5 0 1 0 0 5 2.5 2.5 0 1 0 0-5',
      fill: '#ffffff',
      stroke: ink,
      strokeWidth: '1.7',
    },
    {
      d: 'M17.7 5.8 6.3 18.8',
      fill: 'none',
      stroke: ink,
      strokeWidth: '2',
      strokeLinecap: 'round',
    },
  ],
  comma: [
    {
      d: 'M12.6 8.55m-2.5 0a2.5 2.5 0 1 0 5 0a2.5 2.5 0 1 0-5 0',
      fill: ink,
    },
    {
      d: 'M13.5 10.55c.7 2.6-.6 4.7-3.2 6',
      fill: 'none',
      stroke: ink,
      strokeWidth: '2.8',
      strokeLinecap: 'round',
    },
  ],
  decDown: [
    {
      d: 'M4.6 14.7m-.95 0a.95 .95 0 1 0 1.9 0a.95 .95 0 1 0-1.9 0',
      fill: ink,
    },
    {
      d: 'M8.1 9.5a1.7 2.7 0 1 0 0 5.4 1.7 2.7 0 1 0 0-5.4M12.9 9.5a1.7 2.7 0 1 0 0 5.4 1.7 2.7 0 1 0 0-5.4',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.25',
    },
    {
      d: 'M18.2 15.4V8.6M16.4 13.6l1.8 1.8 1.8-1.8',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.6',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  decUp: [
    {
      d: 'M4.6 14.7m-.95 0a.95 .95 0 1 0 1.9 0a.95 .95 0 1 0-1.9 0',
      fill: ink,
    },
    {
      d: 'M8.1 9.5a1.7 2.7 0 1 0 0 5.4 1.7 2.7 0 1 0 0-5.4M12.9 9.5a1.7 2.7 0 1 0 0 5.4 1.7 2.7 0 1 0 0-5.4',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.25',
    },
    {
      d: 'M18.2 8.6v6.8M16.4 10.4l1.8-1.8 1.8 1.8',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.6',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  autosum: [
    {
      d: 'M5.6 4.9h12.8v2.4H9.6l5.9 4.7-5.9 4.7h8.8v2.4H5.6v-2.2l7-4.9-7-4.9V4.9Z',
      fill: excelBlue,
    },
  ],
  fill: [
    { d: 'M4.5 6h10v10h-10z', fill: '#ffffff', stroke: grid, strokeWidth: '1.15' },
    {
      d: 'M4.5 9.33h10M4.5 12.67h10M7.83 6v10M11.17 6v10',
      fill: 'none',
      stroke: gridLight,
    },
    {
      d: 'M17.5 7v9.5M14.9 13.9l2.6 2.6 2.6-2.6',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.6',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  fillDown: [
    { d: 'M6 5h12v14H6z', fill: '#ffffff', stroke: grid, strokeWidth: '1.1' },
    {
      d: 'M6 8.5h12M6 12h12M6 15.5h12M10 5v14M14 5v14',
      fill: 'none',
      stroke: gridLight,
    },
    {
      d: 'M12 7.8v8.4M9.4 13.6l2.6 2.6 2.6-2.6',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.45',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  fillRight: [
    { d: 'M6 5h12v14H6z', fill: '#ffffff', stroke: grid, strokeWidth: '1.1' },
    {
      d: 'M6 8.5h12M6 12h12M6 15.5h12M10 5v14M14 5v14',
      fill: 'none',
      stroke: gridLight,
    },
    {
      d: 'M7.8 12h8.4M13.6 9.4l2.6 2.6-2.6 2.6',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.45',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  fillUp: [
    { d: 'M6 5h12v14H6z', fill: '#ffffff', stroke: grid, strokeWidth: '1.1' },
    {
      d: 'M6 8.5h12M6 12h12M6 15.5h12M10 5v14M14 5v14',
      fill: 'none',
      stroke: gridLight,
    },
    {
      d: 'M12 16.2V7.8M9.4 10.4 12 7.8l2.6 2.6',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.45',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  fillLeft: [
    { d: 'M6 5h12v14H6z', fill: '#ffffff', stroke: grid, strokeWidth: '1.1' },
    {
      d: 'M6 8.5h12M6 12h12M6 15.5h12M10 5v14M14 5v14',
      fill: 'none',
      stroke: gridLight,
    },
    {
      d: 'M16.2 12H7.8M10.4 9.4 7.8 12l2.6 2.6',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.45',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  fillGroup: [
    { d: 'M7.5 5.5h11v10', fill: 'none', stroke: gridLight, strokeWidth: '1.3' },
    { d: 'M5 7.5h11v10H5z', fill: '#ffffff', stroke: grid, strokeWidth: '1.2' },
    { d: 'M5 11h11M5 14.5h11M8.5 7.5v10M12 7.5v10', fill: 'none', stroke: gridLight },
  ],
  fillSeries: [
    { d: 'M5.5 5.5h8v13h-8z', fill: '#ffffff', stroke: grid, strokeWidth: '1.1' },
    {
      d: 'M7.4 8.2h4.5M7.4 11.8h4.5M7.4 15.4h4.5',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.2',
      strokeLinecap: 'round',
    },
    {
      d: 'M16.2 8.2v8.6M13.7 14.4l2.5 2.5 2.5-2.5',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.35',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  fillJustify: [
    { d: 'M5.5 6.2h13v11.5h-13z', fill: '#ffffff', stroke: grid, strokeWidth: '1.1' },
    {
      d: 'M7.5 8.9h9M7.5 11.4h9M7.5 13.9h9M7.5 16.4h7',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.05',
      strokeLinecap: 'round',
    },
  ],
  flashFill: [
    { d: 'M4.5 5.5h15v13h-15z', fill: '#ffffff', stroke: grid, strokeWidth: '1.15' },
    {
      d: 'M4.5 9.75h15M4.5 14h15M9.5 5.5v13M14.5 5.5v13',
      fill: 'none',
      stroke: gridLight,
    },
    {
      d: 'M13.4 6.4 9.2 12.5h2.6l-1.2 5.1 4.6-6.6h-2.7z',
      fill: '#ffffff',
      stroke: '#ffffff',
      strokeWidth: '2.4',
      strokeLinejoin: 'round',
    },
    {
      d: 'M13.4 6.4 9.2 12.5h2.6l-1.2 5.1 4.6-6.6h-2.7z',
      fill: excelOrange,
      stroke: '#b35a00',
      strokeWidth: '0.8',
      strokeLinejoin: 'round',
    },
  ],
  clear: [
    {
      d: 'M18.5 11.3 15.2 5.7 5.5 11.3 8.8 16.9z',
      fill: '#ffffff',
      stroke: ink,
      strokeWidth: '1.3',
      strokeLinejoin: 'round',
    },
    { d: 'M11.8 7.7 15.1 13.3', fill: 'none', stroke: grid, strokeWidth: '1.15' },
    {
      d: 'M5 19.5h14',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.3',
      strokeLinecap: 'round',
    },
  ],
  clearAll: [
    {
      d: 'M18.5 11.3 15.2 5.7 5.5 11.3 8.8 16.9z',
      fill: '#f7e1ff',
      stroke: ink,
      strokeWidth: '1.3',
      strokeLinejoin: 'round',
    },
    { d: 'M11.8 7.7 15.1 13.3', fill: 'none', stroke: '#8a4fb7', strokeWidth: '1.15' },
    {
      d: 'M5 19.5h14',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.3',
      strokeLinecap: 'round',
    },
  ],
  clearFormats: [
    {
      d: 'M16.4 13.3 13.3 8.1 4.5 13.3l3.1 5.2z',
      fill: '#f7e1ff',
      stroke: ink,
      strokeWidth: '1.3',
      strokeLinejoin: 'round',
    },
    { d: 'M10.2 9.9 13.3 15.1', fill: 'none', stroke: '#8a4fb7', strokeWidth: '1.15' },
    {
      d: 'M13.9 9.4 16.4 4.2l2.5 5.2M14.8 7.6h3.2',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.25',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  clearContents: [
    { d: 'M5.5 5.5h13v13h-13z', fill: '#ffffff', stroke: grid, strokeWidth: '1.1' },
    {
      d: 'M8 8.4h8M8 11.7h6.2M8 15h7.2',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.35',
      strokeLinecap: 'round',
    },
    {
      d: 'M14.9 4.9 19 9M19 4.9 14.9 9',
      fill: 'none',
      stroke: excelRed,
      strokeWidth: '1.35',
      strokeLinecap: 'round',
    },
  ],
  clearComments: [
    {
      d: 'M6 6h11.5v8.4H10l-3.1 3.1v-3.1H6z',
      fill: '#fff8cc',
      stroke: '#b79500',
      strokeWidth: '1.15',
    },
    {
      d: 'M8.4 8.8h6.3M8.4 11.2h4.5',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.05',
      strokeLinecap: 'round',
    },
    {
      d: 'M14.9 4.9 19 9M19 4.9 14.9 9',
      fill: 'none',
      stroke: excelRed,
      strokeWidth: '1.35',
      strokeLinecap: 'round',
    },
  ],
  clearHyperlinks: [
    {
      d: 'M6.2 13.6a3.2 3.2 0 0 1 0-4.5l1.7-1.7a3.2 3.2 0 0 1 4.5 0',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.45',
      strokeLinecap: 'round',
    },
    {
      d: 'M11.6 16.6a3.2 3.2 0 0 0 4.5 0l1.7-1.7a3.2 3.2 0 0 0 0-4.5',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.45',
      strokeLinecap: 'round',
    },
    {
      d: 'M9.6 14.4 14.4 9.6',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.25',
      strokeLinecap: 'round',
    },
    {
      d: 'M14.9 4.9 19 9M19 4.9 14.9 9',
      fill: 'none',
      stroke: excelRed,
      strokeWidth: '1.35',
      strokeLinecap: 'round',
    },
  ],
  clearConditional: [
    { d: 'M5.5 5.5h13v13h-13z', fill: '#ffffff', stroke: grid, strokeWidth: '1.1' },
    { d: 'M7 7h10v3.2H7z', fill: '#63be7b' },
    { d: 'M7 10.2h10v3.2H7z', fill: '#ffeb84' },
    { d: 'M7 13.4h10v3.2H7z', fill: '#f8696b' },
    {
      d: 'M14.9 4.9 19 9M19 4.9 14.9 9',
      fill: 'none',
      stroke: excelRed,
      strokeWidth: '1.35',
      strokeLinecap: 'round',
    },
  ],
  sortAsc: [
    { d: 'M4.7 11 7.3 4.5h1.6L11.5 11H9.8l-.5-1.4H6.9L6.4 11H4.7Z', fill: ink },
    { d: 'M7.3 8.3h1.6L8.1 6.1 7.3 8.3Z', fill: '#ffffff' },
    { d: 'M4.9 12.6h6.4v1.3l-4.1 4.2h4.3v1.4H4.8v-1.3l4.1-4.2H4.9v-1.4Z', fill: excelBlue },
    {
      d: 'M16.2 5.5v12M13.6 14.9l2.6 2.6 2.6-2.6',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.6',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  sortDesc: [
    { d: 'M4.9 4.5h6.4v1.3l-4.1 4.2h4.3v1.4H4.8V10.1l4.1-4.2H4.9V4.5Z', fill: excelBlue },
    { d: 'M4.7 19.5 7.3 13h1.6l2.6 6.5H9.8l-.5-1.4H6.9l-.5 1.4H4.7Z', fill: ink },
    { d: 'M7.3 16.8h1.6L8.1 14.6 7.3 16.8Z', fill: '#ffffff' },
    {
      d: 'M16.2 5.5v12M13.6 14.9l2.6 2.6 2.6-2.6',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.6',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  sortFilter: [
    {
      d: 'M4.5 6h11l-4.1 4.9v5l-2.8-1.7v-3.3L4.5 6Z',
      fill: excelGreen,
      stroke: '#0b5a2f',
      strokeWidth: '1.1',
      strokeLinejoin: 'round',
    },
    {
      d: 'M18 7v10M16.2 8.8 18 7l1.8 1.8M16.2 15.2 18 17l1.8-1.8',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.35',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  sortCustom: [
    { d: 'M4.5 6h9v9h-9z', fill: '#ffffff', stroke: grid, strokeWidth: '1.1' },
    { d: 'M4.5 9h9M4.5 12h9M7.5 6v9M10.5 6v9', fill: 'none', stroke: gridLight },
    {
      d: 'M17 5.7v11.8M14.9 7.8 17 5.7l2.1 2.1M14.9 15.4l2.1 2.1 2.1-2.1',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.25',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  filterToggle: [
    {
      d: 'M5.5 6h13l-5 5.8v5.3l-3 1.5v-6.8L5.5 6Z',
      fill: '#ffffff',
      stroke: ink,
      strokeWidth: '1.2',
      strokeLinejoin: 'round',
    },
  ],
  filterByValue: [
    {
      d: 'M4.5 6h11l-4.2 5v4.6l-2.6 1.3V11L4.5 6Z',
      fill: '#ffffff',
      stroke: ink,
      strokeWidth: '1.15',
      strokeLinejoin: 'round',
    },
    {
      d: 'M14.5 14.6h5M14.5 17.2h5',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.4',
      strokeLinecap: 'round',
    },
  ],
  filterClear: [
    {
      d: 'M4.5 6h11l-4.2 5v4.6l-2.6 1.3V11L4.5 6Z',
      fill: '#ffffff',
      stroke: '#b8b8b8',
      strokeWidth: '1.15',
      strokeLinejoin: 'round',
    },
    {
      d: 'M14.9 13.9 19 18M19 13.9 14.9 18',
      fill: 'none',
      stroke: excelRed,
      strokeWidth: '1.4',
      strokeLinecap: 'round',
    },
  ],
  filterReapply: [
    {
      d: 'M4.5 6h11l-4.2 5v4.6l-2.6 1.3V11L4.5 6Z',
      fill: '#ffffff',
      stroke: '#b8b8b8',
      strokeWidth: '1.15',
      strokeLinejoin: 'round',
    },
    {
      d: 'M19.4 15.9a2.9 2.9 0 1 1-1.1-2.3M18.8 11.2v2.5h-2.5',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.3',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  filterAdvanced: [
    {
      d: 'M4.5 6h11l-4.2 5v4.6l-2.6 1.3V11L4.5 6Z',
      fill: '#ffffff',
      stroke: ink,
      strokeWidth: '1.15',
      strokeLinejoin: 'round',
    },
    {
      d: 'M14.5 13.6h5M14.5 16h5M14.5 18.4h5',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.15',
      strokeLinecap: 'round',
    },
  ],
  find: [
    {
      d: 'M5.5 10.5a5 5 0 1 0 10 0 5 5 0 1 0-10 0',
      fill: '#ffffff',
      stroke: ink,
      strokeWidth: '1.4',
    },
    {
      d: 'M14.2 14.2 19 19',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.8',
      strokeLinecap: 'round',
    },
  ],
  replaceFind: [
    {
      d: 'M5.5 9.5a4 4 0 1 0 8 0 4 4 0 1 0-8 0',
      fill: '#ffffff',
      stroke: ink,
      strokeWidth: '1.4',
    },
    {
      d: 'M12.4 12.4 15 15',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.7',
      strokeLinecap: 'round',
    },
    {
      d: 'M13.8 20h5.6l-1.7-1.7M19.4 16.2h-5.6l1.7 1.7',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.4',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  findFormulas: [
    {
      d: 'M5.5 6.5h13v11.8h-13z',
      fill: '#ffffff',
      stroke: grid,
      strokeWidth: '1.1',
    },
    {
      d: 'M9.4 16.4V11c0-1 .55-1.5 1.45-1.4M8.2 12.4h2.9M12.4 11.9l2.9 4.5M15.3 11.9l-2.9 4.5',
      fill: 'none',
      stroke: '#8764b8',
      strokeWidth: '1.3',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  findConstants: [
    { d: 'M5.5 6.5h13v11.8h-13z', fill: '#ffffff', stroke: grid, strokeWidth: '1.1' },
    {
      d: 'M8.75 9h1.6v7H8.85v-5.3l-1.1.7-.6-1.1L8.75 9ZM12.35 9h2.2c1.4 0 2.3.7 2.3 1.8 0 .7-.4 1.2-1 1.5.8.3 1.2.9 1.2 1.8 0 1.2-1 2-2.5 2h-2.2v-1.2h2.1c.7 0 1.1-.3 1.1-.9s-.4-.9-1.1-.9h-1.3v-1.1h1.2c.6 0 1-.3 1-.8s-.4-.8-1-.8h-2V9Z',
      fill: excelGreen,
    },
  ],
  findConditional: [
    { d: 'M5.5 6.5h13v11.8h-13z', fill: '#ffffff', stroke: grid, strokeWidth: '1.1' },
    { d: 'M7 8h10v2.7H7z', fill: '#63be7b' },
    { d: 'M7 10.9h10v2.7H7z', fill: '#ffeb84' },
    { d: 'M7 13.8h10v2.7H7z', fill: '#f8696b' },
  ],
  findValidation: [
    { d: 'M5.5 6.5h13v11.8h-13z', fill: '#ffffff', stroke: grid, strokeWidth: '1.1' },
    {
      d: 'M8 9.4h5.7M8 12h4.6M8 14.6h3.5',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.05',
      strokeLinecap: 'round',
    },
    {
      d: 'M12.8 13.9 14.6 15.7 17.7 12.2',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.45',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  findComments: [
    {
      d: 'M5.8 7h11.8v8.2h-7.1l-3.3 2.9v-2.9H5.8z',
      fill: '#fff8cc',
      stroke: '#b79500',
      strokeWidth: '1.1',
    },
    {
      d: 'M8.2 9.6h6.8M8.2 12h4.9',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.05',
      strokeLinecap: 'round',
    },
  ],
  top: [
    { d: 'M4.2 5h15.6v3H4.2z', fill: excelGreen },
    { d: 'M6 10.7h12v2.8H6zM6 15.7h8.8v2.8H6z', fill: ink },
    { d: 'M17 15.7h1.8v2.8H17z', fill: grid },
  ],
  middle: [
    { d: 'M4.2 10.5h15.6v3H4.2z', fill: excelGreen },
    { d: 'M6 4.7h12v2.8H6zM6 16.5h8.8v2.8H6z', fill: ink },
    { d: 'M17 16.5h1.8v2.8H17z', fill: grid },
  ],
  bottomAlign: [
    { d: 'M4.2 16h15.6v3H4.2z', fill: excelGreen },
    { d: 'M6 5.5h12v2.8H6zM6 10.5h8.8v2.8H6z', fill: ink },
    { d: 'M17 10.5h1.8v2.8H17z', fill: grid },
  ],
  alignLeft: [
    { d: 'M4.1 4.3h3v15.4h-3z', fill: excelGreen },
    {
      d: 'M8.8 5.2h11v2.8h-11zM8.8 9.6h7.6v2.8H8.8zM8.8 14h10.2v2.8H8.8z',
      fill: ink,
    },
  ],
  alignCenter: [
    { d: 'M10.5 4.3h3v15.4h-3z', fill: excelGreen },
    {
      d: 'M4.7 5.2h14.6v2.8H4.7zM8 9.6h8v2.8H8zM5.8 14h12.4v2.8H5.8z',
      fill: ink,
    },
  ],
  alignRight: [
    { d: 'M16.9 4.3h3v15.4h-3z', fill: excelGreen },
    {
      d: 'M4.2 5.2h11v2.8h-11zM7.6 9.6h7.6v2.8H7.6zM5 14h10.2v2.8H5z',
      fill: ink,
    },
  ],
  textOrientation: [
    {
      d: 'M5.5 16 9.7 5.2h2.6L16.5 16h-2.4l-.9-2.4H8.8L7.9 16H5.5Z',
      fill: ink,
    },
    { d: 'M9.5 11.4h3.4L11.2 6.8 9.5 11.4Z', fill: '#ffffff' },
    {
      d: 'M4.5 18.5h15',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.4',
      strokeLinecap: 'round',
    },
    {
      d: 'M15.4 12.1 19.3 8.2M16.2 8.2h3.1v3.1',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.6',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  wrap: [
    { d: 'M4.5 4.7h15v2.8H4.5z', fill: ink },
    {
      d: 'M4.8 10.3h11a2.8 2.8 0 0 1 0 5.6h-4.4M11.9 13.7 9.7 15.9l2.2 2.2',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '2.2',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  indentDecrease: [
    {
      d: 'M10.4 5.2h9.2V8h-9.2zM10.4 9.6h6.8v2.8h-6.8zM10.4 14h9.2v2.8h-9.2z',
      fill: ink,
    },
    {
      d: 'M8.7 8.2 4.7 12l4 3.8M5.1 12h5.6',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '2.65',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  indentIncrease: [
    {
      d: 'M10.4 5.2h9.2V8h-9.2zM10.4 9.6h6.8v2.8h-6.8zM10.4 14h9.2v2.8h-9.2z',
      fill: ink,
    },
    {
      d: 'M4.7 8.2 8.7 12l-4 3.8M4.7 12h5.8',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '2.65',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  merge: [
    { d: 'M4.8 5.4h14.4v13.2H4.8z', fill: '#ffffff', stroke: ink, strokeWidth: '1.35' },
    {
      d: 'M5.2 10.1h13.6M5.2 13.9h13.6M10.2 5.8v12.4M13.8 5.8v12.4',
      fill: 'none',
      stroke: gridLight,
      strokeWidth: '1.2',
    },
    {
      d: 'M7.5 12h9M10.4 9.1 7.5 12l2.9 2.9M13.6 9.1 16.5 12l-2.9 2.9',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '2',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  table: [
    { d: 'M4.5 5.5h15v13h-15z', fill: '#ffffff', stroke: ink },
    { d: 'M4.5 5.5h15v3.5h-15z', fill: excelGreen },
    { d: 'M4.5 12h15M4.5 15.3h15M9.5 5.5v13M14.5 5.5v13', fill: 'none', stroke: gridLight },
    {
      d: 'M6.2 7.2h2M11.2 7.2h2M16.2 7.2h1.4',
      fill: 'none',
      stroke: '#ffffff',
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
  ],
  pivotTable: [
    { d: 'M4.5 5.5h11.5v12h-11.5z', fill: '#ffffff', stroke: ink },
    { d: 'M4.5 5.5h11.5v3.3h-11.5z', fill: excelBlue },
    { d: 'M4.5 12h11.5M8.3 5.5v12M12.1 5.5v12', fill: 'none', stroke: gridLight },
    { d: 'M14.5 12.5h5.5v6h-5.5z', fill: '#eef6ee', stroke: excelGreen, strokeWidth: '1.2' },
    {
      d: 'M15.8 14.5h2.9M15.8 16.5h2.9',
      fill: 'none',
      stroke: '#0b5a2f',
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
  ],
  pivotRecommended: [
    { d: 'M4.5 6h11.5v11h-11.5z', fill: '#ffffff', stroke: ink },
    { d: 'M4.5 6h11.5v3.1h-11.5z', fill: excelBlue },
    { d: 'M4.5 12.2h11.5M8.3 6v11M12.1 6v11', fill: 'none', stroke: gridLight },
    {
      d: 'M17.8 5.2 18.4 6.8 20.1 6.9 18.8 7.9 19.2 9.5 17.8 8.6 16.4 9.5 16.9 7.9 15.5 6.9 17.2 6.8Z',
      fill: excelYellow,
      stroke: '#b79500',
      strokeLinejoin: 'round',
    },
    { d: 'M15.8 13.5h4.4v4h-4.4z', fill: excelGreen, stroke: '#0b5a2f' },
  ],
  pivotExistingSheet: [
    { d: 'M7 7.5h11.5v11h-11.5z', fill: '#ffffff', stroke: ink },
    { d: 'M7 7.5h11.5v3h-11.5z', fill: excelBlue },
    { d: 'M7 13.5h11.5M10.8 7.5v11M14.6 7.5v11', fill: 'none', stroke: gridLight },
    { d: 'M4.5 4.5h11.5v11h-11.5z', fill: '#ffffff', stroke: ink },
    { d: 'M4.5 4.5h11.5v3h-11.5z', fill: excelGreen },
    { d: 'M4.5 10.5h11.5M8.3 4.5v11M12.1 4.5v11', fill: 'none', stroke: gridLight },
    { d: 'M15.5 14.5h4.5v4.5h-4.5z', fill: excelBlue, stroke: '#1f4e79' },
  ],
  picture: [
    { d: 'M4.5 5.5h15v13h-15z', fill: '#ffffff', stroke: ink },
    { d: 'M6.4 16.2 10 12.3l2.5 2.6 2.2-2.1 3 3.4H6.4Z', fill: excelGreen, stroke: '#0b5a2f' },
    {
      d: 'M14.9 8.2m-1.4 0a1.4 1.4 0 1 0 2.8 0a1.4 1.4 0 1 0 -2.8 0',
      fill: excelYellow,
      stroke: '#b79500',
    },
  ],
  shapes: [
    { d: 'M5.2 6.2h6.4v6.4H5.2z', fill: '#ffffff', stroke: excelBlue, strokeWidth: '1.2' },
    {
      d: 'M14.4 5.8 19.4 14H9.4l5-8.2Z',
      fill: '#ffffff',
      stroke: excelGreen,
      strokeWidth: '1.2',
      strokeLinejoin: 'round',
    },
    {
      d: 'M8.8 13.2a3.1 3.1 0 1 0 0 6.2 3.1 3.1 0 1 0 0-6.2',
      fill: excelYellow,
      stroke: '#b79500',
      strokeWidth: '1.2',
    },
    { d: 'M14.4 15.7h5v3.1h-5z', fill: excelOrange, stroke: '#8f4a12', strokeWidth: '1.2' },
  ],
  screenshot: [
    { d: 'M4.8 6.3h14.4v10.2H4.8z', fill: '#ffffff', stroke: ink, strokeWidth: '1.1' },
    { d: 'M6.6 8h10.8v6.8H6.6z', fill: '#eaf2fb', stroke: gridLight },
    { d: 'M6.6 8h10.8v1.8H6.6z', fill: '#d7e9fb', stroke: gridLight },
    {
      d: 'M9.3 18.5h5.4M12 16.5v2',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    {
      d: 'M14.5 14 18.6 9.9M15.2 9.9h3.4v3.4',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.65',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  chart: [
    { d: 'M5 18.5h14', fill: 'none', stroke: ink, strokeWidth: '1.2', strokeLinecap: 'round' },
    { d: 'M6.2 11.5h2.8v7h-2.8z', fill: excelBlue, stroke: '#1f4e79' },
    { d: 'M10.6 8.5h2.8v10h-2.8z', fill: excelGreen, stroke: '#0b5a2f' },
    { d: 'M15 5.8h2.8v12.7H15z', fill: excelOrange, stroke: '#8f4a12' },
  ],
  chartColumn: [
    { d: 'M5 18.5h14', fill: 'none', stroke: ink, strokeWidth: '1.2', strokeLinecap: 'round' },
    { d: 'M6.3 12.2h2.4v6.3H6.3z', fill: excelBlue, stroke: '#1f4e79' },
    { d: 'M10.7 8.5h2.4v10h-2.4z', fill: excelGreen, stroke: '#0b5a2f' },
    { d: 'M15.1 5.8h2.4v12.7h-2.4z', fill: excelOrange, stroke: '#8f4a12' },
  ],
  chartBar: [
    { d: 'M5.5 5.2v13.4', fill: 'none', stroke: ink, strokeWidth: '1.2', strokeLinecap: 'round' },
    { d: 'M5.5 7h9.6v2.3H5.5z', fill: excelBlue, stroke: '#1f4e79' },
    { d: 'M5.5 11h12.8v2.3H5.5z', fill: excelGreen, stroke: '#0b5a2f' },
    { d: 'M5.5 15h7.4v2.3H5.5z', fill: excelOrange, stroke: '#8f4a12' },
  ],
  chartLine: [
    {
      d: 'M5 18.5h14M5 5.5v13',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    {
      d: 'M6.6 15.3 10 11.4l3.3 2.2 4.3-6.2',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.7',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
    {
      d: 'M6.6 15.3m-1.1 0a1.1 1.1 0 1 0 2.2 0a1.1 1.1 0 1 0 -2.2 0',
      fill: '#ffffff',
      stroke: excelBlue,
    },
    {
      d: 'M10 11.4m-1.1 0a1.1 1.1 0 1 0 2.2 0a1.1 1.1 0 1 0 -2.2 0',
      fill: '#ffffff',
      stroke: excelBlue,
    },
    {
      d: 'M17.6 7.4m-1.1 0a1.1 1.1 0 1 0 2.2 0a1.1 1.1 0 1 0 -2.2 0',
      fill: '#ffffff',
      stroke: excelBlue,
    },
  ],
  chartArea: [
    {
      d: 'M5 18.5h14M5 5.5v13',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    {
      d: 'M6.3 17.8V14l3.7-4.1 3.2 2.4 4.4-5.6v11.1Z',
      fill: '#cfe8d6',
      stroke: excelGreen,
      strokeWidth: '1.1',
      strokeLinejoin: 'round',
    },
    {
      d: 'M6.3 14 10 9.9l3.2 2.4 4.4-5.6',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.4',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  chartPie: [
    { d: 'M12 4.8v7.2h7.2a7.2 7.2 0 0 0-7.2-7.2Z', fill: excelBlue, stroke: '#1f4e79' },
    { d: 'M12 12V4.8a7.2 7.2 0 1 0 6.4 10.5Z', fill: excelGreen, stroke: '#0b5a2f' },
    { d: 'M12 12h7.2a7.2 7.2 0 0 1-.8 3.3Z', fill: excelOrange, stroke: '#8f4a12' },
  ],
  chartScatter: [
    {
      d: 'M5 18.5h14M5 5.5v13',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    {
      d: 'M8 14.6m-1.4 0a1.4 1.4 0 1 0 2.8 0a1.4 1.4 0 1 0 -2.8 0',
      fill: excelBlue,
      stroke: '#1f4e79',
    },
    {
      d: 'M11.5 10.2m-1.4 0a1.4 1.4 0 1 0 2.8 0a1.4 1.4 0 1 0 -2.8 0',
      fill: excelGreen,
      stroke: '#0b5a2f',
    },
    {
      d: 'M16.3 7.5m-1.4 0a1.4 1.4 0 1 0 2.8 0a1.4 1.4 0 1 0 -2.8 0',
      fill: excelOrange,
      stroke: '#8f4a12',
    },
    {
      d: 'M15.1 15.4m-1.2 0a1.2 1.2 0 1 0 2.4 0a1.2 1.2 0 1 0 -2.4 0',
      fill: excelYellow,
      stroke: '#b79500',
    },
  ],
  chartRecommended: [
    { d: 'M4 18.5h12.5', fill: 'none', stroke: ink, strokeWidth: '1.2', strokeLinecap: 'round' },
    { d: 'M5 12.5h2.3v6H5z', fill: excelBlue, stroke: '#1f4e79' },
    { d: 'M9.1 8.7h2.3v9.8H9.1z', fill: excelGreen, stroke: '#0b5a2f' },
    { d: 'M13.2 6.3h2.3v12.2h-2.3z', fill: excelOrange, stroke: '#8f4a12' },
    {
      d: 'M18 4.7 18.6 6.3 20.3 6.4 19 7.4 19.4 9 18 8.1 16.6 9 17.1 7.4 15.7 6.4 17.4 6.3Z',
      fill: excelYellow,
      stroke: '#b79500',
      strokeLinejoin: 'round',
    },
  ],
  devicePicture: [
    { d: 'M4.5 6h15v10.5h-15z', fill: '#ffffff', stroke: ink },
    { d: 'M6.3 14.6 9 11.9l2.1 2.1 2-1.9 3.5 2.5Z', fill: excelGreen, stroke: '#0b5a2f' },
    {
      d: 'M15.5 8.5m-1.2 0a1.2 1.2 0 1 0 2.4 0a1.2 1.2 0 1 0 -2.4 0',
      fill: excelYellow,
      stroke: '#b79500',
    },
    { d: 'M9.2 18.5h5.6M12 16.5v2', fill: 'none', stroke: ink, strokeLinecap: 'round' },
  ],
  onlinePicture: [
    { d: 'M4.5 6h15v11.5h-15z', fill: '#ffffff', stroke: ink },
    { d: 'M6.2 15.6 9.4 12l2 2.2 2.4-2.1 3.5 3.5Z', fill: excelGreen, stroke: '#0b5a2f' },
    {
      d: 'M15.8 8.6m-1.1 0a1.1 1.1 0 1 0 2.2 0a1.1 1.1 0 1 0 -2.2 0',
      fill: excelYellow,
      stroke: '#b79500',
    },
    {
      d: 'M17.4 13.3a2.7 2.7 0 1 0 0 5.4 2.7 2.7 0 1 0 0-5.4',
      fill: '#ffffff',
      stroke: excelBlue,
      strokeWidth: '1.2',
    },
    {
      d: 'M17.4 13.3a1.35 2.7 0 0 0 0 5.4M17.4 13.3a1.35 2.7 0 0 1 0 5.4M14.7 16h5.4',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1',
      strokeLinecap: 'round',
    },
  ],
  stockPicture: [
    { d: 'M4.5 6h15v11.5h-15z', fill: '#ffffff', stroke: ink },
    { d: 'M6.2 15.6 9.4 12l2 2.2 2.4-2.1 3.5 3.5Z', fill: excelGreen, stroke: '#0b5a2f' },
    {
      d: 'M15.8 8.6m-1.1 0a1.1 1.1 0 1 0 2.2 0a1.1 1.1 0 1 0 -2.2 0',
      fill: excelYellow,
      stroke: '#b79500',
    },
    { d: 'M6.2 7.8h4.8v3.1H6.2z', fill: '#f3f8ff', stroke: excelBlue },
    {
      d: 'M7.3 8.7h2.6M7.3 10h1.7',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '.8',
      strokeLinecap: 'round',
    },
  ],
  screenshotWindow: [
    { d: 'M4.5 6h15v10.5h-15z', fill: '#ffffff', stroke: ink },
    { d: 'M6.2 8h11.6v6.5H6.2z', fill: '#f3f8ff', stroke: gridLight },
    {
      d: 'M8 9.5h5.5M8 11.5h8',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1',
      strokeLinecap: 'round',
    },
    { d: 'M9.3 18.5h5.4M12 16.5v2', fill: 'none', stroke: ink, strokeLinecap: 'round' },
  ],
  screenClipping: [
    { d: 'M5.5 6.5h13v10h-13z', fill: '#ffffff', stroke: ink, strokeDasharray: '2 1' },
    { d: 'M7.5 8.5h9v6h-9z', fill: '#f3f8ff', stroke: gridLight },
    {
      d: 'M16.2 14.7 19 17.5M17.1 14.1l-1.5 1.5 2.8 2.8 1.5-1.5Z',
      fill: excelGreen,
      stroke: '#0b5a2f',
      strokeLinejoin: 'round',
    },
  ],
  shapeLine: [
    {
      d: 'M5.5 18.5 18.5 5.5',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.8',
      strokeLinecap: 'round',
    },
  ],
  shapeArrow: [
    {
      d: 'M5.5 18.5 18.2 5.8M13.8 5.6h4.7v4.7',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.8',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  shapeRectangle: [
    { d: 'M5.5 7h13v10h-13z', fill: '#ffffff', stroke: excelBlue, strokeWidth: '1.5' },
  ],
  shapeRoundedRectangle: [
    {
      d: 'M6 7h12a1.5 1.5 0 0 1 1.5 1.5v7A1.5 1.5 0 0 1 18 17H6a1.5 1.5 0 0 1-1.5-1.5v-7A1.5 1.5 0 0 1 6 7Z',
      fill: '#ffffff',
      stroke: excelGreen,
      strokeWidth: '1.5',
    },
  ],
  shapeOval: [
    {
      d: 'M12 6.5a7 5.5 0 1 0 0 11 7 5.5 0 1 0 0-11',
      fill: '#ffffff',
      stroke: excelOrange,
      strokeWidth: '1.5',
    },
  ],
  shapeTriangle: [
    {
      d: 'M12 5.8 19 18H5Z',
      fill: '#ffffff',
      stroke: excelGreen,
      strokeWidth: '1.5',
      strokeLinejoin: 'round',
    },
  ],
  shapeDiamond: [
    {
      d: 'M12 5.5 19 12l-7 6.5L5 12Z',
      fill: '#ffffff',
      stroke: excelBlue,
      strokeWidth: '1.5',
      strokeLinejoin: 'round',
    },
  ],
  themeLight: [
    { d: 'M5 5.5h14v13H5z', fill: '#ffffff', stroke: ink },
    { d: 'M5 5.5h14v3.5H5z', fill: '#f3f2f1' },
    {
      d: 'M7 11h4.7M7 13.5h7.8M7 16h5.8',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    { d: 'M15.2 10.8h2.5v2.5h-2.5z', fill: excelGreen, stroke: '#0b5a2f' },
    { d: 'M15.2 14.7h2.5v1.8h-2.5z', fill: excelBlue, stroke: '#1f4e79' },
  ],
  themeDark: [
    { d: 'M5 5.5h14v13H5z', fill: '#2d2d2d', stroke: ink },
    { d: 'M5 5.5h14v3.5H5z', fill: '#1f1f1f' },
    {
      d: 'M7 11h4.7M7 13.5h7.8M7 16h5.8',
      fill: 'none',
      stroke: '#f3f2f1',
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    { d: 'M15.2 10.8h2.5v2.5h-2.5z', fill: excelBlue, stroke: '#8db3e2' },
    { d: 'M15.2 14.7h2.5v1.8h-2.5z', fill: excelOrange, stroke: '#f4b183' },
  ],
  themeContrast: [
    { d: 'M5 5.5h14v13H5z', fill: '#ffffff', stroke: ink },
    { d: 'M5 5.5h14v3.5H5z', fill: ink },
    {
      d: 'M7 11h4.7M7 13.5h7.8M7 16h5.8',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    { d: 'M15.2 10.8h2.5v2.5h-2.5z', fill: excelYellow, stroke: '#b79500' },
    { d: 'M15.2 14.7h2.5v1.8h-2.5z', fill: excelRed, stroke: '#7f0000' },
  ],
  link: [
    {
      d: 'M9.6 14.4 8.5 15.5a2.4 2.4 0 0 1-3.4-3.4l2.2-2.2a2.4 2.4 0 0 1 3.4 0',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.5',
      strokeLinecap: 'round',
    },
    {
      d: 'M14.4 9.6 15.5 8.5a2.4 2.4 0 0 1 3.4 3.4l-2.2 2.2a2.4 2.4 0 0 1-3.4 0',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.5',
      strokeLinecap: 'round',
    },
    {
      d: 'M10.2 13.8 13.8 10.2',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.3',
      strokeLinecap: 'round',
    },
  ],
  commentAdd: [
    { d: 'M5 6.5h12.5v8.5h-6.7L7.2 18v-3H5z', fill: '#ffffff', stroke: ink },
    {
      d: 'M7.2 9h6.8M7.2 11.5h4.6',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    { d: 'M15.8 14.5h4.5v4.5h-4.5z', fill: excelGreen, stroke: '#0b5a2f' },
    {
      d: 'M18.05 15.8v1.9M17.1 16.75h1.9',
      fill: 'none',
      stroke: '#ffffff',
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
  ],
  function: [
    {
      d: 'M12.9 6.4c-1.9-.6-3 .3-3.4 2.3l-2 9.6M7.9 10.4h4.4',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.9',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
    {
      d: 'M13.1 13.4 16.3 17.2M16.3 13.4 13.1 17.2',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.6',
      strokeLinecap: 'round',
    },
  ],
  pageTheme: [
    { d: 'M5 5.5h14v13H5z', fill: '#ffffff', stroke: ink },
    { d: 'M5 5.5h14v4H5z', fill: excelGreen },
    {
      d: 'M7 12h5.5M7 15h4',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    {
      d: 'M16.4 10.9 19.5 14l-3.1 3.1L13.3 14l3.1-3.1Z',
      fill: excelYellow,
      stroke: '#b79500',
      strokeLinejoin: 'round',
    },
  ],
  pageSetup: [
    { d: 'M5.5 4.8H15l2.5 2.5v12H5.5z', fill: '#ffffff', stroke: ink },
    { d: 'M15 4.8v2.6h2.5', fill: '#eef6ee', stroke: grid },
    {
      d: 'M7.4 8.8h7.2M7.4 11.6h7.2M7.4 14.4h4.8',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    {
      d: 'M16.4 15.5m-2 0a2 2 0 1 0 4 0a2 2 0 1 0-4 0',
      fill: '#eef6ee',
      stroke: excelGreen,
      strokeWidth: '1.3',
    },
    {
      d: 'M16.4 12.9v1.1M16.4 17v1.1M13.8 15.5h1.1M17.9 15.5h1.1M14.6 13.7l.8.8M18.2 17.3l-.8-.8M14.6 17.3l.8-.8M18.2 13.7l-.8.8',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
  ],
  printArea: [
    { d: 'M5 6h14v11.5H5z', fill: '#ffffff', stroke: ink },
    { d: 'M5 9.8h14M9.7 6v11.5M14.4 6v11.5', fill: 'none', stroke: gridLight },
    { d: 'M7 11.5h10v4.2H7z', fill: '#eef6ee', stroke: excelGreen, strokeWidth: '1.25' },
    {
      d: 'M6.8 4.2h3.4M13.8 4.2h3.4M6.8 19.8h3.4M13.8 19.8h3.4',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.2',
      strokeLinecap: 'round',
    },
  ],
  pageBreaks: [
    { d: 'M6 4.8h9.3l2.7 2.7v12H6z', fill: '#ffffff', stroke: ink },
    { d: 'M15.3 4.8v2.7H18', fill: '#eef6ee', stroke: grid },
    {
      d: 'M6 12.1h12',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.6',
      strokeLinecap: 'round',
      strokeDasharray: '2 1.4',
    },
    {
      d: 'M12 4.8v14.7',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.3',
      strokeLinecap: 'round',
      strokeDasharray: '2 1.4',
    },
  ],
  sheetBackground: [
    { d: 'M4.5 5.5h15v13h-15z', fill: '#ffffff', stroke: ink },
    {
      d: 'M4.5 5.5h15M4.5 9.8h15M4.5 14.1h15M9.5 5.5v13M14.5 5.5v13',
      fill: 'none',
      stroke: gridLight,
    },
    { d: 'M6.2 16.4 9.8 12.7l2.5 2.5 2.1-2 3.3 3.2H6.2Z', fill: excelGreen, stroke: '#0b5a2f' },
    {
      d: 'M15.3 8.2m-1.3 0a1.3 1.3 0 1 0 2.6 0a1.3 1.3 0 1 0 -2.6 0',
      fill: excelYellow,
      stroke: '#b79500',
    },
  ],
  printTitles: [
    { d: 'M4.8 6h14.4v12H4.8z', fill: '#ffffff', stroke: ink },
    { d: 'M4.8 6h14.4v3.6H4.8z', fill: excelBlue },
    { d: 'M8.8 6v12M13.2 6v12M4.8 12.4h14.4M4.8 15.2h14.4', fill: 'none', stroke: gridLight },
    {
      d: 'M6.4 7.8h3M11 7.8h2.8M15 7.8h2.3',
      fill: 'none',
      stroke: '#ffffff',
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
  ],
  filter: [
    {
      d: 'M5.5 6h13l-5 5.8v5.3l-3 1.5v-6.8L5.5 6Z',
      fill: excelGreen,
      stroke: '#0b5a2f',
      strokeWidth: '1.1',
      strokeLinejoin: 'round',
    },
  ],
  textToColumns: [
    { d: 'M4.8 6h14.4v12H4.8z', fill: '#ffffff', stroke: ink },
    { d: 'M8.6 6v12M13.4 6v12M4.8 10h14.4M4.8 14h14.4', fill: 'none', stroke: gridLight },
    {
      d: 'M6.2 8h1.2M10.2 8h1.7M15.2 8h2M6.2 12h1.2M10.2 12h1.7M15.2 12h2',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    {
      d: 'M12 4.5v15',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.5',
      strokeLinecap: 'round',
    },
  ],
  removeDuplicates: [
    { d: 'M5 6h9.5v9.5H5z', fill: '#ffffff', stroke: ink },
    { d: 'M8.5 9.5H18v9.5H8.5z', fill: '#ffffff', stroke: ink },
    {
      d: 'M6.8 8h5M10.3 11.5h5M10.3 14.2h5',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.05',
      strokeLinecap: 'round',
    },
    {
      d: 'M14.9 4.9 19 9M19 4.9 14.9 9',
      fill: 'none',
      stroke: excelRed,
      strokeWidth: '1.35',
      strokeLinecap: 'round',
    },
  ],
  dataValidation: [
    { d: 'M4.5 6h11v12h-11z', fill: '#ffffff', stroke: ink },
    {
      d: 'M6.7 9h6.6M6.7 12h4.6',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    {
      d: 'M14.2 13.6 16 15.4l3.2-3.6',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.8',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  dataValidationCircle: [
    { d: 'M6.5 6h11v12h-11z', fill: '#ffffff', stroke: ink },
    {
      d: 'M8.7 9h6.6M8.7 12h4.6M8.7 15h5.4',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    {
      d: 'M12.5 12.4a4.6 2.6 0 1 0 0 5.2 4.6 2.6 0 1 0 0-5.2',
      fill: 'none',
      stroke: excelRed,
      strokeWidth: '1.7',
    },
  ],
  dataValidationClearCircles: [
    { d: 'M4.5 6h11v12h-11z', fill: '#ffffff', stroke: ink },
    {
      d: 'M6.7 9h6.6M6.7 12h4.6M6.7 15h5.4',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    {
      d: 'M10.5 12.4a4.6 2.6 0 1 0 0 5.2 4.6 2.6 0 1 0 0-5.2',
      fill: 'none',
      stroke: excelRed,
      strokeWidth: '1.7',
    },
    {
      d: 'M15.9 5.9 19 9M19 5.9 15.9 9',
      fill: 'none',
      stroke: excelRed,
      strokeWidth: '1.35',
      strokeLinecap: 'round',
    },
  ],
  dataValidationClearRules: [
    { d: 'M4.5 6h11v12h-11z', fill: '#ffffff', stroke: ink },
    {
      d: 'M6.7 9h6.6M6.7 12h4.6M6.7 15h5.4',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    {
      d: 'M13.6 14.2 15.4 16l3.2-3.6',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.8',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
    {
      d: 'M15.9 5.9 19 9M19 5.9 15.9 9',
      fill: 'none',
      stroke: excelRed,
      strokeWidth: '1.35',
      strokeLinecap: 'round',
    },
  ],
  outlineGroup: [
    { d: 'M5.5 6.5h13v11h-13z', fill: '#ffffff', stroke: ink },
    { d: 'M5.5 10.2h13M5.5 13.8h13M9.5 6.5v11M14.5 6.5v11', fill: 'none', stroke: gridLight },
    { d: 'M4.3 3.6h5.4v2.8H4.3zM14.3 17.6h5.4v2.8h-5.4z', fill: excelGreen, stroke: '#0b5a2f' },
    {
      d: 'M7 4.2v1.8M6.1 5.1h1.8M17 18.2v1.8M16.1 19.1h1.8',
      fill: 'none',
      stroke: '#ffffff',
      strokeWidth: '1.2',
      strokeLinecap: 'round',
    },
  ],
  outlineUngroup: [
    { d: 'M5.5 6.5h13v11h-13z', fill: '#ffffff', stroke: ink },
    { d: 'M5.5 10.2h13M5.5 13.8h13M9.5 6.5v11M14.5 6.5v11', fill: 'none', stroke: gridLight },
    { d: 'M4.3 3.6h5.4v2.8H4.3zM14.3 17.6h5.4v2.8h-5.4z', fill: excelRed, stroke: '#7f0000' },
    {
      d: 'M6.1 5.1h1.8M16.1 19.1h1.8',
      fill: 'none',
      stroke: '#ffffff',
      strokeWidth: '1.2',
      strokeLinecap: 'round',
    },
  ],
  outlineShow: [
    { d: 'M4.5 7h10v10h-10z', fill: '#ffffff', stroke: ink },
    { d: 'M4.5 10.3h10M4.5 13.7h10M7.8 7v10M11.2 7v10', fill: 'none', stroke: gridLight },
    {
      d: 'M18 10.2v3.6M16.2 12h3.6',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.5',
      strokeLinecap: 'round',
    },
  ],
  outlineHide: [
    { d: 'M4.5 7h10v10h-10z', fill: '#ffffff', stroke: ink },
    { d: 'M4.5 10.3h10M4.5 13.7h10M7.8 7v10M11.2 7v10', fill: 'none', stroke: gridLight },
    {
      d: 'M16.2 12h3.6',
      fill: 'none',
      stroke: excelRed,
      strokeWidth: '1.5',
      strokeLinecap: 'round',
    },
  ],
  names: [
    { d: 'M5 6h14v12H5z', fill: '#ffffff', stroke: ink },
    { d: 'M5 6h14v3.4H5z', fill: excelGreen },
    {
      d: 'M7.1 12h4.6M7.1 14.8h3.3',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    { d: 'M14 11.3h4.2v4.4H14z', fill: '#eef6ee', stroke: excelGreen },
    {
      d: 'M15 13.5h2.2',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
  ],
  namesCreateTop: [
    { d: 'M5 6h14v12H5z', fill: '#ffffff', stroke: ink },
    { d: 'M5 6h14v3.2H5z', fill: excelGreen },
    { d: 'M5 12h14M5 15h14M9.6 6v12M14.2 6v12', fill: 'none', stroke: gridLight },
    { d: 'M9 7.6h6', fill: 'none', stroke: '#ffffff', strokeWidth: '1.3', strokeLinecap: 'round' },
  ],
  namesCreateBottom: [
    { d: 'M5 6h14v12H5z', fill: '#ffffff', stroke: ink },
    { d: 'M5 14.8h14V18H5z', fill: excelGreen },
    { d: 'M5 9.2h14M5 12h14M9.6 6v12M14.2 6v12', fill: 'none', stroke: gridLight },
    { d: 'M9 16.4h6', fill: 'none', stroke: '#ffffff', strokeWidth: '1.3', strokeLinecap: 'round' },
  ],
  namesCreateLeft: [
    { d: 'M5 6h14v12H5z', fill: '#ffffff', stroke: ink },
    { d: 'M5 6h4v12H5z', fill: excelGreen },
    { d: 'M9.6 6v12M14.2 6v12M5 9.2h14M5 12.4h14M5 15.6h14', fill: 'none', stroke: gridLight },
    {
      d: 'M7 9.4v5.2',
      fill: 'none',
      stroke: '#ffffff',
      strokeWidth: '1.3',
      strokeLinecap: 'round',
    },
  ],
  namesCreateRight: [
    { d: 'M5 6h14v12H5z', fill: '#ffffff', stroke: ink },
    { d: 'M15 6h4v12h-4z', fill: excelGreen },
    { d: 'M9.6 6v12M14.2 6v12M5 9.2h14M5 12.4h14M5 15.6h14', fill: 'none', stroke: gridLight },
    {
      d: 'M17 9.4v5.2',
      fill: 'none',
      stroke: '#ffffff',
      strokeWidth: '1.3',
      strokeLinecap: 'round',
    },
  ],
  trace: [
    { d: 'M5 8h5v5H5zM14.5 5h5v5h-5zM14.5 14h5v5h-5z', fill: '#ffffff', stroke: ink },
    {
      d: 'M14.2 7.5h-1.6v2h-2.3M11.5 8.3 10.3 9.5l1.2 1.2',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.3',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
    {
      d: 'M14.2 16.5h-2.8v-5h-1.1M11.5 10.3 10.3 11.5l1.2 1.2',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.3',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  dependents: [
    { d: 'M5 8h5v5H5zM14.5 5h5v5h-5zM14.5 14h5v5h-5z', fill: '#ffffff', stroke: ink },
    {
      d: 'M10.3 9.5h2.3v-2h1.6M13.2 6.3 14.4 7.5l-1.2 1.2',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.3',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
    {
      d: 'M10.3 11.5h1.1v5h2.8M13.2 15.3 14.4 16.5l-1.2 1.2',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.3',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  clearArrows: [
    { d: 'M5 7.5h5v5H5zM14 7.5h5v5h-5z', fill: '#ffffff', stroke: ink },
    {
      d: 'M10.4 10h3.3M13.7 10l-1.6-1.6M13.7 10 12.1 11.6',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.35',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
    {
      d: 'M7 17.5 17.5 7',
      fill: 'none',
      stroke: excelRed,
      strokeWidth: '1.8',
      strokeLinecap: 'round',
    },
  ],
  errorChecking: [
    { d: 'M4.5 6h11v12h-11z', fill: '#ffffff', stroke: ink },
    {
      d: 'M6.7 9h6.6M6.7 12h4.6M6.7 15h5.4',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    {
      d: 'M17.5 5.6 20.1 10.2h-5.2l2.6-4.6Z',
      fill: excelYellow,
      stroke: '#b79500',
      strokeLinejoin: 'round',
    },
    { d: 'M17.5 7.4v1.4', fill: 'none', stroke: ink, strokeWidth: '1.1', strokeLinecap: 'round' },
  ],
  calcOptions: [
    { d: 'M5 6h14v12H5z', fill: '#ffffff', stroke: ink },
    {
      d: 'M7 9h4M7 12h4M7 15h4',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    {
      d: 'M13.4 9h4.2M13.4 12h4.2M13.4 15h4.2',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
  ],
  watch: [
    { d: 'M4.8 12s2.8-5 7.2-5 7.2 5 7.2 5-2.8 5-7.2 5-7.2-5-7.2-5Z', fill: '#ffffff', stroke: ink },
    { d: 'M12 9.2a2.8 2.8 0 1 0 0 5.6 2.8 2.8 0 1 0 0-5.6', fill: '#eef6ee', stroke: excelGreen },
    { d: 'M12 12m-1.2 0a1.2 1.2 0 1 0 2.4 0a1.2 1.2 0 1 0-2.4 0', fill: excelGreen },
  ],
  spelling: [
    { d: 'M5 16 9.3 5.5h2.4L16 16h-2l-.85-2.3H7.85L7 16H5Z', fill: ink },
    { d: 'M8.4 12.1h4.4L10.6 6.5 8.4 12.1Z', fill: '#ffffff' },
    {
      d: 'M13.8 17 15.8 19l3.7-4.6',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.7',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  accessibility: [
    {
      d: 'M12 5.4m-1.7 0a1.7 1.7 0 1 0 3.4 0a1.7 1.7 0 1 0 -3.4 0',
      fill: excelBlue,
      stroke: '#1f4e79',
    },
    {
      d: 'M5.5 9h13M12 9v4.2M8.2 19l2.2-5.8h3.2l2.2 5.8',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.5',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  translate: [
    { d: 'M5 6h8v8H5z', fill: '#ffffff', stroke: ink },
    { d: 'M11 10h8v8h-8z', fill: '#ffffff', stroke: ink },
    {
      d: 'M7 9h4M9 7v5M7.5 12c1.7-1.2 2.7-2.7 3-4',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    {
      d: 'M13.2 15.8 15 11.8l1.8 4M14 14.5h2',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  protect: [
    { d: 'M7 10h10v8.5H7z', fill: '#ffffff', stroke: ink },
    {
      d: 'M9 10V8a3 3 0 0 1 6 0v2',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.4',
      strokeLinecap: 'round',
    },
    {
      d: 'M12 13.1m-1.2 0a1.2 1.2 0 1 0 2.4 0a1.2 1.2 0 1 0 -2.4 0',
      fill: excelYellow,
      stroke: '#b79500',
    },
    {
      d: 'M12 14.3v2',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.2',
      strokeLinecap: 'round',
    },
  ],
  print: [
    { d: 'M7 4.8h10v5.2H7z', fill: '#ffffff', stroke: ink },
    { d: 'M5 9.2h14v6.2H5z', fill: '#eef6ee', stroke: ink },
    { d: 'M7.5 13.5h9v5.2h-9z', fill: '#ffffff', stroke: ink },
    {
      d: 'M9 15.4h6M9 17h4.2',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    {
      d: 'M16.5 11.3h1.2',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.3',
      strokeLinecap: 'round',
    },
  ],
  freeze: [
    { d: 'M5 5.5h14v13H5z', fill: '#ffffff', stroke: ink },
    { d: 'M5 5.5h14v4H5zM5 5.5h4v13H5z', fill: '#dceef8', stroke: excelBlue },
    { d: 'M9 9.5h10M9 13.5h10M13 5.5v13', fill: 'none', stroke: gridLight },
    {
      d: 'M6.1 7.5h1.8M6.1 11.5h1.8',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
  ],
  zoom: [
    {
      d: 'M5.5 10.5a5 5 0 1 0 10 0 5 5 0 1 0-10 0',
      fill: '#ffffff',
      stroke: ink,
      strokeWidth: '1.4',
    },
    {
      d: 'M14.2 14.2 19 19',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.8',
      strokeLinecap: 'round',
    },
    {
      d: 'M8.2 10.5h4.6M10.5 8.2v4.6',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.2',
      strokeLinecap: 'round',
    },
  ],
  page: [
    { d: 'M6.5 4.8h9.5l2.5 2.5v12H6.5z', fill: '#ffffff', stroke: ink },
    { d: 'M16 4.8v2.6h2.5', fill: '#eef6ee', stroke: grid },
    {
      d: 'M8.5 9.4h7M8.5 12.2h7M8.5 15h4.8',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
  ],
  goTo: [
    { d: 'M5 6h10v12H5z', fill: '#ffffff', stroke: ink },
    {
      d: 'M7 9h5M7 12h4M7 15h5',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    {
      d: 'M13.5 12h6M17.2 9.8 19.5 12l-2.3 2.2',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.6',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
  ],
  goToSpecial: [
    { d: 'M5 6h10v12H5z', fill: '#ffffff', stroke: ink },
    {
      d: 'M7 9h5M7 12h4M7 15h5',
      fill: 'none',
      stroke: grid,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
    {
      d: 'M17.8 5.2 18.4 6.8 20.1 6.9 18.8 7.9 19.2 9.5 17.8 8.6 16.4 9.5 16.9 7.9 15.5 6.9 17.2 6.8Z',
      fill: '#d9c2f0',
      stroke: '#8764b8',
      strokeWidth: '1',
      strokeLinejoin: 'round',
    },
  ],
  objectSelect: [
    {
      d: 'M7 5v13l3.4-3.3 2.4 4.9 2.4-1.1-2.3-4.7h4.5L7 5Z',
      fill: '#ffffff',
      stroke: ink,
      strokeWidth: '1.2',
      strokeLinejoin: 'round',
    },
  ],
  selectionPane: [
    { d: 'M5 5.5h13v13H5z', fill: '#ffffff', stroke: ink, strokeWidth: '1.1' },
    { d: 'M14 5.5v13', fill: 'none', stroke: grid, strokeWidth: '1.1' },
    {
      d: 'M7 8.2h4.5M7 11h4.5M7 13.8h4.5M15.5 8.2h1.2M15.5 11h1.2M15.5 13.8h1.2',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.1',
      strokeLinecap: 'round',
    },
  ],
  options: [
    {
      d: 'M12 5.7a6.3 6.3 0 1 0 0 12.6 6.3 6.3 0 1 0 0-12.6',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '2.3',
      strokeDasharray: '1.9 3.05',
    },
    {
      d: 'M12 6.4a5.6 5.6 0 1 0 0 11.2 5.6 5.6 0 1 0 0-11.2',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.4',
    },
    {
      d: 'M12 9.6a2.4 2.4 0 1 0 0 4.8 2.4 2.4 0 1 0 0-4.8',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.4',
    },
  ],
  pen: [
    { d: 'M6 16.7 15.4 7.3l2.3 2.3-9.4 9.4-3 .8.7-3.1Z', fill: excelYellow, stroke: ink },
    { d: 'M15.1 5.5 18.5 8.9l1.4-1.4-3.4-3.4-1.4 1.4Z', fill: excelBlue, stroke: '#1f4e79' },
    { d: 'M5 20h10', fill: 'none', stroke: excelGreen, strokeWidth: '1.4', strokeLinecap: 'round' },
  ],
  eraser: [
    {
      d: 'M5.5 11.3 8.8 5.7l9.7 5.6-3.3 5.6z',
      fill: '#ffffff',
      stroke: ink,
      strokeWidth: '1.3',
      strokeLinejoin: 'round',
    },
    { d: 'M12.2 7.7 8.9 13.3', fill: 'none', stroke: excelOrange, strokeWidth: '1.15' },
    {
      d: 'M5 19.5h14',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.3',
      strokeLinecap: 'round',
    },
  ],
  script: [
    { d: 'M6 5.5h12v13H6z', fill: '#ffffff', stroke: ink },
    {
      d: 'M8.2 9.2 6.8 11l1.4 1.8M15.8 9.2l1.4 1.8-1.4 1.8',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.3',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
    {
      d: 'M13.1 8.8 10.9 13.2',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.3',
      strokeLinecap: 'round',
    },
    { d: 'M8.2 16h7.6', fill: 'none', stroke: grid, strokeWidth: '1.1', strokeLinecap: 'round' },
  ],
  addIn: [
    {
      d: 'M5.5 5.5h5.2v5.2H5.5zM13.3 5.5h5.2v5.2h-5.2zM5.5 13.3h5.2v5.2H5.5zM13.3 13.3h5.2v5.2h-5.2z',
      fill: '#ffffff',
      stroke: ink,
    },
    { d: 'M7 7h2.2v2.2H7z', fill: excelGreen },
    { d: 'M14.8 7h2.2v2.2h-2.2z', fill: excelBlue },
    { d: 'M7 14.8h2.2V17H7z', fill: excelOrange },
    {
      d: 'M16 14.4v3M14.5 15.9h3',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.3',
      strokeLinecap: 'round',
    },
  ],
  pdf: [
    { d: 'M6.5 4.8h9.5l2.5 2.5v12H6.5z', fill: '#ffffff', stroke: ink },
    { d: 'M16 4.8v2.6h2.5', fill: '#ffe5e5', stroke: grid },
    { d: 'M7.8 11h8.4v4.5H7.8z', fill: excelRed, stroke: '#7f0000' },
    {
      d: 'M8.7 14v-2h.9a.7.7 0 0 1 0 1.4h-.9M11 14v-2h.7a1 1 0 0 1 0 2H11M13.5 14v-2h1.6M13.5 13h1.2',
      fill: 'none',
      stroke: '#ffffff',
      strokeWidth: '.75',
      strokeLinecap: 'round',
    },
  ],
  save: [
    { d: 'M5.5 5.5h12l2 2v11h-14z', fill: '#ffffff', stroke: ink },
    { d: 'M8 5.5h7.5v4.5H8z', fill: excelBlue, stroke: '#1f4e79' },
    { d: 'M8 14h9v4.5H8z', fill: '#eef6ee', stroke: grid },
    { d: 'M10 15.7h5', fill: 'none', stroke: grid, strokeWidth: '1.1', strokeLinecap: 'round' },
  ],
  saveAs: [
    { d: 'M5.5 5.5h12l2 2v11h-14z', fill: '#ffffff', stroke: ink },
    { d: 'M8 5.5h7.5v4.5H8z', fill: excelBlue, stroke: '#1f4e79' },
    { d: 'M8 14h7v4.5H8z', fill: '#eef6ee', stroke: grid },
    { d: 'M14.4 17.2 18 13.6l1.4 1.4-3.6 3.6-1.7.3.3-1.7Z', fill: excelYellow, stroke: '#806000' },
  ],
  autosave: [
    { d: 'M7 9h10a3 3 0 0 1 0 6H7a3 3 0 0 1 0-6Z', fill: '#ffffff', stroke: ink },
    { d: 'M7 9h6v6H7a3 3 0 0 1 0-6Z', fill: excelGreen, stroke: '#0b5a2f' },
    {
      d: 'M16 12m-2.2 0a2.2 2.2 0 1 0 4.4 0a2.2 2.2 0 1 0 -4.4 0',
      fill: '#ffffff',
      stroke: grid,
    },
  ],
  share: [
    {
      d: 'M6.5 13.5v4h11v-4',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1.3',
      strokeLinecap: 'round',
    },
    {
      d: 'M12 16V5.5M8.8 8.7 12 5.5l3.2 3.2',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.7',
      strokeLinecap: 'round',
      strokeLinejoin: 'round',
    },
    { d: 'M8.5 18.5h7', fill: 'none', stroke: grid, strokeWidth: '1.1', strokeLinecap: 'round' },
  ],
  insertRows: [
    { d: 'M4 8.5h12v11H4z', fill: '#ffffff', stroke: ink },
    { d: 'M4 8.5h12v3.7H4z', fill: '#e2f0d9' },
    { d: 'M4 12.2h12M4 15.8h12M8 8.5v11M12 8.5v11', fill: 'none', stroke: grid },
    { d: 'M14.5 3.5h5.5V9h-5.5z', fill: excelGreen, stroke: '#0b5a2f' },
    {
      d: 'M17.25 4.9v2.7M15.9 6.25h2.7',
      fill: 'none',
      stroke: '#ffffff',
      strokeWidth: '1.3',
      strokeLinecap: 'round',
    },
  ],
  insertCells: [
    { d: 'M4 8.5h12v11H4z', fill: '#ffffff', stroke: ink },
    { d: 'M8 12.2h4v3.6H8z', fill: '#d9eaf7' },
    { d: 'M4 12.2h12M4 15.8h12M8 8.5v11M12 8.5v11', fill: 'none', stroke: grid },
    { d: 'M14.5 3.5h5.5V9h-5.5z', fill: excelGreen, stroke: '#0b5a2f' },
    {
      d: 'M17.25 4.9v2.7M15.9 6.25h2.7',
      fill: 'none',
      stroke: '#ffffff',
      strokeWidth: '1.3',
      strokeLinecap: 'round',
    },
  ],
  insertCols: [
    { d: 'M4 8.5h12v11H4z', fill: '#ffffff', stroke: ink },
    { d: 'M4 8.5h4v11H4z', fill: '#e2f0d9' },
    { d: 'M4 12.2h12M4 15.8h12M8 8.5v11M12 8.5v11', fill: 'none', stroke: grid },
    { d: 'M14.5 3.5h5.5V9h-5.5z', fill: excelGreen, stroke: '#0b5a2f' },
    {
      d: 'M17.25 4.9v2.7M15.9 6.25h2.7',
      fill: 'none',
      stroke: '#ffffff',
      strokeWidth: '1.3',
      strokeLinecap: 'round',
    },
  ],
  deleteRows: [
    { d: 'M4 8.5h12v11H4z', fill: '#ffffff', stroke: ink },
    { d: 'M4 8.5h12v3.7H4z', fill: '#f4cccc' },
    { d: 'M4 12.2h12M4 15.8h12M8 8.5v11M12 8.5v11', fill: 'none', stroke: grid },
    { d: 'M14.5 3.5h5.5V9h-5.5z', fill: excelRed, stroke: '#7f0000' },
    {
      d: 'M15.9 4.9 18.6 7.6M18.6 4.9 15.9 7.6',
      fill: 'none',
      stroke: '#ffffff',
      strokeWidth: '1.3',
      strokeLinecap: 'round',
    },
  ],
  deleteCells: [
    { d: 'M4 8.5h12v11H4z', fill: '#ffffff', stroke: ink },
    { d: 'M8 12.2h4v3.6H8z', fill: '#f4cccc' },
    { d: 'M4 12.2h12M4 15.8h12M8 8.5v11M12 8.5v11', fill: 'none', stroke: grid },
    { d: 'M14.5 3.5h5.5V9h-5.5z', fill: excelRed, stroke: '#7f0000' },
    {
      d: 'M15.9 4.9 18.6 7.6M18.6 4.9 15.9 7.6',
      fill: 'none',
      stroke: '#ffffff',
      strokeWidth: '1.3',
      strokeLinecap: 'round',
    },
  ],
  deleteCols: [
    { d: 'M4 8.5h12v11H4z', fill: '#ffffff', stroke: ink },
    { d: 'M4 8.5h4v11H4z', fill: '#f4cccc' },
    { d: 'M4 12.2h12M4 15.8h12M8 8.5v11M12 8.5v11', fill: 'none', stroke: grid },
    { d: 'M14.5 3.5h5.5V9h-5.5z', fill: excelRed, stroke: '#7f0000' },
    {
      d: 'M15.9 4.9 18.6 7.6M18.6 4.9 15.9 7.6',
      fill: 'none',
      stroke: '#ffffff',
      strokeWidth: '1.3',
      strokeLinecap: 'round',
    },
  ],
  formatCells: [
    { d: 'M5 5h11v11H5z', fill: '#ffffff', stroke: ink },
    { d: 'M5 5h11v3.4H5z', fill: excelGreen },
    { d: 'M5 11.1h11M8.7 5v11M12.3 5v11', fill: 'none', stroke: gridLight },
    { d: 'M14 13.5h5v5h-5z', fill: '#ffffff', stroke: ink, strokeWidth: '1.1' },
    {
      d: 'M15.1 15h2.8M15.1 17h2.8M15.8 14.4v1.2M17.2 16.4v1.2',
      fill: 'none',
      stroke: excelBlue,
      strokeWidth: '1.2',
      strokeLinecap: 'round',
    },
  ],
  conditional: [
    { d: 'M4 5.5h11v12H4z', fill: '#ffffff', stroke: ink },
    { d: 'M4 9.5h11M4 13.5h11M7.7 5.5v12M11.3 5.5v12', fill: 'none', stroke: gridLight },
    { d: 'M16.5 7h4v2h-4z', fill: excelRed },
    { d: 'M16.5 11h4v2h-4z', fill: excelYellow },
    { d: 'M16.5 15h4v2h-4z', fill: excelGreen },
  ],
  tableStyle: [
    { d: 'M4.5 5.5h13v12h-13z', fill: '#ffffff', stroke: ink },
    { d: 'M4.5 5.5h13v3.5h-13z', fill: excelBlue },
    { d: 'M4.5 12.5h13M8.8 5.5v12M13.1 5.5v12', fill: 'none', stroke: gridLight },
    {
      d: 'M14.6 15.9 18.8 11.7l1.2 1.2-4.2 4.2-1.6.4.4-1.6Z',
      fill: excelOrange,
      stroke: '#8f4a12',
      strokeLinejoin: 'round',
    },
  ],
  cellStyles: [
    { d: 'M4.6 6h7.2v4.2H4.6z', fill: '#e2f0d9', stroke: excelGreen, strokeWidth: '1.1' },
    { d: 'M12.8 6h6.6v4.2h-6.6z', fill: '#fff2cc', stroke: '#b79500', strokeWidth: '1.1' },
    { d: 'M4.6 11.8h7.2V16H4.6z', fill: '#ddebf7', stroke: excelBlue, strokeWidth: '1.1' },
    { d: 'M12.8 11.8h6.6V16h-6.6z', fill: '#fce4d6', stroke: excelOrange, strokeWidth: '1.1' },
    {
      d: 'M6.1 8.1h4.2M14 8.1h4M6.1 13.9h4.2M14 13.9h4',
      fill: 'none',
      stroke: ink,
      strokeWidth: '1',
      strokeLinecap: 'round',
    },
    {
      d: 'M8 18.3h8',
      fill: 'none',
      stroke: excelGreen,
      strokeWidth: '1.4',
      strokeLinecap: 'round',
    },
  ],
} as const satisfies Record<string, ExcelRibbonIconDefinition>;

export type ExcelRibbonIconName = keyof typeof EXCEL_RIBBON_ICON_PATHS;

export const excelRibbonIconPaths = (name: string): ExcelRibbonIconDefinition | null =>
  Object.hasOwn(EXCEL_RIBBON_ICON_PATHS, name)
    ? EXCEL_RIBBON_ICON_PATHS[name as ExcelRibbonIconName]
    : null;

export const createExcelRibbonSvg = (name: string, className: string): SVGSVGElement | null => {
  const segments = excelRibbonIconPaths(name);
  if (!segments) return null;

  const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.classList.add(className);
  svg.setAttribute('viewBox', '0 0 24 24');
  svg.setAttribute('focusable', 'false');
  svg.setAttribute('aria-hidden', 'true');

  for (const segment of segments) {
    const path = document.createElementNS('http://www.w3.org/2000/svg', 'path');
    path.setAttribute('d', segment.d);
    path.setAttribute('fill', segment.fill ?? 'none');
    if (segment.stroke) path.setAttribute('stroke', segment.stroke);
    if (segment.strokeWidth) path.setAttribute('stroke-width', segment.strokeWidth);
    if (segment.strokeLinecap) path.setAttribute('stroke-linecap', segment.strokeLinecap);
    if (segment.strokeLinejoin) path.setAttribute('stroke-linejoin', segment.strokeLinejoin);
    if (segment.strokeDasharray) path.setAttribute('stroke-dasharray', segment.strokeDasharray);
    svg.appendChild(path);
  }

  return svg;
};
