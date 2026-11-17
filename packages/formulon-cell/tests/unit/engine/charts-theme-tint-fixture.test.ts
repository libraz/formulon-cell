/**
 * Charts Excel saved with theme colours under lumMod/lumOff, opened in the cell. The expectations
 * are the RGB values Excel reports for each series fill, not the adapter's own output.
 */
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { afterEach, describe, expect, it } from 'vitest';
import { WorkbookHandle } from '../../../src/engine/workbook-handle.js';

type Rgb = readonly [number, number, number];

const open: WorkbookHandle[] = [];
afterEach(() => {
  for (const wb of open.splice(0)) wb.dispose();
});

/** Every series fill of the sheet's charts, keyed by series name, as getChart renders it. */
async function seriesFills(file: string): Promise<Map<string, Rgb>> {
  const bytes = new Uint8Array(readFileSync(join(__dirname, '../../fixtures/charts', file)));
  const wb = await WorkbookHandle.loadBytes(bytes);
  open.push(wb);
  const fills = new Map<string, Rgb>();
  for (const object of wb.listDrawingObjects(0)) {
    const info = wb.getChart(0, object.objectId);
    const data = wb.getChartData(0, object.objectId);
    if (!info || !data) continue;
    for (const series of data.series) {
      const fill = info.plots[series.plotIndex]?.series.find(
        (entry) => entry.seriesId === series.seriesId,
      )?.fill?.color;
      if (!fill || series.name === null) continue;
      const argb = fill.resolved?.argb;
      if (argb === undefined) continue;
      fills.set(series.name, [(argb >> 16) & 0xff, (argb >> 8) & 0xff, argb & 0xff]);
    }
  }
  return fills;
}

const within1 = (actual: Rgb | undefined, expected: Rgb): boolean =>
  actual?.every((channel, i) => Math.abs(channel - (expected[i] ?? 0)) <= 1) === true;

/** Series named "slot|lumMod|lumOff", with the fill RGB Excel reports, where the cell matches it. */
const EXACT: Record<string, Rgb> = {
  'tx1|20000|80000': [204, 204, 204],
  'tx1|40000|60000': [153, 153, 153],
  'tx1|60000|40000': [102, 102, 102],
  'tx1|65000|35000': [89, 89, 89],
  'tx1|75000|25000': [64, 64, 64],
  'tx1|85000|15000': [38, 38, 38],
  'tx1|95000|5000': [13, 13, 13],
  'tx1|95000|0': [0, 0, 0],
  'tx1|85000|0': [0, 0, 0],
  'tx1|75000|0': [0, 0, 0],
  'tx1|65000|0': [0, 0, 0],
  'tx1|50000|0': [0, 0, 0],
  'tx1|25000|0': [0, 0, 0],
  'bg1|20000|80000': [255, 255, 255],
  'bg1|40000|60000': [255, 255, 255],
  'bg1|60000|40000': [255, 255, 255],
  'bg1|50000|50000': [255, 255, 255],
  'bg1|65000|35000': [255, 255, 255],
  'bg1|75000|25000': [255, 255, 255],
  'bg1|85000|15000': [255, 255, 255],
  'bg1|95000|5000': [255, 255, 255],
  'bg1|95000|0': [242, 242, 242],
  'bg1|85000|0': [217, 217, 217],
  'bg1|75000|0': [191, 191, 191],
  'bg1|65000|0': [166, 166, 166],
  'bg1|25000|0': [64, 64, 64],
  'tx2|20000|80000': [184, 212, 240],
  'tx2|40000|60000': [113, 170, 224],
  'tx2|60000|40000': [45, 127, 207],
  'tx2|50000|50000': [78, 149, 217],
  'tx2|65000|35000': [41, 116, 189],
  'tx2|75000|25000': [33, 95, 154],
  'tx2|85000|15000': [25, 73, 118],
  'tx2|95000|5000': [18, 51, 83],
  'tx2|95000|0': [13, 38, 62],
  'tx2|85000|0': [12, 34, 55],
  'tx2|75000|0': [11, 30, 49],
  'tx2|65000|0': [9, 26, 42],
  'bg2|20000|80000': [250, 250, 250],
  'bg2|40000|60000': [246, 246, 246],
  'bg2|60000|40000': [241, 241, 241],
  'bg2|65000|35000': [240, 240, 240],
  'bg2|75000|25000': [238, 238, 238],
  'bg2|85000|15000': [235, 235, 235],
  'bg2|95000|5000': [233, 233, 233],
  'bg2|95000|0': [220, 220, 220],
  'bg2|85000|0': [197, 197, 197],
  'bg2|75000|0': [174, 174, 174],
  'bg2|65000|0': [151, 151, 151],
  'bg2|50000|0': [116, 116, 116],
  'bg2|25000|0': [58, 58, 58],
  'accent1|20000|80000': [193, 229, 245],
  'accent1|40000|60000': [131, 203, 235],
  'accent1|60000|40000': [70, 177, 225],
  'accent1|50000|50000': [100, 190, 230],
  'accent1|65000|35000': [54, 170, 223],
  'accent1|75000|25000': [33, 153, 207],
  'accent1|85000|15000': [28, 130, 176],
  'accent1|75000|0': [16, 72, 98],
  'accent1|65000|0': [14, 62, 85],
  'accent1|50000|0': [11, 48, 65],
  'accent2|20000|80000': [251, 227, 214],
  'accent2|40000|60000': [246, 198, 173],
  'accent2|60000|40000': [242, 170, 132],
  'accent2|50000|50000': [244, 184, 153],
  'accent2|65000|35000': [241, 163, 122],
  'accent2|75000|25000': [239, 148, 101],
  'accent2|85000|15000': [236, 134, 81],
  'accent2|95000|5000': [234, 120, 60],
  'accent2|95000|0': [232, 104, 37],
  'accent2|85000|0': [217, 90, 23],
  'accent2|75000|0': [192, 79, 21],
  'accent2|65000|0': [166, 69, 18],
  'accent2|50000|0': [128, 53, 14],
  'accent2|25000|0': [64, 26, 7],
  'accent3|20000|80000': [194, 241, 200],
  'accent3|40000|60000': [132, 226, 145],
  'accent3|60000|40000': [71, 212, 90],
  'accent3|50000|50000': [102, 219, 118],
  'accent3|65000|35000': [56, 208, 76],
  'accent3|75000|25000': [43, 184, 62],
  'accent3|85000|15000': [36, 153, 51],
  'accent3|95000|5000': [29, 122, 41],
  'accent3|95000|0': [24, 102, 34],
  'accent3|85000|0': [21, 91, 31],
  'accent3|75000|0': [19, 80, 27],
  'accent3|65000|0': [16, 70, 23],
  'accent3|25000|0': [6, 27, 9],
  'accent4|20000|80000': [202, 238, 251],
  'accent4|40000|60000': [150, 220, 248],
  'accent4|60000|40000': [97, 203, 244],
  'accent4|50000|50000': [123, 212, 246],
  'accent4|65000|35000': [84, 199, 243],
  'accent4|75000|25000': [57, 190, 241],
  'accent4|85000|15000': [31, 181, 239],
  'accent4|95000|5000': [16, 168, 226],
  'accent4|95000|0': [14, 150, 202],
  'accent4|85000|0': [13, 134, 181],
  'accent4|75000|0': [11, 118, 160],
  'accent4|65000|0': [10, 103, 138],
  'accent4|50000|0': [8, 79, 106],
  'accent4|25000|0': [4, 39, 53],
  'accent5|20000|80000': [242, 207, 238],
  'accent5|40000|60000': [229, 158, 221],
  'accent5|60000|40000': [216, 110, 204],
  'accent5|50000|50000': [222, 134, 213],
  'accent5|65000|35000': [213, 98, 200],
  'accent5|75000|25000': [206, 74, 191],
  'accent5|85000|15000': [196, 53, 180],
  'accent5|95000|5000': [172, 46, 158],
  'accent5|95000|0': [152, 41, 140],
  'accent5|85000|0': [136, 37, 125],
  'accent5|75000|0': [120, 32, 110],
  'accent5|65000|0': [104, 28, 96],
  'accent5|25000|0': [40, 11, 37],
  'accent6|20000|80000': [217, 242, 208],
  'accent6|40000|60000': [180, 229, 162],
  'accent6|60000|40000': [142, 217, 115],
  'accent6|50000|50000': [161, 223, 139],
  'accent6|65000|35000': [133, 213, 104],
  'accent6|75000|25000': [114, 207, 80],
  'accent6|85000|15000': [95, 200, 57],
  'accent6|95000|5000': [83, 179, 49],
  'accent6|95000|0': [74, 159, 44],
  'accent6|85000|0': [66, 142, 39],
  'accent6|75000|0': [59, 125, 35],
  'accent6|65000|0': [51, 109, 30],
};

/** Rows the cell renders one off on a channel. The engine owns the exact arithmetic for these. */
const NEAR: Record<string, Rgb> = {
  'tx1|50000|50000': [127, 127, 127],
  'bg1|50000|0': [127, 127, 127],
  'tx2|50000|0': [7, 20, 32],
  'tx2|25000|0': [3, 10, 16],
  'bg2|50000|50000': [244, 244, 244],
  'accent1|95000|5000': [24, 107, 145],
  'accent1|95000|0': [20, 91, 124],
  'accent1|85000|0': [18, 82, 111],
  'accent1|25000|0': [5, 24, 32],
  'accent3|50000|0': [13, 53, 18],
  'accent5|50000|0': [80, 22, 74],
  'accent6|50000|0': [39, 83, 23],
  'accent6|25000|0': [19, 42, 12],
};

describe('chart theme colours under lumMod/lumOff', () => {
  it('fills each series as Excel does', async () => {
    const fills = await seriesFills('excel-chart-theme-tints.xlsx');

    expect(Object.fromEntries(fills)).toEqual({
      S1: [70, 177, 225],
      S2: [192, 79, 21],
      S3: [217, 242, 208],
    });
  });

  it('renders the theme palette as Excel does', async () => {
    const fills = await seriesFills('excel-chart-theme-palette.xlsx');

    expect([...fills.keys()].sort()).toEqual([...Object.keys(EXACT), ...Object.keys(NEAR)].sort());
    for (const [name, rgb] of Object.entries(EXACT))
      expect([name, fills.get(name)]).toEqual([name, rgb]);
    for (const [name, rgb] of Object.entries(NEAR))
      expect([name, within1(fills.get(name), rgb)]).toEqual([name, true]);
  });
});
