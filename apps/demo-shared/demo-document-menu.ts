const WORKBOOK_EXTENSION = /\.(xlsx|xlsm|xlsb|xls)$/i;

/** The book name a rename from the document menu commits: trimmed and without
 *  a workbook extension, since the title bar shows the bare name. An empty
 *  entry keeps the current name. */
export const commitDemoBookName = (input: string, current: string): string => {
  const name = input.trim().replace(WORKBOOK_EXTENSION, '').trim();
  return name || current;
};
