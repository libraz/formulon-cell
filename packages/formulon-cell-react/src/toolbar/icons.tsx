import {
  Accessibility24Regular,
  Add24Regular,
  AlignCenterHorizontal24Regular,
  AlignLeft24Regular,
  AlignRight24Regular,
  AlignSpaceEvenlyVertical24Regular,
  AppsAddIn24Regular,
  ArrowRedo24Regular,
  ArrowSortDown24Regular,
  ArrowSortUp24Regular,
  ArrowTrendingLines24Regular,
  ArrowUndo24Regular,
  Autosum24Regular,
  Book24Regular,
  BorderAll24Regular,
  Calculator24Regular,
  ChartMultiple24Regular,
  ClipboardPaste24Regular,
  Code24Regular,
  Comma24Regular,
  Comment24Regular,
  CommentAdd24Regular,
  CommentMultiple24Regular,
  Copy24Regular,
  CurrencyDollarEuro24Regular,
  Cut24Regular,
  DecimalArrowLeft24Regular,
  DecimalArrowRight24Regular,
  Delete24Regular,
  DocumentMargins24Regular,
  DocumentOnePage24Regular,
  DocumentPdf24Regular,
  Eraser24Regular,
  Filter24Regular,
  FontDecrease24Regular,
  FontIncrease24Regular,
  Highlight24Regular,
  Link24Regular,
  LockClosed24Regular,
  Merge24Regular,
  Orientation24Regular,
  PaintBrush24Regular,
  Pen24Regular,
  Print24Regular,
  Search24Regular,
  SearchSquare24Regular,
  Settings24Regular,
  Shield24Regular,
  TableDeleteColumn24Regular,
  TableDeleteRow24Regular,
  TableDismiss24Regular,
  TableFreezeColumnAndRow24Regular,
  TableInsertColumn24Regular,
  TableInsertRow24Regular,
  TableLightning24Regular,
  TableSettings24Regular,
  TableSimple24Regular,
  TagPercent24Regular,
  TextAlignCenter24Regular,
  TextBold24Regular,
  TextBulletListSquare24Regular,
  TextColor24Regular,
  TextIndentDecrease24Regular,
  TextIndentIncrease24Regular,
  TextItalic24Regular,
  TextProofingTools24Regular,
  TextStrikethrough24Regular,
  TextUnderline24Regular,
  TextWrap24Regular,
  Translate24Regular,
  Window24Regular,
  ZoomFit24Regular,
} from '@fluentui/react-icons';
import type { IconName } from '@libraz/formulon-cell';
import type { ReactElement } from 'react';

export type { IconName };

export const Icon = ({ name }: { name: IconName }): ReactElement => {
  const common = { className: 'demo__rb-icon', 'aria-hidden': true };
  switch (name) {
    case 'accessibility':
      return <Accessibility24Regular {...common} />;
    case 'add':
      return <Add24Regular {...common} />;
    case 'paste':
      return <ClipboardPaste24Regular {...common} />;
    case 'cut':
      return <Cut24Regular {...common} />;
    case 'copy':
      return <Copy24Regular {...common} />;
    case 'paint':
      return <PaintBrush24Regular {...common} />;
    case 'undo':
      return <ArrowUndo24Regular {...common} />;
    case 'redo':
      return <ArrowRedo24Regular {...common} />;
    case 'fontGrow':
      return <FontIncrease24Regular {...common} />;
    case 'fontShrink':
      return <FontDecrease24Regular {...common} />;
    case 'bold':
      return <TextBold24Regular {...common} />;
    case 'italic':
      return <TextItalic24Regular {...common} />;
    case 'underline':
      return <TextUnderline24Regular {...common} />;
    case 'strike':
      return <TextStrikethrough24Regular {...common} />;
    case 'fontColor':
      return <TextColor24Regular {...common} />;
    case 'fillColor':
      return <Highlight24Regular {...common} />;
    case 'top':
      return <AlignSpaceEvenlyVertical24Regular {...common} />;
    case 'middle':
      return <TextAlignCenter24Regular {...common} />;
    case 'bottomAlign':
      return pathIcon(common, [
        'M4 20.25c0-.41.34-.75.75-.75h14.5a.75.75 0 0 1 0 1.5H4.75a.75.75 0 0 1-.75-.75ZM6.25 3h11.5C18.99 3 20 4 20 5.25v1.5C20 7.99 19 9 17.75 9H6.25C5.01 9 4 8 4 6.75v-1.5C4 4.01 5 3 6.25 3Zm0 1.5a.75.75 0 0 0-.75.75v1.5c0 .41.34.75.75.75h11.5c.41 0 .75-.34.75-.75v-1.5a.75.75 0 0 0-.75-.75H6.25Zm0 7h11.5c1.24 0 2.25 1 2.25 2.25v1.5c0 1.24-1 2.25-2.25 2.25H6.25C5.01 17.5 4 16.5 4 15.25v-1.5c0-1.24 1-2.25 2.25-2.25Zm0 1.5a.75.75 0 0 0-.75.75v1.5c0 .41.34.75.75.75h11.5c.41 0 .75-.34.75-.75v-1.5a.75.75 0 0 0-.75-.75H6.25Z',
      ]);
    case 'textOrientation':
      return pathIcon(common, [
        'M4.22 17.72 15.72 6.22a.75.75 0 1 1 1.06 1.06l-11.5 11.5a.75.75 0 0 1-1.06-1.06Zm2.06-8.44a.75.75 0 0 1 0-1.06l3-3a.75.75 0 0 1 1.06 0l5.44 5.44a.75.75 0 0 1-1.06 1.06L13.5 10.5 10 14l1.22 1.22a.75.75 0 0 1-1.06 1.06L4.72 10.84a.75.75 0 1 1 1.06-1.06L7 11l3.5-3.5-1.22-1.22-1.94 1.94a.75.75 0 0 1-1.06 0ZM18.25 4a.75.75 0 0 1 .75.75V9.5a.75.75 0 0 1-1.5 0V6.56l-2.22 2.22a.75.75 0 0 1-1.06-1.06l2.22-2.22H13.5a.75.75 0 0 1 0-1.5h4.75Z',
      ]);
    case 'currency':
      return <CurrencyDollarEuro24Regular {...common} />;
    case 'percent':
      return <TagPercent24Regular {...common} />;
    case 'comma':
      return <Comma24Regular {...common} />;
    case 'decDown':
      return <DecimalArrowLeft24Regular {...common} />;
    case 'decUp':
      return <DecimalArrowRight24Regular {...common} />;
    case 'autosum':
      return <Autosum24Regular {...common} />;
    case 'alignLeft':
      return <AlignLeft24Regular {...common} />;
    case 'alignCenter':
      return <AlignCenterHorizontal24Regular {...common} />;
    case 'alignRight':
      return <AlignRight24Regular {...common} />;
    case 'indentDecrease':
      return <TextIndentDecrease24Regular {...common} />;
    case 'indentIncrease':
      return <TextIndentIncrease24Regular {...common} />;
    case 'borders':
      return <BorderAll24Regular {...common} />;
    case 'merge':
      return <Merge24Regular {...common} />;
    case 'wrap':
      return <TextWrap24Regular {...common} />;
    case 'freeze':
      return <TableFreezeColumnAndRow24Regular {...common} />;
    case 'insertRows':
      return <TableInsertRow24Regular {...common} />;
    case 'deleteRows':
      return <TableDeleteRow24Regular {...common} />;
    case 'insertCols':
      return <TableInsertColumn24Regular {...common} />;
    case 'deleteCols':
      return <TableDeleteColumn24Regular {...common} />;
    case 'filter':
      return <Filter24Regular {...common} />;
    case 'sortAsc':
      return <ArrowSortUp24Regular {...common} />;
    case 'sortDesc':
      return <ArrowSortDown24Regular {...common} />;
    case 'table':
      return <TableSimple24Regular {...common} />;
    case 'tableStyle':
      return <TableSettings24Regular {...common} />;
    case 'conditional':
      return <TableLightning24Regular {...common} />;
    case 'formatCells':
      return <TableSettings24Regular {...common} />;
    case 'removeDuplicates':
      return <TableDismiss24Regular {...common} />;
    case 'link':
      return <Link24Regular {...common} />;
    case 'pen':
      return <Pen24Regular {...common} />;
    case 'eraser':
      return <Eraser24Regular {...common} />;
    case 'page':
      return <DocumentOnePage24Regular {...common} />;
    case 'margins':
      return <DocumentMargins24Regular {...common} />;
    case 'orientation':
      return <Orientation24Regular {...common} />;
    case 'scale':
      return <ZoomFit24Regular {...common} />;
    case 'print':
      return <Print24Regular {...common} />;
    case 'function':
      return <Calculator24Regular {...common} />;
    case 'names':
      return <Book24Regular {...common} />;
    case 'trace':
      return <ArrowTrendingLines24Regular {...common} />;
    case 'dependents':
      return <TextBulletListSquare24Regular {...common} />;
    case 'clearArrows':
      return <Delete24Regular {...common} />;
    case 'options':
      return <Settings24Regular {...common} />;
    case 'watch':
      return <Window24Regular {...common} />;
    case 'comment':
      return <Comment24Regular {...common} />;
    case 'protect':
      return <LockClosed24Regular {...common} />;
    case 'zoom':
      return <ZoomFit24Regular {...common} />;
    case 'script':
      return <Code24Regular {...common} />;
    case 'addIn':
      return <AppsAddIn24Regular {...common} />;
    case 'pdf':
      return <DocumentPdf24Regular {...common} />;
    case 'goTo':
      return <Search24Regular {...common} />;
    case 'find':
      return <Search24Regular {...common} />;
    case 'findSelect':
      return <SearchSquare24Regular {...common} />;
    case 'translate':
      return <Translate24Regular {...common} />;
    case 'spelling':
      return <TextProofingTools24Regular {...common} />;
    case 'chart':
      return <ChartMultiple24Regular {...common} />;
    case 'commentAdd':
      return <CommentAdd24Regular {...common} />;
    case 'commentMultiple':
      return <CommentMultiple24Regular {...common} />;
    case 'clear':
      return <Shield24Regular {...common} />;
  }
};

const pathIcon = (
  common: { className: string; 'aria-hidden': boolean },
  paths: readonly string[],
): ReactElement => (
  <svg {...common} viewBox="0 0 24 24" fill="currentColor" focusable="false">
    {paths.map((d) => (
      <path key={d} d={d} />
    ))}
  </svg>
);
