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
  TextItalic24Regular,
  TextProofingTools24Regular,
  TextStrikethrough24Regular,
  TextUnderline24Regular,
  TextWrap24Regular,
  Translate24Regular,
  Window24Regular,
  ZoomFit24Regular,
} from '@fluentui/react-icons';
import type { ReactElement } from 'react';

export type IconName =
  | 'accessibility'
  | 'add'
  | 'paste'
  | 'cut'
  | 'copy'
  | 'paint'
  | 'undo'
  | 'redo'
  | 'fontGrow'
  | 'fontShrink'
  | 'bold'
  | 'italic'
  | 'underline'
  | 'strike'
  | 'fontColor'
  | 'fillColor'
  | 'top'
  | 'middle'
  | 'currency'
  | 'percent'
  | 'comma'
  | 'decDown'
  | 'decUp'
  | 'autosum'
  | 'alignLeft'
  | 'alignCenter'
  | 'alignRight'
  | 'borders'
  | 'merge'
  | 'wrap'
  | 'freeze'
  | 'insertRows'
  | 'deleteRows'
  | 'insertCols'
  | 'deleteCols'
  | 'filter'
  | 'sortAsc'
  | 'sortDesc'
  | 'table'
  | 'tableStyle'
  | 'conditional'
  | 'formatCells'
  | 'removeDuplicates'
  | 'link'
  | 'pen'
  | 'eraser'
  | 'page'
  | 'margins'
  | 'orientation'
  | 'scale'
  | 'print'
  | 'function'
  | 'names'
  | 'trace'
  | 'dependents'
  | 'clearArrows'
  | 'options'
  | 'watch'
  | 'comment'
  | 'commentAdd'
  | 'commentMultiple'
  | 'protect'
  | 'zoom'
  | 'script'
  | 'addIn'
  | 'pdf'
  | 'goTo'
  | 'find'
  | 'findSelect'
  | 'translate'
  | 'spelling'
  | 'chart'
  | 'clear';

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
