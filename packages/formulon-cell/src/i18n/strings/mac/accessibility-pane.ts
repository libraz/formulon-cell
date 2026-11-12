/** The Mac accessibility checker task pane: summary, rule rows and drill-in cards. */
export interface MacAccessibilityPaneStrings {
  title: string;
  close: string;
  help: string;
  back: string;
  cleanTitle: string;
  cleanText: string;
  issuesTitle: string;
  issuesText: string;
  /** Drill-in header: `{index}/{total} - {section}`. */
  counter: string;
  tryFix: string;
  /** Rule-row state read by assistive tech: `{count}` issues, or none. */
  issueCount: string;
  noIssues: string;
  sections: {
    colorContrast: string;
    media: string;
    tables: string;
    structure: string;
    access: string;
  };
  rules: {
    contrast: string;
    redFormat: string;
    altText: string;
    tableHeader: string;
    mergedCells: string;
    defaultSheetName: string;
    restrictedAccess: string;
  };
  contrastTitle: string;
  contrastText: string;
  /** A suggested text colour swatch: `{color}`. */
  contrastSwatch: string;
  moreFontColors: string;
  fillColor: string;
  formatCellFont: string;
  tableHeaderTitle: string;
  tableHeaderText: string;
  insertHeaderRow: string;
  /** The status-bar button beside Ready. */
  statusClean: string;
  statusIssues: string;
}

/** Mac wording for this locale; not measured. */
export const macAccessibilityPaneEn: MacAccessibilityPaneStrings = {
  title: 'Accessibility Assistant',
  close: 'Close Accessibility Assistant',
  help: 'Help',
  back: 'Back',
  cleanTitle: 'Looks good.',
  cleanText: 'No issues found.',
  issuesTitle: 'Keep going.',
  issuesText: 'Fix the remaining issues in your document so everyone can access it.',
  counter: '{index}/{total} - {section}',
  tryFix: 'To fix, try one of these:',
  issueCount: '{count} issues',
  noIssues: 'No issues',
  sections: {
    colorContrast: 'Color and contrast',
    media: 'Media and illustrations',
    tables: 'Tables',
    structure: 'Document structure',
    access: 'Document access',
  },
  rules: {
    contrast: 'Hard-to-read text contrast',
    redFormat: "Don't use red formatting",
    altText: 'Missing alternative text',
    tableHeader: 'Missing table header',
    mergedCells: 'Merged cells',
    defaultSheetName: 'Default sheet names',
    restrictedAccess: 'Restricted access',
  },
  contrastTitle: 'Hard-to-read text contrast',
  contrastText:
    'The current text color is hard to see. Consider a high-contrast color so the text is clearly visible.',
  contrastSwatch: 'Font color {color}',
  moreFontColors: 'More Font Colors',
  fillColor: 'Fill Color',
  formatCellFont: 'Format Cell Font',
  tableHeaderTitle: 'Missing table header',
  tableHeaderText:
    'Table headers are essential to make the data in a table clear and easy to navigate and understand.',
  insertHeaderRow: 'Insert Header Row',
  statusClean: 'Accessibility: Good to go',
  statusIssues: 'Accessibility: Investigate',
};

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macAccessibilityPaneJa: MacAccessibilityPaneStrings = {
  title: 'ユーザー補助アシスタント',
  close: 'ユーザー補助アシスタント を閉じる',
  help: 'ヘルプ',
  back: '戻る',
  cleanTitle: '問題ないようです。',
  cleanText: '問題は見つかりませんでした。',
  issuesTitle: '引き続き頑張ってください。',
  issuesText: 'ドキュメントの残りの問題を修正して、すべてのユーザーがアクセスできるようにします。',
  counter: '{index}/{total} - {section}',
  tryFix: '修正するには、次のいずれかを試してください:',
  issueCount: '{count} 件の問題',
  noIssues: '問題なし',
  sections: {
    colorContrast: '色とコントラスト',
    media: 'メディアとイラスト',
    tables: 'テーブル',
    structure: 'ドキュメント構造',
    access: 'ドキュメント アクセス',
  },
  rules: {
    contrast: '読み取りにくいテキストのコントラスト',
    redFormat: '赤の書式を使用しない',
    altText: '代替テキストなし',
    tableHeader: 'テーブル ヘッダーがありません',
    mergedCells: '結合されたセルの使用',
    defaultSheetName: '既定のシート名',
    restrictedAccess: '制限付きアクセス',
  },
  contrastTitle: '読み取りにくいテキストのコントラストです',
  contrastText:
    '現在のテキストの色は見えにくいです。テキストが明確に表示されるように、ハイ コントラストの色を考慮してください。',
  contrastSwatch: 'フォントの色 {color}',
  moreFontColors: 'その他のフォントの色',
  fillColor: '塗りつぶしの色',
  formatCellFont: 'セルのフォントの書式設定',
  tableHeaderTitle: 'テーブル ヘッダーがありません',
  tableHeaderText:
    'テーブル ヘッダーは、テーブル内のデータをわかりやすくし、簡単に移動して理解を容易にするために不可欠です。',
  insertHeaderRow: 'ヘッダー行の挿入',
  statusClean: 'アクセシビリティ: 問題ありません',
  statusIssues: 'アクセシビリティ: 検討が必要です',
};

/** Mac app wording from the Excel bundle; not answered by it: close, help, cleanTitle, cleanText, issuesTitle, issuesText, and 16 more. */
export const macAccessibilityPaneDe: MacAccessibilityPaneStrings = {
  title: 'Barrierefreiheit-Assistent',
  close: 'Barrierefreiheits-Assistent schließen',
  help: 'Hilfe',
  back: 'Zurück',
  cleanTitle: 'Sieht gut aus.',
  cleanText: 'Keine Probleme gefunden.',
  issuesTitle: 'Weiter so.',
  issuesText:
    'Beheben Sie die verbleibenden Probleme im Dokument, damit alle Benutzer darauf zugreifen können.',
  counter: '{index}/{total} - {section}',
  tryFix: 'Versuchen Sie zum Beheben eine der folgenden Optionen:',
  issueCount: '{count} Probleme',
  noIssues: 'Keine Probleme',
  sections: {
    colorContrast: 'Farbe und Kontrast',
    media: 'Medien und Illustrationen',
    tables: 'Tabellen',
    structure: 'Dokumentstruktur',
    access: 'Dokumentzugriff',
  },
  rules: {
    contrast: 'Schwer lesbarer Textkontrast',
    redFormat: 'Keine rote Formatierung verwenden',
    altText: 'Fehlender Alternativtext',
    tableHeader: 'Fehlende Tabellenüberschrift',
    mergedCells: 'Verbundene Zellen',
    defaultSheetName: 'Standardblattnamen',
    restrictedAccess: 'Eingeschränkter Zugriff',
  },
  contrastTitle: 'Schwer lesbarer Textkontrast',
  contrastText:
    'Die aktuelle Textfarbe ist schwer zu erkennen. Verwenden Sie eine kontrastreiche Farbe, damit der Text gut sichtbar ist.',
  contrastSwatch: 'Schriftfarbe {color}',
  moreFontColors: 'Weitere Schriftfarben',
  fillColor: 'Füllfarbe',
  formatCellFont: 'Zellenschriftart formatieren',
  tableHeaderTitle: 'Fehlende Tabellenüberschrift',
  tableHeaderText:
    'Tabellenüberschriften machen die Daten einer Tabelle verständlich und leicht navigierbar.',
  insertHeaderRow: 'Kopfzeile einfügen',
  statusClean: 'Barrierefreiheit: Alles in Ordnung',
  statusIssues: 'Barrierefreiheit: Untersuchen',
};

/** Mac app wording from the Excel bundle; not answered by it: close, back, cleanTitle, cleanText, issuesTitle, issuesText, and 16 more. */
export const macAccessibilityPaneFr: MacAccessibilityPaneStrings = {
  title: 'Assistant d’accessibilité',
  close: 'Fermer Assistant Accessibilité',
  help: 'Aide',
  back: 'Précédent',
  cleanTitle: 'Tout semble correct.',
  cleanText: 'Aucun problème détecté.',
  issuesTitle: 'Continuez.',
  issuesText:
    'Corrigez les problèmes restants du document pour que tout le monde puisse y accéder.',
  counter: '{index}/{total} - {section}',
  tryFix: 'Pour corriger, essayez l’une des options suivantes :',
  issueCount: '{count} problèmes',
  noIssues: 'Aucun problème',
  sections: {
    colorContrast: 'Couleur et contraste',
    media: 'Média et illustrations',
    tables: 'Tableaux',
    structure: 'Document Structure',
    access: 'Accès au document',
  },
  rules: {
    contrast: 'Contraste de texte difficile à lire',
    redFormat: 'N’utilisez pas de mise en forme rouge',
    altText: 'Texte de remplacement manquant',
    tableHeader: 'En-tête de tableau manquant',
    mergedCells: 'Cellules fusionnées',
    defaultSheetName: 'Noms de feuille par défaut',
    restrictedAccess: 'Accès restreint',
  },
  contrastTitle: 'Contraste de texte dur à lire',
  contrastText:
    'La couleur actuelle du texte est difficile à voir. Choisissez une couleur à fort contraste pour que le texte soit bien visible.',
  contrastSwatch: 'Couleur de police {color}',
  moreFontColors: 'Autres couleurs de police',
  fillColor: 'Couleur de remplissage',
  formatCellFont: 'Mettre en forme la police de la cellule',
  tableHeaderTitle: 'En-tête de tableau manquant',
  tableHeaderText:
    'Les en-têtes de tableau rendent les données claires, faciles à parcourir et à comprendre.',
  insertHeaderRow: 'Insérer une ligne d’en-tête',
  statusClean: 'Accessibilité : vous pouvez continuer',
  statusIssues: 'Accessibilité : examen requis',
};

/** Mac app wording from the Excel bundle; not answered by it: close, back, cleanTitle, cleanText, issuesTitle, issuesText, and 16 more. */
export const macAccessibilityPaneZh: MacAccessibilityPaneStrings = {
  title: '辅助功能助手',
  close: '关闭辅助功能助手',
  help: '帮助',
  back: '返回',
  cleanTitle: '看起来不错。',
  cleanText: '未发现问题。',
  issuesTitle: '继续加油。',
  issuesText: '修复文档中剩余的问题，使所有人都能访问。',
  counter: '{index}/{total} - {section}',
  tryFix: '若要修复，请尝试以下操作之一:',
  issueCount: '{count} 个问题',
  noIssues: '无问题',
  sections: {
    colorContrast: '颜色和对比度',
    media: '媒体和插图',
    tables: '表格',
    structure: '文档结构',
    access: '文档访问',
  },
  rules: {
    contrast: '文本对比度难以阅读',
    redFormat: '不要使用红色格式',
    altText: '缺少替代文字',
    tableHeader: '缺少表格标题',
    mergedCells: '合并的单元格',
    defaultSheetName: '默认工作表名称',
    restrictedAccess: '限制访问',
  },
  contrastTitle: '难以阅读文本对比度',
  contrastText: '当前文本颜色难以看清。请考虑使用高对比度颜色，使文本清晰可见。',
  contrastSwatch: '字体颜色 {color}',
  moreFontColors: '其他字体颜色',
  fillColor: '填充颜色',
  formatCellFont: '设置单元格字体格式',
  tableHeaderTitle: '缺少表格标题',
  tableHeaderText: '表格标题对于使表格中的数据清晰、易于浏览和理解至关重要。',
  insertHeaderRow: '插入标题行',
  statusClean: '辅助功能: 一切就绪',
  statusIssues: '辅助功能: 需要调查',
};

/** Mac app wording from the Excel bundle; not answered by it: close, cleanTitle, cleanText, issuesTitle, issuesText, counter, and 15 more. */
export const macAccessibilityPaneKo: MacAccessibilityPaneStrings = {
  title: '접근성 도우미',
  close: '접근성 도우미 닫기',
  help: '도움말',
  back: '뒤로',
  cleanTitle: '문제가 없습니다.',
  cleanText: '문제를 찾지 못했습니다.',
  issuesTitle: '계속 진행하세요.',
  issuesText: '모든 사용자가 액세스할 수 있도록 문서의 남은 문제를 해결하세요.',
  counter: '{index}/{total} - {section}',
  tryFix: '해결하려면 다음 중 하나를 시도하세요:',
  issueCount: '문제 {count}개',
  noIssues: '문제 없음',
  sections: {
    colorContrast: '색 및 대비',
    media: '미디어 및 일러스트레이션',
    tables: '표',
    structure: '문서 구조',
    access: '문서 액세스',
  },
  rules: {
    contrast: '읽기 어려운 텍스트 대비',
    redFormat: '빨간색 서식 사용 안 함',
    altText: '대체 텍스트 없음',
    tableHeader: '표 머리글 누락',
    mergedCells: '병합된 셀',
    defaultSheetName: '기본 시트 이름',
    restrictedAccess: '제한된 액세스',
  },
  contrastTitle: '읽기 어려운 텍스트 대비',
  contrastText: '현재 텍스트 색은 잘 보이지 않습니다. 텍스트가 잘 보이도록 고대비 색을 고려하세요.',
  contrastSwatch: '글꼴 색 {color}',
  moreFontColors: '다른 글꼴 색',
  fillColor: '채우기 색',
  formatCellFont: '셀 글꼴 서식 지정',
  tableHeaderTitle: '표 머리글 누락',
  tableHeaderText: '표 머리글은 표의 데이터를 명확하고 쉽게 탐색하고 이해하는 데 필수적입니다.',
  insertHeaderRow: '머리글 행 삽입',
  statusClean: '접근성: 문제 없음',
  statusIssues: '접근성: 조사 필요',
};

/** Mac app wording from the Excel bundle; not answered by it: close, help, cleanTitle, cleanText, issuesTitle, issuesText, and 16 more. */
export const macAccessibilityPaneTh: MacAccessibilityPaneStrings = {
  title: 'ตัวช่วยสำหรับการเข้าถึง',
  close: 'ปิดผู้ช่วยการช่วยสำหรับการเข้าถึง',
  help: 'วิธีใช้',
  back: 'ย้อนกลับ',
  cleanTitle: 'ดูดีแล้ว',
  cleanText: 'ไม่พบปัญหา',
  issuesTitle: 'ทำต่อไป',
  issuesText: 'แก้ไขปัญหาที่เหลือในเอกสารเพื่อให้ทุกคนเข้าถึงได้',
  counter: '{index}/{total} - {section}',
  tryFix: 'เพื่อแก้ไข ให้ลองตัวเลือกใดตัวเลือกหนึ่งต่อไปนี้:',
  issueCount: '{count} ปัญหา',
  noIssues: 'ไม่มีปัญหา',
  sections: {
    colorContrast: 'สีและความคมชัด',
    media: 'สื่อและภาพประกอบ',
    tables: 'ตาราง',
    structure: 'โครงสร้างเอกสาร',
    access: 'การเข้าถึงเอกสาร',
  },
  rules: {
    contrast: 'คอนทราสต์ข้อความอ่านยาก',
    redFormat: 'อย่าใช้การจัดรูปแบบสีแดง',
    altText: 'ไม่มีข้อความแสดงแทน',
    tableHeader: 'ส่วนหัวของตารางขาดหายไป',
    mergedCells: 'เซลล์ที่ผสาน',
    defaultSheetName: 'ชื่อแผ่นงานเริ่มต้น',
    restrictedAccess: 'การเข้าถึงที่จำกัด',
  },
  contrastTitle: 'ความคมชัดของข้อความที่อ่านยาก',
  contrastText: 'สีข้อความปัจจุบันมองเห็นได้ยาก ลองใช้สีที่มีคอนทราสต์สูงเพื่อให้ข้อความชัดเจน',
  contrastSwatch: 'สีแบบอักษร {color}',
  moreFontColors: 'สีแบบอักษรเพิ่มเติม',
  fillColor: 'สีเติม',
  formatCellFont: 'จัดรูปแบบฟอนต์ของเซลล์',
  tableHeaderTitle: 'ส่วนหัวของตารางขาดหายไป',
  tableHeaderText: 'ส่วนหัวของตารางจำเป็นต่อการทำให้ข้อมูลในตารางชัดเจน นำทางและเข้าใจได้ง่าย',
  insertHeaderRow: 'แทรกแถวส่วนหัว',
  statusClean: 'การช่วยสำหรับการเข้าถึง: พร้อมใช้งาน',
  statusIssues: 'การช่วยสำหรับการเข้าถึง: ตรวจสอบ',
};
