/** Wording of the PivotTable Options dialog. */
export interface MacPivotOptionsStrings {
  title: string;
  name: string;
  tabDisplay: string;
  tabLayout: string;
  tabData: string;
  tabAltText: string;
  display: {
    group: string;
    expandButtons: string;
    printExpandButtons: string;
    contextualTooltips: string;
    fieldCaptions: string;
    classicLayout: string;
    errorValues: string;
    emptyCells: string;
    noDataGroup: string;
    noDataRows: string;
    noDataColumns: string;
    printGroup: string;
    repeatRowLabels: string;
    printTitles: string;
    itemLabels: string;
  };
  layout: {
    group: string;
    multipleFilters: string;
    autofitColumns: string;
    preserveFormatting: string;
    mergeLabels: string;
    compactIndent: string;
    reportFilterGroup: string;
    fieldOrder: string;
    downThenOver: string;
    overThenDown: string;
    fieldsPerColumn: string;
    sortGroup: string;
    customLists: string;
    sortAscending: string;
    sortDataSource: string;
  };
  data: {
    group: string;
    saveSourceData: string;
    enableShowDetails: string;
    retainDeletedItems: string;
    itemsToRetain: string;
    retainAutomatic: string;
    retainMax: string;
    refreshOnOpen: string;
    refreshOnChange: string;
    connectionGroup: string;
    savePassword: string;
    backgroundQuery: string;
  };
  altText: {
    title: string;
    description: string;
    help: string;
  };
  ok: string;
  cancel: string;
  /** Reason shown on an option this build cannot apply. */
  unavailable: string;
}

/** Mac wording for this locale; not measured. */
export const macPivotOptionsEn: MacPivotOptionsStrings = {
  title: 'PivotTable Options',
  name: 'PivotTable Name:',
  tabDisplay: 'Display',
  tabLayout: 'Layout',
  tabData: 'Data',
  tabAltText: 'Alt Text',
  display: {
    group: 'Display:',
    expandButtons: 'Expand/collapse buttons',
    printExpandButtons: 'Print expand/collapse buttons',
    contextualTooltips: 'Contextual tooltips',
    fieldCaptions: 'Field captions and filters',
    classicLayout: 'Classic PivotTable layout',
    errorValues: 'For error values show:',
    emptyCells: 'For empty cells show:',
    noDataGroup: 'Show items with no data:',
    noDataRows: 'On rows',
    noDataColumns: 'On columns',
    printGroup: 'Print:',
    repeatRowLabels: 'Repeat row labels on each page',
    printTitles: 'Page, row and column headings',
    itemLabels: 'Item labels when no fields are on the table',
  },
  layout: {
    group: 'Layout:',
    multipleFilters: 'Allow multiple filters per field',
    autofitColumns: 'Autofit column widths on update',
    preserveFormatting: 'Preserve cell formatting on update',
    mergeLabels: 'Merge and center cells with labels',
    compactIndent: 'When in compact form indent row labels:',
    reportFilterGroup: 'Report filter:',
    fieldOrder: 'Display fields:',
    downThenOver: 'Down, then over',
    overThenDown: 'Over, then down',
    fieldsPerColumn: 'Fields per column:',
    sortGroup: 'Sort:',
    customLists: 'Use custom lists when sorting',
    sortAscending: 'Sort A to Z',
    sortDataSource: 'Sort in data source order',
  },
  data: {
    group: 'Data:',
    saveSourceData: 'Save source data with file',
    enableShowDetails: 'Enable show details',
    retainDeletedItems: 'Retain items deleted from the data source',
    itemsToRetain: 'Number to retain per field:',
    retainAutomatic: 'Automatic',
    retainMax: 'Max',
    refreshOnOpen: 'Refresh data when opening the file',
    refreshOnChange: 'Automatically refresh when the source data changes',
    connectionGroup: 'Connection:',
    savePassword: 'Save password with file',
    backgroundQuery: 'Run queries in the background',
  },
  altText: {
    title: 'Title',
    description: 'Description',
    help: 'Titles and descriptions provide alternative, text-based representations of the information contained in tables, diagrams, images, and other objects. This information is useful for people with vision or cognitive impairments who may not be able to see or understand the object.\n\nA title can be read to a person with a disability and is used to determine whether they wish to hear the description of the content.',
  },
  ok: 'OK',
  cancel: 'Cancel',
  unavailable: 'Not available in this version.',
};

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macPivotOptionsJa: MacPivotOptionsStrings = {
  title: 'ピボットテーブル オプション',
  name: 'ピボットテーブル名:',
  tabDisplay: '表示形式',
  tabLayout: 'レイアウト',
  tabData: 'データ',
  tabAltText: '代替テキスト',
  display: {
    group: '表示:',
    expandButtons: '展開/折りたたみボタン',
    printExpandButtons: '展開/折りたたみボタンをプリントする',
    contextualTooltips: '操作に応じたヒント',
    fieldCaptions: 'フィールドのタイトルとフィルター',
    classicLayout: '従来のピボットテーブル レイアウト',
    errorValues: 'エラー値に表示する値:',
    emptyCells: '空白セルに表示する値:',
    noDataGroup: 'データのないアイテム:',
    noDataRows: '行に表示する',
    noDataColumns: '列に表示する',
    printGroup: '印刷:',
    repeatRowLabels: '各印刷ページに行ラベルを繰り返す',
    printTitles: '印刷タイトルを設定する',
    itemLabels: '値の領域にフィールドがない場合のアイテムのラベル',
  },
  layout: {
    group: 'レイアウト:',
    multipleFilters: '1 つのフィールドに複数のフィルターを使用可能にする',
    autofitColumns: '更新時に列幅を自動調整する',
    preserveFormatting: '更新時にセル書式を保持する',
    mergeLabels: 'セルとラベルを結合して中央揃えにする',
    compactIndent: 'コンパクト形式での行ラベルのインデント:',
    reportFilterGroup: 'レポート フィルター:',
    fieldOrder: 'フィールドの表示:',
    downThenOver: '下から上へ',
    overThenDown: '上から下へ',
    fieldsPerColumn: '列ごとのフィールド数:',
    sortGroup: '並べ替え:',
    customLists: '並べ替え時にユーザー設定リストを使用する',
    sortAscending: '昇順で並べ替え',
    sortDataSource: 'データ ソース順で並べ替え',
  },
  data: {
    group: 'データ:',
    saveSourceData: 'ファイルに元のデータを保存する',
    enableShowDetails: '詳細を表示可能にする',
    retainDeletedItems: 'データ ソースから削除されたアイテムの保持',
    itemsToRetain: '1 フィールドに保持する数:',
    retainAutomatic: '自動',
    retainMax: '最大値',
    refreshOnOpen: 'ファイルを開くときにデータを更新する',
    refreshOnChange: 'ソース データが変更されたときに自動的に更新する',
    connectionGroup: '接続:',
    savePassword: 'パスワードをファイルに保存する',
    backgroundQuery: 'バックグラウンドでクエリを実行する',
  },
  altText: {
    title: 'タイトル',
    description: '説明',
    help: 'タイトルと説明は、表、図表、画像、その他のオブジェクト内の情報をテキストで表した代替情報です。この情報は、オブジェクトを見たり理解したりするのが困難な、視覚障碍や認識障碍のある方にとって役立ちます。\n\nタイトルは障碍のある方が読み取れるようになっており、コンテンツの説明を音声で聞くかどうかを判断するために使用されます。',
  },
  ok: 'OK',
  cancel: 'キャンセル',
  unavailable: 'このバージョンでは使用できません。',
};

/** Mac app wording from the Excel bundle; not answered by it: display.group, display.noDataGroup, display.noDataRows, display.noDataColumns, display.repeatRowLabels, display.printTitles, and 9 more. */
export const macPivotOptionsDe: MacPivotOptionsStrings = {
  title: 'PivotTable-Optionen',
  name: 'PivotTable-Name:',
  tabDisplay: 'Anzeigen',
  tabLayout: 'Layout',
  tabData: 'Daten',
  tabAltText: 'Alternativtext',
  display: {
    group: 'Anzeige:',
    expandButtons: 'Schaltflächen zum Erweitern/Reduzieren',
    printExpandButtons: 'Schaltflächen zum Erweitern/Reduzieren drucken',
    contextualTooltips: 'Kontextabhängige QuickInfo',
    fieldCaptions: 'Feldbeschriftungen und Filter',
    classicLayout: 'Klassisches PivotTable-Layout',
    errorValues: 'Fehlerwerte anzeigen als:',
    emptyCells: 'Leere Zellen anzeigen als:',
    noDataGroup: 'Elemente ohne Daten anzeigen:',
    noDataRows: 'In Zeilen',
    noDataColumns: 'In Spalten',
    printGroup: 'Drucken:',
    repeatRowLabels: 'Zeilenbeschriftungen auf jeder Seite wiederholen',
    printTitles: 'Seiten-, Zeilen- und Spaltenüberschriften',
    itemLabels: 'Elementbeschriftungen, wenn sich keine Felder in der Tabelle befinden',
  },
  layout: {
    group: 'Layout:',
    multipleFilters: 'Mehrere Filter pro Feld zulassen',
    autofitColumns: 'Spaltenbreiten bei Aktualisierung automatisch anpassen',
    preserveFormatting: 'Zellformatierung bei Aktualisierung beibehalten',
    mergeLabels: 'Zellen mit Beschriftungen verbinden und zentrieren',
    compactIndent: 'Zeilenbeschriftungen in Kurzformat einrücken:',
    reportFilterGroup: 'Berichtsfilter:',
    fieldOrder: 'Felder anzeigen:',
    downThenOver: 'Zuerst nach unten',
    overThenDown: 'Zuerst nach rechts',
    fieldsPerColumn: 'Felder pro Spalte:',
    sortGroup: 'Sortieren:',
    customLists: 'Beim Sortieren benutzerdefinierte Listen verwenden',
    sortAscending: 'Von A bis Z sortieren',
    sortDataSource: 'In Datenquellenreihenfolge sortieren',
  },
  data: {
    group: 'Daten:',
    saveSourceData: 'Quelldaten mit der Datei speichern',
    enableShowDetails: "'Details anzeigen' aktivieren",
    retainDeletedItems: 'Elemente beibehalten, die aus der Datenquelle gelöscht wurden',
    itemsToRetain: 'Anzahl der beizubehaltenden Elemente pro Feld:',
    retainAutomatic: 'Automatisch',
    retainMax: 'Maximum',
    refreshOnOpen: 'Daten beim Öffnen der Datei aktualisieren',
    refreshOnChange: 'Automatisch aktualisieren, wenn sich die Quelldaten ändern',
    connectionGroup: 'Verbindung:',
    savePassword: 'Kennwort mit der Datei speichern',
    backgroundQuery: 'Abfragen im Hintergrund ausführen',
  },
  altText: {
    title: 'Titel',
    description: 'Beschreibung',
    help: 'Titel und Beschreibungen stellen alternative, textbasierte Darstellungen der in Tabellen, Diagrammen, Bildern und anderen Objekten enthaltenen Informationen zur Verfügung. Diese Informationen sind für Personen mit Sehschwächen oder kognitiven Störungen nützlich, die ggf. nicht in der Lage sind, das Objekt zu erkennen oder zu verstehen.\n\nEin Titel kann Personen mit Beeinträchtigungen vorgelesen werden. Diese können dann entscheiden, ob sie die Beschreibung des Inhalts hören möchten.',
  },
  ok: 'OK',
  cancel: 'Abbrechen',
  unavailable: 'In dieser Version nicht verfügbar.',
};

/** Mac app wording from the Excel bundle; not answered by it: display.group, display.noDataGroup, display.noDataRows, display.noDataColumns, display.repeatRowLabels, display.printTitles, and 8 more. */
export const macPivotOptionsFr: MacPivotOptionsStrings = {
  title: 'Options du tableau croisé dynamique',
  name: 'Nom du tableau croisé dynamique :',
  tabDisplay: 'Affichage',
  tabLayout: 'Disposition',
  tabData: 'Données',
  tabAltText: 'Texte de remplacement',
  display: {
    group: 'Affichage :',
    expandButtons: 'Boutons Développer/Réduire',
    printExpandButtons: 'Imprimer les boutons Développer/Réduire',
    contextualTooltips: 'Info-bulles contextuelles',
    fieldCaptions: 'Légendes des champs et filtres',
    classicLayout: 'Disposition classique de tableau croisé dynamique',
    errorValues: 'Pour les valeurs d’erreur, afficher :',
    emptyCells: 'Pour les cellules vides, afficher :',
    noDataGroup: 'Afficher les éléments sans données :',
    noDataRows: 'Dans des lignes',
    noDataColumns: 'Dans des colonnes',
    printGroup: 'Imprimer :',
    repeatRowLabels: 'Répéter les étiquettes de lignes sur chaque page',
    printTitles: 'Page, en-têtes de ligne et de colonne',
    itemLabels: "Étiquettes d'élément si le tableau ne contient pas de champs",
  },
  layout: {
    group: 'Disposition :',
    multipleFilters: 'Autoriser plusieurs filtres par champ',
    autofitColumns: 'Ajuster automatiquement la largeur des colonnes lors de la mise à jour',
    preserveFormatting: 'Conserver la mise en forme des cellules lors de la mise à jour',
    mergeLabels: 'Fusionner et centrer les cellules avec les étiquettes',
    compactIndent: 'En forme compacte, mettre en retrait les étiquettes de lignes :',
    reportFilterGroup: 'Filtre du rapport :',
    fieldOrder: 'Afficher les champs :',
    downThenOver: 'Vers le bas, puis à droite',
    overThenDown: 'À droite, puis vers le bas',
    fieldsPerColumn: 'Champs par colonne :',
    sortGroup: 'Trier :',
    customLists: 'Utiliser des listes personnalisées lors du tri',
    sortAscending: 'Trier de A à Z',
    sortDataSource: 'Trier par ordre de source de données',
  },
  data: {
    group: 'Données :',
    saveSourceData: 'Enregistrer les données sources avec le fichier',
    enableShowDetails: "Activer l'affichage des détails",
    retainDeletedItems: 'Conserver les éléments supprimés de la source de données',
    itemsToRetain: 'Nombre à conserver par champ :',
    retainAutomatic: 'Automatique',
    retainMax: 'Max',
    refreshOnOpen: 'Actualiser les données lors de l’ouverture du fichier',
    refreshOnChange: 'Actualiser automatiquement en cas de modification des données sources',
    connectionGroup: 'Connexion :',
    savePassword: 'Enregistrer le mot de passe avec le fichier',
    backgroundQuery: 'Exécuter les requêtes en arrière-plan',
  },
  altText: {
    title: 'Titre',
    description: 'Description',
    help: 'Les titres et les descriptions offrent des représentations textuelles alternatives des informations contenues dans les tableaux, diagrammes, images et autres objets. Ces informations sont utiles aux personnes souffrant d’un trouble visuel ou cognitif, dans la mesure où elles ne peuvent peut-être pas voir ou comprendre l’objet. Un titre peut être lu à une personne souffrant d’un handicap et lui permet de déterminer si elle souhaite entendre la description du contenu.',
  },
  ok: 'OK',
  cancel: 'Annuler',
  unavailable: 'Non disponible dans cette version.',
};

/** Mac app wording from the Excel bundle; not answered by it: display.group, display.noDataGroup, display.noDataRows, display.noDataColumns, display.repeatRowLabels, display.printTitles, and 8 more. */
export const macPivotOptionsZh: MacPivotOptionsStrings = {
  title: '数据透视表选项',
  name: '数据透视表名称:',
  tabDisplay: '显示',
  tabLayout: '布局',
  tabData: '数据',
  tabAltText: '替换文字',
  display: {
    group: '显示:',
    expandButtons: '展开/折叠按钮',
    printExpandButtons: '打印展开/折叠按钮',
    contextualTooltips: '上下文工具提示',
    fieldCaptions: '字段标题和筛选',
    classicLayout: '经典数据透视表布局',
    errorValues: '对于错误值显示:',
    emptyCells: '对于空单元格显示:',
    noDataGroup: '显示没有数据的项目:',
    noDataRows: '行上',
    noDataColumns: '列上',
    printGroup: '打印:',
    repeatRowLabels: '重复每页上的行标签',
    printTitles: '页、行和列标题',
    itemLabels: '表上无字段时的项目标签',
  },
  layout: {
    group: '布局:',
    multipleFilters: '每个字段允许多个筛选',
    autofitColumns: '更新时自动调整列宽',
    preserveFormatting: '更新时保留单元格格式',
    mergeLabels: '合并且使带标签的单元格居中',
    compactIndent: '压缩表单中缩进行标签:',
    reportFilterGroup: '报表筛选:',
    fieldOrder: '显示域:',
    downThenOver: '先列后行',
    overThenDown: '先行后列',
    fieldsPerColumn: '每列字段数:',
    sortGroup: '排序:',
    customLists: '排序时使用自定义列表',
    sortAscending: '升序',
    sortDataSource: '按数据源顺序排序',
  },
  data: {
    group: '数据:',
    saveSourceData: '保存源数据及文件',
    enableShowDetails: '启用显示详细信息',
    retainDeletedItems: '保留从数据源删除的项目',
    itemsToRetain: '每个字段需保留的数字:',
    retainAutomatic: '自动',
    retainMax: '最大值',
    refreshOnOpen: '打开文件时刷新数据',
    refreshOnChange: '源数据更改时自动刷新',
    connectionGroup: '连接:',
    savePassword: '保存密码及文件',
    backgroundQuery: '在后台运行查询',
  },
  altText: {
    title: '标题',
    description: '描述',
    help: '标题和描述为表格、图示、图像和其他对象中包含的信息提供替代性的文字表示。此信息对于有视觉或认知障碍而无法看清楚或理解对象的人十分有用。\n\n系统可为残障人士朗读标题，让其判断是否希望听到内容的描述。',
  },
  ok: '确定',
  cancel: '取消',
  unavailable: '此版本中不可用。',
};

/** Mac app wording from the Excel bundle; not answered by it: display.group, display.noDataGroup, display.noDataRows, display.noDataColumns, display.repeatRowLabels, display.printTitles, and 9 more. */
export const macPivotOptionsKo: MacPivotOptionsStrings = {
  title: '피벗 테이블 옵션',
  name: '피벗 테이블 이름:',
  tabDisplay: '화면 표시',
  tabLayout: '레이아웃',
  tabData: '데이터',
  tabAltText: '대체 텍스트',
  display: {
    group: '표시:',
    expandButtons: '확장/축소 단추',
    printExpandButtons: '확장/축소 단추 인쇄',
    contextualTooltips: '상황별 도구 설명',
    fieldCaptions: '필드 캡션 및 필터',
    classicLayout: '클래식 피벗 테이블 모양',
    errorValues: '오류 값 표시:',
    emptyCells: '빈 셀 표시:',
    noDataGroup: '데이터가 없는 항목 표시:',
    noDataRows: '행',
    noDataColumns: '열',
    printGroup: '인쇄:',
    repeatRowLabels: '각 페이지에 행 레이블 반복',
    printTitles: '페이지, 행 및 열 머리글',
    itemLabels: '테이블에 필드가 없는 경우 항목 레이블',
  },
  layout: {
    group: '모양:',
    multipleFilters: '필드 하나에 여러 필터 허용',
    autofitColumns: '업데이트 시 열 자동 맞춤',
    preserveFormatting: '업데이트 시 셀 서식 유지',
    mergeLabels: '레이블이 있는 셀 병합 및 가운데 맞춤',
    compactIndent: '압축 형식일 때 행 레이블 들여쓰기:',
    reportFilterGroup: '보고서 필터',
    fieldOrder: '필드 표시:',
    downThenOver: '행 우선',
    overThenDown: '열 우선',
    fieldsPerColumn: '열당 필드 수:',
    sortGroup: '정렬:',
    customLists: '정렬에 사용자 지정 목록 사용',
    sortAscending: '텍스트 오름차순 정렬',
    sortDataSource: '데이터 원본 순서로 정렬',
  },
  data: {
    group: '데이터:',
    saveSourceData: '파일에 원본 데이터 저장',
    enableShowDetails: '하위 수준 표시 사용',
    retainDeletedItems: '데이터 원본에서 삭제된 항목 유지',
    itemsToRetain: '필드별로 유지할 수:',
    retainAutomatic: '자동',
    retainMax: '최대값',
    refreshOnOpen: '파일을 열 때 데이터 새로 고침',
    refreshOnChange: '원본 데이터 변경 시 자동으로 새로 고침',
    connectionGroup: '연결:',
    savePassword: '파일과 함께 암호 저장',
    backgroundQuery: '백그라운드에서 쿼리 실행',
  },
  altText: {
    title: '제목',
    description: '설명',
    help: '제목 및 설명에서는 표, 다이어그램, 이미지, 기타 개체에 들어 있는 정보를 텍스트로 대체하여 표현할 수 있습니다. 이 정보는 개체를 보거나 이해할 수 없는 시각 또는 인지 장애가 있는 사용자에게 유용합니다.',
  },
  ok: '확인',
  cancel: '취소',
  unavailable: '이 버전에서는 사용할 수 없습니다.',
};

/** Mac app wording from the Excel bundle; not answered by it: display.group, display.noDataGroup, display.noDataRows, display.noDataColumns, display.repeatRowLabels, display.printTitles, and 9 more. */
export const macPivotOptionsTh: MacPivotOptionsStrings = {
  title: 'ตัวเลือก PivotTable',
  name: 'ชื่อ PivotTable:',
  tabDisplay: 'แสดง',
  tabLayout: 'เค้าโครง',
  tabData: 'ข้อมูล',
  tabAltText: 'ข้อความแสดงแทน',
  display: {
    group: 'การแสดง:',
    expandButtons: 'ปุ่มขยาย/ยุบ',
    printExpandButtons: 'พิมพ์ปุ่มขยาย/ยุบ',
    contextualTooltips: 'คำแนะนำเครื่องมือตามบริบท',
    fieldCaptions: 'คำอธิบายเขตข้อมูลและตัวกรอง',
    classicLayout: 'เค้าโครง PivotTable แบบคลาสสิก',
    errorValues: 'สำหรับการแสดงค่าความผิดพลาด:',
    emptyCells: 'สำหรับการแสดงเซลล์ว่าง:',
    noDataGroup: 'แสดงรายการที่ไม่มีข้อมูล:',
    noDataRows: 'บนแถว',
    noDataColumns: 'ในคอลัมน์',
    printGroup: 'การพิมพ์:',
    repeatRowLabels: 'ทำซ้ำป้ายชื่อแถวบนแต่ละหน้า',
    printTitles: 'ส่วนหัวของหน้า แถว และคอลัมน์',
    itemLabels: 'ป้ายรายการเมื่อไม่มีเขตข้อมูลในตาราง',
  },
  layout: {
    group: 'เค้าโครง:',
    multipleFilters: 'อนุญาตให้ใช้หลายตัวกรองกับแต่ละเขตข้อมูล',
    autofitColumns: 'ปรับความกว้างของคอลัมน์พอดีอัตโนมัติเมื่ออัปเดต',
    preserveFormatting: 'รักษาการจัดรูปแบบเซลล์เมื่ออัปเดต',
    mergeLabels: 'ผสานและจัดกึ่งกลางเซลล์ที่มีป้ายชื่อ',
    compactIndent: 'เยื้องป้ายชื่อแถวในรูปแบบกะทัดรัด:',
    reportFilterGroup: 'ตัวกรองรายการ:',
    fieldOrder: 'เขตข้อมูลที่แสดงอยู่:',
    downThenOver: 'ลงแล้วซ้ายไปขวา',
    overThenDown: 'ซ้ายไปขวาแล้วลง',
    fieldsPerColumn: 'เขตข้อมูลต่อคอลัมน์:',
    sortGroup: 'เรียงลำดับ:',
    customLists: 'ใช้รายการแบบกำหนดเองเมื่อเรียงลำดับ',
    sortAscending: 'เรียงลำดับจาก ก ถึง ฮ',
    sortDataSource: 'เรียงตามลำดับแหล่งข้อมูล',
  },
  data: {
    group: 'ข้อมูล:',
    saveSourceData: 'บันทึกข้อมูลต้นฉบับพร้อมกับไฟล์',
    enableShowDetails: 'เปิดใช้งานการแสดงรายละเอียด',
    retainDeletedItems: 'คงรายการที่ลบออกจากแหล่งข้อมูลไว้',
    itemsToRetain: 'จำนวนที่จะเก็บรักษาต่อเขตข้อมูล:',
    retainAutomatic: 'อัตโนมัติ',
    retainMax: 'ค่ามากที่สุด',
    refreshOnOpen: 'รีเฟรชข้อมูลเมื่อเปิดไฟล์',
    refreshOnChange: 'รีเฟรชโดยอัตโนมัติเมื่อข้อมูลต้นฉบับเปลี่ยนแปลง',
    connectionGroup: 'การเชื่อมต่อ:',
    savePassword: 'บันทึกรหัสผ่านพร้อมไฟล์',
    backgroundQuery: 'เรียกใช้แบบสอบถามในเบื้องหลัง',
  },
  altText: {
    title: 'ชื่อเรื่อง',
    description: 'คำอธิบาย',
    help: 'ชื่อเรื่องและคำอธิบายเป็นสิ่งที่ใช้แสดงแทนข้อความของข้อมูลที่มีอยู่ในตาราง ไดอะแกรม รูปภาพ และวัตถุอื่นๆ ข้อมูลนี้จะเป็นประโยชน์สำหรับบุคคลที่มีความบกพร่องด้านการมองเห็นหรือการรู้คิด ซึ่งอาจไม่สามารถมองเห็นหรือเข้าใจวัตถุได้\n\nระบบจะอ่านชื่อเรื่องให้แก่บุคคลที่มีความบกพร่อง ซึ่งจะช่วยให้บุคคลเหล่านั้นสามารถพิจารณาได้ว่าต้องการฟังคำอธิบายของเนื้อหาหรือไม่',
  },
  ok: 'ตกลง',
  cancel: 'ยกเลิก',
  unavailable: 'ไม่พร้อมใช้งานในเวอร์ชันนี้',
};
