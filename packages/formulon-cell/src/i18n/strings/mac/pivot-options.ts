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
    noDataGroup: 'Items with no data:',
    noDataRows: 'Show on rows',
    noDataColumns: 'Show on columns',
    printGroup: 'Print:',
    repeatRowLabels: 'Repeat row labels on each printed page',
    printTitles: 'Set print titles',
    itemLabels: 'Item labels when no fields are in the values area',
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
    itemsToRetain: 'Number of items to retain per field:',
    retainAutomatic: 'Automatic',
    retainMax: 'Max',
    refreshOnOpen: 'Refresh data when opening the file',
    refreshOnChange: 'Refresh automatically when source data changes',
    connectionGroup: 'Connection:',
    savePassword: 'Save password',
    backgroundQuery: 'Enable background refresh',
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

/** Mac wording for this locale; not measured. */
export const macPivotOptionsDe: MacPivotOptionsStrings = {
  title: 'PivotTable-Optionen',
  name: 'PivotTable-Name:',
  tabDisplay: 'Anzeige',
  tabLayout: 'Layout',
  tabData: 'Daten',
  tabAltText: 'Alternativtext',
  display: {
    group: 'Anzeige:',
    expandButtons: 'Schaltflächen zum Erweitern/Reduzieren',
    printExpandButtons: 'Schaltflächen zum Erweitern/Reduzieren drucken',
    contextualTooltips: 'Kontextbezogene QuickInfos',
    fieldCaptions: 'Feldbeschriftungen und Filter',
    classicLayout: 'Klassisches PivotTable-Layout',
    errorValues: 'Für Fehlerwerte anzeigen:',
    emptyCells: 'Für leere Zellen anzeigen:',
    noDataGroup: 'Elemente ohne Daten:',
    noDataRows: 'In Zeilen anzeigen',
    noDataColumns: 'In Spalten anzeigen',
    printGroup: 'Drucken:',
    repeatRowLabels: 'Zeilenbeschriftungen auf jeder Druckseite wiederholen',
    printTitles: 'Drucktitel festlegen',
    itemLabels: 'Elementbeschriftungen, wenn keine Felder im Wertebereich sind',
  },
  layout: {
    group: 'Layout:',
    multipleFilters: 'Mehrere Filter pro Feld zulassen',
    autofitColumns: 'Spaltenbreiten bei Aktualisierung automatisch anpassen',
    preserveFormatting: 'Zellformatierung bei Aktualisierung beibehalten',
    mergeLabels: 'Zellen mit Beschriftungen zusammenführen und zentrieren',
    compactIndent: 'Zeilenbeschriftungen in Kurzformat einrücken:',
    reportFilterGroup: 'Berichtsfilter:',
    fieldOrder: 'Felder anzeigen:',
    downThenOver: 'Nach unten, dann quer',
    overThenDown: 'Quer, dann nach unten',
    fieldsPerColumn: 'Felder pro Spalte:',
    sortGroup: 'Sortieren:',
    customLists: 'Beim Sortieren benutzerdefinierte Listen verwenden',
    sortAscending: 'Aufsteigend sortieren',
    sortDataSource: 'In Datenquellenreihenfolge sortieren',
  },
  data: {
    group: 'Daten:',
    saveSourceData: 'Quelldaten mit Datei speichern',
    enableShowDetails: 'Details anzeigen aktivieren',
    retainDeletedItems: 'Aus der Datenquelle gelöschte Elemente beibehalten',
    itemsToRetain: 'Anzahl beizubehaltender Elemente pro Feld:',
    retainAutomatic: 'Automatisch',
    retainMax: 'Maximum',
    refreshOnOpen: 'Daten beim Öffnen der Datei aktualisieren',
    refreshOnChange: 'Automatisch aktualisieren, wenn sich die Quelldaten ändern',
    connectionGroup: 'Verbindung:',
    savePassword: 'Kennwort speichern',
    backgroundQuery: 'Hintergrundaktualisierung aktivieren',
  },
  altText: {
    title: 'Titel',
    description: 'Beschreibung',
    help: 'Titel und Beschreibungen sind textbasierte Alternativen zu den Informationen in Tabellen, Diagrammen, Bildern und anderen Objekten. Sie helfen Menschen mit Seh- oder kognitiven Beeinträchtigungen, die das Objekt nicht sehen oder verstehen können.\n\nEin Titel kann einer Person mit Beeinträchtigung vorgelesen werden, damit sie entscheiden kann, ob sie die Beschreibung des Inhalts hören möchte.',
  },
  ok: 'OK',
  cancel: 'Abbrechen',
  unavailable: 'In dieser Version nicht verfügbar.',
};

/** Mac wording for this locale; not measured. */
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
    classicLayout: 'Disposition classique du tableau croisé dynamique',
    errorValues: 'Pour les valeurs d’erreur, afficher :',
    emptyCells: 'Pour les cellules vides, afficher :',
    noDataGroup: 'Éléments sans données :',
    noDataRows: 'Afficher sur les lignes',
    noDataColumns: 'Afficher sur les colonnes',
    printGroup: 'Impression :',
    repeatRowLabels: 'Répéter les étiquettes de lignes sur chaque page imprimée',
    printTitles: 'Définir les titres à imprimer',
    itemLabels: 'Étiquettes des éléments lorsque la zone des valeurs ne contient aucun champ',
  },
  layout: {
    group: 'Disposition :',
    multipleFilters: 'Autoriser plusieurs filtres par champ',
    autofitColumns: 'Ajuster automatiquement la largeur des colonnes lors de la mise à jour',
    preserveFormatting: 'Conserver la mise en forme des cellules lors de la mise à jour',
    mergeLabels: 'Fusionner et centrer les cellules avec étiquettes',
    compactIndent: 'En forme compacte, mettre en retrait les étiquettes de lignes :',
    reportFilterGroup: 'Filtre du rapport :',
    fieldOrder: 'Afficher les champs :',
    downThenOver: 'Vers le bas, puis à droite',
    overThenDown: 'À droite, puis vers le bas',
    fieldsPerColumn: 'Champs par colonne :',
    sortGroup: 'Tri :',
    customLists: 'Utiliser les listes personnalisées lors du tri',
    sortAscending: 'Trier dans l’ordre croissant',
    sortDataSource: 'Trier dans l’ordre de la source de données',
  },
  data: {
    group: 'Données :',
    saveSourceData: 'Enregistrer les données sources avec le fichier',
    enableShowDetails: 'Activer l’affichage des détails',
    retainDeletedItems: 'Conserver les éléments supprimés de la source de données',
    itemsToRetain: 'Nombre d’éléments à conserver par champ :',
    retainAutomatic: 'Automatique',
    retainMax: 'Max',
    refreshOnOpen: 'Actualiser les données lors de l’ouverture du fichier',
    refreshOnChange: 'Actualiser automatiquement quand les données sources changent',
    connectionGroup: 'Connexion :',
    savePassword: 'Enregistrer le mot de passe',
    backgroundQuery: 'Activer l’actualisation en arrière-plan',
  },
  altText: {
    title: 'Titre',
    description: 'Description',
    help: 'Les titres et descriptions fournissent une représentation textuelle des informations contenues dans les tableaux, diagrammes, images et autres objets. Ils sont utiles aux personnes ayant une déficience visuelle ou cognitive qui ne peuvent pas voir ou comprendre l’objet.\n\nLe titre peut être lu à une personne handicapée pour l’aider à décider si elle souhaite entendre la description du contenu.',
  },
  ok: 'OK',
  cancel: 'Annuler',
  unavailable: 'Non disponible dans cette version.',
};

/** Mac wording for this locale; not measured. */
export const macPivotOptionsZh: MacPivotOptionsStrings = {
  title: '数据透视表选项',
  name: '数据透视表名称:',
  tabDisplay: '显示',
  tabLayout: '布局',
  tabData: '数据',
  tabAltText: '可选文字',
  display: {
    group: '显示:',
    expandButtons: '展开/折叠按钮',
    printExpandButtons: '打印展开/折叠按钮',
    contextualTooltips: '上下文工具提示',
    fieldCaptions: '字段标题和筛选',
    classicLayout: '经典数据透视表布局',
    errorValues: '对于错误值，显示:',
    emptyCells: '对于空单元格，显示:',
    noDataGroup: '无数据的项目:',
    noDataRows: '在行中显示',
    noDataColumns: '在列中显示',
    printGroup: '打印:',
    repeatRowLabels: '在每一打印页上重复行标签',
    printTitles: '设置打印标题',
    itemLabels: '值区域中没有字段时的项目标签',
  },
  layout: {
    group: '布局:',
    multipleFilters: '每个字段允许多个筛选',
    autofitColumns: '更新时自动调整列宽',
    preserveFormatting: '更新时保留单元格格式',
    mergeLabels: '合并且居中排列带标签的单元格',
    compactIndent: '压缩表单中缩进行标签:',
    reportFilterGroup: '报表筛选:',
    fieldOrder: '显示字段:',
    downThenOver: '垂直并排',
    overThenDown: '水平并排',
    fieldsPerColumn: '每列字段数:',
    sortGroup: '排序:',
    customLists: '排序时使用自定义列表',
    sortAscending: '升序排序',
    sortDataSource: '按数据源顺序排序',
  },
  data: {
    group: '数据:',
    saveSourceData: '保存文件及源数据',
    enableShowDetails: '启用显示明细数据',
    retainDeletedItems: '保留从数据源删除的项目',
    itemsToRetain: '每个字段保留的项数:',
    retainAutomatic: '自动',
    retainMax: '最大值',
    refreshOnOpen: '打开文件时刷新数据',
    refreshOnChange: '源数据更改时自动刷新',
    connectionGroup: '连接:',
    savePassword: '保存密码',
    backgroundQuery: '启用后台刷新',
  },
  altText: {
    title: '标题',
    description: '说明',
    help: '标题和说明以文本形式表示表格、图表、图像和其他对象中包含的信息。对于无法查看或理解对象的视觉或认知障碍人士，这些信息非常有用。\n\n标题可以读给残障人士听，用于帮助其判断是否要听取内容说明。',
  },
  ok: '确定',
  cancel: '取消',
  unavailable: '此版本中不可用。',
};

/** Mac wording for this locale; not measured. */
export const macPivotOptionsKo: MacPivotOptionsStrings = {
  title: '피벗 테이블 옵션',
  name: '피벗 테이블 이름:',
  tabDisplay: '표시',
  tabLayout: '레이아웃',
  tabData: '데이터',
  tabAltText: '대체 텍스트',
  display: {
    group: '표시:',
    expandButtons: '확장/축소 단추',
    printExpandButtons: '확장/축소 단추 인쇄',
    contextualTooltips: '상황에 맞는 도구 설명',
    fieldCaptions: '필드 캡션 및 필터',
    classicLayout: '클래식 피벗 테이블 레이아웃',
    errorValues: '오류 값 표시:',
    emptyCells: '빈 셀 표시:',
    noDataGroup: '데이터가 없는 항목:',
    noDataRows: '행에 표시',
    noDataColumns: '열에 표시',
    printGroup: '인쇄:',
    repeatRowLabels: '인쇄되는 각 페이지에 행 레이블 반복',
    printTitles: '인쇄 제목 설정',
    itemLabels: '값 영역에 필드가 없을 때 항목 레이블',
  },
  layout: {
    group: '레이아웃:',
    multipleFilters: '필드당 여러 필터 허용',
    autofitColumns: '업데이트 시 열 너비 자동 맞춤',
    preserveFormatting: '업데이트 시 셀 서식 유지',
    mergeLabels: '레이블이 있는 셀 병합 및 가운데 맞춤',
    compactIndent: '압축 형식일 때 행 레이블 들여쓰기:',
    reportFilterGroup: '보고서 필터:',
    fieldOrder: '필드 표시:',
    downThenOver: '아래로, 그 다음 옆으로',
    overThenDown: '옆으로, 그 다음 아래로',
    fieldsPerColumn: '열당 필드 수:',
    sortGroup: '정렬:',
    customLists: '정렬할 때 사용자 지정 목록 사용',
    sortAscending: '오름차순 정렬',
    sortDataSource: '데이터 원본 순서로 정렬',
  },
  data: {
    group: '데이터:',
    saveSourceData: '파일에 원본 데이터 저장',
    enableShowDetails: '세부 정보 표시 사용',
    retainDeletedItems: '데이터 원본에서 삭제된 항목 유지',
    itemsToRetain: '필드당 유지할 항목 수:',
    retainAutomatic: '자동',
    retainMax: '최대',
    refreshOnOpen: '파일을 열 때 데이터 새로 고침',
    refreshOnChange: '원본 데이터가 변경되면 자동으로 새로 고침',
    connectionGroup: '연결:',
    savePassword: '암호 저장',
    backgroundQuery: '백그라운드 새로 고침 사용',
  },
  altText: {
    title: '제목',
    description: '설명',
    help: '제목과 설명은 표, 다이어그램, 이미지 및 기타 개체에 포함된 정보를 텍스트로 표현한 대체 정보입니다. 이 정보는 개체를 보거나 이해하기 어려운 시각 또는 인지 장애가 있는 사용자에게 유용합니다.\n\n제목은 장애가 있는 사용자에게 읽어 줄 수 있으며, 콘텐츠 설명을 들을지 결정하는 데 사용됩니다.',
  },
  ok: '확인',
  cancel: '취소',
  unavailable: '이 버전에서는 사용할 수 없습니다.',
};

/** Mac wording for this locale; not measured. */
export const macPivotOptionsTh: MacPivotOptionsStrings = {
  title: 'ตัวเลือก PivotTable',
  name: 'ชื่อ PivotTable:',
  tabDisplay: 'การแสดง',
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
    errorValues: 'สำหรับค่าข้อผิดพลาด ให้แสดง:',
    emptyCells: 'สำหรับเซลล์ว่าง ให้แสดง:',
    noDataGroup: 'รายการที่ไม่มีข้อมูล:',
    noDataRows: 'แสดงในแถว',
    noDataColumns: 'แสดงในคอลัมน์',
    printGroup: 'การพิมพ์:',
    repeatRowLabels: 'ทำซ้ำป้ายชื่อแถวในแต่ละหน้าที่พิมพ์',
    printTitles: 'ตั้งค่าชื่อเรื่องที่จะพิมพ์',
    itemLabels: 'ป้ายชื่อรายการเมื่อไม่มีเขตข้อมูลในพื้นที่ค่า',
  },
  layout: {
    group: 'เค้าโครง:',
    multipleFilters: 'อนุญาตให้มีหลายตัวกรองต่อเขตข้อมูล',
    autofitColumns: 'ปรับความกว้างคอลัมน์อัตโนมัติเมื่ออัปเดต',
    preserveFormatting: 'คงการจัดรูปแบบเซลล์ไว้เมื่ออัปเดต',
    mergeLabels: 'ผสานและจัดกึ่งกลางเซลล์ที่มีป้ายชื่อ',
    compactIndent: 'เยื้องป้ายชื่อแถวในรูปแบบกะทัดรัด:',
    reportFilterGroup: 'ตัวกรองรายงาน:',
    fieldOrder: 'แสดงเขตข้อมูล:',
    downThenOver: 'ลงล่าง แล้วไปทางขวา',
    overThenDown: 'ไปทางขวา แล้วลงล่าง',
    fieldsPerColumn: 'เขตข้อมูลต่อคอลัมน์:',
    sortGroup: 'การเรียงลำดับ:',
    customLists: 'ใช้รายการแบบกำหนดเองเมื่อเรียงลำดับ',
    sortAscending: 'เรียงจากน้อยไปมาก',
    sortDataSource: 'เรียงตามลำดับแหล่งข้อมูล',
  },
  data: {
    group: 'ข้อมูล:',
    saveSourceData: 'บันทึกข้อมูลต้นฉบับพร้อมไฟล์',
    enableShowDetails: 'เปิดใช้งานการแสดงรายละเอียด',
    retainDeletedItems: 'เก็บรายการที่ถูกลบออกจากแหล่งข้อมูล',
    itemsToRetain: 'จำนวนรายการที่จะเก็บต่อเขตข้อมูล:',
    retainAutomatic: 'อัตโนมัติ',
    retainMax: 'สูงสุด',
    refreshOnOpen: 'รีเฟรชข้อมูลเมื่อเปิดไฟล์',
    refreshOnChange: 'รีเฟรชอัตโนมัติเมื่อข้อมูลต้นฉบับเปลี่ยนแปลง',
    connectionGroup: 'การเชื่อมต่อ:',
    savePassword: 'บันทึกรหัสผ่าน',
    backgroundQuery: 'เปิดใช้งานการรีเฟรชในพื้นหลัง',
  },
  altText: {
    title: 'ชื่อเรื่อง',
    description: 'คำอธิบาย',
    help: 'ชื่อเรื่องและคำอธิบายเป็นข้อความแทนข้อมูลที่อยู่ในตาราง ไดอะแกรม รูปภาพ และวัตถุอื่นๆ ซึ่งเป็นประโยชน์สำหรับผู้ที่มีความบกพร่องทางการมองเห็นหรือการรับรู้ที่อาจมองไม่เห็นหรือไม่เข้าใจวัตถุนั้น\n\nชื่อเรื่องสามารถอ่านให้ผู้พิการฟังได้ และใช้เพื่อตัดสินใจว่าต้องการฟังคำอธิบายเนื้อหาหรือไม่',
  },
  ok: 'ตกลง',
  cancel: 'ยกเลิก',
  unavailable: 'ไม่พร้อมใช้งานในเวอร์ชันนี้',
};
