/** Why a Mac ribbon command or menu item is unavailable. The app has no wording for these. */
export interface MacRibbonReasonStrings {
  /** The command needs a cloud service connection. */
  officeRequired: string;
  /** The command is not implemented yet. */
  notImplemented: string;
  /** The command is not available in the embedded workbook. */
  embeddedUnsupported: string;
  /** The engine cannot edit workbook themes. */
  themeEngineUnavailable: string;
  /** Theme edits are blocked by protected sheets. */
  themeProtected: string;
  /** The workbook theme cannot be read. */
  themeUnreadable: string;
  /** The engine cannot edit threaded comments. */
  threadedCommentsUnavailable: string;
  /** No filter range is active. */
  filterRangeRequired: string;
  /** The selection is outside every outline group. */
  outlineRequired: string;
  /** A slicer needs a table or PivotTable selected. */
  slicerRequired: string;
  /** Symbol insertion needs a cell in edit mode. */
  symbolEditRequired: string;
  /** The workbook has no external links. */
  externalLinksRequired: string;
  /** No temporary sheet view is shown. */
  temporarySheetViewRequired: string;
  /** The Default sheet view is shown. */
  sheetViewRequired: string;
  /** No filter criteria can be reapplied. */
  filterCriteriaRequired: string;
  /** Ink is not available. */
  inkUnavailable: string;
  /** Notice for an unknown ribbon command; `{id}` is its id. */
  unavailableCommand: string;
  /** The calculation engine lacks the function. */
  functionUnavailable: string;
}

export const macRibbonReasonsEn: MacRibbonReasonStrings = {
  officeRequired: 'Requires a cloud service connection.',
  notImplemented: 'This feature is not implemented yet.',
  embeddedUnsupported: 'This feature is not available in the embedded workbook.',
  themeEngineUnavailable: 'This engine cannot edit workbook themes.',
  themeProtected: 'The workbook contains protected sheets.',
  themeUnreadable: 'The workbook theme cannot be read.',
  threadedCommentsUnavailable: 'This engine cannot edit threaded comments.',
  filterRangeRequired: 'There is no active filter range.',
  outlineRequired: 'The selection is not in an outline group.',
  slicerRequired: 'Select a table or PivotTable first.',
  symbolEditRequired: 'Available only while editing a cell.',
  externalLinksRequired: 'This workbook has no external links.',
  temporarySheetViewRequired: 'There is no temporary view to keep.',
  sheetViewRequired: 'The Default view is shown.',
  filterCriteriaRequired: 'There are no filter criteria to reapply.',
  inkUnavailable: 'Ink is unavailable for this spreadsheet.',
  unavailableCommand: 'Unavailable ribbon command: {id}',
  functionUnavailable: 'This function is unavailable in the current calculation engine.',
};

export const macRibbonReasonsJa: MacRibbonReasonStrings = {
  officeRequired: 'クラウドサービスへの接続が必要です。',
  notImplemented: 'この機能は未実装です。',
  embeddedUnsupported: 'この埋め込み環境では未対応です。',
  themeEngineUnavailable: 'このエンジンはブックのテーマを編集できません。',
  themeProtected: '保護されたシートがあります。',
  themeUnreadable: 'ブックのテーマを読み取れません。',
  threadedCommentsUnavailable: 'このエンジンはスレッドコメントを編集できません。',
  filterRangeRequired: '適用中のフィルター範囲はありません。',
  outlineRequired: '選択範囲にアウトライン グループはありません。',
  slicerRequired: 'テーブルまたはピボットテーブルを選択してください。',
  symbolEditRequired: 'セルを編集中にのみ使用できます。',
  externalLinksRequired: 'このブックには外部リンクがありません。',
  temporarySheetViewRequired: '保持できる一時ビューはありません。',
  sheetViewRequired: '既定のビューが表示されています。',
  filterCriteriaRequired: '再適用できるフィルター条件はありません。',
  inkUnavailable: 'このスプレッドシートではインクを使用できません。',
  unavailableCommand: '使用できないリボン コマンド: {id}',
  functionUnavailable: 'この関数は現在の計算エンジンでは使用できません。',
};

export const macRibbonReasonsDe: MacRibbonReasonStrings = {
  officeRequired: 'Eine Verbindung mit einem Clouddienst ist erforderlich.',
  notImplemented: 'Diese Funktion ist noch nicht implementiert.',
  embeddedUnsupported: 'Diese Funktion ist in der eingebetteten Arbeitsmappe nicht verfügbar.',
  themeEngineUnavailable: 'Diese Engine kann keine Arbeitsmappendesigns bearbeiten.',
  themeProtected: 'Die Arbeitsmappe enthält geschützte Blätter.',
  themeUnreadable: 'Das Arbeitsmappendesign kann nicht gelesen werden.',
  threadedCommentsUnavailable: 'Diese Engine kann Threadkommentare nicht bearbeiten.',
  filterRangeRequired: 'Es gibt keinen Filter zum Löschen.',
  outlineRequired: 'Die Auswahl befindet sich in keiner Gliederungsgruppe.',
  slicerRequired: 'Wählen Sie zuerst eine Tabelle oder PivotTable aus.',
  symbolEditRequired: 'Nur beim Bearbeiten einer Zelle verfügbar.',
  externalLinksRequired: 'Diese Arbeitsmappe enthält keine externen Verknüpfungen.',
  temporarySheetViewRequired: 'Es gibt keine temporäre Ansicht zum Beibehalten.',
  sheetViewRequired: 'Die Standardansicht wird angezeigt.',
  filterCriteriaRequired: 'Es gibt keine Filterkriterien zum erneuten Anwenden.',
  inkUnavailable: 'Freihand ist für diese Tabelle nicht verfügbar.',
  unavailableCommand: 'Nicht verfügbarer Menübandbefehl: {id}',
  functionUnavailable: 'Diese Funktion ist in der aktuellen Berechnungsengine nicht verfügbar.',
};

export const macRibbonReasonsFr: MacRibbonReasonStrings = {
  officeRequired: 'Une connexion à un service cloud est requise.',
  notImplemented: 'Cette fonctionnalité n’est pas encore implémentée.',
  embeddedUnsupported: 'Cette fonctionnalité n’est pas disponible dans le classeur incorporé.',
  themeEngineUnavailable: 'Ce moteur ne peut pas modifier les thèmes du classeur.',
  themeProtected: 'Le classeur contient des feuilles protégées.',
  themeUnreadable: 'Le thème du classeur ne peut pas être lu.',
  threadedCommentsUnavailable: 'Ce moteur ne peut pas modifier les commentaires avec conversation.',
  filterRangeRequired: 'Il n’y a aucun filtre à effacer.',
  outlineRequired: 'La sélection ne se trouve dans aucun groupe du plan.',
  slicerRequired: 'Sélectionnez d’abord un tableau ou un tableau croisé dynamique.',
  symbolEditRequired: 'Disponible uniquement pendant la modification d’une cellule.',
  externalLinksRequired: 'Ce classeur ne contient aucun lien externe.',
  temporarySheetViewRequired: 'Aucun affichage temporaire à conserver.',
  sheetViewRequired: 'L’affichage par défaut est affiché.',
  filterCriteriaRequired: 'Aucun critère de filtre à réappliquer.',
  inkUnavailable: 'L’entrée manuscrite n’est pas disponible pour cette feuille de calcul.',
  unavailableCommand: 'Commande du ruban indisponible : {id}',
  functionUnavailable: 'Cette fonction n’est pas disponible dans le moteur de calcul actuel.',
};

export const macRibbonReasonsZh: MacRibbonReasonStrings = {
  officeRequired: '需要连接云服务。',
  notImplemented: '此功能尚未实现。',
  embeddedUnsupported: '嵌入式工作簿不支持此功能。',
  themeEngineUnavailable: '此引擎无法编辑工作簿主题。',
  themeProtected: '工作簿包含受保护的工作表。',
  themeUnreadable: '无法读取工作簿主题。',
  threadedCommentsUnavailable: '此引擎无法编辑线程批注。',
  filterRangeRequired: '没有可清除的筛选。',
  outlineRequired: '所选内容不在任何分级显示组中。',
  slicerRequired: '请先选择表格或数据透视表。',
  symbolEditRequired: '仅在编辑单元格时可用。',
  externalLinksRequired: '此工作簿没有外部链接。',
  temporarySheetViewRequired: '没有可保留的临时视图。',
  sheetViewRequired: '正在显示默认视图。',
  filterCriteriaRequired: '没有可重新应用的筛选条件。',
  inkUnavailable: '此电子表格无法使用墨迹。',
  unavailableCommand: '不可用的功能区命令：{id}',
  functionUnavailable: '此函数在当前计算引擎中不可用。',
};

export const macRibbonReasonsKo: MacRibbonReasonStrings = {
  officeRequired: '클라우드 서비스에 연결해야 합니다.',
  notImplemented: '이 기능은 아직 구현되지 않았습니다.',
  embeddedUnsupported: '포함된 통합 문서에서는 이 기능을 사용할 수 없습니다.',
  themeEngineUnavailable: '이 엔진에서는 통합 문서 테마를 편집할 수 없습니다.',
  themeProtected: '통합 문서에 보호된 시트가 있습니다.',
  themeUnreadable: '통합 문서 테마를 읽을 수 없습니다.',
  threadedCommentsUnavailable: '이 엔진에서는 스레드 댓글을 편집할 수 없습니다.',
  filterRangeRequired: '지울 필터가 없습니다.',
  outlineRequired: '선택 영역에 윤곽 그룹이 없습니다.',
  slicerRequired: '먼저 표 또는 피벗 테이블을 선택하세요.',
  symbolEditRequired: '셀을 편집하는 동안에만 사용할 수 있습니다.',
  externalLinksRequired: '이 통합 문서에는 외부 링크가 없습니다.',
  temporarySheetViewRequired: '유지할 임시 보기가 없습니다.',
  sheetViewRequired: '기본 보기가 표시되어 있습니다.',
  filterCriteriaRequired: '다시 적용할 필터 조건이 없습니다.',
  inkUnavailable: '이 스프레드시트에서는 잉크를 사용할 수 없습니다.',
  unavailableCommand: '사용할 수 없는 리본 명령: {id}',
  functionUnavailable: '현재 계산 엔진에서는 이 함수를 사용할 수 없습니다.',
};

export const macRibbonReasonsTh: MacRibbonReasonStrings = {
  officeRequired: 'ต้องเชื่อมต่อบริการคลาวด์',
  notImplemented: 'ฟีเจอร์นี้ยังไม่ได้ใช้งาน',
  embeddedUnsupported: 'ฟีเจอร์นี้ไม่พร้อมใช้งานในเวิร์กบุ๊กแบบฝังตัว',
  themeEngineUnavailable: 'กลไกนี้ไม่สามารถแก้ไขธีมเวิร์กบุ๊กได้',
  themeProtected: 'เวิร์กบุ๊กมีแผ่นงานที่ได้รับการป้องกัน',
  themeUnreadable: 'ไม่สามารถอ่านธีมเวิร์กบุ๊กได้',
  threadedCommentsUnavailable: 'กลไกนี้ไม่สามารถแก้ไขข้อคิดเห็นแบบเธรดได้',
  filterRangeRequired: 'ไม่มีตัวกรองให้ล้าง',
  outlineRequired: 'การเลือกไม่ได้อยู่ในกลุ่มเค้าร่าง',
  slicerRequired: 'เลือกตารางหรือพีวอตเทเบิลก่อน',
  symbolEditRequired: 'ใช้ได้เฉพาะขณะแก้ไขเซลล์',
  externalLinksRequired: 'สมุดงานนี้ไม่มีการเชื่อมโยงภายนอก',
  temporarySheetViewRequired: 'ไม่มีมุมมองชั่วคราวให้เก็บไว้',
  sheetViewRequired: 'กำลังแสดงมุมมองเริ่มต้น',
  filterCriteriaRequired: 'ไม่มีเกณฑ์ตัวกรองให้นำไปใช้ใหม่',
  inkUnavailable: 'หมึกไม่พร้อมใช้งานสำหรับสเปรดชีตนี้',
  unavailableCommand: 'คำสั่ง Ribbon ไม่พร้อมใช้งาน: {id}',
  functionUnavailable: 'ฟังก์ชันนี้ไม่พร้อมใช้งานในกลไกการคำนวณปัจจุบัน',
};
