/** The Mac error-check smart tag menu: rule headings and the actions under them. */
export interface MacErrorMenuStrings {
  /** Heading of the #DIV/0! rule. */
  divZeroHeading: string;
  /** Heading of the number-stored-as-text rule. */
  numberAsTextHeading: string;
  /** Heading of a stale formula value. */
  staleHeading: string;
  /** Hover tip body of a stale formula value. */
  staleBody: string;
  /** The stale rule's actions. */
  calculateNow: string;
  switchToAutomatic: string;
  disableFormatStale: string;
  convertToNumber: string;
  help: string;
  trace: string;
  ignore: string;
  editInFormulaBar: string;
  options: string;
  /** Why the options entry cannot be used. */
  optionsUnavailable: string;
}

/** Mac wording for this locale; not measured. */
export const macErrorMenuEn: MacErrorMenuStrings = {
  divZeroHeading: 'Divide by Zero Error',
  numberAsTextHeading: 'Number Stored as Text',
  staleHeading: 'Stale Value',
  staleBody: 'The value in this cell is stale. Calculate the workbook to see an up-to-date value.',
  calculateNow: 'Calculate Now (F9)',
  switchToAutomatic: 'Switch to Automatic Calculation',
  disableFormatStale: 'Disable Format Stale Cells',
  convertToNumber: 'Convert to Number',
  help: 'Help on This Error',
  trace: 'Trace Error',
  ignore: 'Ignore Error',
  editInFormulaBar: 'Edit in Formula Bar',
  options: 'Error Checking Options...',
  optionsUnavailable: 'This feature is not implemented yet.',
};

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macErrorMenuJa: MacErrorMenuStrings = {
  divZeroHeading: '0 除算のエラー',
  numberAsTextHeading: '数値が文字列として保存されています',
  staleHeading: '古くなった値',
  staleBody: 'このセルの値は古くなっています。ブックを計算して、最新の値を表示します。',
  calculateNow: '今すぐ計算(C) (F9)',
  switchToAutomatic: '自動計算に切り替える(A)',
  disableFormatStale: '古くなったセルのフォーマットを無効にする(D)',
  convertToNumber: '数値に変換する(C)',
  help: 'このエラーに関するヘルプ(H)',
  trace: 'エラーのトレース(T)',
  ignore: 'エラーを無視する(I)',
  editInFormulaBar: '数式バーで編集(F)',
  options: 'エラー チェック オプション...',
  optionsUnavailable: 'この機能は未実装です。',
};

/** Mac app wording from the Excel bundle; not answered by it: convertToNumber, help, trace, ignore, editInFormulaBar, optionsUnavailable. */
export const macErrorMenuDe: MacErrorMenuStrings = {
  divZeroHeading: 'Division durch Null',
  numberAsTextHeading: 'Als Text gespeicherte Zahlen',
  staleHeading: 'Veralteter Wert',
  staleBody:
    'Der Wert in dieser Zelle ist veraltet. Berechnen Sie die Arbeitsmappe, um einen aktuellen Wert anzuzeigen.',
  calculateNow: 'Jetzt berechnen (F9)',
  switchToAutomatic: 'Zur automatischen Berechnung wechseln',
  disableFormatStale: 'Veraltete Zellenformatierung deaktivieren',
  convertToNumber: 'In Zahl umwandeln',
  help: 'Hilfe zu diesem Fehler',
  trace: 'Fehler überprüfen',
  ignore: 'Fehler ignorieren',
  editInFormulaBar: 'In Bearbeitungsleiste bearbeiten',
  options: 'Optionen zur Fehlerüberprüfung...',
  optionsUnavailable: 'Diese Funktion ist noch nicht implementiert.',
};

/** Mac app wording from the Excel bundle; not answered by it: convertToNumber, help, trace, ignore, editInFormulaBar, optionsUnavailable. */
export const macErrorMenuFr: MacErrorMenuStrings = {
  divZeroHeading: 'Erreur lors de la division par zéro',
  numberAsTextHeading: 'Nombre stocké sous forme de texte',
  staleHeading: 'Valeur périmée',
  staleBody:
    'La valeur de cette cellule est périmée. Calculez le classeur pour voir une valeur à jour.',
  calculateNow: 'Calculer maintenant (F9)',
  switchToAutomatic: 'Basculer vers le calcul automatique',
  disableFormatStale: 'Désactiver le format des cellules périmées',
  convertToNumber: 'Convertir en nombre',
  help: 'Aide sur cette erreur',
  trace: 'Tracer l’erreur',
  ignore: 'Ignorer l’erreur',
  editInFormulaBar: 'Modifier dans la barre de formule',
  options: 'Options de vérification des erreurs...',
  optionsUnavailable: 'Cette fonctionnalité n’est pas encore implémentée.',
};

/** Mac app wording from the Excel bundle; not answered by it: divZeroHeading, convertToNumber, help, trace, ignore, editInFormulaBar, and 1 more. */
export const macErrorMenuZh: MacErrorMenuStrings = {
  divZeroHeading: '除以零错误',
  numberAsTextHeading: '以文本形式存储的数字',
  staleHeading: '过时值',
  staleBody: '此单元格中的值已过时。请计算工作簿以查看最新值。',
  calculateNow: '立即计算(F9)',
  switchToAutomatic: '切换到自动计算',
  disableFormatStale: '禁用过时单元格格式',
  convertToNumber: '转换为数字',
  help: '有关此错误的帮助',
  trace: '追踪错误',
  ignore: '忽略错误',
  editInFormulaBar: '在编辑栏中编辑',
  options: '错误检查选项...',
  optionsUnavailable: '此功能尚未实现。',
};

/** Mac app wording from the Excel bundle; not answered by it: convertToNumber, help, trace, ignore, editInFormulaBar, optionsUnavailable. */
export const macErrorMenuKo: MacErrorMenuStrings = {
  divZeroHeading: '0으로 나누기 오류',
  numberAsTextHeading: '텍스트 형식으로 저장된 숫자',
  staleHeading: '부실 값',
  staleBody: '이 셀의 값이 부실합니다. 통합 문서를 계산하여 최신 값을 확인하세요.',
  calculateNow: '지금 계산(F9)',
  switchToAutomatic: '자동 계산으로 전환',
  disableFormatStale: '부실 셀 서식 사용 안 함',
  convertToNumber: '숫자로 변환',
  help: '이 오류에 대한 도움말',
  trace: '오류 추적',
  ignore: '오류 무시',
  editInFormulaBar: '수식 입력줄에서 편집',
  options: '오류 검사 옵션...',
  optionsUnavailable: '이 기능은 아직 구현되지 않았습니다.',
};

/** Mac app wording from the Excel bundle; not answered by it: convertToNumber, help, trace, ignore, editInFormulaBar, optionsUnavailable. */
export const macErrorMenuTh: MacErrorMenuStrings = {
  divZeroHeading: 'ข้อผิดพลาดจากการหารด้วยศูนย์',
  numberAsTextHeading: 'ตัวเลขที่ถูกเก็บในรูปข้อความ',
  staleHeading: 'ค่าเก่า',
  staleBody: 'ค่าในเซลล์นี้เก่า คํานวณเวิร์กบุ๊กเพื่อดูค่าล่าสุด',
  calculateNow: 'คำนวณทันที (F9)',
  switchToAutomatic: 'สลับไปยังการคำนวณอัตโนมัติ',
  disableFormatStale: 'ปิดใช้งานการจัดรูปแบบเซลล์เก่า',
  convertToNumber: 'แปลงเป็นตัวเลข',
  help: 'วิธีใช้เกี่ยวกับข้อผิดพลาดนี้',
  trace: 'ติดตามข้อผิดพลาด',
  ignore: 'ละเว้นข้อผิดพลาด',
  editInFormulaBar: 'แก้ไขในแถบสูตร',
  options: 'ตัวเลือกการตรวจสอบข้อผิดพลาด...',
  optionsUnavailable: 'ฟีเจอร์นี้ยังไม่ได้ใช้งาน',
};
