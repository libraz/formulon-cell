/** The Mac error-check smart tag menu: rule headings and the actions under them. */
export interface MacErrorMenuStrings {
  /** Heading of the #DIV/0! rule. */
  divZeroHeading: string;
  /** Heading of the number-stored-as-text rule. */
  numberAsTextHeading: string;
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
  convertToNumber: '数値に変換する(C)',
  help: 'このエラーに関するヘルプ(H)',
  trace: 'エラーのトレース(T)',
  ignore: 'エラーを無視する(I)',
  editInFormulaBar: '数式バーで編集(F)',
  options: 'エラー チェック オプション...',
  optionsUnavailable: 'この機能は未実装です。',
};

/** Mac wording for this locale; not measured. */
export const macErrorMenuDe: MacErrorMenuStrings = {
  divZeroHeading: 'Fehler: Division durch Null',
  numberAsTextHeading: 'Zahl als Text gespeichert',
  convertToNumber: 'In Zahl umwandeln',
  help: 'Hilfe zu diesem Fehler',
  trace: 'Fehler überprüfen',
  ignore: 'Fehler ignorieren',
  editInFormulaBar: 'In Bearbeitungsleiste bearbeiten',
  options: 'Optionen für Fehlerüberprüfung...',
  optionsUnavailable: 'Diese Funktion ist noch nicht implementiert.',
};

/** Mac wording for this locale; not measured. */
export const macErrorMenuFr: MacErrorMenuStrings = {
  divZeroHeading: 'Erreur de division par zéro',
  numberAsTextHeading: 'Nombre stocké sous forme de texte',
  convertToNumber: 'Convertir en nombre',
  help: 'Aide sur cette erreur',
  trace: 'Tracer l’erreur',
  ignore: 'Ignorer l’erreur',
  editInFormulaBar: 'Modifier dans la barre de formule',
  options: 'Options de vérification des erreurs...',
  optionsUnavailable: 'Cette fonctionnalité n’est pas encore implémentée.',
};

/** Mac wording for this locale; not measured. */
export const macErrorMenuZh: MacErrorMenuStrings = {
  divZeroHeading: '除以零错误',
  numberAsTextHeading: '以文本形式存储的数字',
  convertToNumber: '转换为数字',
  help: '有关此错误的帮助',
  trace: '追踪错误',
  ignore: '忽略错误',
  editInFormulaBar: '在编辑栏中编辑',
  options: '错误检查选项...',
  optionsUnavailable: '此功能尚未实现。',
};

/** Mac wording for this locale; not measured. */
export const macErrorMenuKo: MacErrorMenuStrings = {
  divZeroHeading: '0으로 나누기 오류',
  numberAsTextHeading: '텍스트로 저장된 숫자',
  convertToNumber: '숫자로 변환',
  help: '이 오류에 대한 도움말',
  trace: '오류 추적',
  ignore: '오류 무시',
  editInFormulaBar: '수식 입력줄에서 편집',
  options: '오류 검사 옵션...',
  optionsUnavailable: '이 기능은 아직 구현되지 않았습니다.',
};

/** Mac wording for this locale; not measured. */
export const macErrorMenuTh: MacErrorMenuStrings = {
  divZeroHeading: 'ข้อผิดพลาดหารด้วยศูนย์',
  numberAsTextHeading: 'ตัวเลขที่จัดเก็บเป็นข้อความ',
  convertToNumber: 'แปลงเป็นตัวเลข',
  help: 'วิธีใช้เกี่ยวกับข้อผิดพลาดนี้',
  trace: 'ติดตามข้อผิดพลาด',
  ignore: 'ละเว้นข้อผิดพลาด',
  editInFormulaBar: 'แก้ไขในแถบสูตร',
  options: 'ตัวเลือกการตรวจสอบข้อผิดพลาด...',
  optionsUnavailable: 'ฟีเจอร์นี้ยังไม่ได้ใช้งาน',
};
