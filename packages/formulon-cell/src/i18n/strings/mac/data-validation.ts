/** Captions only the Mac Data Validation message tabs show. */
export interface MacDataValidationStrings {
  /** Caption above the input message title and text. */
  inputCaption: string;
  /** Caption above the error alert style, title and text. */
  errorCaption: string;
}

/** Japanese Mac app wording, measured from Excel for Mac. */
export const macDataValidationJa: MacDataValidationStrings = {
  inputCaption: 'セルを選択したときに、この入力時メッセージを表示:',
  errorCaption: '無効なデータが入力されたときに表示するエラー メッセージ:',
};
