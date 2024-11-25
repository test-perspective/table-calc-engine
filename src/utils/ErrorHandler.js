export class ErrorHandler {
  // エラーコード定数
  static DIV_ZERO = '#DIV/0!';  // ゼロ除算エラー
  static ERROR = '#ERROR!';      // 一般的なエラー
  static REF = '#REF!';         // 参照エラー
  static NAME = '#NAME?';        // 関数名エラー
  static VALUE = '#VALUE!';      // 値エラー
  static NA = '#N/A';           // データなしエラー
  static NUM = '#NUM!';         // 数値エラー
  static PARSE = '#PARSE!';     // 構文解析エラー

  /**
   * 値がエラーかどうかを判定
   * @param {*} value - 検査する値
   * @returns {boolean} エラーの場合true
   */
  static isError(value) {
    if (typeof value !== 'string') return false;
    return value.startsWith('#') && value.endsWith('!');
  }

  /**
   * エラーメッセージを生成
   * @param {string} code - エラーコード
   * @param {string} [detail] - 詳細メッセージ
   * @returns {string} エラーメッセージ
   */
  static formatError(code, detail = '') {
    return detail ? `${code} (${detail})` : code;
  }

  /**
   * エラーの種類を判定
   * @param {string} error - エラー文字列
   * @returns {string} エラーの種類
   */
  static getErrorType(error) {
    if (!this.isError(error)) return 'NOT_ERROR';
    
    switch (error) {
      case this.DIV_ZERO: return 'DIVISION_BY_ZERO';
      case this.REF: return 'REFERENCE';
      case this.NAME: return 'NAME';
      case this.VALUE: return 'VALUE';
      case this.NA: return 'NOT_AVAILABLE';
      case this.NUM: return 'NUMERIC';
      case this.PARSE: return 'PARSE';
      default: return 'UNKNOWN';
    }
  }

  /**
   * エラーの重大度を判定
   * @param {string} error - エラー文字列
   * @returns {string} 重大度 ('FATAL', 'ERROR', 'WARNING')
   */
  static getErrorSeverity(error) {
    if (!this.isError(error)) return 'NONE';

    switch (error) {
      case this.DIV_ZERO:
      case this.REF:
      case this.PARSE:
        return 'FATAL';
      case this.NAME:
      case this.VALUE:
      case this.NUM:
        return 'ERROR';
      case this.NA:
        return 'WARNING';
      default:
        return 'UNKNOWN';
    }
  }
} 