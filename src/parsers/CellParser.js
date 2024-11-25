import { ErrorHandler } from '../utils/ErrorHandler.js';

export class CellParser {
  /**
   * セル参照文字列をパースして行と列のインデックスを返す
   * @param {string} ref - セル参照文字列 (例: "A1", "$B$2")
   * @returns {Object|null} パース結果 ({ row, col, isRowAbsolute, isColAbsolute })
   */
  static parse(ref) {
    try {
      const match = ref.trim().match(/^(\$?)([A-Za-z]+)(\$?)(\d+)$/i);
      if (!match) return null;

      const [, colAbs, colStr, rowAbs, rowStr] = match;
      const col = this.columnNameToIndex(colStr.toUpperCase());
      const row = parseInt(rowStr) - 1;

      if (row < 0 || col < 0) return null;

      return {
        row,
        col,
        isRowAbsolute: rowAbs === '$',
        isColAbsolute: colAbs === '$'
      };
    } catch (error) {
      return null;
    }
  }

  /**
   * 列名をインデックスに変換
   * @param {string} colName - 列名 (例: "A", "BC")
   * @returns {number} 列インデックス
   */
  static columnNameToIndex(colName) {
    let index = 0;
    for (let i = 0; i < colName.length; i++) {
      index = index * 26 + colName.charCodeAt(i) - 'A'.charCodeAt(0);
    }
    return index;
  }

  /**
   * インデックスを列名に変換
   * @param {number} index - 列インデックス
   * @returns {string} 列名
   */
  static indexToColumnName(index) {
    let name = '';
    while (index >= 0) {
      name = String.fromCharCode('A'.charCodeAt(0) + (index % 26)) + name;
      index = Math.floor(index / 26) - 1;
    }
    return name;
  }

  /**
   * セル参照の範囲をパース
   * @param {string} range - 範囲文字列 (例: "A1:B2")
   * @returns {Object|null} パース結果 ({ start: { row, col }, end: { row, col } })
   */
  static parseRange(range) {
    try {
      const [start, end] = range.split(':');
      const startRef = this.parse(start);
      const endRef = this.parse(end);

      if (!startRef || !endRef) return null;

      return {
        start: { row: startRef.row, col: startRef.col },
        end: { row: endRef.row, col: endRef.col }
      };
    } catch (error) {
      return null;
    }
  }

  /**
   * セル参照が有効かどうかを検証
   * @param {Object} ref - パース済みのセル参照
   * @param {Object} tableSize - テーブルのサイズ ({ rows, cols })
   * @returns {boolean} 有効な場合true
   */
  static isValidReference(ref, tableSize) {
    if (!ref) return false;
    return ref.row >= 0 && 
           ref.row < tableSize.rows && 
           ref.col >= 0 && 
           ref.col < tableSize.cols;
  }

  /**
   * セル参照文字列を生成
   * @param {number} row - 行インデックス
   * @param {number} col - 列インデックス
   * @param {boolean} [isRowAbsolute=false] - 行の絶対参照
   * @param {boolean} [isColAbsolute=false] - 列の絶対参照
   * @returns {string} セル参照文字列
   */
  static formatReference(row, col, isRowAbsolute = false, isColAbsolute = false) {
    const colStr = this.indexToColumnName(col);
    const rowStr = row + 1;
    return `${isColAbsolute ? '$' : ''}${colStr}${isRowAbsolute ? '$' : ''}${rowStr}`;
  }
} 