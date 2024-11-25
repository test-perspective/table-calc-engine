import { ErrorHandler } from '../utils/ErrorHandler.js';
import { CellParser } from '../parsers/CellParser.js';

export class BasicEvaluator {
  /**
   * 数式を評価
   * @param {string} expression - 評価する数式
   * @param {Array} table - テーブルデータ
   * @returns {number|string} 計算結果またはエラー
   */
  static evaluate(expression, table) {
    try {
      // ゼロ除算のチェック
      if (expression.includes('/0')) {
        return ErrorHandler.DIV_ZERO;
      }

      // 数式を部分式に分割
      const parts = this._splitExpression(expression);
      
      // 各部分式を評価
      const evaluatedParts = this._evaluateParts(parts, table);
      
      // エラーチェック
      if (evaluatedParts.some(part => ErrorHandler.isError(part))) {
        return evaluatedParts.find(part => ErrorHandler.isError(part));
      }

      // 最終的な計算を実行
      return this._calculateResult(evaluatedParts);
    } catch (error) {
      return ErrorHandler.ERROR;
    }
  }

  /**
   * 数式を演算子で分割
   * @private
   * @param {string} expression - 数式
   * @returns {Array} 分割された部分式
   */
  static _splitExpression(expression) {
    return expression.split(/([+\-*\/])/).map(part => part.trim());
  }

  /**
   * 各部分式を評価
   * @private
   * @param {Array} parts - 部分式の配列
   * @param {Array} table - テーブルデータ
   * @returns {Array} 評価された値の配列
   */
  static _evaluateParts(parts, table) {
    return parts.map(part => {
      // 演算子はそのまま返す
      if (['+', '-', '*', '/'].includes(part)) {
        return part;
      }

      // セル参照の場合
      const cellRef = CellParser.parse(part);
      if (cellRef) {
        return this._evaluateCellReference(cellRef, table);
      }

      // 数値の場合
      const numValue = Number(part);
      if (isNaN(numValue)) {
        return ErrorHandler.VALUE;
      }
      return numValue;
    });
  }

  /**
   * セル参照を評価
   * @private
   * @param {Object} cellRef - セル参照情報
   * @param {Array} table - テーブルデータ
   * @returns {number|string} セルの値またはエラー
   */
  static _evaluateCellReference(cellRef, table) {
    if (!table[cellRef.row] || !table[cellRef.row][cellRef.col]) {
      return ErrorHandler.REF;
    }
    
    const value = table[cellRef.row][cellRef.col].value;
    
    // セルの値が数値でない場合
    if (typeof value !== 'number' && !this._isNumericString(value)) {
      return ErrorHandler.VALUE;
    }
    
    return Number(value);
  }

  /**
   * 文字列が数値として解釈可能か判定
   * @private
   * @param {string} value - 検査する値
   * @returns {boolean} 数値として解釈可能な場合true
   */
  static _isNumericString(value) {
    return !isNaN(value) && !isNaN(parseFloat(value));
  }

  /**
   * 最終的な計算を実行
   * @private
   * @param {Array} parts - 評価済みの部分式配列
   * @returns {number|string} 計算結果またはエラー
   */
  static _calculateResult(parts) {
    let result = Number(parts[0]);
    
    for (let i = 1; i < parts.length; i += 2) {
      const operator = parts[i];
      const operand = Number(parts[i + 1]);

      switch (operator) {
        case '+':
          result += operand;
          break;
        case '-':
          result -= operand;
          break;
        case '*':
          result *= operand;
          break;
        case '/':
          if (operand === 0) return ErrorHandler.DIV_ZERO;
          result /= operand;
          break;
        default:
          return ErrorHandler.ERROR;
      }
    }

    return result;
  }
} 