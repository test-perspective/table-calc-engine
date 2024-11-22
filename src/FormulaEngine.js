import { FormulaParser } from './parser/FormulaParser.js';
import { ExcelFunctions } from './functions/ExcelFunctions.js';
import { CellReference } from '../utils/CellReference.js';
import { FormulaError } from './types/index.js';

export class FormulaEngine {
  constructor() {
    this.parser = new FormulaParser();
    this.cache = new Map();
    this.functions = ExcelFunctions;
  }

  processData(tables) {
    if (!Array.isArray(tables)) {
      throw new Error('Input must be an array of tables');
    }

    try {
      // すべてのテーブルを正規化
      const normalizedTables = tables.map(table =>
        table.map(row =>
          row.map(cell => ({
            ...cell,
            value: !isNaN(Number(cell.value)) ? Number(cell.value) : cell.value
          }))
        )
      );

      // 各テーブルの数式を処理
      const processedTables = normalizedTables.map((table, tableIndex) =>
        this.processTableData(table, normalizedTables, tableIndex)
      );

      // 全テーブルから数式を抽出
      const extractedFormulas = this.extractFormulas(processedTables);

      return {
        tables: processedTables,
        formulas: extractedFormulas
      };
    } catch (error) {
      console.error('Processing error:', error);
      // handleErrorの代わりにエラーオブジェクトを返す
      return {
        tables: [],
        formulas: [],
        error: error.message
      };
    }
  }

  processTableData(table, allTables, currentTableIndex) {
    return table.map(row =>
      row.map(cell => {
        // 値が文字列で、かつ'='で始まる場合のみ数式として処理
        if (cell.value && typeof cell.value === 'string' && cell.value.startsWith('=')) {
          const result = this.evaluateFormula(cell.value, allTables, currentTableIndex);
          const formattedResult = this.formatValue(result, cell.excelFormat);
          return {
            ...cell,
            value: cell.value,
            resolvedValue: result,
            displayValue: formattedResult,
            resolved: true
          };
        }
        // 数値の場合は数値のまま返す
        if (!isNaN(Number(cell.value))) {
          return {
            ...cell,
            value: Number(cell.value),
            resolvedValue: Number(cell.value),
            displayValue: this.formatValue(Number(cell.value), cell.excelFormat),
            resolved: true
          };
        }
        // その他の場合は元の値をそのまま返す
        return {
          ...cell,
          resolvedValue: cell.value,
          displayValue: cell.value,
          resolved: true
        };
      })
    );
  }

  processFormulas(formulas, tableData) {
    return formulas.map(formula => {
      if (!formula.resolved) {
        const result = this.evaluateFormula(formula.value, tableData);
        return {
          ...formula,
          value: this.formatValue(result, formula.excelFormat),
          resolved: true
        };
      }
      return formula;
    });
  }

  evaluateFormula(formula, allTables, currentTableIndex) {
    try {
      if (!formula.startsWith('=')) {
        return formula;
      }

      const formulaWithoutEquals = formula.substring(1);
      const ast = this.parser.parse(formulaWithoutEquals);

      if (ast.type === 'function') {
        const values = ast.arguments.map(arg => {
          if (arg.type === 'range') {
            return this.getRangeValues(arg.reference, allTables, currentTableIndex);
          }
          if (arg.type === 'cell') {
            return [this.getCellValue(arg.reference, allTables, currentTableIndex)];
          }
          return [arg.value];
        }).flat();

        const funcName = ast.name.toUpperCase();
        if (!(funcName in this.functions)) {
          throw new Error(`Unknown function: ${funcName}`);
        }

        return this.functions[funcName](values);
      } else if (ast.type === 'literal' && typeof ast.value === 'string' && ast.value.includes('!')) {
        // テーブル参照の処理
        const [tableRef, cellRef] = this.parseTableReference(ast.value);
        if (tableRef !== null) {
          return this.getCellValue(cellRef, allTables, tableRef);
        }
      }

      return ast.value ?? '#ERROR!';
    } catch (error) {
      console.error('Formula evaluation error:', error);
      return '#ERROR!';
    }
  }

  getRangeValues(reference, allTables, currentTableIndex) {
    try {
      const [tableRef, rangeRef] = this.parseTableReference(reference);
      const targetTableIndex = tableRef !== null ? tableRef : currentTableIndex;
      const targetTable = allTables[targetTableIndex];

      if (!targetTable) {
        throw new Error(`Invalid table reference: ${tableRef}`);
      }

      const { start, end } = CellReference.parseRange(rangeRef);
      const values = [];

      for (let row = start.row; row <= end.row; row++) {
        for (let col = start.column; col <= end.column; col++) {
          if (!targetTable[row] || !targetTable[row][col]) {
            throw new Error(`Invalid range reference: ${rangeRef}`);
          }
          const cellValue = targetTable[row][col].resolvedValue ?? targetTable[row][col].value;
          const value = !isNaN(Number(cellValue)) ? Number(cellValue) : cellValue;
          values.push(value);
        }
      }

      return values;
    } catch (error) {
      console.error('Error in getRangeValues:', error);
      throw error;
    }
  }

  parseTableReference(reference) {
    if (typeof reference !== 'string') {
      return [null, reference];
    }
    const parts = reference.split('!');
    if (parts.length === 2) {
      return [parseInt(parts[0]), parts[1]];
    }
    return [null, reference];
  }

  extractFormulas(tables) {
    const formulas = [];
    tables.forEach((table, tableIndex) => {
      table.forEach((row, rowIndex) => {
        row.forEach((cell, colIndex) => {
          if (typeof cell.value === 'string' && cell.value.startsWith('=')) {
            formulas.push({
              value: cell.value,
              resolvedValue: cell.resolvedValue,
              displayValue: cell.displayValue,
              tableIndex,
              position: `${CellReference.columnToLetter(colIndex)}${rowIndex + 1}`,
              resolved: true,
              excelFormat: cell.excelFormat,
              macroId: cell.macroId
            });
          }
        });
      });
    });
    return formulas;
  }

  getCellValue(reference, allTables, tableIndex) {
    try {
      const [targetTableIndex, cellRef] = this.parseTableReference(reference);
      const targetTable = allTables[targetTableIndex ?? tableIndex];
      
      if (!targetTable) {
        throw new Error(`Invalid table reference: ${targetTableIndex}`);
      }

      const { row, column } = CellReference.parse(cellRef);
      if (!targetTable[row] || !targetTable[row][column]) {
        throw new Error(`Invalid cell reference: ${cellRef}`);
      }

      const cell = targetTable[row][column];
      return cell.resolvedValue ?? cell.value;
    } catch (error) {
      console.error('Error in getCellValue:', error);
      throw error;
    }
  }

  formatValue(value, format) {
    if (!format || value == null || typeof value === 'string') {
      return value;
    }

    try {
      if (typeof value === 'number') {
        // フォーマットパターンを解析
        const parts = format.split('.');
        const hasDecimal = parts.length > 1;
        
        // 小数点以下の桁数を取得
        const decimals = hasDecimal ? parts[1].length : 0;
        
        // 数値を指定された桁数でフォーマット
        // Excelの仕様に従い、先頭のスペースパディングは行わない
        return value.toFixed(decimals);
      }

      if (value instanceof Date) {
        return value.toISOString().split('T')[0];
      }

      return value;
    } catch (error) {
      console.error('Format error:', error);
      return '#FORMAT_ERROR';
    }
  }

  getDecimalPlaces(format) {
    const match = format.match(/\.(\d+)/);
    return match ? match[1].length : 0;
  }
} 