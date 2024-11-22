import { FormulaParser } from './parser/FormulaParser.js';
import { ExcelFunctions } from './functions/ExcelFunctions.js';
import { CellReference } from '../utils/CellReference.js';
import { FormulaError } from './types/index.js';

export class FormulaEngine {
  constructor() {
    this.parser = new FormulaParser();
    this.functions = {
      'SUM': this.sumFunction.bind(this)
    };
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

      const ast = this.parser.parse(formula);
      return this.evaluateAst(ast, allTables, currentTableIndex);
    } catch (error) {
      console.error('Formula evaluation error:', error);
      return '#ERROR!';
    }
  }

  evaluateAst(ast, allTables, currentTableIndex) {
    if (!ast) return null;

    switch (ast.type) {
      case 'function':
        if (!(ast.name in this.functions)) {
          throw new Error(`Unknown function: ${ast.name}`);
        }
        return this.functions[ast.name](ast.arguments, allTables, currentTableIndex);

      case 'cell':
        return this.getCellValue(ast.reference, allTables, currentTableIndex);

      case 'range':
        return this.getRangeValues(ast.reference, allTables, currentTableIndex);

      case 'literal':
        return ast.value;

      default:
        throw new Error(`Unknown AST type: ${ast.type}`);
    }
  }

  sumFunction(args, allTables, currentTableIndex) {
    try {
      const values = args.map(arg => {
        if (arg.type === 'range') {
          return this.getRangeValues(arg.reference, allTables, currentTableIndex);
        } else if (arg.type === 'cell') {
          return [this.getCellValue(arg.reference, allTables, currentTableIndex)];
        } else {
          return [this.evaluateAst(arg, allTables, currentTableIndex)];
        }
      }).flat();

      return values.reduce((sum, value) => {
        const num = Number(value);
        return isNaN(num) ? sum : sum + num;
      }, 0);
    } catch (error) {
      console.error('Error in SUM function:', error);
      return '#ERROR!';
    }
  }

  getRangeValues(range, allTables, currentTableIndex) {
    try {
      const { start, end } = CellReference.parseRange(range);
      const table = allTables[currentTableIndex];
      
      const values = [];
      for (let row = start.row; row <= end.row; row++) {
        for (let col = start.column; col <= end.column; col++) {
          if (table?.[row]?.[col]) {
            const cell = table[row][col];
            const value = typeof cell.value === 'string' && cell.value.startsWith('=')
              ? cell.resolvedValue
              : cell.value;
            values.push(value);
          }
        }
      }
      return values;
    } catch (error) {
      console.error('Error in getRangeValues:', error);
      throw error;
    }
  }

  getCellValue(reference, allTables, currentTableIndex) {
    try {
      const { row, column } = CellReference.parse(reference);
      const table = allTables[currentTableIndex];
      
      if (!table?.[row]?.[column]) {
        return '#REF!';
      }

      const cell = table[row][column];
      return typeof cell.value === 'string' && cell.value.startsWith('=')
        ? cell.resolvedValue
        : cell.value;
    } catch (error) {
      console.error('Error in getCellValue:', error);
      return '#REF!';
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
        // Excelの仕様に従い、先頭のスペスパディングは行わない
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