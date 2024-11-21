import { FormulaParser } from './parser/FormulaParser.js';
import { ExcelFunctions } from './functions/ExcelFunctions.js';
import { CellReference } from './utils/CellReference.js';
import { FormulaError } from './types/index.js';

export class FormulaEngine {
  constructor() {
    this.parser = new FormulaParser();
    this.cache = new Map();
  }

  processData(data) {
    try {
      console.log('Processing formulas:', data.formulas);
      
      // テーブルデータの正規化
      const normalizedTableData = data.tableData.map(row =>
        row.map(cell => ({
          ...cell,
          value: !isNaN(Number(cell.value)) ? Number(cell.value) : cell.value
        }))
      );

      // formulasの処理
      const processedFormulas = data.formulas.map(formula => {
        console.log('\nProcessing formula:', formula.value);
        
        if (!formula.resolved) {
          const result = this.evaluateFormula(formula.value, normalizedTableData);
          console.log('Formula result:', result);
          
          return {
            ...formula,
            value: result,
            displayValue: this.formatValue(result, formula.excelFormat),
            resolved: true
          };
        }
        return formula;
      });

      return {
        tableData: normalizedTableData,
        formulas: processedFormulas
      };
    } catch (error) {
      console.error('Processing error:', error);
      return this.handleError(error);
    }
  }

  processTableData(tableData) {
    return tableData.map(row =>
      row.map(cell => {
        if (!cell.resolved && cell.value.startsWith('=')) {
          const result = this.evaluateFormula(cell.value, tableData);
          return {
            ...cell,
            value: this.formatValue(result, cell.excelFormat),
            resolved: true
          };
        }
        return cell;
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

  evaluateFormula(formula, tableData) {
    try {
      // '=' で始まらない場合は直接値を返す
      if (!formula.startsWith('=')) {
        return formula;
      }

      const formulaWithoutEquals = formula.substring(1); // '=' を削除
      const ast = this.parser.parse(formulaWithoutEquals);
      console.log('Formula:', formula);
      console.log('AST:', JSON.stringify(ast, null, 2));
      
      if (ast.type === 'function') {
        const values = ast.arguments.map(arg => {
          if (arg.type === 'range') {
            return this.getRangeValues(arg.reference, tableData);
          }
          if (arg.type === 'cell') {
            return [this.getCellValue(arg.reference, tableData)];
          }
          return [arg.value];
        }).flat();

        console.log('Values for', ast.name, ':', values);
        
        const funcName = ast.name.toUpperCase();
        if (!(funcName in ExcelFunctions)) {
          throw new Error(`Unknown function: ${funcName}`);
        }
        
        const result = ExcelFunctions[funcName](values);
        console.log('Function result:', result);
        
        return result;
      }
      
      return '#ERROR!';
    } catch (error) {
      console.error('Formula evaluation error:', error);
      return '#ERROR!';
    }
  }

  getRangeValues(range, tableData) {
    try {
      const { start, end } = CellReference.parseRange(range);
      const values = [];
      
      console.log('Getting range values for', range);
      console.log('Start:', start, 'End:', end);
      
      for (let row = start.row; row <= end.row; row++) {
        for (let col = start.column; col <= end.column; col++) {
          if (!tableData[row] || !tableData[row][col]) {
            throw new Error(`Invalid range reference: ${range}`);
          }
          const value = Number(tableData[row][col].value);
          console.log(`Cell [${row},${col}] value:`, value);
          values.push(value);
        }
      }
      
      console.log('Collected values:', values);
      return values;
    } catch (error) {
      console.error('Error in getRangeValues:', error);
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