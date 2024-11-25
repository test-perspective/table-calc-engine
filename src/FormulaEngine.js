import { FormulaParser } from './parsers/FormulaParser.js';
import { ExcelFunctions } from './functions/ExcelFunctions.js';
import { ExcelFormatter } from './formatter/ExcelFormatter.js';
import { CellFormatter } from './formatter/CellFormatter.js';

export class FormulaEngine {
  constructor() {
    this.parser = new FormulaParser();
    this.formatter = new ExcelFormatter();
    this.excelFunctions = new ExcelFunctions(this);
    this.cellFormatter = new CellFormatter(this.formatter);
  }

  processData(data) {
    const result = {
      tables: Array.isArray(data) ? data : [data.tables],
      formulas: []
    };

    result.tables.forEach((table, tableIndex) => {
      table.forEach((rows, rowIndex) => {
        rows.forEach((cell, colIndex) => {
          // 数式の場合は評価
          if (typeof cell.value === 'string' && cell.value.startsWith('=')) {
            cell.resolvedValue = this.evaluateFormula(cell.value, result.tables, tableIndex);
          }
          
          // フォーマット処理
          const formulaCell = this.cellFormatter.processCell(cell, tableIndex, rowIndex, colIndex);
          if (formulaCell) {
            result.formulas.push(formulaCell);
          }
        });
      });
    });

    return result;
  }

  evaluateFormula(formula, allTables, currentTableIndex, visited = new Set()) {
    try {
      if (!formula.startsWith('=')) {
        return formula;
      }

      const formulaKey = `${currentTableIndex}:${formula}`;
      if (visited.has(formulaKey)) {
        return '#CIRCULAR!';
      }
      visited.add(formulaKey);

      const ast = this.parser.parse(formula);
      const result = this.evaluateAst(ast, allTables, currentTableIndex, visited);
      
      visited.delete(formulaKey);
      return result;
    } catch (error) {
      return '#ERROR!';
    }
  }

  evaluateAst(ast, allTables, currentTableIndex) {
    if (!ast) return null;

    try {
      switch (ast.type) {
        case 'operation':
          const left = this.evaluateAst(ast.left, allTables, currentTableIndex);
          const right = this.evaluateAst(ast.right, allTables, currentTableIndex);
          
          // 数値に変換
          const leftNum = Number(left);
          const rightNum = Number(right);
          
          if (isNaN(leftNum) || isNaN(rightNum)) {
            return '#ERROR!';
          }

          switch (ast.operator) {
            case '+': return leftNum + rightNum;
            case '-': return leftNum - rightNum;
            case '*': return leftNum * rightNum;
            case '/': return rightNum === 0 ? '#DIV/0!' : leftNum / rightNum;
            case '^': return Math.pow(leftNum, rightNum);
            default: throw new Error(`Unknown operator: ${ast.operator}`);
          }

        case 'function':
          if (!(ast.name in this.excelFunctions.functions)) {
            throw new Error(`Unknown function: ${ast.name}`);
          }
          return this.excelFunctions.functions[ast.name](ast.arguments, allTables, currentTableIndex);

        case 'cell':
          const tableIndex = ast.tableId !== undefined ? ast.tableId : currentTableIndex;
          return this.getCellValue(ast.reference, allTables, tableIndex);

        case 'range':
          const rangeTableIndex = ast.tableId !== undefined ? ast.tableId : currentTableIndex;
          return this.getRangeValues(ast.reference, allTables, rangeTableIndex);

        case 'literal':
          return ast.value;

        default:
          throw new Error(`Unknown AST type: ${ast.type}`);
      }
    } catch (error) {
      return '#ERROR!';
    }
  }

  getRangeValues(range, data, tableIndex) {
    try {
      if (tableIndex >= data.length) {
        return null;
      }

      const table = data[tableIndex];
      const normalizedRange = range.replace(/([a-z]+)/gi, match => match.toUpperCase());
      const [startRef, endRef] = normalizedRange.split(':');
      const start = this.parseCellReference(startRef);
      const end = this.parseCellReference(endRef);
      
      if (!start || !end) {
        return null;
      }

      const values = [];
      for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
        for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col++) {
          if (table?.[row]?.[col]) {
            const cell = table[row][col];
            const value = cell.resolvedValue !== undefined ? cell.resolvedValue : cell.value;
            const numValue = Number(value);
            if (!isNaN(numValue)) {
              values.push(numValue);
            }
          }
        }
      }
      
      return values;
    } catch (error) {
      return null;
    }
  }

  getCellValue(ref, data, tableIndex) {
    if (tableIndex >= data.length) {
      return '#REF!';
    }

    const cellRef = this.parseCellReference(ref);
    if (!cellRef) {
      return '#REF!';
    }

    try {
      const cell = data[tableIndex][cellRef.row][cellRef.col];
      return cell.resolvedValue !== undefined ? cell.resolvedValue : cell.value;
    } catch (e) {
      return '#REF!';
    }
  }

  parseCellReference(ref) {
    const match = ref.match(/^(\$?)([A-Za-z]+)(\$?)(\d+)$/i);
    if (!match) {
      return null;
    }

    const [_, colAbs, colStr, rowAbs, rowStr] = match;
    
    let column = 0;
    const upperColStr = colStr.toUpperCase();
    for (let i = 0; i < upperColStr.length; i++) {
      column = column * 26 + (upperColStr.charCodeAt(i) - 'A'.charCodeAt(0));
    }

    return {
      col: column,
      row: parseInt(rowStr) - 1,
      isAbsoluteCol: colAbs === '$',
      isAbsoluteRow: rowAbs === '$'
    };
  }

} 