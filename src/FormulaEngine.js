import { FormulaParser } from './parser/FormulaParser.js';
import { ExcelFunctions } from './functions/ExcelFunctions.js';
import { CellReference } from '../utils/CellReference.js';
import { FormulaError } from './types/index.js';

export class FormulaEngine {
  constructor() {
    this.parser = new FormulaParser();
    this.functions = {
      'SUM': (args, allTables, currentTableIndex) => {
        let sum = 0;
        for (const arg of args) {
          if (arg.type === 'range') {
            const values = this.getRangeValues(arg.reference, allTables, currentTableIndex);
            if (values && values.length > 0) {
              sum += values.reduce((acc, val) => acc + val, 0);
            }
          } else if (arg.type === 'cell') {
            const value = this.getCellValue(arg.reference, allTables, currentTableIndex);
            const numValue = Number(value);
            if (!isNaN(numValue)) {
              sum += numValue;
            }
          } else if (arg.type === 'literal') {
            sum += arg.value;
          }
        }
        return sum;
      },
      'AVERAGE': (args, allTables, currentTableIndex) => { 
        const values = this.getFunctionArgValues(args, allTables, currentTableIndex);
        return ExcelFunctions.AVERAGE(values);
      },
      'COUNT': (args, allTables, currentTableIndex) => {
        const values = this.getFunctionArgValues(args, allTables, currentTableIndex);
        return ExcelFunctions.COUNT(values);
      },
      'MAX': (args, allTables, currentTableIndex) => {
        const values = this.getFunctionArgValues(args, allTables, currentTableIndex);
        return ExcelFunctions.MAX(values);
      },
      'MIN': (args, allTables, currentTableIndex) => {
        const values = this.getFunctionArgValues(args, allTables, currentTableIndex);
        return ExcelFunctions.MIN(values);
      }
    };
  }

  getFunctionArgValues(args, allTables, currentTableIndex) {
    try {
      return args.map(arg => {
        if (arg.type === 'range') {
          return this.getRangeValues(arg.reference, allTables, currentTableIndex);
        } else if (arg.type === 'cell') {
          return [this.getCellValue(arg.reference, allTables, currentTableIndex)];
        } else {
          return [this.evaluateAst(arg, allTables, currentTableIndex)];
        }
      }).flat();
    } catch (error) {
      console.error('Error in getFunctionArgValues:', error);
      throw error;
    }
  }

  processData(data) {
    console.log('Starting processData');
    const result = {
      tables: JSON.parse(JSON.stringify(data)),
      formulas: []
    };

    result.tables.forEach((table, tableIndex) => {
      table.forEach((rows, sheetIndex) => {
        rows.forEach((cell, colIndex) => {
          const originalCell = { ...cell };

          if (typeof cell.value === 'string' && cell.value.startsWith('=')) {
            try {
              const formula = cell.value;
              const resolvedValue = this.evaluateFormula(formula, result.tables, tableIndex);
              
              cell.resolvedValue = resolvedValue;
              cell.displayValue = this.formatValue(resolvedValue, cell.excelFormat);
              cell.resolved = true;

              result.formulas.push({
                tableIndex,
                sheetIndex,
                colIndex,
                formula,
                resolvedValue,
                displayValue: cell.displayValue,
                excelFormat: cell.excelFormat,
                macroId: originalCell.macroId,
                style: originalCell.style,
                metadata: originalCell.metadata,
                validation: originalCell.validation,
                comment: originalCell.comment,
                originalCell: originalCell
              });
            } catch (error) {
              console.error('Formula evaluation error:', error);
              cell.resolvedValue = '#ERROR!';
              cell.displayValue = '#ERROR!';
              cell.resolved = false;
            }
          } else {
            cell.resolvedValue = cell.value;
            cell.displayValue = this.formatValue(cell.value, cell.excelFormat);
            cell.resolved = true;
          }
        });
      });
    });

    console.log('ProcessData completed successfully');
    return result;
  }

  processTableData(table, allTables, currentTableIndex) {
    return table.map(row =>
      row.map(cell => {
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
        if (!isNaN(Number(cell.value))) {
          return {
            ...cell,
            value: Number(cell.value),
            resolvedValue: Number(cell.value),
            displayValue: this.formatValue(Number(cell.value), cell.excelFormat),
            resolved: true
          };
        }
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

      case 'operation':
        const left = this.evaluateAst(ast.left, allTables, currentTableIndex);
        const right = this.evaluateAst(ast.right, allTables, currentTableIndex);
        
        const leftNum = Number(left);
        const rightNum = Number(right);
        
        if (isNaN(leftNum) || isNaN(rightNum)) {
          return '#ERROR!';
        }

        switch (ast.operator) {
          case '+': return leftNum + rightNum;
          case '-': return leftNum - rightNum;
          case '*': return leftNum * rightNum;
          case '/':
            if (rightNum === 0) return '#DIV/0!';
            return leftNum / rightNum;
          case '^': return Math.pow(leftNum, rightNum);
          default:
            throw new Error(`Unknown operator: ${ast.operator}`);
        }

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

  sumFunction(args, allTables, currentTableIndex, visited) {
    try {
      const values = args.map(arg => {
        if (arg.type === 'range') {
          return this.getRangeValues(arg.reference, allTables, currentTableIndex);
        } else if (arg.type === 'cell') {
          return [this.getCellValue(arg.reference, allTables, currentTableIndex)];
        } else {
          return [this.evaluateAst(arg, allTables, currentTableIndex, visited)];
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
      const normalizedRange = range.replace(/[a-z]+/g, match => match.toUpperCase());
      const [startRef, endRef] = normalizedRange.split(':');
      const start = this.parseCellReference(startRef);
      const end = this.parseCellReference(endRef);
      
      if (!start || !end) {
        return null;
      }

      const table = allTables[currentTableIndex];
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
      console.error('Error in getRangeValues:', error);
      return null;
    }
  }

  getCellValue(ref, data, sheetIndex) {
    const cellRef = this.parseCellReference(ref);
    if (!cellRef) {
      return '#REF!';
    }

    try {
      const cell = data[sheetIndex][cellRef.row][cellRef.col];
      return cell.resolvedValue !== undefined ? cell.resolvedValue : cell.value;
    } catch (e) {
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
    if (value === null || value === undefined) return '';
    
    if (typeof value === 'number') {
      if (!format) {
        return value.toString();
      }
      if (format.includes('.')) {
        const decimals = format.split('.')[1].length;
        return value.toFixed(decimals);
      }
      return value.toString();
    }
    
    return value.toString();
  }

  getDecimalPlaces(format) {
    const match = format.match(/\.(\d+)/);
    return match ? match[1].length : 0;
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

  parseRange(range) {
    const [start, end] = range.split(':');
    const startCell = this.parseCellReference(start);
    const endCell = this.parseCellReference(end);

    if (!startCell || !endCell) {
      return null;
    }

    return {
      startCol: startCell.col,
      startRow: startCell.row,
      endCol: endCell.col,
      endRow: endCell.row,
      isAbsoluteStartCol: startCell.isAbsoluteCol,
      isAbsoluteStartRow: startCell.isAbsoluteRow,
      isAbsoluteEndCol: endCell.isAbsoluteCol,
      isAbsoluteEndRow: endCell.isAbsoluteRow
    };
  }
} 