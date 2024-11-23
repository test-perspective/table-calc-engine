import { FormulaParser } from './parser/FormulaParser.js';
import { ExcelFunctions } from './functions/ExcelFunctions.js';
import { CellReference } from '../utils/CellReference.js';
import { FormulaError } from './types/index.js';

export class FormulaEngine {
  constructor() {
    this.parser = new FormulaParser();
    this.functions = {
      'SUM': (args, allTables, currentTableIndex) => {
        const values = this.getFunctionArgValues(args, allTables, currentTableIndex);
        return ExcelFunctions.SUM(values);
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
      
      if (typeof cell.value === 'string' && cell.value.startsWith('=')) {
        if (cell.resolvedValue === undefined) {
          cell.resolvedValue = this.evaluateFormula(cell.value, allTables, currentTableIndex);
          cell.displayValue = cell.resolvedValue;
        }
        return cell.resolvedValue;
      }
      
      return cell.value;
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
} 