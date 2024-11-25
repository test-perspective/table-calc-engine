export class ExcelFunctions {
  constructor(engine) {
    this.engine = engine;
    this.functions = {
      'SUM': (args, allTables, currentTableIndex) => this.SUM(args, allTables, currentTableIndex),
      'AVERAGE': (args, allTables, currentTableIndex) => this.AVERAGE(args, allTables, currentTableIndex),
      'COUNT': (args, allTables, currentTableIndex) => this.COUNT(args, allTables, currentTableIndex),
      'MAX': (args, allTables, currentTableIndex) => this.MAX(args, allTables, currentTableIndex),
      'MIN': (args, allTables, currentTableIndex) => this.MIN(args, allTables, currentTableIndex)
    };
  }

  SUM(args, allTables, currentTableIndex) {
    let sum = 0;
    for (const arg of args) {
      if (arg.type === 'range') {
        const tableIndex = arg.tableId !== undefined ? arg.tableId : currentTableIndex;
        const values = this.engine.getRangeValues(arg.reference, allTables, tableIndex);
        if (values && values.length > 0) {
          sum += values.reduce((acc, val) => acc + val, 0);
        }
      } else if (arg.type === 'cell') {
        const tableIndex = arg.tableId !== undefined ? arg.tableId : currentTableIndex;
        const value = this.engine.getCellValue(arg.reference, allTables, tableIndex);
        const numValue = Number(value);
        if (!isNaN(numValue)) {
          sum += numValue;
        }
      } else if (arg.type === 'literal') {
        sum += arg.value;
      }
    }
    return sum;
  }

  AVERAGE(args, allTables, currentTableIndex) {
    try {
      const values = this.getFunctionArgValues(args, allTables, currentTableIndex);
      return ExcelFunctions.AVERAGE(values);
    } catch (error) {
      return '#ERROR!';
    }
  }

  COUNT(args, allTables, currentTableIndex) {
    try {
      const values = this.getFunctionArgValues(args, allTables, currentTableIndex);
      return ExcelFunctions.COUNT(values);
    } catch (error) {
      return '#ERROR!';
    }
  }

  MAX(args, allTables, currentTableIndex) {
    try {
      const values = this.getFunctionArgValues(args, allTables, currentTableIndex);
      return ExcelFunctions.MAX(values);
    } catch (error) {
      return '#ERROR!';
    }
  }

  MIN(args, allTables, currentTableIndex) {
    try {
      const values = this.getFunctionArgValues(args, allTables, currentTableIndex);
      return ExcelFunctions.MIN(values);
    } catch (error) {
      return '#ERROR!';
    }
  }

  getFunctionArgValues(args, allTables, currentTableIndex) {
    return args.map(arg => {
      if (arg.type === 'range') {
        const tableIndex = arg.tableId !== undefined ? arg.tableId : currentTableIndex;
        return this.engine.getRangeValues(arg.reference, allTables, tableIndex);
      } else if (arg.type === 'cell') {
        const tableIndex = arg.tableId !== undefined ? arg.tableId : currentTableIndex;
        return [this.engine.getCellValue(arg.reference, allTables, tableIndex)];
      } else if (arg.type === 'literal') {
        return [arg.value];
      }
    }).flat();
  }

  static AVERAGE(values) {
    if (!Array.isArray(values) || values.length === 0) {
      return 0;
    }
    const numbers = values.filter(value =>
      typeof value === 'number' && !isNaN(value)
    );
    if (numbers.length === 0) {
      return 0;
    }
    return numbers.reduce((acc, val) => acc + val, 0) / numbers.length;
  }

  static COUNT(values) {
    return values.filter(value => 
      typeof value === 'number' || 
      (typeof value === 'string' && !isNaN(value))
    ).length;
  }

  static MAX(values) {
    const numbers = values
      .map(v => Number(v))
      .filter(n => !isNaN(n));
    if (numbers.length === 0) return '#ERROR!';
    return Math.max(...numbers);
  }

  static MIN(values) {
    const numbers = values
      .map(v => Number(v))
      .filter(n => !isNaN(n));
    if (numbers.length === 0) return '#ERROR!';
    return Math.min(...numbers);
  }
}