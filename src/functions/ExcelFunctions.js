export class ExcelFunctions {
  constructor(engine) {
    this.engine = engine;
    this.functions = Object.fromEntries(
        Object.getOwnPropertyNames(Object.getPrototypeOf(this))
            .filter((methodName) => typeof this[methodName] === 'function' && methodName !== 'constructor')
            .map((methodName) => [methodName.toUpperCase(), (...args) => this[methodName](...args)]))
  }

  SUM(args, allTables, currentTableIndex, visited) {
    try {
      const values = this.getFunctionArgValues(args, allTables, currentTableIndex, visited);
      const numbers = values
        .map(v => Number(v))
        .filter(n => !isNaN(n));
      return numbers.reduce((sum, val) => sum + val, 0);
    } catch (error) {
      return '#ERROR!';
    }
  }

  AVERAGE(args, allTables, currentTableIndex, visited) {
    try {
      const values = this.getFunctionArgValues(args, allTables, currentTableIndex, visited);
      const numbers = values
        .map(v => Number(v))
        .filter(n => !isNaN(n));
      if (numbers.length === 0) return 0;
      return numbers.reduce((sum, val) => sum + val, 0) / numbers.length;
    } catch (error) {
      return '#ERROR!';
    }
  }

  COUNT(args, allTables, currentTableIndex, visited) {
    try {
      const values = this.getFunctionArgValues(args, allTables, currentTableIndex, visited);
      return values.filter(value => 
        typeof value === 'number' || 
        (typeof value === 'string' && !isNaN(value))
      ).length;
    } catch (error) {
      return '#ERROR!';
    }
  }

  MAX(args, allTables, currentTableIndex, visited) {
    try {
      const values = this.getFunctionArgValues(args, allTables, currentTableIndex, visited);
      const numbers = values
        .map(v => Number(v))
        .filter(n => !isNaN(n));
      if (numbers.length === 0) return '#ERROR!';
      return Math.max(...numbers);
    } catch (error) {
      return '#ERROR!';
    }
  }

  MIN(args, allTables, currentTableIndex, visited) {
    try {
      const values = this.getFunctionArgValues(args, allTables, currentTableIndex, visited);
      const numbers = values
        .map(v => Number(v))
        .filter(n => !isNaN(n));
      if (numbers.length === 0) return '#ERROR!';
      return Math.min(...numbers);
    } catch (error) {
      return '#ERROR!';
    }
  }

  DATE(args, allTables, currentTableIndex, visited) {
    try {
      const values = this.getFunctionArgValues(args, allTables, currentTableIndex, visited);
      
      // 引数が3つであることを確認
      if (values.length !== 3) {
        return '#VALUE!';
      }

      // 各引数を数値に変換（小数点以下は切り捨て）
      const [yearNum, monthNum, dayNum] = values.map(v => {
        const num = Number(v);
        if (isNaN(num)) {
          throw new Error('#VALUE!');
        }
        return Math.floor(num);
      });

      // 年の処理
      let year = yearNum;
      if (year >= 0 && year < 1900) {
        year += 1900;
      }

      // 年の範囲チェック
      if (year < 0 || year > 9999) {
        return '#NUM!';
      }

      try {
        const date = new Date(year, monthNum - 1, dayNum);
        if (isNaN(date.getTime())) {
          return '#VALUE!';
        }

        // Excel基準日からの経過日数を計算
        const excelBaseDate = new Date(1900, 0, 1);
        const daysSinceBase = Math.floor((date - excelBaseDate) / (1000 * 60 * 60 * 24)) + 1;
        
        return daysSinceBase;
      } catch (error) {
        return '#VALUE!';
      }
    } catch (error) {
      if (typeof error === 'object' && error.message) {
        return error.message;
      }
      return '#VALUE!';
    }
  }

  TODAY(args, allTables, currentTableIndex, visited) {
    try {
      // 引数は無視する（Excelの仕様に準拠）
      const now = new Date();
      const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
      
      // Excel基準日（1900年1月1日）
      const excelBaseDate = new Date(1900, 0, 1);
      
      // 経過日数を計算（ミリ秒を日数に変換）
      const daysSinceBase = Math.floor((today - excelBaseDate) / (1000 * 60 * 60 * 24)) + 1;
      
      return daysSinceBase;
    } catch (error) {
      return '#ERROR!';
    }
  }

  getFunctionArgValues(args, allTables, currentTableIndex, visited) {
    return args.map(arg => {
      if (arg.type === 'range') {
        const tableIndex = arg.tableId !== undefined ? arg.tableId : currentTableIndex;
        return this.engine.getRangeValues(arg.reference, allTables, tableIndex, visited);
      } else if (arg.type === 'cell') {
        const tableIndex = arg.tableId !== undefined ? arg.tableId : currentTableIndex;
        return [this.engine.getCellValue(arg.reference, allTables, tableIndex, visited)];
      } else if (arg.type === 'literal') {
        return [arg.value];
      }
    }).flat();
  }
}