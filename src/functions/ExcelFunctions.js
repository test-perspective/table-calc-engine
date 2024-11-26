export class ExcelFunctions {
  constructor(engine) {
    this.engine = engine;
    this.functions = Object.fromEntries(
        Object.getOwnPropertyNames(Object.getPrototypeOf(this))
            .filter((methodName) => typeof this[methodName] === 'function' && methodName !== 'constructor')
            .map((methodName) => [methodName.toUpperCase(), (...args) => this[methodName](...args)]))
  }

  SUM(args, allTables, currentTableIndex) {
    try {
      const values = this.getFunctionArgValues(args, allTables, currentTableIndex);
      const numbers = values
        .map(v => Number(v))
        .filter(n => !isNaN(n));
      return numbers.reduce((sum, val) => sum + val, 0);
    } catch (error) {
      return '#ERROR!';
    }
  }

  AVERAGE(args, allTables, currentTableIndex) {
    try {
      const values = this.getFunctionArgValues(args, allTables, currentTableIndex);
      const numbers = values
        .map(v => Number(v))
        .filter(n => !isNaN(n));
      if (numbers.length === 0) return 0;
      return numbers.reduce((sum, val) => sum + val, 0) / numbers.length;
    } catch (error) {
      return '#ERROR!';
    }
  }

  COUNT(args, allTables, currentTableIndex) {
    try {
      const values = this.getFunctionArgValues(args, allTables, currentTableIndex);
      return values.filter(value => 
        typeof value === 'number' || 
        (typeof value === 'string' && !isNaN(value))
      ).length;
    } catch (error) {
      return '#ERROR!';
    }
  }

  MAX(args, allTables, currentTableIndex) {
    try {
      const values = this.getFunctionArgValues(args, allTables, currentTableIndex);
      const numbers = values
        .map(v => Number(v))
        .filter(n => !isNaN(n));
      if (numbers.length === 0) return '#ERROR!';
      return Math.max(...numbers);
    } catch (error) {
      return '#ERROR!';
    }
  }

  MIN(args, allTables, currentTableIndex) {
    try {
      const values = this.getFunctionArgValues(args, allTables, currentTableIndex);
      const numbers = values
        .map(v => Number(v))
        .filter(n => !isNaN(n));
      if (numbers.length === 0) return '#ERROR!';
      return Math.min(...numbers);
    } catch (error) {
      return '#ERROR!';
    }
  }

  DATE(args, allTables, currentTableIndex) {
    try {
      const values = this.getFunctionArgValues(args, allTables, currentTableIndex);
      
      // 引数が3つであることを確認
      if (values.length !== 3) {
        return '#VALUE!';
      }

      // 各引数を数値に変換（小数点以下は切り捨て）
      const [yearNum, monthNum, dayNum] = values.map(v => {
        // 文字列が数値に変換できない場合は #VALUE! を返す
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
        // JavaScriptのDateオブジェクトを使用して日付を作成
        // monthは0-basedなので1を引く
        const date = new Date(year, monthNum - 1, dayNum);

        // 有効な日付かどうかを確認
        if (isNaN(date.getTime())) {
          return '#VALUE!';
        }

        return date;
      } catch (error) {
        return '#VALUE!';
      }
    } catch (error) {
      // エラーメッセージをそのまま返す
      if (typeof error === 'object' && error.message) {
        return error.message;
      }
      return '#VALUE!';
    }
  }

  TODAY(args, allTables, currentTableIndex) {
    try {
      // 引数は無視する（Excelの仕様に準拠）
      const now = new Date();
      
      // 時刻部分を0にリセットした新しい日付オブジェクトを返す
      return new Date(now.getFullYear(), now.getMonth(), now.getDate());
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
}