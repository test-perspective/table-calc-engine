class ExcelFunctions {
    static SUM(values) {
      return values.reduce((sum, val) => sum + (Number(val) || 0), 0);
    }
  
    static AVERAGE(values) {
      if (values.length === 0) return 0;
      return this.SUM(values) / values.length;
    }
  
    static COUNT(values) {
      return values.filter(val => typeof val === 'number').length;
    }
  
    static MAX(values) {
      return Math.max(...values.filter(val => typeof val === 'number'));
    }
  
    static MIN(values) {
      return Math.min(...values.filter(val => typeof val === 'number'));
    }
  
    static IF(condition, trueValue, falseValue) {
      return condition ? trueValue : falseValue;
    }
  
    static VLOOKUP(lookupValue, tableArray, colIndex, exactMatch = true) {
      // VLOOKUP実装
      for (const row of tableArray) {
        if (exactMatch ? row[0] === lookupValue : row[0] >= lookupValue) {
          return row[colIndex - 1];
        }
      }
      return null;
    }
  
    // 日付関数
    static TODAY() {
      return new Date();
    }
  
    static DATE(year, month, day) {
      return new Date(year, month - 1, day);
    }
  
    // 統計関数
    static MEDIAN(values) {
      const sorted = [...values].sort((a, b) => a - b);
      const mid = Math.floor(sorted.length / 2);
      return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
    }
  }