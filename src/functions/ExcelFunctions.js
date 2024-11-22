export class ExcelFunctions {
  static SUM(values) {
    return values.reduce((sum, value) => {
      const num = Number(value);
      return isNaN(num) ? sum : sum + num;
    }, 0);
  }

  static AVERAGE(values) {
    if (values.length === 0) return '#DIV/0!';
    const sum = this.SUM(values);
    if (typeof sum !== 'number') return sum;
    return sum / values.length;
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

  static IF(condition, trueValue, falseValue) {
    return condition ? trueValue : falseValue;
  }

  static TODAY() {
    return new Date();
  }

  static DATE(year, month, day) {
    return new Date(year, month - 1, day);
  }
} 