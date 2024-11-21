export class ExcelFunctions {
  static SUM(values) {
    const numbers = values.filter(val => typeof val === 'number');
    return numbers.reduce((sum, val) => sum + val, 0);
  }

  static AVERAGE(values) {
    const numbers = values.filter(val => typeof val === 'number');
    if (numbers.length === 0) return 0;
    return this.SUM(numbers) / numbers.length;
  }

  static COUNT(values) {
    return values.filter(val => !isNaN(Number(val))).length;
  }

  static MAX(values) {
    const numbers = values.filter(val => !isNaN(Number(val)));
    return Math.max(...numbers);
  }

  static MIN(values) {
    const numbers = values.filter(val => !isNaN(Number(val)));
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