export class ExcelFunctions {
  static AVERAGE(values) {
    if (!Array.isArray(values) || values.length === 0) {
      return 0;
    }

    // 数値のみをフィルタリング
    const numbers = values.filter(value =>
      typeof value === 'number' && !isNaN(value)
    );

    if (numbers.length === 0) {
      return 0;
    }

    // 合計を計算
    const sum = numbers.reduce((acc, val) => acc + val, 0);

    // 平均を返す
    return sum / numbers.length;
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