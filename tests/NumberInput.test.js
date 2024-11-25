import { FormulaEngine } from '../src/FormulaEngine.js';

describe('Number Input Tests', () => {
  let engine;

  beforeEach(() => {
    engine = new FormulaEngine();
  });

  test('should handle numbers with commas', () => {
    const testData = [[
      [
        { value: '1,234.567' },  // カンマ付き数値
        { value: '1,234,567.89' },  // 複数のカンマ
        { value: '=A1 + 1' },  // カンマ付き数値を計算に使用
        { value: '=B1 * 2' }   // 複数カンマ付き数値を計算に使用
      ]
    ]];

    const result = engine.processData(testData);

    expect(result.tables[0][0][0].resolvedValue).toBe(1234.567);  // カンマが除去されて数値に
    expect(result.tables[0][0][1].resolvedValue).toBe(1234567.89);  // 複数のカンマが除去されて数値に
    expect(result.tables[0][0][2].resolvedValue).toBe(1235.567);  // 計算結果が正しい
    expect(result.tables[0][0][3].resolvedValue).toBe(2469135.78);  // 計算結果が正しい
  });

  test('should handle mixed format numbers', () => {
    const testData = [[
      [
        { value: '1,234' },      // 整数部のみカンマ
        { value: '1,234.00' },   // 小数部を含むカンマ付き
        { value: '1234.567' },   // カンマなし
        { value: '=SUM(A1:C1)' }  // 全ての数値を合計
      ]
    ]];

    const result = engine.processData(testData);

    expect(result.tables[0][0][0].resolvedValue).toBe(1234);
    expect(result.tables[0][0][1].resolvedValue).toBe(1234.00);
    expect(result.tables[0][0][2].resolvedValue).toBe(1234.567);
    expect(result.tables[0][0][3].resolvedValue).toBe(3702.567);
  });

}); 