import { FormulaEngine } from '../src/FormulaEngine.js';

describe('FormulaEngine', () => {
  let engine;
  let testData;

  beforeEach(() => {
    engine = new FormulaEngine();
    
    // テストデータの準備
    testData = [
      [
        [
          { value: "A", resolved: true, resolvedValue: "A", displayValue: "A" },
          { value: "B", resolved: true, resolvedValue: "B", displayValue: "B" },
          { value: "C", resolved: true, resolvedValue: "C", displayValue: "C" }
        ],
        [
          { value: 1, resolved: true, resolvedValue: 1, displayValue: 1 },
          { value: 2, resolved: true, resolvedValue: 2, displayValue: 2 },
          { value: 3, resolved: true, resolvedValue: 3, displayValue: 3 }
        ],
        [
          { value: 4, resolved: true, resolvedValue: 4, displayValue: 4 },
          { value: 5, resolved: true, resolvedValue: 5, displayValue: 5 },
          { value: 6, resolved: true, resolvedValue: 6, displayValue: 6 }
        ]
      ]
    ];
  });

  // 基本的なセル参照のテスト
  test('should handle simple cell reference', () => {
    const result = engine.evaluateFormula('=A1', testData, 0);
    expect(result).toBe('A');
  });

  // SUM関数のテスト
  test('should calculate SUM correctly', () => {
    const result = engine.evaluateFormula('=SUM(A2:B3)', testData, 0);
    expect(result).toBe(12); // 1 + 2 + 4 + 5 = 12
  });

  // 算術演算のテスト
  test('should handle basic arithmetic', () => {
    const result = engine.evaluateFormula('=1+2', testData, 0);
    expect(result).toBe(3);
  });

  // セル参照を含む算術演算のテスト
  test('should handle arithmetic with cell references', () => {
    const result = engine.evaluateFormula('=A2+B2', testData, 0);
    expect(result).toBe(3); // 1 + 2 = 3
  });

  // エラー処理のテスト
  test('should handle invalid cell reference', () => {
    const result = engine.evaluateFormula('=Z99', testData, 0);
    expect(result).toBe('#REF!');
  });

  test('should handle division by zero', () => {
    const result = engine.evaluateFormula('=1/0', testData, 0);
    expect(result).toBe('#DIV/0!');
  });
}); 