import { FormulaEngine } from '../src/FormulaEngine.js';

describe('Excel Functions', () => {
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

  test('should calculate SUM correctly', () => {
    const result = engine.evaluateFormula('=SUM(A2:B3)', testData, 0);
    expect(result).toBe(12); // 1 + 2 + 4 + 5 = 12
  });

  test('should calculate AVERAGE correctly', () => {
    const result = engine.evaluateFormula('=AVERAGE(A2:B3)', testData, 0);
    expect(result).toBe(3); // (1 + 2 + 4 + 5) / 4 = 3
  });

  test('should calculate COUNT correctly', () => {
    const result = engine.evaluateFormula('=COUNT(A2:B3)', testData, 0);
    expect(result).toBe(4); // 4 numeric values
  });

  test('should calculate MAX correctly', () => {
    const result = engine.evaluateFormula('=MAX(A2:B3)', testData, 0);
    expect(result).toBe(5); // highest value is 5
  });

  test('should calculate MIN correctly', () => {
    const result = engine.evaluateFormula('=MIN(A2:B3)', testData, 0);
    expect(result).toBe(1); // lowest value is 1
  });
}); 