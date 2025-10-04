import { FormulaEngine } from '../src/FormulaEngine.js';

describe('Cell Reference Tests', () => {
  let engine;
  let testData;

  beforeEach(() => {
    engine = new FormulaEngine();
    
    // セル参照テスト用のデータ
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

  describe('Basic Cell References', () => {
    test('should handle simple cell reference', () => {
      const result = engine.evaluateFormula('=A1', testData, 0);
      expect(result).toBe('A');
    });

    test('should handle arithmetic with cell references', () => {
      const result = engine.evaluateFormula('=A2+B2', testData, 0);
      expect(result).toBe(3); // 1 + 2 = 3
    });

    test('should handle arithmetic with lowercase cell references', () => {
      const result = engine.evaluateFormula('=a2+b2', testData, 0);
      expect(result).toBe(3); // 1 + 2 = 3
    });

    test('should handle invalid cell reference', () => {
      const result = engine.evaluateFormula('=Z99', testData, 0);
      expect(result).toBe('#REF!');
    });

    test('should handle lowercase cell reference', () => {
      const result = engine.evaluateFormula('=a1', testData, 0);
      expect(result).toBe('A');
    });

    test('should handle mixed case cell reference', () => {
      const result = engine.evaluateFormula('=a2', testData, 0);
      expect(result).toBe(1);
    });

    test('should handle arithmetic with lowercase cell references', () => {
      const result = engine.evaluateFormula('=a2+b2', testData, 0);
      expect(result).toBe(3); // 1 + 2 = 3
    });

    test('should add two numeric cells with lowercase refs', () => {
      const result = engine.evaluateFormula('=a2+a3', testData, 0);
      expect(result).toBe(5);
    });
  });

  describe('Absolute Cell References', () => {
    test('should handle fully absolute cell reference', () => {
      const result = engine.evaluateFormula('=$A$2', testData, 0);
      expect(result).toBe(1);
    });

    test('should handle column absolute reference', () => {
      const result = engine.evaluateFormula('=$A2', testData, 0);
      expect(result).toBe(1);
    });

    test('should handle row absolute reference', () => {
      const result = engine.evaluateFormula('=A$2', testData, 0);
      expect(result).toBe(1);
    });

    test('should handle absolute references in range operations', () => {
      const result = engine.evaluateFormula('=SUM($A$2:$B$3)', testData, 0);
      expect(result).toBe(12); // 1 + 2 + 4 + 5 = 12
    });

    test('should handle mixed absolute and relative references in range', () => {
      const result = engine.evaluateFormula('=SUM($A2:B$3)', testData, 0);
      expect(result).toBe(12); // 1 + 2 + 4 + 5 = 12
    });

    test('should handle invalid absolute cell reference', () => {
      const result = engine.evaluateFormula('=$Z$99', testData, 0);
      expect(result).toBe('#REF!');
    });

    test('should handle lowercase absolute cell reference', () => {
      const result = engine.evaluateFormula('=$a$2', testData, 0);
      expect(result).toBe(1);
    });

    test('should handle mixed case absolute references', () => {
      const result = engine.evaluateFormula('=$A$2+$b$2', testData, 0);
      expect(result).toBe(3); // 1 + 2 = 3
    });

    test('should handle lowercase range references', () => {
      const result = engine.evaluateFormula('=SUM(a2:b3)', testData, 0);
      expect(result).toBe(12); // 1 + 2 + 4 + 5 = 12
    });

    test('should handle mixed case absolute range references', () => {
      const result = engine.evaluateFormula('=SUM($a$2:$B$3)', testData, 0);
      expect(result).toBe(12); // 1 + 2 + 4 + 5 = 12
    });

    test('should handle invalid lowercase cell reference', () => {
      const result = engine.evaluateFormula('=$z$99', testData, 0);
      expect(result).toBe('#REF!');
    });
  });

  describe('Function Case Sensitivity', () => {
    test('should handle lowercase function name', () => {
      const result = engine.evaluateFormula('=sum(A2:B3)', testData, 0);
      expect(result).toBe(12); // 1 + 2 + 4 + 5 = 12
    });

    test('should handle mixed case function name', () => {
      const result = engine.evaluateFormula('=SuM(A2:B3)', testData, 0);
      expect(result).toBe(12);
    });

    test('should handle lowercase function name with lowercase references', () => {
      const result = engine.evaluateFormula('=sum(a2:b3)', testData, 0);
      expect(result).toBe(12);
    });

    test('should handle mixed case function with mixed case references', () => {
      const result = engine.evaluateFormula('=sUm($A$2:$b$3)', testData, 0);
      expect(result).toBe(12);
    });

    test('should handle lowercase function with absolute references', () => {
      const result = engine.evaluateFormula('=sum($A$2:$B$3)', testData, 0);
      expect(result).toBe(12);
    });
  });
}); 