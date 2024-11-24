import { FormulaEngine } from '../src/FormulaEngine.js';

describe.only('Table Reference Tests', () => {
  let engine;
  let testData;

  beforeEach(() => {
    engine = new FormulaEngine();
    
    // 複数のテーブルを含むテストデータ
    testData = [
      [ // table 0
        [
          { value: "A", resolved: true, resolvedValue: "A", displayValue: "A" },
          { value: "B", resolved: true, resolvedValue: "B", displayValue: "B" }
        ],
        [
          { value: 1, resolved: true, resolvedValue: 1, displayValue: 1 },
          { value: 2, resolved: true, resolvedValue: 2, displayValue: 2 }
        ]
      ],
      [ // table 1
        [
          { value: "X", resolved: true, resolvedValue: "X", displayValue: "X" },
          { value: "Y", resolved: true, resolvedValue: "Y", displayValue: "Y" }
        ],
        [
          { value: 10, resolved: true, resolvedValue: 10, displayValue: 10 },
          { value: 20, resolved: true, resolvedValue: 20, displayValue: 20 }
        ]
      ]
    ];
  });

  describe('Basic Table References', () => {
    test('should handle simple table reference', () => {
      const result = engine.evaluateFormula('=0!A1', testData, 0);
      expect(result).toBe('A');
    });

    test('should handle reference to other table', () => {
      const result = engine.evaluateFormula('=1!A1', testData, 0);
      expect(result).toBe('X');
    });

    test('should handle lowercase table reference', () => {
      const result = engine.evaluateFormula('=0!a1', testData, 0);
      expect(result).toBe('A');
    });

    test('should handle invalid table reference', () => {
      const result = engine.evaluateFormula('=99!A1', testData, 0);
      expect(result).toBe('#REF!');
    });
  });

  describe('Table Range References', () => {
    test('should handle range reference within same table', () => {
      const result = engine.evaluateFormula('=SUM(0!A2:B2)', testData, 0);
      expect(result).toBe(3); // 1 + 2 = 3
    });

    test('should handle range reference to other table', () => {
      const result = engine.evaluateFormula('=SUM(1!A2:B2)', testData, 0);
      expect(result).toBe(30); // 10 + 20 = 30
    });

    test('should handle lowercase range reference', () => {
      const result = engine.evaluateFormula('=SUM(1!a2:b2)', testData, 0);
      expect(result).toBe(30);
    });
  });

  describe('Mixed Table References', () => {
    test('should handle arithmetic between tables', () => {
      const result = engine.evaluateFormula('=0!A2+1!A2', testData, 0);
      expect(result).toBe(11); // 1 + 10 = 11
    });

    test('should handle mixed absolute references', () => {
      const result = engine.evaluateFormula('=SUM(0!$A$2:$B$2)', testData, 0);
      expect(result).toBe(3);
    });

    test('should handle complex table references', () => {
      const result = engine.evaluateFormula('=SUM(0!A2:B2)+SUM(1!A2:B2)', testData, 0);
      expect(result).toBe(33); // (1 + 2) + (10 + 20) = 33
    });
  });

  describe('Error Handling', () => {
    test('should handle missing table reference', () => {
      const result = engine.evaluateFormula('=!A1', testData, 0);
      expect(result).toBe('#ERROR!');
    });

    test('should handle invalid table format', () => {
      const result = engine.evaluateFormula('=0A!A1', testData, 0);
      expect(result).toBe('#ERROR!');
    });

    test('should handle out of range table reference', () => {
      const result = engine.evaluateFormula('=2!A1', testData, 0);
      expect(result).toBe('#REF!');
    });

    test('should handle out of range cell in valid table', () => {
      const result = engine.evaluateFormula('=1!Z99', testData, 0);
      expect(result).toBe('#REF!');
    });
  });

  describe('Function with Table References', () => {
    test('should handle SUM across tables', () => {
      const result = engine.evaluateFormula('=SUM(0!A2:B2, 1!A2:B2)', testData, 0);
      expect(result).toBe(33); // (1 + 2) + (10 + 20) = 33
    });

    test('should handle AVERAGE with table reference', () => {
      const result = engine.evaluateFormula('=AVERAGE(1!A2:B2)', testData, 0);
      expect(result).toBe(15); // (10 + 20) / 2 = 15
    });

    test('should handle nested functions with table references', () => {
      const result = engine.evaluateFormula('=SUM(0!A2:B2) + AVERAGE(1!A2:B2)', testData, 0);
      expect(result).toBe(18); // (1 + 2) + ((10 + 20) / 2) = 18
    });
  });
}); 