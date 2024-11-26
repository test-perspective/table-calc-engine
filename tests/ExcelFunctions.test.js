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
          { value: "A" },
          { value: "B" },
          { value: "C" }
        ],
        [
          { value: 1 },
          { value: 2 },
          { value: 3 }
        ],
        [
          { value: 4 },
          { value: 5 },
          { value: 6 }
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

  describe('DATE Function', () => {
    test('should create date from year, month, day', () => {
      const result = engine.evaluateFormula('=DATE(2024,1,15)', testData, 0);
      expect(result instanceof Date).toBe(true);
      expect(result.getFullYear()).toBe(2024);
      expect(result.getMonth()).toBe(0); // 0-based month
      expect(result.getDate()).toBe(15);
    });

    test('should handle year between 0 and 1899', () => {
      const result = engine.evaluateFormula('=DATE(99,1,1)', testData, 0);
      expect(result.getFullYear()).toBe(1999);
    });

    test('should handle year between 1900 and 9999', () => {
      const result = engine.evaluateFormula('=DATE(1901,1,1)', testData, 0);
      expect(result.getFullYear()).toBe(1901);
    });

    test('should handle month overflow', () => {
      const result = engine.evaluateFormula('=DATE(2024,13,1)', testData, 0);
      expect(result.getFullYear()).toBe(2025);
      expect(result.getMonth()).toBe(0); // 13月は次年の1月
    });

    test.only('should handle invalid parameters', () => {
      console.log('zzzzzeee...')
      expect(engine.evaluateFormula('=DATE("invalid",1,1)', testData, 0)).toBe('#VALUE!');
      expect(engine.evaluateFormula('=DATE(2024,"invalid",1)', testData, 0)).toBe('#VALUE!');
      expect(engine.evaluateFormula('=DATE(2024,1,"invalid")', testData, 0)).toBe('#VALUE!');
      expect(engine.evaluateFormula('=DATE(10000,1,1)', testData, 0)).toBe('#NUM!');
      expect(engine.evaluateFormula('=DATE(-1,1,1)', testData, 0)).toBe('#NUM!');
    });

    test('should handle cell references as parameters', () => {
      testData[0][0][0] = { value: 2024 };
      testData[0][0][1] = { value: 1 };
      testData[0][0][2] = { value: 15 };
      
      const result = engine.evaluateFormula('=DATE(A1,B1,C1)', testData, 0);
      expect(result instanceof Date).toBe(true);
      expect(result.getFullYear()).toBe(2024);
      expect(result.getMonth()).toBe(0);
      expect(result.getDate()).toBe(15);
    });

    test('should handle string literals correctly', () => {
      expect(engine.evaluateFormula('=SUM(1,"2",3,"invalid")', testData, 0)).toBe(6);
      expect(engine.evaluateFormula('=AVERAGE(1,"2",3,"invalid")', testData, 0)).toBe(2);
      expect(engine.evaluateFormula('=COUNT(1,"2",3,"invalid")', testData, 0)).toBe(3);
      expect(engine.evaluateFormula('=MAX(1,"2",3,"invalid")', testData, 0)).toBe(3);
      expect(engine.evaluateFormula('=MIN(1,"2",3,"invalid")', testData, 0)).toBe(1);
    });
  });
}); 