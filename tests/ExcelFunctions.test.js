import { FormulaEngine } from '../src/FormulaEngine.js';
import { ExcelFormatter } from '../src/formatter/ExcelFormatter.js';

describe('Excel Functions', () => {
  let engine;
  let testData;
  let formatter;

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

    formatter = new ExcelFormatter();
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

  describe('TODAY Function', () => {
    test('should return current date as Excel serial number', () => {
      const result = engine.evaluateFormula('=TODAY()', testData, 0);
      const today = new Date();
      today.setHours(0, 0, 0, 0);
      const excelBaseDate = new Date(1900, 0, 1);
      const expectedSerial = Math.floor((today - excelBaseDate) / (1000 * 60 * 60 * 24)) + 1;
      
      expect(typeof result).toBe('number');
      expect(result).toBe(expectedSerial);
    });

    test('should ignore arguments if provided', () => {
      const result1 = engine.evaluateFormula('=TODAY()', testData, 0);
      const result2 = engine.evaluateFormula('=TODAY(1,2,3)', testData, 0);
      
      expect(result1).toBe(result2);
    });

    test('should work with arithmetic operations', () => {
      const today = engine.evaluateFormula('=TODAY()', testData, 0);
      
      // 明日
      const tomorrow = engine.evaluateFormula('=TODAY()+1', testData, 0);
      expect(tomorrow).toBe(today + 1);
      
      // 昨日
      const yesterday = engine.evaluateFormula('=TODAY()-1', testData, 0);
      expect(yesterday).toBe(today - 1);
      
      // 来週
      const nextWeek = engine.evaluateFormula('=TODAY()+7', testData, 0);
      expect(nextWeek).toBe(today + 7);
      
      // 先月（おおよそ）
      const lastMonth = engine.evaluateFormula('=TODAY()-30', testData, 0);
      expect(lastMonth).toBe(today - 30);
    });

    test('should work in complex formulas', () => {
      const today = engine.evaluateFormula('=TODAY()', testData, 0);
      
      // 日付の差分
      const diffDays = engine.evaluateFormula('=TODAY()-DATE(2024,1,1)', testData, 0);
      const date20240101 = engine.evaluateFormula('=DATE(2024,1,1)', testData, 0);
      expect(diffDays).toBe(today - date20240101);
      
      // 複数の演算
      const complexCalc = engine.evaluateFormula('=(TODAY()+1)*2-TODAY()', testData, 0);
      // (today + 1) * 2 - today
      // = (today * 2 + 2) - today
      // = today * 2 - today + 2
      // = today + 2
      expect(complexCalc).toBe(today + 2);

      // 別の複数の演算
      const anotherCalc = engine.evaluateFormula('=TODAY()*2-TODAY()', testData, 0);
      // today * 2 - today = today
      expect(anotherCalc).toBe(today);
    });

    test('should format as date when specified', () => {
      const testData = [[
        [
          { value: '=TODAY()', excelFormat: 'yyyy/mm/dd' }  // フォーマットを指定
        ]
      ]];

      const result = engine.processData(testData);
      const formatted = result.tables[0][0][0];
      const today = new Date();
      const expectedFormat = `${today.getFullYear()}/${String(today.getMonth() + 1).padStart(2, '0')}/${String(today.getDate()).padStart(2, '0')}`;

      expect(formatted.displayValue).toBe(expectedFormat);
    });

    test('should format as date in complex formulas', () => {
      const testData = [[
        [
          { value: '=TODAY()+1', excelFormat: 'yyyy/mm/dd' }  // フォーマットを指定
        ]
      ]];

      const result = engine.processData(testData);
      const formatted = result.tables[0][0][0];
      const tomorrowDate = new Date();
      tomorrowDate.setDate(tomorrowDate.getDate() + 1);
      const expectedFormat = `${tomorrowDate.getFullYear()}/${String(tomorrowDate.getMonth() + 1).padStart(2, '0')}/${String(tomorrowDate.getDate()).padStart(2, '0')}`;

      expect(formatted.displayValue).toBe(expectedFormat);
    });
  });

  describe('DATE Function', () => {
    test('should create date as Excel serial number', () => {
      const result = engine.evaluateFormula('=DATE(2024,1,15)', testData, 0);
      const date = new Date(2024, 0, 15);
      const excelBaseDate = new Date(1900, 0, 1);
      const expectedSerial = Math.floor((date - excelBaseDate) / (1000 * 60 * 60 * 24)) + 1;
      
      expect(typeof result).toBe('number');
      expect(result).toBe(expectedSerial);
    });

    test('should handle year between 0 and 1899', () => {
      const result = engine.evaluateFormula('=DATE(99,1,1)', testData, 0);
      const date = new Date(1999, 0, 1);
      const excelBaseDate = new Date(1900, 0, 1);
      const expectedSerial = Math.floor((date - excelBaseDate) / (1000 * 60 * 60 * 24)) + 1;
      
      expect(result).toBe(expectedSerial);
    });

    test('should handle year between 1900 and 9999', () => {
      const result = engine.evaluateFormula('=DATE(1901,1,1)', testData, 0);
      const date = new Date(1901, 0, 1);
      const excelBaseDate = new Date(1900, 0, 1);
      const expectedSerial = Math.floor((date - excelBaseDate) / (1000 * 60 * 60 * 24)) + 1;
      
      expect(result).toBe(expectedSerial);
    });

    test('should handle month overflow', () => {
      const result = engine.evaluateFormula('=DATE(2024,13,1)', testData, 0);
      const date = new Date(2025, 0, 1); // 2024年13月1日 = 2025年1月1日
      const excelBaseDate = new Date(1900, 0, 1);
      const expectedSerial = Math.floor((date - excelBaseDate) / (1000 * 60 * 60 * 24)) + 1;
      
      expect(result).toBe(expectedSerial);
    });

    test('should handle invalid parameters', () => {
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
      const date = new Date(2024, 0, 15);
      const excelBaseDate = new Date(1900, 0, 1);
      const expectedSerial = Math.floor((date - excelBaseDate) / (1000 * 60 * 60 * 24)) + 1;
      
      expect(typeof result).toBe('number');
      expect(result).toBe(expectedSerial);
    });

    test('should handle string literals correctly', () => {
      expect(engine.evaluateFormula('=SUM(1,"2",3,"invalid")', testData, 0)).toBe(6);
      expect(engine.evaluateFormula('=AVERAGE(1,"2",3,"invalid")', testData, 0)).toBe(2);
      expect(engine.evaluateFormula('=COUNT(1,"2",3,"invalid")', testData, 0)).toBe(3);
      expect(engine.evaluateFormula('=MAX(1,"2",3,"invalid")', testData, 0)).toBe(3);
      expect(engine.evaluateFormula('=MIN(1,"2",3,"invalid")', testData, 0)).toBe(1);
      expect(engine.evaluateFormula('=MIN(-1,"2",3,"invalid")', testData, 0)).toBe(-1);
    });
  });
}); 