import { FormulaEngine } from '../src/FormulaEngine';

describe('Cell Format Tests', () => {
  let engine;

  beforeEach(() => {
    engine = new FormulaEngine();
  });

  describe('Number Formats', () => {
    const testData = [[
      [
        { value: 1234.567, excelFormat: '#,##0.00' },
        { value: -1234.567, excelFormat: '#,##0.00' },
        { value: 1234.567, excelFormat: '$#,##0.00' },
        { value: 0.123, excelFormat: '0%' },
        { value: 1234.567, excelFormat: '0.00E+00' },
      ]
    ]];

    test('should format numbers with thousand separators and decimals', () => {
      const result = engine.processData(testData);
      expect(result.tables[0][0][0].displayValue).toBe('1,234.57');
    });

    test('should format negative numbers', () => {
      const result = engine.processData(testData);
      expect(result.tables[0][0][1].displayValue).toBe('-1,234.57');
    });

    test('should format currency', () => {
      const result = engine.processData(testData);
      expect(result.tables[0][0][2].displayValue).toBe('$1,234.57');
    });

    test('should format percentages', () => {
      const result = engine.processData(testData);
      expect(result.tables[0][0][3].displayValue).toBe('12%');
    });

    test('should format scientific notation', () => {
      const result = engine.processData(testData);
      expect(result.tables[0][0][4].displayValue).toBe('1.23E+03');
    });

    test('should round to specified decimal places with # format', () => {
      const testData = [[
        [
          { value: 1.2345, excelFormat: '#.#' },    // 1.2を期待
          { value: 1.2345, excelFormat: '#.##' },   // 1.23を期待
          { value: 1.2345, excelFormat: '#.0' },    // 1.2を期待
          { value: 1.2345, excelFormat: '#.00' },   // 1.23を期待
        ]
      ]];

      const result = engine.processData(testData);
      expect(result.tables[0][0][0].displayValue).toBe('1.2');
      expect(result.tables[0][0][1].displayValue).toBe('1.23');
      expect(result.tables[0][0][2].displayValue).toBe('1.2');
      expect(result.tables[0][0][3].displayValue).toBe('1.23');
    });
  });

  describe('Conditional Formats', () => {
    const testData = [[
      [
        { value: 1234, excelFormat: '[>1000]#,##0;"小さい"' },
        { value: 999, excelFormat: '[>1000]#,##0;"小さい"' },
        { value: -50, excelFormat: '[Red]#,##0;[Blue]-#,##0' },
      ]
    ]];

    test('should apply conditional format when condition is met', () => {
      const result = engine.processData(testData);
      expect(result.tables[0][0][0].displayValue).toBe('1,234');
    });

    test('should apply alternative format when condition is not met', () => {
      const result = engine.processData(testData);
      expect(result.tables[0][0][1].displayValue).toBe('小さい');
    });

    test('should apply color format based on value', () => {
      const result = engine.processData(testData);
      expect(result.tables[0][0][2].displayValue).toBe('-50');
      expect(result.tables[0][0][2].textColor).toBe('Blue');
    });
  });

  describe('Formula Results with Formats', () => {
    test('should format formula results', () => {
      const testData = [[
        [
          { value: '=10+20', excelFormat: '#.00' },
          { value: '=25*2', excelFormat: '0.00%' },
          { value: '=1+2', excelFormat: '@' }
        ]
      ]];
      
      const result = engine.processData(testData);
      expect(result.tables[0][0][0].displayValue).toBe('30.00');
      expect(result.tables[0][0][1].displayValue).toBe('50.00%');
      expect(result.tables[0][0][2].displayValue).toBe('=1+2');
    });
  });

  describe('Special Format Cases', () => {
    const testData = [[
      [
        { value: 0, excelFormat: '#,##0.00;-#,##0.00;"Zero"' },
        { value: '', excelFormat: '@' },
        { value: 1234.567, excelFormat: '#,##0.00' },
        { value: 1234.567, excelFormat: 'Number' },
        { value: 1234.567, excelFormat: '#,##0.00' }
      ]
    ]];

    test('should handle zero values with special format', () => {
      const result = engine.processData(testData);
      expect(result.tables[0][0][0].displayValue).toBe('Zero');
    });

    test('should handle empty values with text format', () => {
      const result = engine.processData(testData);
      expect(result.tables[0][0][1].displayValue).toBe('');
    });

    test('should handle locale-specific formats', () => {
      const result = engine.processData(testData);
      expect(result.tables[0][0][2].displayValue).toBe('1,234.57');
    });
  });

  describe('Error Handling', () => {
    const testData = [[
      [
        { value: '#DIV/0!', excelFormat: '#,##0.00' },
        { value: '=1/0', excelFormat: '#,##0.00' },
        { value: 'Invalid', excelFormat: '#,##0.00' },
      ]
    ]];

    test('should preserve error values regardless of format', () => {
      const result = engine.processData(testData);
      expect(result.tables[0][0][0].displayValue).toBe('#DIV/0!');
      expect(result.tables[0][0][1].displayValue).toBe('#DIV/0!');
    });

    test('should handle invalid numeric values', () => {
      const result = engine.processData(testData);
      expect(result.tables[0][0][2].displayValue).toBe('Invalid');
    });
  });

  describe('Predefined Format Names', () => {
    test('should handle predefined format names', () => {
      const testData = [[
        [
          { value: 1234.567, excelFormat: 'Number' },      // 1,234.57
          { value: 1234.567, excelFormat: 'Currency' },    // $1,234.57
          { value: 0.1234, excelFormat: 'Percentage' },    // 12.34%
          { value: 1234.567, excelFormat: 'Scientific' },  // 1.23E+03
          { value: "ABC", excelFormat: 'Text' },           // ABC
          { value: "=SUM(A1:B1)", excelFormat: 'Text' }    // =SUM(A1:B1) として表示
        ]
      ]];

      const result = engine.processData(testData);
      
      expect(result.tables[0][0][0].displayValue).toBe('1,234.57');
      expect(result.tables[0][0][1].displayValue).toBe('$1,234.57');
      expect(result.tables[0][0][2].displayValue).toBe('12.34%');
      expect(result.tables[0][0][3].displayValue).toBe('1.23E+03');
      expect(result.tables[0][0][4].displayValue).toBe('ABC');
      expect(result.tables[0][0][5].displayValue).toBe('=SUM(A1:B1)');  // 数式がそのまま表示される
    });

    test('should fallback to custom format if name is not predefined', () => {
      const testData = [[
        [
          { value: 1234.567, excelFormat: 'NotExist' },
          { value: 1234.567, excelFormat: '#,##0.00' }
        ]
      ]];

      const result = engine.processData(testData);
      
      expect(result.tables[0][0][0].displayValue).toBe('1234.567');
      expect(result.tables[0][0][1].displayValue).toBe('1,234.57');
    });

    test('should handle General format', () => {
      const testData = [[
        [
          { value: 1234.567, excelFormat: 'General' },
          { value: 0.000123, excelFormat: 'General' },
          { value: "ABC", excelFormat: 'General' }
        ]
      ]];

      const result = engine.processData(testData);
      
      expect(result.tables[0][0][0].displayValue).toBe('1234.567');
      expect(result.tables[0][0][1].displayValue).toBe('0.000123');
      expect(result.tables[0][0][2].displayValue).toBe('ABC');
    });
  });

  describe('Fraction Format', () => {
    test('should format decimal numbers as fractions', () => {
      const testData = [[
        [
          { value: 0.5, excelFormat: 'Fraction' },      // 1/2
          { value: 0.25, excelFormat: 'Fraction' },     // 1/4
          { value: 0.125, excelFormat: 'Fraction' },    // 1/8
          { value: 0.333333, excelFormat: 'Fraction' }, // 1/3
          { value: 1.25, excelFormat: 'Fraction' },     // 1 1/4
          { value: 2.5, excelFormat: 'Fraction' },      // 2 1/2
          { value: -0.5, excelFormat: 'Fraction' },     // -1/2
        ]
      ]];

      const result = engine.processData(testData);
      
      expect(result.tables[0][0][0].displayValue).toBe('1/2');
      expect(result.tables[0][0][1].displayValue).toBe('1/4');
      expect(result.tables[0][0][2].displayValue).toBe('1/8');
      expect(result.tables[0][0][3].displayValue).toBe('1/3');
      expect(result.tables[0][0][4].displayValue).toBe('1 1/4');
      expect(result.tables[0][0][5].displayValue).toBe('2 1/2');
      expect(result.tables[0][0][6].displayValue).toBe('-1/2');
    });

    test('should handle custom fraction formats', () => {
      const testData = [[
        [
          { value: 0.5, excelFormat: '# ?/2' },       // 分母が2
          { value: 0.33, excelFormat: '# ?/3' },      // 分母が3
          { value: 0.125, excelFormat: '# ?/8' },     // 分母が8
          { value: 0.167, excelFormat: '# ??/??' },   // 自動で最適な分数
          { value: 2.5, excelFormat: '# ?/2' },       // 2 1/2
        ]
      ]];

      const result = engine.processData(testData);
      
      expect(result.tables[0][0][0].displayValue).toBe('1/2');
      expect(result.tables[0][0][1].displayValue).toBe('1/3');
      expect(result.tables[0][0][2].displayValue).toBe('1/8');
      expect(result.tables[0][0][3].displayValue).toBe('1/6');
      expect(result.tables[0][0][4].displayValue).toBe('2 1/2');
    });

    test('should handle edge cases in fraction format', () => {
      const testData = [[
        [
          { value: 0, excelFormat: 'Fraction' },          // 0
          { value: 1, excelFormat: 'Fraction' },          // 1
          { value: -1, excelFormat: 'Fraction' },         // -1
          { value: 0.000001, excelFormat: 'Fraction' },   // とても小さい数
          { value: 'invalid', excelFormat: 'Fraction' },  // 無効な入力
        ]
      ]];

      const result = engine.processData(testData);
      
      expect(result.tables[0][0][0].displayValue).toBe('0');
      expect(result.tables[0][0][1].displayValue).toBe('1');
      expect(result.tables[0][0][2].displayValue).toBe('-1');
      expect(result.tables[0][0][3].displayValue).toBe('0');  // 小さすぎる数は0として表示
      expect(result.tables[0][0][4].displayValue).toBe('invalid');
    });

    test('should handle complex mixed numbers', () => {
      const testData = [[
        [
          { value: 3.75, excelFormat: 'Fraction' },     // 3 3/4
          { value: 2.6666, excelFormat: 'Fraction' },   // 2 2/3
          { value: -1.25, excelFormat: 'Fraction' },    // -1 1/4
          { value: 4.125, excelFormat: 'Fraction' },    // 4 1/8
        ]
      ]];

      const result = engine.processData(testData);
      
      expect(result.tables[0][0][0].displayValue).toBe('3 3/4');
      expect(result.tables[0][0][1].displayValue).toBe('2 2/3');
      expect(result.tables[0][0][2].displayValue).toBe('-1 1/4');
      expect(result.tables[0][0][3].displayValue).toBe('4 1/8');
    });
  });
}); 