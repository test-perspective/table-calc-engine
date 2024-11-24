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
          { value: '=25*2', excelFormat: '0.00%' }
        ]
      ]];
      
      const result = engine.processData(testData);
      expect(result.tables[0][0][0].displayValue).toBe('30.00');
      expect(result.tables[0][0][1].displayValue).toBe('50.00%');
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
          { value: "ABC", excelFormat: 'Text' }            // ABC
        ]
      ]];

      const result = engine.processData(testData);
      
      expect(result.tables[0][0][0].displayValue).toBe('1,234.57');
      expect(result.tables[0][0][1].displayValue).toBe('$1,234.57');
      expect(result.tables[0][0][2].displayValue).toBe('12.34%');
      expect(result.tables[0][0][3].displayValue).toBe('1.23E+03');
      expect(result.tables[0][0][4].displayValue).toBe('ABC');
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
}); 