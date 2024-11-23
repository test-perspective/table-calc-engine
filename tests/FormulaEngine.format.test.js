import { FormulaEngine } from '../src/FormulaEngine.js';

describe('FormulaEngine Format Tests', () => {
  let engine;

  beforeEach(() => {
    engine = new FormulaEngine();
  });

  test('should process simple table data with correct format', () => {
    const testData = [
      [
        [
          { value: 1, excelFormat: null },
          { value: 2, excelFormat: null },
          { value: "=SUM(2,3)", excelFormat: "###.##" }
        ]
      ]
    ];

    const result = engine.processData(testData);

    expect(result).toHaveProperty('tables');
    expect(result).toHaveProperty('formulas');
    expect(result.tables).toHaveLength(1);

    const cell = result.tables[0][0][0];
    expect(cell).toEqual({
      value: 1,
      resolvedValue: 1,
      displayValue: "1",
      excelFormat: null,
      resolved: true
    });

    const formulaCell = result.tables[0][0][2];
    expect(formulaCell).toEqual({
      value: "=SUM(2,3)",
      resolvedValue: 5,
      displayValue: "5.00",
      excelFormat: "###.##",
      resolved: true
    });
  });
}); 