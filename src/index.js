import { FormulaEngine } from './FormulaEngine.js';

const testData = {
  tableData: [
    [
      { value: "1", resolved: true, tableNumber: 0, macroId: null, excelFormat: null },
      { value: "2", resolved: true, tableNumber: 0, macroId: null, excelFormat: null },
      { value: "=SUM(A1:B2)", resolved: false, tableNumber: 0, macroId: "xxx", excelFormat: "###.##" }
    ],
    [
      { value: "3", resolved: true, tableNumber: 0, macroId: null, excelFormat: null },
      { value: "4", resolved: true, tableNumber: 0, macroId: null, excelFormat: null },
      { value: "=AVERAGE(A1:A2)", resolved: false, tableNumber: 0, macroId: "yyy", excelFormat: "###.##" }
    ]
  ],
  formulas: [
    { value: "=SUM(A1:B2)", resolved: false, tableNumber: 0, macroId: "xxx", excelFormat: "###.##" },
    { value: "=AVERAGE(A1:A2)", resolved: false, tableNumber: 0, macroId: "yyy", excelFormat: "###.##" }
  ]
};

const engine = new FormulaEngine();
const result = engine.processData(testData);
console.log('Final result:', JSON.stringify(result, null, 2)); 