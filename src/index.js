import { FormulaEngine } from './FormulaEngine.js';

const testData = [
  // テーブル0
  [
    [
      { value: "1", excelFormat: null },
      { value: "2", excelFormat: null },
      { value: "=SUM(A1:B2)", excelFormat: "###.##" }
    ],
    [
      { value: "3", excelFormat: null },
      { value: "4", excelFormat: null },
      { value: "=AVERAGE(A1:A2)", excelFormat: "###.##" }
    ]
  ],
  // テーブル1（他のテーブルの参照テスト用）
  [
    [
      { value: "10", excelFormat: null },
      { value: "=0!A1", excelFormat: null },  // テーブル0のA1を参照
      { value: "=SUM(0!A1:B2)", excelFormat: "###.##" }  // テーブル0の範囲を参照
    ]
  ]
];

const engine = new FormulaEngine();
const result = engine.processData(testData);
console.log('Final result:', JSON.stringify(result, null, 2)); 