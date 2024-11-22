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

const testData2 = [
    [
        [
            {"value":"A","resolved":true,"macroId":null,"excelFormat":null},
            {"value":"B","resolved":true,"macroId":null,"excelFormat":null},
            {"value":"C","resolved":true,"macroId":null,"excelFormat":null}
        ],
        [
            {"value":"1","resolved":true,"macroId":null,"excelFormat":null},
            {"value":"2","resolved":true,"macroId":null,"excelFormat":null},
            {"value":"=sum(a2:b3)","resolved":false,"macroId":"331acef9-8841-401c-ba97-3c64a506865f","excelFormat":""}
        ],
        [
            {"value":"3","resolved":true,"macroId":null,"excelFormat":null},
            {"value":"4","resolved":true,"macroId":null,"excelFormat":null},
            {"value":"=sum(a1:b2)","resolved":false,"macroId":"bc35cfc0-e1ed-4624-84cf-17855bcfa9a7","excelFormat":"####.#"}                
              // {"value":"=1+2","resolved":false,"macroId":"331acef9-8841-401c-ba97-3c64a506865f","excelFormat":""}],[{"value":"3","resolved":true,"macroId":null,"excelFormat":null},{"value":"4","resolved":true,"macroId":null,"excelFormat":null},{"value":"=sum(a1:b2)","resolved":false,"macroId":"bc35cfc0-e1ed-4624-84cf-17855bcfa9a7","excelFormat":"####.#"}
        ]
    ]
]

const engine = new FormulaEngine();
const result = engine.processData(testData2);
console.log('Final result:', JSON.stringify(result, null, 2)); 