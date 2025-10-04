import { FormulaEngine } from '../../src/FormulaEngine.js';

describe('Error Case - compute D3 and D4', () => {
  let engine;
  let testData;

  beforeEach(() => {
    engine = new FormulaEngine();

    // テーブル（1枚）: A〜D 列、1〜3 行（D1 は空白）
    //       |  A |  B |  C |     D
    //   1   |  1 |  2 |  3 |  (blank)
    //   2   |  4 |  5 |  6 |  =d4+1
    //   3   |  7 |  8 |  9 |  =a2+b3
    //
    // 期待値:
    // - D3 = A2 + B3 = 4 + 8 = 12
    // - D2 = D4 + 1 = 13
    testData = [[
      [ { value: undefined }, { value: undefined }, { value: undefined }, { value: undefined } ],
      [ { value: 4 }, { value: 5 }, { value: 6 }, { value: '=d4+1' } ],
      [ { value: 7 }, { value: 8 }, { value: 9 } ],
      [ { value: undefined }, { value: undefined }, { value: undefined }, { value: '=a2+b3' } ]
    ]];
  });

  test('D3 should evaluate A2+B3 and return 12', () => {
    const result = engine.evaluateFormula('=a2+b3', testData, 0);
    expect(result).toBe(12);
  });

  test('Expression referencing D4 should return 13', () => {
    const result = engine.evaluateFormula('=d4+1', testData, 0);
    expect(result).toBe(13);
  });
});


