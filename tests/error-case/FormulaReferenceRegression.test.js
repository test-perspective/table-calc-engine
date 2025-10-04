import { FormulaEngine } from '../../src/FormulaEngine.js';

describe('Formula reference regression - directional and cross-table', () => {
  let engine;
  let testData;

  beforeEach(() => {
    engine = new FormulaEngine();

    // 2テーブル構成
    // table 0
    //       |  A |  B |   C
    //   1   |  1 |  2 | =A1+B1      (右方向参照)
    //   2   |  3 | =A1+C1 |  6      (上・左混在参照)
    //   3   | =A2+B1 |  5 |  9      (上方向参照と固定値)
    //
    // table 1
    //       |  A |  B
    //   1   | =0!A1 | =0!C1         (他テーブル参照)
    //   2   |  10  | =A1+B1         (同表内の式参照)
    testData = [
      [
        [ { value: 1 }, { value: 2 }, { value: '=A1+B1' } ],
        [ { value: 3 }, { value: '=A1+C1' }, { value: 6 } ],
        [ { value: '=A2+B1' }, { value: 5 }, { value: 9 } ]
      ],
      [
        [ { value: '=0!A1' }, { value: '=0!C1' } ],
        [ { value: 10 }, { value: '=A1+B1' } ]
      ]
    ];
  });

  test('rightward reference within same table (C1 = A1+B1 = 3)', () => {
    expect(engine.evaluateFormula('=C1', testData, 0)).toBe(3);
  });

  test('upward and left mixed reference (B2 = A1+C1 = 1+3 = 4)', () => {
    expect(engine.evaluateFormula('=B2', testData, 0)).toBe(4);
  });

  test('upward reference (A3 = A2+B1 = 3+2 = 5)', () => {
    expect(engine.evaluateFormula('=A3', testData, 0)).toBe(5);
  });

  test('cross-table references (table1 B1 = table0 C1 = 3)', () => {
    expect(engine.evaluateFormula('=1!B1', testData, 0)).toBe(3);
  });

  test('same-table formula chain (table1 B2 = A1+B1 = 1!A1 + 1!B1 = 1 + 3 = 4)', () => {
    expect(engine.evaluateFormula('=1!B2', testData, 0)).toBe(4);
  });
});


