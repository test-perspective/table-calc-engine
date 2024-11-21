export class FormulaParser {
  constructor() {
    this.tokens = [];
    this.position = 0;
  }

  parse(formula) {
    // 関数呼び出しのパターンをチェック
    const functionMatch = formula.match(/^([A-Z]+)\((.*)\)$/i);
    if (functionMatch) {
      const [_, functionName, args] = functionMatch;
      
      // 引数を解析
      const parsedArgs = args.split(',').map(arg => {
        const trimmedArg = arg.trim();
        // セル範囲の参照をチェック
        if (trimmedArg.includes(':')) {
          return {
            type: 'range',
            reference: trimmedArg
          };
        }
        // 単一セルの参照をチェック
        if (/^[A-Z]+\d+$/i.test(trimmedArg)) {
          return {
            type: 'cell',
            reference: trimmedArg
          };
        }
        // 数値をチェック
        if (!isNaN(trimmedArg)) {
          return {
            type: 'number',
            value: Number(trimmedArg)
          };
        }
        // その他は文字列として扱う
        return {
          type: 'literal',
          value: trimmedArg
        };
      });

      return {
        type: 'function',
        name: functionName.toUpperCase(),
        arguments: parsedArgs
      };
    }

    // 関数呼び出しでない場合
    if (!isNaN(formula)) {
      return {
        type: 'number',
        value: Number(formula)
      };
    }

    return {
      type: 'literal',
      value: formula
    };
  }
} 