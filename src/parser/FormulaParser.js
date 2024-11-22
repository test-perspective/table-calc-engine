export class FormulaParser {
  parse(formula) {
    const expression = formula.startsWith('=') ? formula.substring(1) : formula;
    
    // 関数呼び出しの場合（例：SUM(A1:B2)）
    const funcMatch = expression.match(/^(\w+)\((.*)\)$/);
    if (funcMatch) {
      const [, funcName, args] = funcMatch;
      return {
        type: 'function',
        name: funcName.toUpperCase(),
        arguments: args.split(',').map(arg => this.parseArgument(arg.trim()))
      };
    }
    
    // セル参照の場合（例：A1）
    if (expression.match(/^[A-Za-z]+[0-9]+$/)) {
      return {
        type: 'cell',
        reference: expression.toUpperCase()
      };
    }
    
    // 算術演算の場合
    return this.parseArithmetic(expression);
  }

  parseArgument(arg) {
    // 範囲参照（例：A1:B2）
    if (arg.includes(':')) {
      return {
        type: 'range',
        reference: arg.toUpperCase()
      };
    }
    
    // セル参照（例：A1）
    if (arg.match(/^[A-Za-z]+[0-9]+$/)) {
      return {
        type: 'cell',
        reference: arg.toUpperCase()
      };
    }
    
    // 数値
    if (!isNaN(arg)) {
      return {
        type: 'literal',
        value: Number(arg)
      };
    }
    
    throw new Error(`Invalid argument: ${arg}`);
  }

  parseArithmetic(expression) {
    // 単純な数値の場合
    if (!isNaN(expression)) {
      return {
        type: 'literal',
        value: Number(expression)
      };
    }
    
    // その他の算術演算は必要に応じて実装
    throw new Error(`Unsupported arithmetic expression: ${expression}`);
  }
} 