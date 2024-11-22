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

  parseArithmetic(expression) {
    // 単純な数値の場合
    if (!isNaN(expression)) {
      return {
        type: 'literal',
        value: Number(expression)
      };
    }

    // 演算子を探す
    const operators = ['+', '-', '*', '/', '^'];
    let operator = null;
    let splitIndex = -1;

    for (const op of operators) {
      splitIndex = expression.lastIndexOf(op);
      if (splitIndex !== -1) {
        operator = op;
        break;
      }
    }

    if (splitIndex === -1) {
      throw new Error(`Invalid arithmetic expression: ${expression}`);
    }

    const left = expression.substring(0, splitIndex).trim();
    const right = expression.substring(splitIndex + 1).trim();

    return {
      type: 'operation',
      operator: operator,
      left: this.parse(left),
      right: this.parse(right)
    };
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
    
    // その他の式
    return this.parseArithmetic(arg);
  }
} 