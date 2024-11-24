export class FormulaParser {
  parse(formula) {
    if (!formula.startsWith('=')) {
      throw new Error('Formula must start with =');
    }

    const cleanFormula = formula.substring(1);
    
    // 関数呼び出しを最初にチェック
    if (this.isFunction(cleanFormula)) {
      return this.parseFunction(cleanFormula);
    }

    // セル参照や範囲の処理
    if (this.isRange(cleanFormula)) {
      return {
        type: 'range',
        reference: this.normalizeReference(cleanFormula)
      };
    }

    if (this.isCellReference(cleanFormula)) {
      return {
        type: 'cell',
        reference: this.normalizeReference(cleanFormula)
      };
    }

    // 算術式の場合、セル参照を正規化してから処理
    return this.parseArithmetic(this.normalizeArithmetic(cleanFormula));
  }

  normalizeReference(ref) {
    // $A$1 -> $A$1, a1 -> A1, $a$1 -> $A$1 のように正規化
    return ref.replace(/[a-z]+/g, match => match.toUpperCase());
  }

  normalizeArithmetic(formula) {
    // 算術式内のセル参照を正規化
    return formula.replace(/(\$?[a-zA-Z]+\$?\d+)/g, match => this.normalizeReference(match));
  }

  isCellReference(value) {
    return /^\$?[a-zA-Z]+\$?\d+$/i.test(value);
  }

  isRange(value) {
    return /^\$?[a-zA-Z]+\$?\d+:\$?[a-zA-Z]+\$?\d+$/i.test(value);
  }

  parseFunction(formula) {
    const match = formula.match(/^([A-Za-z]+)\((.*)\)$/i);
    if (!match) {
      throw new Error('Invalid function call');
    }

    const [_, name, argsString] = match;
    // 関数名を大文字に正規化
    const normalizedName = name.toUpperCase();
    const args = this.parseFunctionArguments(argsString);

    return {
      type: 'function',
      name: normalizedName,
      arguments: args
    };
  }

  parseFunctionArguments(argsString) {
    if (!argsString.trim()) {
      return [];
    }

    const args = [];
    let currentArg = '';
    let parenCount = 0;

    for (let i = 0; i < argsString.length; i++) {
      const char = argsString[i];
      
      if (char === '(') {
        parenCount++;
        currentArg += char;
      } else if (char === ')') {
        parenCount--;
        currentArg += char;
      } else if (char === ',' && parenCount === 0) {
        args.push(this.parseArgument(currentArg.trim()));
        currentArg = '';
      } else {
        currentArg += char;
      }
    }

    if (currentArg) {
      args.push(this.parseArgument(currentArg.trim()));
    }

    return args;
  }

  parseArgument(arg) {
    if (this.isRange(arg)) {
      return {
        type: 'range',
        reference: arg
      };
    }

    if (this.isCellReference(arg)) {
      return {
        type: 'cell',
        reference: arg
      };
    }

    if (!isNaN(arg)) {
      return {
        type: 'literal',
        value: Number(arg)
      };
    }

    // 再帰的に解析（関数の入れ子に対応）
    return this.parse('=' + arg);
  }

  isFunction(value) {
    return /^[A-Za-z]+\(.*\)$/i.test(value);
  }

  parseArithmetic(formula) {
    // 演算子を含む式の処理
    const operators = ['+', '-', '*', '/', '^'];
    for (const operator of operators) {
      const parts = formula.split(operator);
      if (parts.length === 2) {
        const left = this.parse('=' + parts[0].trim());
        const right = this.parse('=' + parts[1].trim());
        return {
          type: 'operation',
          operator,
          left,
          right
        };
      }
    }

    // 数値リテラルの処理
    if (!isNaN(formula)) {
      return {
        type: 'literal',
        value: Number(formula)
      };
    }

    throw new Error(`Invalid arithmetic expression: ${formula}`);
  }
} 