export class FormulaParser {
  constructor() {
    this.tokens = [];
    this.position = 0;
  }

  parse(formula) {
    // 数式から'='を除去して処理
    const expression = formula.startsWith('=') ? formula.substring(1) : formula;
    
    // 単純な算術演算の場合
    if (expression.match(/^[\d\s+\-*/^()]+$/)) {
      const tokens = this.tokenizeArithmetic(expression);
      return this.parseArithmetic(tokens);
    }

    // その他の場合（関数呼び出しなど）
    return this.parseFunction(expression);
  }

  tokenizeArithmetic(expression) {
    const tokens = [];
    let current = '';
    
    for (let i = 0; i < expression.length; i++) {
      const char = expression[i];
      
      if ('+-*/^()'.includes(char)) {
        if (current.trim()) {
          tokens.push({ type: 'number', value: Number(current.trim()) });
          current = '';
        }
        tokens.push({ type: 'operator', value: char });
      } else if (!char.trim()) {
        if (current.trim()) {
          tokens.push({ type: 'number', value: Number(current.trim()) });
          current = '';
        }
      } else {
        current += char;
      }
    }
    
    if (current.trim()) {
      tokens.push({ type: 'number', value: Number(current.trim()) });
    }
    
    return tokens;
  }

  parseArithmetic(tokens) {
    if (tokens.length === 1) {
      return {
        type: 'literal',
        value: tokens[0].value
      };
    }

    // 基本的な算術演算の解析
    let left = {
      type: 'literal',
      value: tokens[0].value
    };

    for (let i = 1; i < tokens.length; i += 2) {
      const operator = tokens[i].value;
      const right = {
        type: 'literal',
        value: tokens[i + 1].value
      };

      left = {
        type: 'operation',
        operator: operator,
        left: left,
        right: right
      };
    }

    return left;
  }

  parseFunction(expression) {
    // 関数呼び出しのパターンをチェック
    const functionMatch = expression.match(/^([A-Z]+)\((.*)\)$/i);
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
    if (!isNaN(expression)) {
      return {
        type: 'number',
        value: Number(expression)
      };
    }

    return {
      type: 'literal',
      value: expression
    };
  }

  tokenize(formula) {
    const expression = formula.substring(1);
    const tokens = [];
    let current = '';
    
    // 演算子リストを拡張
    const operators = '+-*/^&=<>≤≥(),:';
    
    for (let i = 0; i < expression.length; i++) {
      const char = expression[i];
      
      // 複合演算子の処理 (<=, >=, <>, =)
      if (i < expression.length - 1) {
        const nextChar = expression[i + 1];
        const twoChars = char + nextChar;
        if (['<=', '>=', '<>', '='].includes(twoChars)) {
          if (current) {
            tokens.push(current.toUpperCase());
            current = '';
          }
          tokens.push(twoChars);
          i++; // 次の文字をスキップ
          continue;
        }
      }
      
      // 単一演算子の処理
      if (operators.includes(char)) {
        if (current) {
          tokens.push(current.toUpperCase());
          current = '';
        }
        tokens.push(char);
      } else {
        current += char;
      }
    }
    
    if (current) {
      tokens.push(current.toUpperCase());
    }
    return tokens;
  }

  parseExpression(tokens) {
    // ... 既存のコード ...

    // 比較演算子の処理を追加
    const operatorIndex = tokens.findIndex(t => 
      ['+-*/^&', '=', '<>', '<=', '>=', '<', '>'].some(op => op === t)
    );
    if (operatorIndex !== -1) {
      return {
        type: 'operation',
        operator: tokens[operatorIndex],
        left: this.parseExpression(tokens.slice(0, operatorIndex)),
        right: this.parseExpression(tokens.slice(operatorIndex + 1))
      };
    }

    // ... 既存のコード ...
  }
} 