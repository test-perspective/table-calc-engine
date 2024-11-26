export class FormulaParser {
  parse(formula) {
    if (!formula.startsWith('=')) {
      throw new Error('Formula must start with =');
    }

    const cleanFormula = formula.substring(1);
    return this.parseExpression(cleanFormula);
  }

  isTableReference(value) {
    return /^\d+!\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?$/i.test(value);
  }

  isCellReference(value) {
    return /^\$?[A-Za-z]+\$?\d+$/i.test(value);
  }

  isRange(value) {
    return /^\$?[A-Za-z]+\$?\d+:\$?[A-Za-z]+\$?\d+$/i.test(value);
  }

  isFunction(value) {
    return /^[A-Za-z]+\s*\(.*\)$/i.test(value);
  }

  parseFunctionArguments(argsString) {
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
        if (currentArg.trim()) {
          args.push(this.parseExpression(currentArg.trim()));
        }
        currentArg = '';
      } else {
        currentArg += char;
      }
    }

    if (currentArg.trim()) {
      args.push(this.parseExpression(currentArg.trim()));
    }

    return args;
  }

  parseTableReference(formula) {
    const [tableId, reference] = formula.split('!');
    
    if (reference.includes(':')) {
      return {
        type: 'range',
        tableId: parseInt(tableId),
        reference: this.normalizeReference(reference)
      };
    }

    return {
      type: 'cell',
      tableId: parseInt(tableId),
      reference: this.normalizeReference(reference)
    };
  }

  normalizeReference(ref) {
    return ref.replace(/[a-z]+/gi, match => match.toUpperCase());
  }

  parseExpression(formula) {
    const operators = ['+', '-', '*', '/', '^'];
    let lowestPrecedenceOp = null;
    let lowestPrecedenceIndex = -1;
    let parenCount = 0;

    for (let i = formula.length - 1; i >= 0; i--) {
      const char = formula[i];
      if (char === ')') parenCount++;
      else if (char === '(') parenCount--;
      else if (parenCount === 0 && operators.includes(char)) {
        if (lowestPrecedenceOp === null || 
            operators.indexOf(char) <= operators.indexOf(lowestPrecedenceOp)) {
          lowestPrecedenceOp = char;
          lowestPrecedenceIndex = i;
        }
      }
    }

    if (lowestPrecedenceOp) {
      const left = formula.substring(0, lowestPrecedenceIndex).trim();
      const right = formula.substring(lowestPrecedenceIndex + 1).trim();
      return {
        type: 'operation',
        operator: lowestPrecedenceOp,
        left: this.parseExpression(left),
        right: this.parseExpression(right)
      };
    }

    if (formula.startsWith('(') && formula.endsWith(')')) {
      return this.parseExpression(formula.slice(1, -1).trim());
    }

    if (this.isFunction(formula)) {
      return this.parseFunction(formula);
    }

    if (this.isTableReference(formula)) {
      return this.parseTableReference(formula);
    }

    if (this.isRange(formula)) {
      return {
        type: 'range',
        reference: this.normalizeReference(formula)
      };
    }

    if (this.isCellReference(formula)) {
      return {
        type: 'cell',
        reference: this.normalizeReference(formula)
      };
    }

    if (formula.startsWith('"') && formula.endsWith('"')) {
      return {
        type: 'literal',
        value: formula.slice(1, -1)
      };
    }

    if (!isNaN(formula)) {
      return {
        type: 'literal',
        value: Number(formula)
      };
    }

    throw new Error(`Invalid expression: ${formula}`);
  }

  parseFunction(formula) {
    const match = formula.match(/^([A-Za-z]+)\s*\((.*)\)$/i);
    if (!match) {
      throw new Error('Invalid function call');
    }

    const [_, name, argsString] = match;
    return {
      type: 'function',
      name: name.toUpperCase(),
      arguments: this.parseFunctionArguments(argsString)
    };
  }
} 