class FormulaParser {
    constructor() {
      this.tokens = [];
      this.position = 0;
    }
  
    parse(formula) {
      this.tokens = this.tokenize(formula);
      return this.parseExpression();
    }
  
    tokenize(formula) {
      // 数式をトークンに分解
      const tokens = [];
      let current = '';
      
      for (let i = 0; i < formula.length; i++) {
        const char = formula[i];
        
        if ('+-*/(),:'.includes(char)) {
          if (current) tokens.push(current);
          tokens.push(char);
          current = '';
        } else {
          current += char;
        }
      }
      
      if (current) tokens.push(current);
      return tokens;
    }
  
    parseExpression() {
      // 数式の構文解析
      const node = {
        type: 'expression',
        children: []
      };
  
      while (this.position < this.tokens.length) {
        const token = this.tokens[this.position];
        
        if (this.isFunction(token)) {
          node.children.push(this.parseFunction());
        } else if (this.isCellReference(token)) {
          node.children.push(this.parseCellReference());
        } else if (this.isNumber(token)) {
          node.children.push({ type: 'number', value: parseFloat(token) });
          this.position++;
        } else {
          this.position++;
        }
      }
  
      return node;
    }
  }