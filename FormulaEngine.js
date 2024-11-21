class FormulaEngine {
    constructor() {
      this.parser = new FormulaParser();
      this.cache = new Map();
      this.functions = ExcelFunctions;
    }
  
    processData(data) {
      try {
        const processedData = this.processTableData(data.tableData);
        const processedFormulas = this.processFormulas(data.formulas, processedData);
  
        return {
          tableData: processedData,
          formulas: processedFormulas
        };
      } catch (error) {
        console.error('Processing error:', error);
        return this.handleError(error);
      }
    }
  
    processTableData(tableData) {
      return tableData.map(row =>
        row.map(cell => {
          if (!cell.resolved) {
            const result = this.evaluateFormula(cell.value, tableData);
            return {
              ...cell,
              value: this.formatValue(result, cell.excelFormat),
              resolved: true
            };
          }
          return cell;
        })
      );
    }
  
    evaluateFormula(formula, tableData) {
      const cacheKey = `${formula}_${JSON.stringify(tableData)}`;
      if (this.cache.has(cacheKey)) {
        return this.cache.get(cacheKey);
      }
  
      const ast = this.parser.parse(formula);
      const result = this.evaluateAst(ast, tableData);
      
      this.cache.set(cacheKey, result);
      return result;
    }
  
    formatValue(value, format) {
      if (!format || !value) return value;
  
      try {
        if (typeof value === 'number') {
          return new Intl.NumberFormat('ja-JP', {
            minimumFractionDigits: this.getDecimalPlaces(format),
            maximumFractionDigits: this.getDecimalPlaces(format)
          }).format(value);
        }
  
        if (value instanceof Date) {
          return new Intl.DateTimeFormat('ja-JP', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit'
          }).format(value);
        }
  
        return value;
      } catch (error) {
        return '#FORMAT_ERROR';
      }
    }
  
    getDecimalPlaces(format) {
      const match = format.match(/\.(\d+)/);
      return match ? match[1].length : 0;
    }
  }