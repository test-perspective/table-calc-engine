class CellReference {
    static parse(reference) {
      const match = reference.match(/^(\$?)([A-Z]+)(\$?)(\d+)$/);
      if (!match) throw new Error('Invalid cell reference');
  
      const [, colAbs, col, rowAbs, row] = match;
      return {
        column: this.columnToIndex(col),
        row: parseInt(row) - 1,
        isColAbsolute: colAbs === '$',
        isRowAbsolute: rowAbs === '$'
      };
    }
  
    static columnToIndex(col) {
      return col.split('').reduce((acc, char) => 
        acc * 26 + char.charCodeAt(0) - 'A'.charCodeAt(0) + 1, 0) - 1;
    }
  
    static parseRange(range) {
      const [start, end] = range.split(':');
      return {
        start: this.parse(start),
        end: this.parse(end)
      };
    }
  }