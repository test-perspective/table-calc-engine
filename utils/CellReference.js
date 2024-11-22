export class CellReference {
    static parse(reference) {
      const upperRef = reference.toUpperCase();
      const match = upperRef.match(/^([A-Z]+)(\d+)$/);
      if (!match) throw new Error(`Invalid cell reference: ${reference}`);
  
      const [, col, row] = match;
      return {
        column: this.columnToIndex(col),
        row: parseInt(row) - 1
      };
    }
  
    static parseRange(range) {
      const [start, end] = range.split(':');
      return {
        start: this.parse(start),
        end: this.parse(end)
      };
    }
  
    static columnToIndex(col) {
      return col.split('').reduce((acc, char) => 
        acc * 26 + char.charCodeAt(0) - 'A'.charCodeAt(0), 0);
    }
  
    static indexToColumn(index) {
      let column = '';
      let temp = index + 1;
      
      while (temp > 0) {
        temp--;
        column = String.fromCharCode(65 + (temp % 26)) + column;
        temp = Math.floor(temp / 26);
      }
      return column;
    }
  
    static columnToLetter(index) {
      return this.indexToColumn(index);
    }
}