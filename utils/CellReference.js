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
  
    static columnToIndex(col) {
      return col.split('').reduce((acc, char) => 
        acc * 26 + char.charCodeAt(0) - 'A'.charCodeAt(0) + 1, 0) - 1;
    }
  
    static parseRange(range) {
      const upperRange = range.toUpperCase();
      const [start, end] = upperRange.split(':');
      return {
        start: this.parse(start),
        end: this.parse(end)
      };
    }

    static columnToLetter(column) {
      let temp = column + 1;
      let letter = '';
      while (temp > 0) {
        temp--;
        letter = String.fromCharCode(65 + (temp % 26)) + letter;
        temp = Math.floor(temp / 26);
      }
      return letter;
    }
}