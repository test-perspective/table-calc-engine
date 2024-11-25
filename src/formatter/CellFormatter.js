export class CellFormatter {
  constructor(formatter) {
    this.formatter = formatter;
  }

  processCell(cell, tableIndex, rowIndex, colIndex) {
    const originalValue = cell.value;
    
    // 数式の場合
    if (typeof originalValue === 'string' && originalValue.startsWith('=')) {
      cell.resolved = true;
      cell.resolvedValue = this._convertToNumber(cell.resolvedValue);
      
      // フォーマット用の値を計算
      const formatValue = cell.excelFormat && cell.excelFormat.includes('%') 
        ? cell.resolvedValue / 100 
        : cell.resolvedValue;

      this.applyFormat(cell, formatValue);
      
      return {
        table: tableIndex,
        row: rowIndex,
        col: colIndex,
        ...cell
      };
    } else {
      // 通常の値の場合
      cell.resolved = true;
      cell.resolvedValue = originalValue;
      
      this.applyFormat(cell, originalValue);
      return null;
    }
  }

  applyFormat(cell, value) {
    if (cell.excelFormat) {
      const formatted = this.formatter.format(value, cell.excelFormat);
      cell.displayValue = formatted.displayValue;
      if (formatted.textColor) {
        cell.textColor = formatted.textColor;
      }
    } else {
      cell.displayValue = String(value);
    }
  }

  _convertToNumber(value) {
    if (typeof value === 'string' && !isNaN(value) && value.trim() !== '') {
      return Number(value);
    }
    return value;
  }
} 