export class CellFormatter {
  constructor(formatter) {
    this.formatter = formatter;
  }

  processCell(cell, tableIndex, rowIndex, colIndex) {
    if (typeof cell.value === 'string' && !cell.value.startsWith('=')) {
      const cleanValue = cell.value.replace(/,/g, '');
      if (!isNaN(cleanValue) && cleanValue.trim() !== '') {
        cell.value = Number(cleanValue);
      }
    }
    
    const originalValue = cell.value;
    
    if (typeof originalValue === 'string' && originalValue.startsWith('=')) {
      const resolvedFormat = this.formatter._resolvePredefinedFormat(cell.excelFormat);

      if (resolvedFormat === '@') {
        cell.resolved = true;
        cell.resolvedValue = originalValue;
        this.applyFormat(cell, originalValue);
        return {
          table: tableIndex,
          row: rowIndex,
          col: colIndex,
          ...cell
        };
      }

      cell.resolved = true;
      cell.resolvedValue = this._convertToNumber(cell.resolvedValue);
      
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