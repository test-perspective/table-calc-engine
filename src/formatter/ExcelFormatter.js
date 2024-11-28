export class ExcelFormatter {
  constructor() {
    this.predefinedFormats = {
      'Number': '#,##0.00',
      'Currency': '$#,##0.00',
      'Date': 'yyyy/mm/dd',
      'Time': 'hh:mm:ss',
      'Percentage': '0.00%',
      'Fraction': '# ??/??',
      'Scientific': '0.00E+00',
      'Text': '@'
    };
  }

  format(value, formatPattern) {
    let result = {
      displayValue: '',
      textColor: null
    };

    const actualFormat = this._resolvePredefinedFormat(formatPattern);

    if (!actualFormat || actualFormat === 'General') {
      result.displayValue = this._formatGeneral(value);
      return result;
    }

    if (this._isError(value)) {
      result.displayValue = value.toString();
      return result;
    }

    try {
      if (actualFormat.includes(';')) {
        const formatted = this._handleConditionalFormat(value, actualFormat);
        result.displayValue = formatted;
        result.textColor = this.lastColor;
      } else {
        result.displayValue = this._applyFormat(value, actualFormat);
      }
      return result;
    } catch (error) {
      result.displayValue = value.toString();
      return result;
    }
  }

  _applyFormat(value, format) {
    if (value === '' || value === null || value === undefined) {
      return '';
    }

    const numValue = Number(value);
    if (isNaN(numValue) && format !== '@') return String(value);

    if (format.includes('?/')) {
      return this._formatAsFraction(numValue, format);
    }

    if (format === '@') {
      return value === 0 ? '' : String(value);
    }

    if (format.includes('E+00')) {
      const [mantissa, exponent] = numValue.toExponential(2).split('e');
      const expNum = Number(exponent);
      return `${Number(mantissa).toFixed(2)}E${expNum >= 0 ? '+' : ''}${String(Math.abs(expNum)).padStart(2, '0')}`;
    }

    if (format.includes('%')) {
      const percentValue = numValue * 100;
      const decimalPlaces = (format.match(/0\.?(0+)?%/)?.[1] || '').length || 0;
      return `${percentValue.toFixed(decimalPlaces)}%`;
    }

    if (format.includes('$')) {
      const places = (format.match(/0\.?(0+)?/)?.[1] || '').length || 0;
      const formattedNum = this._formatWithThousandSeparator(numValue.toFixed(places));
      return `$${formattedNum}`;
    }

    if (format.includes('#') || format.includes('0')) {
      const decimalPlaces = (format.match(/[#0]\.([#0]+)/)?.[1] || '').length || 0;
      const useThousandSeparator = format.includes(',');
      const fixedNum = numValue.toFixed(decimalPlaces);
      return useThousandSeparator ? this._formatWithThousandSeparator(fixedNum) : fixedNum;
    }

    if (format === 'yyyy/mm/dd') {
      const excelBaseDate = new Date(1900, 0, 1);
      const targetDate = new Date(excelBaseDate.getTime() + (numValue - 1) * 24 * 60 * 60 * 1000);
      
      const year = targetDate.getFullYear();
      const month = String(targetDate.getMonth() + 1).padStart(2, '0');
      const day = String(targetDate.getDate()).padStart(2, '0');
      
      return `${year}/${month}/${day}`;
    }

    return this._formatGeneral(value);
  }

  _handleConditionalFormat(value, format) {
    const sections = format.split(';');
    this.lastColor = null;
    
    if (value === 0 && sections.length >= 3) {
      return this._cleanFormatString(sections[2]);
    }

    for (const section of sections) {
      if (section.startsWith('[')) {
        const match = section.match(/\[(.*?)\](.*)/);
        if (match) {
          const [_, condition, formatPattern] = match;
          if (this._evaluateCondition(value, condition)) {
            if (condition === 'Red' || condition === 'Blue') {
              this.lastColor = condition;
            }
            return this._applyFormat(value, formatPattern);
          }
        }
      }
    }

    const defaultSection = sections.find(s => !s.startsWith('[')) || sections[0];
    return this._cleanFormatString(defaultSection);
  }

  _cleanFormatString(formatString) {
    return formatString.replace(/"/g, '');
  }

  _evaluateCondition(value, condition) {
    if (condition === 'Blue') {
      return value < 0;
    }
    if (condition === 'Red') {
      return value > 0;
    }
    if (condition.startsWith('>')) {
      return value > parseFloat(condition.substring(1));
    }
    if (condition.startsWith('<')) {
      return value < parseFloat(condition.substring(1));
    }
    return false;
  }

  _formatGeneral(value) {
    if (value === '' || value === null || value === undefined) {
      return '';
    }
    const numValue = Number(value);
    return isNaN(numValue) ? String(value) : String(numValue);
  }

  _isError(value) {
    return typeof value === 'string' && value.startsWith('#');
  }

  _resolvePredefinedFormat(formatPattern) {
    const resolved = this.predefinedFormats[formatPattern] || formatPattern;
    return resolved;
  }

  _formatWithThousandSeparator(numStr) {
    const [intPart, decPart] = numStr.split('.');
    const formattedInt = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return decPart ? `${formattedInt}.${decPart}` : formattedInt;
  }

  _formatAsFraction(value, format) {
    if (isNaN(value)) return String(value);
    if (Math.abs(value) < 0.000001) return '0';

    const isNegative = value < 0;
    const absValue = Math.abs(value);
    const wholePart = Math.floor(absValue);
    const fractionalPart = absValue - wholePart;

    if (fractionalPart < 0.000001) {
      return isNegative ? `-${wholePart}` : String(wholePart);
    }

    const denominatorMatch = format.match(/\?\/(\d+)/);
    
    if (denominatorMatch) {
      const denominator = parseInt(denominatorMatch[1]);
      const numerator = Math.round(fractionalPart * denominator);
      
      if (numerator === 0) {
        return isNegative ? `-${wholePart}` : String(wholePart);
      }
      
      if (wholePart === 0) {
        const result = `${numerator}/${denominator}`;
        return isNegative ? `-${result}` : result;
      } else {
        return `${isNegative ? '-' : ''}${wholePart} ${numerator}/${denominator}`;
      }
    } else {
      const commonDenominators = [2, 3, 4, 5, 6, 8, 10, 12, 16];
      let bestError = Number.MAX_VALUE;
      let bestFraction = null;

      for (const denominator of commonDenominators) {
        const numerator = Math.round(fractionalPart * denominator);
        const error = Math.abs(fractionalPart - numerator / denominator);

        if (error < bestError) {
          bestError = error;
          bestFraction = { numerator, denominator };
        }
      }

      if (bestFraction && bestFraction.numerator !== 0) {
        if (wholePart === 0) {
          const result = `${bestFraction.numerator}/${bestFraction.denominator}`;
          return isNegative ? `-${result}` : result;
        } else {
          return `${isNegative ? '-' : ''}${wholePart} ${bestFraction.numerator}/${bestFraction.denominator}`;
        }
      }
      
      return isNegative ? `-${wholePart}` : String(wholePart);
    }
  }
} 