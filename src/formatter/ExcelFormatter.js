export class ExcelFormatter {
  constructor() {
    this.defaultFormat = 'General';
  }

  format(value, formatPattern) {
    let result = {
      displayValue: '',
      textColor: null
    };

    if (!formatPattern || formatPattern === 'General') {
      result.displayValue = this._formatGeneral(value);
      return result;
    }

    if (this._isError(value)) {
      result.displayValue = value.toString();
      return result;
    }

    try {
      if (formatPattern.includes(';')) {
        const formatted = this._handleConditionalFormat(value, formatPattern);
        result.displayValue = formatted;
        result.textColor = this.lastColor;
      } else {
        result.displayValue = this._applyFormat(value, formatPattern);
      }
      return result;
    } catch (error) {
      result.displayValue = value.toString();
      return result;
    }
  }

  _applyFormat(value, format) {
    if (typeof value !== 'number') {
      return value.toString();
    }

    const cleanFormat = format.trim();

    if (cleanFormat.includes('¥')) {
      return `¥${this._formatNumber(value, cleanFormat.replace('¥', ''))}`;
    }

    if (cleanFormat.match(/0\.0+E\+00/)) {
      return this._formatScientific(value);
    }

    if (cleanFormat.includes('%')) {
      return this._formatPercent(value, cleanFormat);
    }

    return this._formatNumber(value, cleanFormat);
  }

  _formatScientific(value) {
    const exponential = value.toExponential(2);
    const [mantissa, exponent] = exponential.split('e');
    const absExponent = Math.abs(parseInt(exponent)).toString().padStart(2, '0');
    return `${mantissa}E+${absExponent}`;
  }

  _formatNumber(value, format) {
    const isNegative = value < 0;
    const absValue = Math.abs(value);
    
    const useThousandSeparator = format.includes('#,##') || format.includes('0,00');
    
    const decimalPlaces = (format.split('.')[1] || '').replace(/[^0#]/g, '').length;
    
    let result = absValue.toFixed(decimalPlaces);
    
    if (useThousandSeparator) {
      const parts = result.split('.');
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
      result = parts.join('.');
    }

    return isNegative ? `-${result}` : result;
  }

  _formatPercent(value, format) {
    const percentValue = value * 100;
    const decimalPlaces = (format.match(/0\.?(0+)?%/)?.[1] || '').length || 0;
    return `${percentValue.toFixed(decimalPlaces)}%`;
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
    if (this._isError(value)) return value.toString();
    if (typeof value === 'number') {
      return value.toString();
    }
    return value.toString();
  }

  _isError(value) {
    return typeof value === 'string' && value.startsWith('#');
  }
} 