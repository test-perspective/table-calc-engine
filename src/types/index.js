export const CellDataType = {
  NUMBER: 'number',
  STRING: 'string',
  BOOLEAN: 'boolean',
  DATE: 'date',
  ERROR: 'error'
};

export class FormulaError extends Error {
  constructor(type, message) {
    super(message);
    this.type = type;
  }
} 