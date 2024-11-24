# Excel Formula Engine

Excel-like formula calculation engine implemented in JavaScript.

## Features

- Cell reference support (A1 notation)
- Basic arithmetic operations (+, -, *, /, ^)
- Excel functions support:
  - SUM - Adds numbers
  - AVERAGE - Calculates arithmetic mean
  - COUNT - Counts numeric values
  - MAX - Returns largest value
  - MIN - Returns smallest value
  - More functions coming soon...
- Number formatting (e.g., ###.##)
- Error handling (#REF!, #NAME?, #DIV/0!, #ERROR!, #CIRCULAR!)
- Cross-table references
  - Reference cells between different tables (e.g., 0!A1, 1!B2)
  - Support range references across tables (e.g., 1!A1:B2)
  - Case-insensitive references (1!a1 equals 1!A1)
  - Absolute references with $ symbol (e.g., 1!$A$1)
  - Complex calculations across tables (e.g., SUM(0!A1:B2) + SUM(1!A1:B2))

## Cross-Table References

The engine supports referencing cells and ranges between different tables using table IDs.

```javascript
const data = [
  [ // table 0
    [
      { value: 1 },
      { value: 2 }
    ]
  ],
  [ // table 1
    [
      { value: 10 },
      { value: 20 }
    ]
  ]
];

const engine = new FormulaEngine();

// Single cell references
engine.evaluateFormula('=0!A1', data, 0);  // 1
engine.evaluateFormula('=1!A1', data, 0);  // 10

// Range references
engine.evaluateFormula('=SUM(0!A1:B1)', data, 0);  // 3
engine.evaluateFormula('=SUM(1!A1:B1)', data, 0);  // 30

// Complex calculations
engine.evaluateFormula('=0!A1 + 1!A1', data, 0);  // 11
engine.evaluateFormula('=SUM(0!A1:B1) + AVERAGE(1!A1:B1)', data, 0);  // 18
```

Error handling for cross-table references:
- #REF! - Invalid table ID or cell reference out of range
- #ERROR! - Invalid table reference format

## Installation 
```bash
npm install table-calc-engine
```

## Usage

```javascript
import { FormulaEngine } from 'table-calc-engine';

const engine = new FormulaEngine();

// Input data structure (3D array: tables -> sheets -> cells)
const data = [
  [  // table
    [  // sheet
      { value: 1, excelFormat: null },
      { value: 2, excelFormat: null },
      { value: "=SUM(A1:B1)", excelFormat: "###.##" }
    ]
  ]
];

const result = engine.processData(data);
console.log(result);

// Result structure
{
  tables: [
    [
      [
        {
          value: 1,
          resolvedValue: 1,      // Maintains original type
          displayValue: "1",     // Always string
          excelFormat: null,
          resolved: true,
          // ... other original cell properties
        }
      ]
    ]
  ],
  formulas: [
    {
      tableIndex: 0,
      sheetIndex: 0,
      colIndex: 2,
      formula: "=SUM(A1:B1)",
      resolvedValue: 3,         // Maintains original type
      displayValue: "3.00",     // Formatted string
      excelFormat: "###.##",
      // Original cell information
      macroId: "macro1",
      style: { /* ... */ },
      metadata: { /* ... */ },
      validation: { /* ... */ },
      comment: "..."
    }
  ]
}
```

## Development Status

Current implementation:
- Basic arithmetic operations (+, -, *, /, ^)
- Cell references (A1 notation)
- Basic functions (SUM, AVERAGE, COUNT, MAX, MIN)
- Basic number formatting
- Error handling (#REF!, #NAME?, #DIV/0!, #ERROR!, #CIRCULAR!)
- Circular reference detection

Upcoming features:
- More Excel functions
- Advanced formatting options
- Performance optimizations
- Comprehensive testing

## Error Handling

The engine handles various error cases:
- #REF! - Invalid cell reference
- #NAME? - Unknown function name
- #DIV/0! - Division by zero
- #ERROR! - Other calculation errors
- #CIRCULAR! - Circular reference

## License

MIT
