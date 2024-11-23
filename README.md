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

## Installation 
```bash
npm install excel-formula-engine
```

## Usage

```javascript
import { FormulaEngine } from 'excel-formula-engine';

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
- Range references (e.g., A1:B3)
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
