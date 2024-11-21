# Excel Formula Engine

Excel-like formula calculation engine implemented in JavaScript.

## Features

- Cell reference support (A1 notation)
- Range reference support (e.g., A1:B3)
- Excel functions support:
  - SUM
  - AVERAGE
  - COUNT
  - MAX
  - MIN
  - and more...
- Number formatting
- Error handling

## Installation 
```bash
bash
npm install
```
## Usage

```javascript
import { FormulaEngine } from './src/FormulaEngine.js';
const engine = new FormulaEngine();
const data = {
tableData: [
[
{ value: "1", resolved: true, tableNumber: 0, macroId: null, excelFormat: null },
{ value: "2", resolved: true, tableNumber: 0, macroId: null, excelFormat: null }
],
[
{ value: "3", resolved: true, tableNumber: 0, macroId: null, excelFormat: null },
{ value: "4", resolved: true, tableNumber: 0, macroId: null, excelFormat: null }
]
],
formulas: [
{ value: "=SUM(A1:B2)", resolved: false, tableNumber: 0, macroId: "xxx", excelFormat: "###.##" }
]
};
const result = engine.processData(data);
console.log(result);
```

## Development Status

Current development focuses on:
- Basic arithmetic operations
- Core Excel functions
- Cell reference handling
- Number formatting

Upcoming features:
- More Excel functions
- Advanced error handling
- Performance optimizations
- Comprehensive testing

## License

MIT
