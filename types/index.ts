interface CellData {
    value: string | number | null;
    resolved: boolean;
    tableNumber: number;
    macroId: string | null;
    excelFormat: string | null;
  }
  
  interface Formula {
    value: string;
    resolved: boolean;
    tableNumber: number;
    macroId: string;
    excelFormat: string | null;
  }
  
  interface ProcessingResult {
    tableData: CellData[][];
    formulas: Formula[];
  } 