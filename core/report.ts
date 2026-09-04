export type Box = { top: number; bottom: number; left: number; right: number };
export type Finding = {
  id: string;
  title: string;
  evidence: string;
  action: string;
  level: string;
};
export type SheetReport = {
  name: string;
  state: string;
  supported: boolean;
  reason?: string;
  cells?: number;
  populatedCount?: number;
  formulaCount?: number;
  ranges?: Record<
    'dimension' | 'storedCells' | 'populated' | 'candidates' | 'drawingBounds',
    Box | null
  >;
  printArea?: { expression: string | null; ranges: Box[]; resolved: boolean };
  pageSetup?: {
    orientation: string;
    paperSize: string | null;
    scale: string | null;
    fitToPage: boolean;
    fitToWidth: string | null;
    fitToHeight: string | null;
    errors: string;
  };
  manualBreaks?: {
    rows: number[];
    columns: number[];
    rowCount: number;
    columnCount: number;
  };
  drawingCount?: number;
  findings?: Finding[];
  warnings?: string[];
};
export type WorkbookReport = {
  version: number;
  sheets: SheetReport[];
  storedCellCount: number;
  findingCount: number;
  limitations: string[];
};
