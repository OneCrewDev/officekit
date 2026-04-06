export interface ExcelCellModel {
  value: string;
  formula?: string;
  styleId?: string;
  type?: "string" | "number" | "boolean" | "date";
}

export interface ExcelWorkbookSettings {
  date1904?: boolean;
  codeName?: string;
  filterPrivacy?: boolean;
  showObjects?: string;
  backupFile?: boolean;
  dateCompatibility?: boolean;
  calcMode?: string;
  iterate?: boolean;
  iterateCount?: number;
  iterateDelta?: number;
  fullPrecision?: boolean;
  fullCalcOnLoad?: boolean;
  refMode?: string;
  lockStructure?: boolean;
  lockWindows?: boolean;
}

export interface ExcelNamedRangeModel {
  name: string;
  ref: string;
  scope?: string;
  comment?: string;
}

export interface ExcelValidationModel {
  sqref: string;
  type?: string;
  formula1?: string;
  formula2?: string;
  operator?: string;
  allowBlank?: boolean;
  showError?: boolean;
  errorTitle?: string;
  error?: string;
  showInput?: boolean;
  promptTitle?: string;
  prompt?: string;
}

export interface ExcelCommentModel {
  ref: string;
  text: string;
  author?: string;
}

export interface ExcelTableModel {
  name?: string;
  displayName?: string;
  ref: string;
  style?: string;
  headerRow?: boolean;
  totalsRow?: boolean;
}

export interface ExcelChartSeriesModel {
  name?: string;
  categories?: string;
  values?: string;
  color?: string;
}

export interface ExcelChartModel {
  title?: string;
  type?: string;
  relTarget?: string;
  series?: ExcelChartSeriesModel[];
}

export interface ExcelPivotTableModel {
  name?: string;
  target?: string;
  cacheId?: string;
}

export interface ExcelSparklineModel {
  location?: string;
  range?: string;
  type?: string;
  color?: string;
  negativeColor?: string;
  markers?: boolean;
}

export interface ExcelShapeModel {
  kind?: "shape" | "picture";
  name?: string;
  text?: string;
  alt?: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  rotation?: number;
}

export interface ExcelSheetModel {
  name: string;
  cells: Record<string, ExcelCellModel>;
  autoFilter?: string;
  freezeTopLeftCell?: string;
  zoom?: number;
  showGridLines?: boolean;
  showHeadings?: boolean;
  showRowColHeaders?: boolean;
  tabColor?: string;
  orientation?: string;
  paperSize?: number;
  fitToPage?: string;
  header?: string;
  footer?: string;
  protection?: boolean;
  protect?: boolean;
  rowBreaks?: number[];
  colBreaks?: number[];
  validations?: ExcelValidationModel[];
  comments?: ExcelCommentModel[];
  tables?: ExcelTableModel[];
  charts?: ExcelChartModel[];
  pivots?: ExcelPivotTableModel[];
  sparklines?: ExcelSparklineModel[];
  shapes?: ExcelShapeModel[];
  pictures?: ExcelShapeModel[];
}

export interface ExcelWorkbookModel {
  sheets: ExcelSheetModel[];
  settings?: ExcelWorkbookSettings;
  styleSheetXml?: string;
  namedRanges?: ExcelNamedRangeModel[];
}
