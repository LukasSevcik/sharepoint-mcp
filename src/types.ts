export type ColumnType =
  | "text"
  | "note"
  | "number"
  | "currency"
  | "date"
  | "boolean"
  | "choice"
  | "multichoice"
  | "person"
  | "url"
  | "calculated";

export interface ColumnSpec {
  name: string;
  displayName?: string;
  type: ColumnType;
  required?: boolean;
  description?: string;
  defaultValue?: string | number | boolean;
  indexed?: boolean;
  // type-specific
  choices?: string[];
  richText?: boolean;
  dateOnly?: boolean;
  multiple?: boolean;
  min?: number;
  max?: number;
  decimals?: number;
  formula?: string;
  outputType?: "text" | "number" | "boolean" | "date" | "currency";
}

export interface ViewSpec {
  name: string;
  isDefault?: boolean;
  columns: string[];
  rowLimit?: number;
  query?: string;
}
