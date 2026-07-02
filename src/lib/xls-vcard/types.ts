export type CellValue = string | number | boolean | Date | null;

export interface ColumnMeta {
  key: string; // stable key = index
  header: string; // display name
  index: number;
}

export interface SheetInfo {
  name: string;
  rowCount: number;
}

export interface ParsedSheet {
  columns: ColumnMeta[];
  rows: Record<string, CellValue>[]; // key -> value
}

export interface WorkbookData {
  sheets: SheetInfo[];
  raw: unknown; // XLSX.WorkBook
}

export type NamePart =
  | { kind: "col"; columnKey: string }
  | { kind: "const"; value: string };

export type LabelStrategy = "suffix" | "column" | "custom";

export interface MultiFieldEntry {
  id: string;
  columnKey: string | null;
  labelStrategy: LabelStrategy;
  labelColumnKey: string | null;
  customLabel: string;
}

export interface AddressEntry {
  id: string;
  street: string | null;
  locality: string | null;
  region: string | null;
  postalCode: string | null;
  country: string | null;
  label: string;
}

export interface CustomPropEntry {
  id: string;
  name: string; // e.g., X-SKYPE
  columnKey: string | null;
  constantValue: string;
  ablLabel: string; // optional Apple label
}

export interface ExtraConstantField {
  id: string;
  target: "note" | "category" | "url" | "org" | "title" | "custom";
  customName?: string;
  value: string;
}

export interface MappingConfig {
  mode: "simple" | "advanced";
  // Name
  nameAssembly: NamePart[]; // FN (backward compat)
  firstNameAssembly: NamePart[]; // given name from parts
  lastNameAssembly: NamePart[]; // family name from parts
  fullName: string | null; // auto-split into given/family
  givenName: string | null;
  familyName: string | null;
  additionalNames: string | null;
  honorificPrefix: string | null;
  honorificSuffix: string | null;
  // Multi
  phones: MultiFieldEntry[];
  emails: MultiFieldEntry[];
  dates: MultiFieldEntry[]; // birthdays/anniversaries via label
  urls: MultiFieldEntry[];
  // Org
  company: string | null;
  department: string | null;
  jobTitle: string | null;
  role: string | null;
  // Address
  addresses: AddressEntry[];
  // Misc
  categories: string | null;
  note: string | null;
  customProps: CustomPropEntry[];
  extraConstants: ExtraConstantField[];
}

export type FilterOp =
  | "equals"
  | "not-equals"
  | "contains"
  | "not-contains"
  | "blank"
  | "not-blank"
  | "starts-with"
  | "ends-with";

export interface FilterRow {
  id: string;
  columnKey: string;
  op: FilterOp;
  value: string;
}
