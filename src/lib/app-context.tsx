import { createContext, useContext, useState, useReducer, useCallback, useEffect, useMemo } from "react";
import type { MappingConfig, WorkbookData, ParsedSheet, FilterRow, ColumnMeta, CellValue, NamePart, MultiFieldEntry, AddressEntry, CustomPropEntry, ExtraConstantField } from "./xls-vcard/types";
import { autoDetect, emptyMultiEntry, newId } from "./xls-vcard/autoDetect";
import { buildVCardText, buildAllVCards } from "./xls-vcard/buildVcard";
import { applyFilters } from "./xls-vcard/applyFilters";
import { readWorkbook, parseSheet } from "./xls-vcard/parseXlsx";
import * as XLSX from "xlsx";

export type Action =
  | { type: "set"; cfg: MappingConfig }
  | { type: "patch"; patch: Partial<MappingConfig> }
  | { type: "multi-add"; group: "phones" | "emails" | "dates" | "urls" }
  | { type: "multi-remove"; group: "phones" | "emails" | "dates" | "urls"; id: string }
  | { type: "multi-update"; group: "phones" | "emails" | "dates" | "urls"; id: string; patch: Partial<MultiFieldEntry> }
  | { type: "name-set"; parts: NamePart[] }
  | { type: "firstName-set"; parts: NamePart[] }
  | { type: "lastName-set"; parts: NamePart[] }
  | { type: "addr-add" }
  | { type: "addr-remove"; id: string }
  | { type: "addr-update"; id: string; patch: Partial<AddressEntry> }
  | { type: "cp-add" }
  | { type: "cp-remove"; id: string }
  | { type: "cp-update"; id: string; patch: Partial<CustomPropEntry> }
  | { type: "ex-add"; kind?: "text" | "column" }
  | { type: "ex-remove"; id: string }
  | { type: "ex-update"; id: string; patch: Partial<ExtraConstantField> };

function reducer(state: MappingConfig, action: Action): MappingConfig {
  switch (action.type) {
    case "set": return action.cfg;
    case "patch": return { ...state, ...action.patch };
    case "multi-add": return { ...state, [action.group]: [...state[action.group], emptyMultiEntry()] };
    case "multi-remove": return { ...state, [action.group]: state[action.group].filter((e) => e.id !== action.id) };
    case "multi-update": return { ...state, [action.group]: state[action.group].map((e) => e.id === action.id ? { ...e, ...action.patch } : e) };
    case "name-set": return { ...state, nameAssembly: action.parts };
    case "firstName-set": return { ...state, firstNameAssembly: action.parts };
    case "lastName-set": return { ...state, lastNameAssembly: action.parts };
    case "addr-add": return { ...state, addresses: [...state.addresses, { id: newId(), street: null, locality: null, region: null, postalCode: null, country: null, label: "" }] };
    case "addr-remove": return { ...state, addresses: state.addresses.filter((a) => a.id !== action.id) };
    case "addr-update": return { ...state, addresses: state.addresses.map((a) => a.id === action.id ? { ...a, ...action.patch } : a) };
    case "cp-add": return { ...state, customProps: [...state.customProps, { id: newId(), name: "X-CUSTOM", columnKey: null, constantValue: "", ablLabel: "" }] };
    case "cp-remove": return { ...state, customProps: state.customProps.filter((c) => c.id !== action.id) };
    case "cp-update": return { ...state, customProps: state.customProps.map((c) => c.id === action.id ? { ...c, ...action.patch } : c) };
    case "ex-add": return { ...state, extraConstants: [...state.extraConstants, action.kind === "column" ? { id: newId(), target: "category" as const, value: "", columnKey: null } : action.kind === "text" ? { id: newId(), target: "category" as const, value: "" } : { id: newId(), target: "note" as const, value: "" }] };
    case "ex-remove": return { ...state, extraConstants: state.extraConstants.filter((e) => e.id !== action.id) };
    case "ex-update": return { ...state, extraConstants: state.extraConstants.map((e) => e.id === action.id ? { ...e, ...action.patch } : e) };
  }
}

const emptyCfg: MappingConfig = {
  mode: "simple", nameAssembly: [], firstNameAssembly: [], lastNameAssembly: [],
  fullName: null, givenName: null, familyName: null, additionalNames: null,
  honorificPrefix: null, honorificSuffix: null, phones: [], emails: [], dates: [],
  urls: [], company: null, department: null, jobTitle: null, role: null,
  addresses: [], categories: null, note: null, customProps: [], extraConstants: [],
};

interface AppState {
  file: File | null;
  wb: WorkbookData | null;
  sheetName: string;
  skipRows: number;
  firstRowIsHeader: boolean;
  parsed: ParsedSheet | null;
  cfg: MappingConfig;
  dispatch: React.Dispatch<Action>;
  filters: FilterRow[];
  setFilters: (f: FilterRow[]) => void;
  previewRowIdx: number;
  setPreviewRowIdx: (n: number) => void;
  error: string | null;
  fileName: string;
  setFileName: (n: string) => void;
  splitPhones: boolean;
  setSplitPhones: (v: boolean) => void;
  filteredRows: Record<string, CellValue>[];
  previewRow: Record<string, CellValue> | null;
  headerMap: Map<string, string>;
  handleFile: (f: File) => Promise<void>;
  enterMapping: () => void;
  download: () => void;
  reset: () => void;
  setSheetName: (n: string) => void;
  setSkipRows: (n: number) => void;
  setFirstRowIsHeader: (v: boolean) => void;
}

const AppContext = createContext<AppState | null>(null);

export function AppProvider({ children }: { children: React.ReactNode }) {
  const [file, setFile] = useState<File | null>(null);
  const [wb, setWb] = useState<WorkbookData | null>(null);
  const [sheetName, setSheetName] = useState<string>("");
  const [skipRows, setSkipRows] = useState(0);
  const [firstRowIsHeader, setFirstRowIsHeader] = useState(true);
  const [parsed, setParsed] = useState<ParsedSheet | null>(null);
  const [cfg, dispatch] = useReducer(reducer, emptyCfg);
  const [filters, setFilters] = useState<FilterRow[]>([]);
  const [previewRowIdx, setPreviewRowIdx] = useState(0);
  const [error, setError] = useState<string | null>(null);
  const [fileName, setFileName] = useState("contacts");
  const [splitPhones, setSplitPhones] = useState(false);

  const STORAGE_KEY = "xls-vcard-session";

  const saveSession = useCallback(() => {
    if (!parsed) return;
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({
        sheetName, skipRows, firstRowIsHeader, cfg, filters,
        previewRowIdx, fileName, splitPhones,
        columns: parsed.columns, rows: parsed.rows,
      }));
    } catch { /* ignore */ }
  }, [sheetName, skipRows, firstRowIsHeader, cfg, filters, previewRowIdx, fileName, splitPhones, parsed]);

  useEffect(() => { const t = setTimeout(saveSession, 500); return () => clearTimeout(t); }, [saveSession]);

  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const saved = JSON.parse(raw);
      setSheetName(saved.sheetName ?? "");
      setSkipRows(saved.skipRows ?? 0);
      setFirstRowIsHeader(saved.firstRowIsHeader ?? true);
      setPreviewRowIdx(saved.previewRowIdx ?? 0);
      setFilters(saved.filters ?? []);
      setFileName(saved.fileName ?? "contacts");
      setSplitPhones(saved.splitPhones ?? false);
      dispatch({ type: "set", cfg: saved.cfg ?? emptyCfg });
      if (saved.columns && saved.rows) setParsed({ columns: saved.columns, rows: saved.rows });
    } catch { /* ignore */ }
  }, []);

  const handleFile = useCallback(async (f: File) => {
    setError(null);
    setFile(f);
    try {
      const w = await readWorkbook(f);
      setWb(w);
      const sheet = w.sheets[0]?.name ?? "";
      setSheetName(sheet);
      const autoSkip = (() => {
        try {
          const rawWb = (w.raw as XLSX.WorkBook);
          const ws = rawWb.Sheets[sheet];
          if (!ws) return 0;
          const arr: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });
          let skip = 0;
          for (const row of arr) {
            if (row.filter((c: any) => c !== null && c !== "").length >= 2) break;
            skip++;
          }
          return skip;
        } catch { return 0; }
      })();
      setSkipRows(autoSkip);
      setFirstRowIsHeader(true);
      setFileName(f.name.replace(/\.[^.]+$/, "").replace(/\s*\(\d+\)\s*/g, " ").replace(/\s+/g, " ").trim() || "contacts");
    } catch (e) { setError((e as Error).message); }
  }, []);

  useEffect(() => {
    if (!wb || !sheetName) return;
    setParsed(parseSheet(wb, sheetName, { skipRows, firstRowIsHeader }));
  }, [wb, sheetName, skipRows, firstRowIsHeader]);

  const enterMapping = useCallback(() => {
    if (!parsed) return;
    const detected = autoDetect(parsed.columns);
    if (fileName && fileName !== "contacts") {
      detected.extraConstants = [{ id: newId(), target: "category" as const, value: fileName }];
    }
    dispatch({ type: "set", cfg: detected });
    setPreviewRowIdx(0);
  }, [parsed, fileName]);

  const headerMap = useMemo(() => {
    const m = new Map<string, string>();
    parsed?.columns.forEach((c) => m.set(c.key, c.header));
    return m;
  }, [parsed]);

  const filteredRows = useMemo(() => {
    if (!parsed) return [];
    return applyFilters(parsed.rows, filters);
  }, [parsed, filters]);

  const previewRow = filteredRows[previewRowIdx] ?? filteredRows[0] ?? null;

  const reset = useCallback(() => {
    setFile(null); setWb(null); setParsed(null);
    setSheetName(""); setFilters([]); setFileName("contacts");
    setSplitPhones(false); setSkipRows(0); setFirstRowIsHeader(true);
    dispatch({ type: "set", cfg: emptyCfg });
    localStorage.removeItem(STORAGE_KEY);
  }, []);

  const cell = (row: Record<string, CellValue>, key: string | null): string => {
    if (!key) return "";
    const v = row[key];
    if (v == null) return "";
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return String(v).trim();
  };

  const download = useCallback(() => {
    if (!parsed) return;
    let text: string;
    if (splitPhones) {
      const parts: string[] = [];
      for (const row of filteredRows) {
        const validPhones = cfg.phones.filter((e) => e.columnKey && cell(row, e.columnKey));
        if (validPhones.length <= 1) { parts.push(buildVCardText(row, cfg, headerMap)); continue; }
        for (const pe of validPhones) {
          const header = headerMap.get(pe.columnKey!) ?? "";
          const splitCfg: MappingConfig = { ...cfg, phones: [pe] };
          const lastNameParts: NamePart[] = [];
          if (cfg.lastNameAssembly.length > 0) lastNameParts.push(...cfg.lastNameAssembly);
          else if (cfg.familyName) lastNameParts.push({ kind: "col", columnKey: cfg.familyName });
          else if (cfg.fullName) lastNameParts.push({ kind: "col", columnKey: cfg.fullName });
          if (lastNameParts.length > 0) { lastNameParts.push({ kind: "const", value: ` (${header})` }); splitCfg.lastNameAssembly = lastNameParts; splitCfg.familyName = null; splitCfg.fullName = null; }
          parts.push(buildVCardText(row, splitCfg, headerMap));
        }
      }
      text = parts.join("\n");
    } else {
      text = buildAllVCards(filteredRows, cfg, headerMap);
    }
    const blob = new Blob([text], { type: "text/vcard;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = (fileName || "contacts") + ".vcf";
    a.click();
    URL.revokeObjectURL(url);
  }, [parsed, filteredRows, cfg, headerMap, splitPhones, fileName]);

  return (
    <AppContext.Provider value={{
      file, wb, sheetName, skipRows, firstRowIsHeader, parsed,
      cfg, dispatch, filters, setFilters, previewRowIdx, setPreviewRowIdx,
      error, fileName, setFileName, splitPhones, setSplitPhones,
      filteredRows, previewRow, headerMap, handleFile, enterMapping,
      download, reset, setSheetName, setSkipRows, setFirstRowIsHeader,
    }}>
      {children}
    </AppContext.Provider>
  );
}

export function useAppState() {
  const ctx = useContext(AppContext);
  if (!ctx) throw new Error("useAppState must be used within AppProvider");
  return ctx;
}