import { useCallback, useEffect, useMemo, useReducer, useRef, useState } from "react";
import {
  Upload,
  Download,
  FileSpreadsheet,
  X,
  Plus,
  Trash2,
  ArrowLeft,
  Phone,
  Mail,
  Calendar,
  Link as LinkIcon,
  MapPin,
  Tag,
  User,
  Filter as FilterIcon,
  Sparkles,
  Settings2,
  GripVertical,
} from "lucide-react";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import { Switch } from "@/components/ui/switch";
import { Tabs, TabsList, TabsTrigger, TabsContent } from "@/components/ui/tabs";
import { Badge } from "@/components/ui/badge";
import { Separator } from "@/components/ui/separator";
import { cn } from "@/lib/utils";
import { readWorkbook, parseSheet } from "@/lib/xls-vcard/parseXlsx";
import { autoDetect, emptyMultiEntry, newId } from "@/lib/xls-vcard/autoDetect";
import { buildVCardText, buildAllVCards } from "@/lib/xls-vcard/buildVcard";
import { applyFilters } from "@/lib/xls-vcard/applyFilters";
import type {
  MappingConfig,
  WorkbookData,
  ParsedSheet,
  FilterRow,
  MultiFieldEntry,
  NamePart,
  ColumnMeta,
  CellValue,
  CustomPropEntry,
  ExtraConstantField,
  AddressEntry,
} from "@/lib/xls-vcard/types";
import * as XLSX from "xlsx";
import type { Action } from "@/lib/app-context";

type Step = "drop" | "preview" | "map" | "export";

export type { Step };

function reducer(state: MappingConfig, action: Action): MappingConfig {
  switch (action.type) {
    case "set":
      return action.cfg;
    case "patch":
      return { ...state, ...action.patch };
    case "multi-add":
      return { ...state, [action.group]: [...state[action.group], emptyMultiEntry()] };
    case "multi-remove":
      return {
        ...state,
        [action.group]: state[action.group].filter((e) => e.id !== action.id),
      };
    case "multi-update":
      return {
        ...state,
        [action.group]: state[action.group].map((e) =>
          e.id === action.id ? { ...e, ...action.patch } : e
        ),
      };
    case "name-set":
      return { ...state, nameAssembly: action.parts };
    case "firstName-set":
      return { ...state, firstNameAssembly: action.parts };
    case "lastName-set":
      return { ...state, lastNameAssembly: action.parts };
    case "addr-add":
      return {
        ...state,
        addresses: [
          ...state.addresses,
          {
            id: newId(),
            street: null,
            locality: null,
            region: null,
            postalCode: null,
            country: null,
            label: "",
          },
        ],
      };
    case "addr-remove":
      return { ...state, addresses: state.addresses.filter((a) => a.id !== action.id) };
    case "addr-update":
      return {
        ...state,
        addresses: state.addresses.map((a) =>
          a.id === action.id ? { ...a, ...action.patch } : a
        ),
      };
    case "cp-add":
      return {
        ...state,
        customProps: [
          ...state.customProps,
          { id: newId(), name: "X-CUSTOM", columnKey: null, constantValue: "", ablLabel: "" },
        ],
      };
    case "cp-remove":
      return { ...state, customProps: state.customProps.filter((c) => c.id !== action.id) };
    case "cp-update":
      return {
        ...state,
        customProps: state.customProps.map((c) =>
          c.id === action.id ? { ...c, ...action.patch } : c
        ),
      };
    case "ex-add":
      return {
        ...state,
        extraConstants: [
          ...state.extraConstants,
          action.kind === "column"
            ? { id: newId(), target: "category" as const, value: "", columnKey: null }
            : action.kind === "text"
              ? { id: newId(), target: "category" as const, value: "" }
              : { id: newId(), target: "note" as const, value: "" },
        ],
      };
    case "ex-remove":
      return {
        ...state,
        extraConstants: state.extraConstants.filter((e) => e.id !== action.id),
      };
    case "ex-update":
      return {
        ...state,
        extraConstants: state.extraConstants.map((e) =>
          e.id === action.id ? { ...e, ...action.patch } : e
        ),
      };
  }
}

const emptyCfg: MappingConfig = {
  mode: "simple",
  nameAssembly: [],
  firstNameAssembly: [],
  lastNameAssembly: [],
  fullName: null,
  givenName: null,
  familyName: null,
  additionalNames: null,
  honorificPrefix: null,
  honorificSuffix: null,
  phones: [],
  emails: [],
  dates: [],
  urls: [],
  company: null,
  department: null,
  jobTitle: null,
  role: null,
  addresses: [],
  categories: null,
  note: null,
  customProps: [],
  extraConstants: [],
};

export function XlsToVcardApp() {
  const [step, setStep] = useState<Step>("drop");
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

  // ── Session persistence ──────────────────────────────────────────
  const STORAGE_KEY = "xls-vcard-session";

  const saveSession = useCallback(() => {
    if (step === "drop" || !parsed) return;
    try {
      const data = {
        step,
        sheetName,
        skipRows,
        firstRowIsHeader,
        cfg,
        filters,
        previewRowIdx,
        fileName,
        splitPhones,
        columns: parsed.columns,
        rows: parsed.rows,
      };
      localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
    } catch { /* quota exceeded or private mode */ }
  }, [step, sheetName, skipRows, firstRowIsHeader, cfg, filters, previewRowIdx, fileName, parsed]);

  // Debounced save
  useEffect(() => {
    const timer = setTimeout(saveSession, 500);
    return () => clearTimeout(timer);
  }, [saveSession]);

  // Restore on mount
  useEffect(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return;
      const saved = JSON.parse(raw);
      setStep(saved.step);
      setSheetName(saved.sheetName ?? "");
      setSkipRows(saved.skipRows ?? 0);
      setFirstRowIsHeader(saved.firstRowIsHeader ?? true);
      setPreviewRowIdx(saved.previewRowIdx ?? 0);
      setFilters(saved.filters ?? []);
      setFileName((saved.fileName ?? "contacts").replace(/\s*\(\d+\)\s*/g, "").trim());
      setSplitPhones(saved.splitPhones ?? false);
      dispatch({ type: "set", cfg: saved.cfg ?? emptyCfg });
      if (saved.columns && saved.rows) {
        setParsed({ columns: saved.columns, rows: saved.rows });
      }
    } catch { /* ignore corrupt data */ }
  }, []);

  const handleFile = useCallback(async (f: File) => {
    setError(null);
    setFile(f);
    try {
      const w = await readWorkbook(f);
      setWb(w);
      const sheet = w.sheets[0]?.name ?? "";
      setSheetName(sheet);
      // Auto-detect skip rows: count leading rows with fewer than 2 non-empty cells
      const autoSkip = (() => {
        try {
          const rawWb = (w.raw as XLSX.WorkBook);
          const ws = rawWb.Sheets[sheet];
          if (!ws) return 0;
          const arr: any[][] = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, blankrows: false });
          let skip = 0;
          for (const row of arr) {
            const filled = row.filter((c) => c !== null && c !== "").length;
            if (filled >= 2) break;
            skip++;
          }
          return skip;
        } catch { return 0; }
      })();
      setSkipRows(autoSkip);
      setFirstRowIsHeader(true);
      setStep("preview");
      setFileName(f.name.replace(/\.[^.]+$/, "").replace(/\s*\(\d+\)\s*/g, " ").replace(/\s+/g, " ").trim() || "contacts");
    } catch (e) {
      setError((e as Error).message);
    }
  }, []);

  // Re-parse when sheet options change
  useEffect(() => {
    if (!wb || !sheetName) return;
    const p = parseSheet(wb, sheetName, { skipRows, firstRowIsHeader });
    setParsed(p);
  }, [wb, sheetName, skipRows, firstRowIsHeader]);

  const enterMapping = useCallback(() => {
    if (!parsed) return;
    const detected = autoDetect(parsed.columns);
    if (fileName && fileName !== "contacts") {
      detected.extraConstants = [
        { id: newId(), target: "category" as const, value: fileName },
      ];
    }
    dispatch({ type: "set", cfg: detected });
    setPreviewRowIdx(0);
    setStep("map");
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

  const reset = () => {
    setStep("drop");
    setFile(null);
    setWb(null);
    setParsed(null);
    setSheetName("");
    setFilters([]);
    setFileName("contacts");
    setSplitPhones(false);
    dispatch({ type: "set", cfg: emptyCfg });
    localStorage.removeItem("xls-vcard-session");
  };

  const handleStepNav = (target: Step) => {
    if (target === "drop") {
      reset();
    } else if (target === "preview" && (step === "map" || step === "export")) {
      setStep("preview");
    }
  };

  const download = () => {
    if (!parsed) return;

    const cell = (row: Record<string, CellValue>, key: string | null): string => {
      if (!key) return "";
      const v = row[key];
      if (v == null) return "";
      if (v instanceof Date) return v.toISOString().slice(0, 10);
      return String(v).trim();
    };

    let text: string;
    if (splitPhones) {
      const parts: string[] = [];
      for (const row of filteredRows) {
        const validPhones = cfg.phones.filter((e) => e.columnKey && cell(row, e.columnKey));
        if (validPhones.length <= 1) {
          parts.push(buildVCardText(row, cfg, headerMap));
        } else {
          for (const pe of validPhones) {
            const header = headerMap.get(pe.columnKey!) ?? "";
            const splitCfg: MappingConfig = {
              ...cfg,
              phones: [pe],
            };
            const lastNameParts: NamePart[] = [];
            if (cfg.lastNameAssembly.length > 0) {
              lastNameParts.push(...cfg.lastNameAssembly);
            } else if (cfg.familyName) {
              lastNameParts.push({ kind: "col", columnKey: cfg.familyName });
            } else if (cfg.fullName) {
              lastNameParts.push({ kind: "col", columnKey: cfg.fullName });
            }
            if (lastNameParts.length > 0) {
              lastNameParts.push({ kind: "const", value: ` (${header})` });
              splitCfg.lastNameAssembly = lastNameParts;
              splitCfg.familyName = null;
              splitCfg.fullName = null;
            }
            parts.push(buildVCardText(row, splitCfg, headerMap));
          }
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
  };

  return (
    <div className="min-h-screen bg-background flex flex-col justify-between">
      <header className="border-b border-border">
        <div className="mx-auto flex max-w-7xl items-center justify-between px-4 sm:px-6 py-4">
          <div className="flex items-center gap-2">
            <div className="flex h-8 w-8 items-center justify-center rounded-md bg-primary text-primary-foreground">
              <FileSpreadsheet className="h-4 w-4" />
            </div>
            <div>
              <h1 className="text-sm font-semibold">XLS → vCard</h1>
              <p className="text-xs text-muted-foreground">
                Drop, map, download. No history kept.
              </p>
            </div>
          </div>
          {step !== "drop" && (
            <Button variant="ghost" size="sm" onClick={reset}>
              <X className="h-4 w-4" /> Start over
            </Button>
          )}
        </div>
        <Stepper step={step} onNavigate={handleStepNav} />
      </header>

      <main className="mx-auto max-w-7xl px-4 sm:px-6 py-8 w-full" style={{ flexGrow: 1 }}>
        {error && (
          <div className="mb-4 rounded-md border border-destructive/50 bg-destructive/10 px-4 py-2 text-sm text-destructive">
            {error}
          </div>
        )}
        {step === "drop" && <DropStep onFile={handleFile} />}
        {step === "preview" && wb && (
          <PreviewStep
            wb={wb}
            sheetName={sheetName}
            setSheetName={setSheetName}
            skipRows={skipRows}
            setSkipRows={setSkipRows}
            firstRowIsHeader={firstRowIsHeader}
            setFirstRowIsHeader={setFirstRowIsHeader}
            parsed={parsed}
            filters={filters}
            setFilters={setFilters}
            filteredRows={filteredRows}
            onBack={() => setStep("drop")}
            onNext={enterMapping}
          />
        )}
        {step === "map" && parsed && (
          <MapStep
            parsed={parsed}
            cfg={cfg}
            dispatch={dispatch}
            filters={filters}
            setFilters={setFilters}
            filteredRows={filteredRows}
            previewRow={previewRow}
            previewRowIdx={previewRowIdx}
            setPreviewRowIdx={setPreviewRowIdx}
            headerMap={headerMap}
            onBack={() => setStep("preview")}
            onNext={() => setStep("export")}
            splitPhones={splitPhones}
            setSplitPhones={setSplitPhones}
          />
        )}
        {step === "export" && parsed && (
          <ExportStep
            count={filteredRows.length}
            fileName={fileName}
            onFileNameChange={setFileName}
            onBack={() => setStep("map")}
            onDownload={download}
            onReset={reset}
            previewRow={previewRow}
            cfg={cfg}
            headerMap={headerMap}
            splitPhones={splitPhones}
            rowIdx={previewRowIdx}
            setRowIdx={setPreviewRowIdx}
            total={filteredRows.length}
          />
        )}
      </main>
      <footer className="border-t border-border py-4 text-center text-xs text-muted-foreground">
        <a
          href="https://github.com/Mohamed-faaris/xls-to-vcard"
          target="_blank"
          rel="noopener noreferrer"
          className="inline-flex items-center gap-1.5 hover:text-foreground transition-colors"
        >
          <svg viewBox="0 0 1024 1024" fill="none" className="h-4 w-4">
            <path fill-rule="evenodd" clip-rule="evenodd" d="M8 0C3.58 0 0 3.58 0 8C0 11.54 2.29 14.53 5.47 15.59C5.87 15.66 6.02 15.42 6.02 15.21C6.02 15.02 6.01 14.39 6.01 13.72C4 14.09 3.48 13.23 3.32 12.78C3.23 12.55 2.84 11.84 2.5 11.65C2.22 11.5 1.82 11.13 2.49 11.12C3.12 11.11 3.57 11.7 3.72 11.94C4.44 13.15 5.59 12.81 6.05 12.6C6.12 12.08 6.33 11.73 6.56 11.53C4.78 11.33 2.92 10.64 2.92 7.58C2.92 6.71 3.23 5.99 3.74 5.43C3.66 5.23 3.38 4.41 3.82 3.31C3.82 3.31 4.49 3.1 6.02 4.13C6.66 3.95 7.34 3.86 8.02 3.86C8.7 3.86 9.38 3.95 10.02 4.13C11.55 3.09 12.22 3.31 12.22 3.31C12.66 4.41 12.38 5.23 12.3 5.43C12.81 5.99 13.12 6.7 13.12 7.58C13.12 10.65 11.25 11.33 9.47 11.53C9.76 11.78 10.01 12.26 10.01 13.01C10.01 14.08 10 14.94 10 15.21C10 15.42 10.15 15.67 10.55 15.59C13.71 14.53 16 11.53 16 8C16 3.58 12.42 0 8 0Z" transform="scale(64)" fill="currentColor"/>
          </svg>
          Star on GitHub — fork, contribute, report issues
        </a>
      </footer>
    </div>
  );
}

export function Stepper({
  step,
  onNavigate,
}: {
  step: Step;
  onNavigate: (s: Step) => void;
}) {
  const steps: { id: Step; label: string }[] = [
    { id: "drop", label: "Upload" },
    { id: "preview", label: "Preview" },
    { id: "map", label: "Map & Filter" },
    { id: "export", label: "Export" },
  ];
  const idx = steps.findIndex((s) => s.id === step);
  return (
    <nav className="mx-auto flex max-w-7xl items-center gap-2 overflow-x-auto px-6 pb-4" aria-label="Step navigation">
      {steps.map((s, i) => (
        <div key={s.id} className="flex items-center gap-1.5 sm:gap-2 shrink-0">
          <button
            type="button"
            disabled={i >= idx}
            onClick={() => onNavigate(s.id)}
            className={cn(
              "flex h-6 w-6 items-center justify-center rounded-full text-xs font-medium transition-colors shrink-0",
              i <= idx
                ? "bg-primary text-primary-foreground"
                : "bg-muted text-muted-foreground",
              i < idx && "cursor-pointer hover:opacity-80"
            )}
          >
            {i + 1}
          </button>
          <button
            type="button"
            disabled={i >= idx}
            onClick={() => onNavigate(s.id)}
            className={cn(
              "text-xs transition-colors whitespace-nowrap",
              i === idx
                ? "font-medium text-foreground"
                : "text-muted-foreground",
              i < idx && "cursor-pointer hover:text-foreground"
            )}
          >
            {s.label}
          </button>
          {i < steps.length - 1 && <div className="h-px w-4 sm:w-8 bg-border shrink-0" />}
        </div>
      ))}
    </nav>
  );
}

export function DropStep({ onFile }: { onFile: (f: File) => void }) {
  const [drag, setDrag] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);
  return (
    <div className="flex flex-col items-center py-16">
      <div
        onDragOver={(e) => {
          e.preventDefault();
          setDrag(true);
        }}
        onDragLeave={() => setDrag(false)}
        onDrop={(e) => {
          e.preventDefault();
          setDrag(false);
          const f = e.dataTransfer.files?.[0];
          if (f) onFile(f);
        }}
        onClick={() => inputRef.current?.click()}
          className={cn(
            "flex w-full max-w-2xl cursor-pointer flex-col items-center justify-center rounded-lg border-2 border-dashed px-4 sm:px-6 py-16 sm:py-20 transition-colors",
          drag
            ? "border-primary bg-primary/5"
            : "border-border hover:border-primary/50 hover:bg-muted/30"
        )}
      >
        <input
          ref={inputRef}
          type="file"
          accept=".xls,.xlsx,.csv"
          className="hidden"
          onChange={(e) => {
            const f = e.target.files?.[0];
            if (f) onFile(f);
          }}
        />
        <div className="flex cursor-pointer flex-col items-center">
          <Upload className="mb-4 h-10 w-10 text-muted-foreground" />
          <p className="text-base font-medium">Drop your spreadsheet here</p>
          <p className="mt-1 text-sm text-muted-foreground">
            .xls, .xlsx, or .csv — everything stays in your browser
          </p>
        </div>
        <Button className="mt-6" variant="outline" onClick={(e) => { e.stopPropagation(); inputRef.current?.click(); }}>
          Choose file
        </Button>
      </div>
    </div>
  );
}

export function PreviewStep(props: {
  wb: WorkbookData | null;
  sheetName: string;
  setSheetName: (s: string) => void;
  skipRows: number;
  setSkipRows: (n: number) => void;
  firstRowIsHeader: boolean;
  setFirstRowIsHeader: (b: boolean) => void;
  parsed: ParsedSheet | null;
  filters: FilterRow[];
  setFilters: (f: FilterRow[]) => void;
  filteredRows: Record<string, CellValue>[];
  onBack: () => void;
  onNext: () => void;
}) {
  const {
    wb,
    sheetName,
    setSheetName,
    skipRows,
    setSkipRows,
    firstRowIsHeader,
    setFirstRowIsHeader,
    parsed,
    filters,
    setFilters,
    filteredRows,
    onBack,
    onNext,
  } = props;
  return (
    <div className="space-y-4">
      <div className="rounded-lg border border-border bg-card p-4">
        <div className="grid gap-4 grid-cols-1 sm:grid-cols-3">
          <div>
            <Label className="text-xs">Sheet</Label>
            <Select value={sheetName} onValueChange={setSheetName}>
              <SelectTrigger className="mt-1">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                {(wb?.sheets ?? []).map((s) => (
                  <SelectItem key={s.name} value={s.name}>
                    {s.name} ({s.rowCount} rows)
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
          </div>
          <div>
            <Label className="text-xs">Skip top rows</Label>
            <Input
              type="number"
              min={0}
              className="mt-1"
              value={skipRows}
              onChange={(e) => setSkipRows(Math.max(0, parseInt(e.target.value) || 0))}
            />
            <p className="mt-1 text-xs text-muted-foreground">
              For merged title rows above your data
            </p>
          </div>
          <div>
            <Label className="text-xs">Header row</Label>
            <div className="mt-2 flex items-center gap-2">
              <Switch
                checked={firstRowIsHeader}
                onCheckedChange={setFirstRowIsHeader}
                id="hdr"
              />
              <Label htmlFor="hdr" className="cursor-pointer text-sm font-normal">
                Use first row as column names
              </Label>
            </div>
          </div>
        </div>
      </div>

      {parsed && (
        <FilterSection
          filters={filters}
          setFilters={setFilters}
          columns={parsed.columns}
          totalCount={parsed.rows.length}
          filteredCount={filteredRows.length}
        />
      )}

      <div className="overflow-hidden rounded-lg border border-border bg-card">
        <div className="border-b border-border px-4 py-2 text-xs text-muted-foreground">
          {filteredRows.length < (parsed?.rows.length ?? 0)
            ? `${filteredRows.length} of ${parsed?.rows.length ?? 0} rows`
            : `Preview — ${parsed?.rows.length ?? 0} data rows detected`}
        </div>
        <div className="max-h-96 overflow-x-auto">
          <table className="w-full text-sm min-w-[500px]">
            <thead className="sticky top-0 bg-muted">
              <tr>
                {parsed?.columns.map((c) => (
                  <th
                    key={c.key}
                    className="border-b border-border px-3 py-2 text-left font-medium"
                  >
                    {c.header}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {filteredRows.slice(0, 20).map((r, i) => (
                <tr key={i} className="border-b border-border last:border-0">
                  {(parsed?.columns ?? []).map((c) => (
                    <td key={c.key} className="px-3 py-2 text-muted-foreground">
                      {formatCell(r[c.key])}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div className="flex justify-between">
        <Button variant="ghost" onClick={onBack}>
          <ArrowLeft className="h-4 w-4" /> Back
        </Button>
        <Button
          onClick={onNext}
          disabled={!parsed || filteredRows.length === 0}
        >
          Continue ({filteredRows.length})
        </Button>
      </div>
    </div>
  );
}

export function formatCell(v: CellValue): string {
  if (v == null) return "";
  if (v instanceof Date) return v.toLocaleDateString();
  return String(v);
}

export function MapStep(props: {
  parsed: ParsedSheet;
  cfg: MappingConfig;
  dispatch: React.Dispatch<Action>;
  filters: FilterRow[];
  setFilters: (f: FilterRow[]) => void;
  filteredRows: Record<string, CellValue>[];
  previewRow: Record<string, CellValue> | null;
  previewRowIdx: number;
  setPreviewRowIdx: (n: number) => void;
  headerMap: Map<string, string>;
  onBack: () => void;
  onNext: () => void;
  splitPhones: boolean;
  setSplitPhones: (v: boolean) => void;
}) {
  const {
    parsed,
    cfg,
    dispatch,
    filters,
    setFilters,
    filteredRows,
    previewRow,
    previewRowIdx,
    setPreviewRowIdx,
    headerMap,
    onBack,
    onNext,
    splitPhones,
    setSplitPhones,
  } = props;

  return (
    <div className="grid gap-6 lg:grid-cols-[1fr_400px]">
      <div className="space-y-4">
        <div className="flex items-center justify-between rounded-lg border border-border bg-card p-3">
          <div className="flex items-center gap-2">
            <Sparkles className="h-4 w-4 text-primary" />
            <span className="text-sm font-medium">
              {cfg.mode === "simple" ? "Simple mode" : "Advanced mode"}
            </span>
            <span className="text-xs text-muted-foreground hidden sm:inline">
              {cfg.mode === "simple"
                ? "Auto-detected mapping — quick tweaks"
                : "Full control over every field"}
            </span>
          </div>
          <div className="flex items-center gap-2">
            <Label htmlFor="mode" className="text-xs">
              Advanced
            </Label>
            <Switch
              id="mode"
              checked={cfg.mode === "advanced"}
              onCheckedChange={(v) =>
                dispatch({ type: "patch", patch: { mode: v ? "advanced" : "simple" } })
              }
            />
          </div>
        </div>

        <Tabs defaultValue="mapping">
          <TabsList>
            <TabsTrigger value="mapping">
              <Settings2 className="h-3 w-3" /> Mapping
            </TabsTrigger>
            <TabsTrigger value="filters">
              <FilterIcon className="h-3 w-3" /> Filters
              {filters.length > 0 && (
                <Badge variant="secondary" className="ml-1">
                  {filters.length}
                </Badge>
              )}
            </TabsTrigger>
          </TabsList>

          <TabsContent value="mapping" className="space-y-4">
            <NameSection cfg={cfg} dispatch={dispatch} columns={parsed.columns} advanced={cfg.mode === "advanced"} />
            <MultiSection
              title="Phone numbers"
              icon={<Phone className="h-4 w-4" />}
              group="phones"
              entries={cfg.phones}
              dispatch={dispatch}
              columns={parsed.columns}
              advanced={cfg.mode === "advanced"}
              footer={
                cfg.phones.length > 1 ? (
                  <div className="flex items-center gap-2 border-t border-border pt-3">
                    <Switch
                      id="split-phones"
                      checked={splitPhones}
                      onCheckedChange={setSplitPhones}
                    />
                    <Label htmlFor="split-phones" className="text-xs cursor-pointer">
                      Split each phone number into a separate contact
                    </Label>
                  </div>
                ) : undefined
              }
            />
            <MultiSection
              title="Email addresses"
              icon={<Mail className="h-4 w-4" />}
              group="emails"
              entries={cfg.emails}
              dispatch={dispatch}
              columns={parsed.columns}
              advanced={cfg.mode === "advanced"}
            />
            <MultiSection
              title="Dates (birthday, anniversary, custom)"
              icon={<Calendar className="h-4 w-4" />}
              group="dates"
              entries={cfg.dates}
              dispatch={dispatch}
              columns={parsed.columns}
              advanced={cfg.mode === "advanced"}
            />
            {cfg.mode === "advanced" && (
              <>
                <MultiSection
                  title="URLs"
                  icon={<LinkIcon className="h-4 w-4" />}
                  group="urls"
                  entries={cfg.urls}
                  dispatch={dispatch}
                  columns={parsed.columns}
                  advanced
                />
                <OrgSection cfg={cfg} dispatch={dispatch} columns={parsed.columns} />
                <AddressSection cfg={cfg} dispatch={dispatch} columns={parsed.columns} />
                <MiscSection cfg={cfg} dispatch={dispatch} columns={parsed.columns} />
                <CustomPropsSection cfg={cfg} dispatch={dispatch} columns={parsed.columns} />
                <ExtraConstantsSection cfg={cfg} dispatch={dispatch} />
              </>
            )}
            {cfg.mode !== "advanced" && (
              <Section
                title="Categories"
                icon={<Tag className="h-4 w-4" />}
                action={
                  <div className="flex gap-1">
                    <Button size="sm" variant="outline" onClick={() => dispatch({ type: "ex-add", kind: "text" })}>
                      <Plus className="h-3 w-3" /> Text
                    </Button>
                    <Button size="sm" variant="outline" onClick={() => dispatch({ type: "ex-add", kind: "column" })}>
                      <Plus className="h-3 w-3" /> Column
                    </Button>
                  </div>
                }
              >
                {cfg.extraConstants.filter((e) => e.target === "category").length === 0 && (
                  <p className="text-xs text-muted-foreground">None added.</p>
                )}
                {cfg.extraConstants.filter((e) => e.target === "category").map((e) => (
                  <div key={e.id} className="flex items-center gap-2 rounded-md border border-border bg-muted/30 p-2">
                    {"columnKey" in e && e.columnKey !== undefined ? (
                      <ColSelect
                        value={e.columnKey ?? null}
                        onChange={(v) => dispatch({ type: "ex-update", id: e.id, patch: { columnKey: v } })}
                        columns={parsed.columns}
                        placeholder="Choose column..."
                      />
                    ) : (
                      <Input
                        className="h-8 text-sm flex-1"
                        placeholder="Category text"
                        value={e.value}
                        onChange={(ev) => dispatch({ type: "ex-update", id: e.id, patch: { value: ev.target.value } })}
                      />
                    )}
                    <Button size="icon" variant="ghost" className="h-8 w-8 shrink-0" onClick={() => dispatch({ type: "ex-remove", id: e.id })}>
                      <Trash2 className="h-3 w-3" />
                    </Button>
                  </div>
                ))}
              </Section>
            )}
          </TabsContent>

          <TabsContent value="filters">
            <FilterSection
              filters={filters}
              setFilters={setFilters}
              columns={parsed.columns}
              totalCount={parsed.rows.length}
              filteredCount={filteredRows.length}
            />
          </TabsContent>
        </Tabs>

        <div className="flex justify-between">
          <Button variant="ghost" onClick={onBack}>
            <ArrowLeft className="h-4 w-4" /> Back
          </Button>
          <Button onClick={onNext} disabled={filteredRows.length === 0}>
            Continue ({filteredRows.length})
          </Button>
        </div>
      </div>

      <LivePreview
        row={previewRow}
        cfg={cfg}
        headerMap={headerMap}
        rowIdx={previewRowIdx}
        setRowIdx={setPreviewRowIdx}
        total={filteredRows.length}
        splitPhones={splitPhones}
      />
    </div>
  );
}

export function Section({
  title,
  icon,
  children,
  action,
}: {
  title: string;
  icon?: React.ReactNode;
  children: React.ReactNode;
  action?: React.ReactNode;
}) {
  return (
    <div className="rounded-lg border border-border bg-card">
      <div className="flex items-center justify-between border-b border-border px-4 py-3">
        <div className="flex items-center gap-2 text-sm font-medium">
          {icon}
          {title}
        </div>
        {action}
      </div>
      <div className="space-y-3 p-4">{children}</div>
    </div>
  );
}

export function ColSelect({
  value,
  onChange,
  columns,
  placeholder = "— none —",
}: {
  value: string | null;
  onChange: (v: string | null) => void;
  columns: ColumnMeta[];
  placeholder?: string;
}) {
  return (
    <Select
      value={value ?? "__none"}
      onValueChange={(v) => onChange(v === "__none" ? null : v)}
    >
      <SelectTrigger className="h-8 text-sm">
        <SelectValue placeholder={placeholder} />
      </SelectTrigger>
      <SelectContent>
        <SelectItem value="__none">— none —</SelectItem>
        {columns.map((c) => (
          <SelectItem key={c.key} value={c.key}>
            {c.header}
          </SelectItem>
        ))}
      </SelectContent>
    </Select>
  );
}

function NameAssembler({
  label,
  parts,
  columns,
  onParts,
}: {
  label: string;
  parts: NamePart[];
  columns: ColumnMeta[];
  onParts: (parts: NamePart[]) => void;
}) {
  const [dragIdx, setDragIdx] = useState<number | null>(null);

  const updatePart = (i: number, patch: Partial<NamePart>) => {
    onParts(parts.map((p, idx) => (idx === i ? ({ ...p, ...patch } as NamePart) : p)));
  };
  const removePart = (i: number) => {
    onParts(parts.filter((_, idx) => idx !== i));
  };
  const movePart = (from: number, to: number) => {
    const next = [...parts];
    const [moved] = next.splice(from, 1);
    next.splice(to, 0, moved);
    onParts(next);
  };
  const addPart = (kind: "col" | "const") => {
    const p: NamePart =
      kind === "col"
        ? { kind: "col", columnKey: columns[0]?.key ?? "" }
        : { kind: "const", value: " " };
    onParts([...parts, p]);
  };

  return (
    <div>
      <Label className="text-xs">{label}</Label>
      <p className="mb-1.5 text-[10px] text-muted-foreground">
        Add columns and constant text in order.
      </p>
      <div className="space-y-1.5">
        {parts.map((p, i) => (
          <div
            key={i}
            draggable
            onDragStart={() => setDragIdx(i)}
            onDragOver={(e) => { e.preventDefault(); e.currentTarget.style.opacity = "0.5"; }}
            onDragLeave={(e) => { e.currentTarget.style.opacity = ""; }}
            onDrop={(e) => {
              e.preventDefault();
              e.currentTarget.style.opacity = "";
              if (dragIdx !== null && dragIdx !== i) movePart(dragIdx, i);
              setDragIdx(null);
            }}
            onDragEnd={() => setDragIdx(null)}
            className={cn("flex items-center gap-1.5", dragIdx === i && "opacity-40")}
          >
            <div className="cursor-grab text-muted-foreground">
              <GripVertical className="h-3.5 w-3.5" />
            </div>
            <Badge variant="outline" className="text-[10px] shrink-0">
              {p.kind === "col" ? "Col" : "Text"}
            </Badge>
            {p.kind === "col" ? (
              <div className="flex-1 min-w-0">
                <ColSelect value={p.columnKey} onChange={(v) => updatePart(i, { columnKey: v ?? "" })} columns={columns} />
              </div>
            ) : (
              <Input className="h-7 flex-1 text-sm" value={p.value} onChange={(e) => updatePart(i, { value: e.target.value })} placeholder="text" />
            )}
            <Button size="icon" variant="ghost" className="h-7 w-7 shrink-0" onClick={() => removePart(i)}>
              <Trash2 className="h-3 w-3" />
            </Button>
          </div>
        ))}
        <div className="flex gap-1.5">
          <Button size="sm" variant="outline" className="h-7 px-2 text-xs" onClick={() => addPart("col")}>
            <Plus className="h-3 w-3" /> Column
          </Button>
          <Button size="sm" variant="outline" className="h-7 px-2 text-xs" onClick={() => addPart("const")}>
            <Plus className="h-3 w-3" /> Text
          </Button>
        </div>
      </div>
    </div>
  );
}

function NameSection({
  cfg,
  dispatch,
  columns,
  advanced,
}: {
  cfg: MappingConfig;
  dispatch: React.Dispatch<Action>;
  columns: ColumnMeta[];
  advanced?: boolean;
}) {
  return (
    <Section title="Name" icon={<User className="h-4 w-4" />}>
      <div className="grid gap-4 grid-cols-1 sm:grid-cols-2">
        <NameAssembler label="First name" parts={cfg.firstNameAssembly} columns={columns} onParts={(p) => dispatch({ type: "firstName-set", parts: p })} />
        <NameAssembler label="Last name" parts={cfg.lastNameAssembly} columns={columns} onParts={(p) => dispatch({ type: "lastName-set", parts: p })} />
      </div>
      {advanced && (
        <>
          <Separator />
          <div className="grid gap-3 grid-cols-1 sm:grid-cols-2">
            <Field label="Full name (auto-split)">
              <ColSelect value={cfg.fullName} onChange={(v) => dispatch({ type: "patch", patch: { fullName: v } })} columns={columns} />
              <p className="mt-0.5 text-[10px] text-muted-foreground">Splits on first space when assemblers above are empty</p>
            </Field>
            <Field label="Prefix (Mr., Dr.)">
              <ColSelect value={cfg.honorificPrefix} onChange={(v) => dispatch({ type: "patch", patch: { honorificPrefix: v } })} columns={columns} />
            </Field>
            <Field label="Middle name(s)">
              <ColSelect value={cfg.additionalNames} onChange={(v) => dispatch({ type: "patch", patch: { additionalNames: v } })} columns={columns} />
            </Field>
            <Field label="Suffix (Jr., PhD)">
              <ColSelect value={cfg.honorificSuffix} onChange={(v) => dispatch({ type: "patch", patch: { honorificSuffix: v } })} columns={columns} />
            </Field>
          </div>
        </>
      )}
    </Section>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div>
      <Label className="text-xs">{label}</Label>
      <div className="mt-1">{children}</div>
    </div>
  );
}

function MultiSection({
  title,
  icon,
  group,
  entries,
  dispatch,
  columns,
  advanced,
  footer,
}: {
  title: string;
  icon: React.ReactNode;
  group: "phones" | "emails" | "dates" | "urls";
  entries: MultiFieldEntry[];
  dispatch: React.Dispatch<Action>;
  columns: ColumnMeta[];
  advanced: boolean;
  footer?: React.ReactNode;
}) {
  return (
    <Section
      title={title}
      icon={icon}
      action={
        <Button size="sm" variant="outline" onClick={() => dispatch({ type: "multi-add", group })}>
          <Plus className="h-3 w-3" /> Add
        </Button>
      }
    >
      {entries.length === 0 && (
        <p className="text-xs text-muted-foreground">None mapped.</p>
      )}
      {entries.map((e) => (
        <div key={e.id} className="rounded-md border border-border bg-muted/30 p-3">
          <div className="flex flex-wrap items-center gap-2">
            <div className="min-w-[140px] flex-1">
              <ColSelect
                value={e.columnKey}
                onChange={(v) =>
                  dispatch({ type: "multi-update", group, id: e.id, patch: { columnKey: v } })
                }
                columns={columns}
              />
            </div>
            {group !== "urls" && (
              <Input
                className="h-8 w-full sm:w-28 text-sm"
                placeholder="Label"
                value={e.customLabel}
                onChange={(ev) =>
                  dispatch({
                    type: "multi-update",
                    group,
                    id: e.id,
                    patch: { customLabel: ev.target.value, labelStrategy: "custom" },
                  })
                }
              />
            )}
            <Button
              size="icon"
              variant="ghost"
              className="h-8 w-8 shrink-0"
              onClick={() => dispatch({ type: "multi-remove", group, id: e.id })}
            >
              <Trash2 className="h-3 w-3" />
            </Button>
          </div>
          {advanced && group !== "urls" && (
            <div className="mt-2 grid gap-2 grid-cols-1 sm:grid-cols-2">
              <div>
                <Label className="text-xs">Label from</Label>
                <Select
                  value={e.labelStrategy}
                  onValueChange={(v) =>
                    dispatch({
                      type: "multi-update",
                      group,
                      id: e.id,
                      patch: { labelStrategy: v as MultiFieldEntry["labelStrategy"] },
                    })
                  }
                >
                  <SelectTrigger className="h-8 text-sm">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="custom">Custom label above</SelectItem>
                    <SelectItem value="suffix">Column name suffix</SelectItem>
                    <SelectItem value="column">Another column</SelectItem>
                  </SelectContent>
                </Select>
              </div>
              {e.labelStrategy === "column" && (
                <div>
                  <Label className="text-xs">Label column</Label>
                  <ColSelect
                    value={e.labelColumnKey}
                    onChange={(v) =>
                      dispatch({
                        type: "multi-update",
                        group,
                        id: e.id,
                        patch: { labelColumnKey: v },
                      })
                    }
                    columns={columns}
                  />
                </div>
              )}
            </div>
          )}
        </div>
      ))}
      {footer}
    </Section>
  );
}

function OrgSection({
  cfg,
  dispatch,
  columns,
}: {
  cfg: MappingConfig;
  dispatch: React.Dispatch<Action>;
  columns: ColumnMeta[];
}) {
  return (
    <Section title="Organization">
      <div className="grid gap-3 grid-cols-1 sm:grid-cols-2">
        <Field label="Company">
          <ColSelect
            value={cfg.company}
            onChange={(v) => dispatch({ type: "patch", patch: { company: v } })}
            columns={columns}
          />
        </Field>
        <Field label="Department">
          <ColSelect
            value={cfg.department}
            onChange={(v) => dispatch({ type: "patch", patch: { department: v } })}
            columns={columns}
          />
        </Field>
        <Field label="Job title">
          <ColSelect
            value={cfg.jobTitle}
            onChange={(v) => dispatch({ type: "patch", patch: { jobTitle: v } })}
            columns={columns}
          />
        </Field>
        <Field label="Role">
          <ColSelect
            value={cfg.role}
            onChange={(v) => dispatch({ type: "patch", patch: { role: v } })}
            columns={columns}
          />
        </Field>
      </div>
    </Section>
  );
}

function AddressSection({
  cfg,
  dispatch,
  columns,
}: {
  cfg: MappingConfig;
  dispatch: React.Dispatch<Action>;
  columns: ColumnMeta[];
}) {
  return (
    <Section
      title="Addresses"
      action={
        <Button size="sm" variant="outline" onClick={() => dispatch({ type: "addr-add" })}>
          <Plus className="h-3 w-3" /> Add
        </Button>
      }
    >
      {cfg.addresses.length === 0 && (
        <p className="text-xs text-muted-foreground">None mapped.</p>
      )}
      {cfg.addresses.map((a) => (
        <div key={a.id} className="rounded-md border border-border bg-muted/30 p-3">
          <div className="grid gap-2 grid-cols-1 sm:grid-cols-2">
            <Field label="Street">
              <ColSelect
                value={a.street}
                onChange={(v) => dispatch({ type: "addr-update", id: a.id, patch: { street: v } })}
                columns={columns}
              />
            </Field>
            <Field label="City">
              <ColSelect
                value={a.locality}
                onChange={(v) => dispatch({ type: "addr-update", id: a.id, patch: { locality: v } })}
                columns={columns}
              />
            </Field>
            <Field label="Region / State">
              <ColSelect
                value={a.region}
                onChange={(v) => dispatch({ type: "addr-update", id: a.id, patch: { region: v } })}
                columns={columns}
              />
            </Field>
            <Field label="Postal code">
              <ColSelect
                value={a.postalCode}
                onChange={(v) => dispatch({ type: "addr-update", id: a.id, patch: { postalCode: v } })}
                columns={columns}
              />
            </Field>
            <Field label="Country">
              <ColSelect
                value={a.country}
                onChange={(v) => dispatch({ type: "addr-update", id: a.id, patch: { country: v } })}
                columns={columns}
              />
            </Field>
          </div>
          <div className="mt-2 flex justify-end">
            <Button
              size="sm"
              variant="ghost"
              onClick={() => dispatch({ type: "addr-remove", id: a.id })}
            >
              <Trash2 className="h-3 w-3" /> Remove
            </Button>
          </div>
        </div>
      ))}
    </Section>
  );
}

function MiscSection({
  cfg,
  dispatch,
  columns,
}: {
  cfg: MappingConfig;
  dispatch: React.Dispatch<Action>;
  columns: ColumnMeta[];
}) {
  return (
    <Section title="Notes & Categories">
      <div className="grid gap-3 grid-cols-1 sm:grid-cols-2">
        <Field label="Categories (comma or semicolon separated)">
          <ColSelect
            value={cfg.categories}
            onChange={(v) => dispatch({ type: "patch", patch: { categories: v } })}
            columns={columns}
          />
        </Field>
        <Field label="Note">
          <ColSelect
            value={cfg.note}
            onChange={(v) => dispatch({ type: "patch", patch: { note: v } })}
            columns={columns}
          />
        </Field>
      </div>
    </Section>
  );
}

function CustomPropsSection({
  cfg,
  dispatch,
  columns,
}: {
  cfg: MappingConfig;
  dispatch: React.Dispatch<Action>;
  columns: ColumnMeta[];
}) {
  return (
    <Section
      title="Custom vCard properties"
      action={
        <Button size="sm" variant="outline" onClick={() => dispatch({ type: "cp-add" })}>
          <Plus className="h-3 w-3" /> Add
        </Button>
      }
    >
      {cfg.customProps.length === 0 && (
        <p className="text-xs text-muted-foreground">
          Add any vCard property (X-SKYPE, X-TWITTER, X-ANNIVERSARY, etc.)
        </p>
      )}
      {cfg.customProps.map((cp) => (
        <div key={cp.id} className="grid gap-2 rounded-md border border-border bg-muted/30 p-3 grid-cols-2 sm:grid-cols-4">
          <Field label="Property name">
            <Input
              className="h-8 text-sm"
              value={cp.name}
              onChange={(e) => dispatch({ type: "cp-update", id: cp.id, patch: { name: e.target.value } })}
            />
          </Field>
          <Field label="From column">
            <ColSelect
              value={cp.columnKey}
              onChange={(v) => dispatch({ type: "cp-update", id: cp.id, patch: { columnKey: v } })}
              columns={columns}
            />
          </Field>
          <Field label="Custom label (optional)">
            <Input
              className="h-8 text-sm"
              value={cp.ablLabel}
              placeholder="e.g. Skype"
              onChange={(e) =>
                dispatch({ type: "cp-update", id: cp.id, patch: { ablLabel: e.target.value } })
              }
            />
          </Field>
          <div className="flex items-end">
            <Button
              size="sm"
              variant="ghost"
              onClick={() => dispatch({ type: "cp-remove", id: cp.id })}
            >
              <Trash2 className="h-3 w-3" /> Remove
            </Button>
          </div>
        </div>
      ))}
    </Section>
  );
}

function ExtraConstantsSection({
  cfg,
  dispatch,
}: {
  cfg: MappingConfig;
  dispatch: React.Dispatch<Action>;
}) {
  return (
    <Section
      title="+ Add to every contact (constants)"
      action={
        <Button size="sm" variant="outline" onClick={() => dispatch({ type: "ex-add" })}>
          <Plus className="h-3 w-3" /> Add
        </Button>
      }
    >
      {cfg.extraConstants.length === 0 && (
        <p className="text-xs text-muted-foreground">
          Same value added to every exported contact (e.g. category "Imported 2026").
        </p>
      )}
      {cfg.extraConstants.map((ex) => (
        <div key={ex.id} className="grid gap-2 rounded-md border border-border bg-muted/30 p-3 grid-cols-2 sm:grid-cols-4">
          <Field label="Target">
            <Select
              value={ex.target}
              onValueChange={(v) =>
                dispatch({ type: "ex-update", id: ex.id, patch: { target: v as ExtraConstantField["target"] } })
              }
            >
              <SelectTrigger className="h-8 text-sm">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                <SelectItem value="note">Note</SelectItem>
                <SelectItem value="category">Category</SelectItem>
                <SelectItem value="url">URL</SelectItem>
                <SelectItem value="org">Company</SelectItem>
                <SelectItem value="title">Job title</SelectItem>
                <SelectItem value="custom">Custom X- property</SelectItem>
              </SelectContent>
            </Select>
          </Field>
          {ex.target === "custom" && (
            <Field label="Property name">
              <Input
                className="h-8 text-sm"
                value={ex.customName ?? ""}
                onChange={(e) =>
                  dispatch({ type: "ex-update", id: ex.id, patch: { customName: e.target.value } })
                }
              />
            </Field>
          )}
          <Field label="Value">
            <Input
              className="h-8 text-sm"
              value={ex.value}
              onChange={(e) => dispatch({ type: "ex-update", id: ex.id, patch: { value: e.target.value } })}
            />
          </Field>
          <div className="flex items-end">
            <Button
              size="sm"
              variant="ghost"
              onClick={() => dispatch({ type: "ex-remove", id: ex.id })}
            >
              <Trash2 className="h-3 w-3" /> Remove
            </Button>
          </div>
        </div>
      ))}
    </Section>
  );
}

export function FilterSection({
  filters,
  setFilters,
  columns,
  totalCount,
  filteredCount,
}: {
  filters: FilterRow[];
  setFilters: (f: FilterRow[]) => void;
  columns: ColumnMeta[];
  totalCount: number;
  filteredCount: number;
}) {
  const add = () =>
    setFilters([
      ...filters,
      { id: newId(), columnKey: columns[0]?.key ?? "", op: "contains", value: "" },
    ]);
  const upd = (id: string, patch: Partial<FilterRow>) =>
    setFilters(filters.map((f) => (f.id === id ? { ...f, ...patch } : f)));
  const rm = (id: string) => setFilters(filters.filter((f) => f.id !== id));
  return (
    <Section
      title="Row filters"
      icon={<FilterIcon className="h-4 w-4" />}
      action={
        <div className="flex items-center gap-3">
          <span className="text-xs text-muted-foreground">
            {filteredCount} of {totalCount} rows
          </span>
          <Button size="sm" variant="outline" onClick={add}>
            <Plus className="h-3 w-3" /> Add filter
          </Button>
        </div>
      }
    >
      {filters.length === 0 && (
        <p className="text-xs text-muted-foreground">
          Include only rows that match. All conditions must pass (AND).
        </p>
      )}
      {filters.map((f) => (
        <div key={f.id} className="flex flex-wrap items-center gap-2">
          <div className="min-w-[120px] flex-1 sm:flex-none sm:min-w-0 sm:w-36">
            <ColSelect
              value={f.columnKey}
              onChange={(v) => upd(f.id, { columnKey: v ?? "" })}
              columns={columns}
            />
          </div>
          <Select value={f.op} onValueChange={(v) => upd(f.id, { op: v as FilterRow["op"] })}>
            <SelectTrigger className="h-8 w-36 sm:w-40 text-sm">
              <SelectValue />
            </SelectTrigger>
            <SelectContent>
              <SelectItem value="contains">contains</SelectItem>
              <SelectItem value="not-contains">does not contain</SelectItem>
              <SelectItem value="equals">equals</SelectItem>
              <SelectItem value="not-equals">not equals</SelectItem>
              <SelectItem value="starts-with">starts with</SelectItem>
              <SelectItem value="ends-with">ends with</SelectItem>
              <SelectItem value="blank">is blank</SelectItem>
              <SelectItem value="not-blank">is not blank</SelectItem>
            </SelectContent>
          </Select>
          {f.op !== "blank" && f.op !== "not-blank" && (
            <Input
              className="h-8 w-full sm:w-40 text-sm"
              value={f.value}
              onChange={(e) => upd(f.id, { value: e.target.value })}
              placeholder="value"
            />
          )}
          <Button size="icon" variant="ghost" className="h-8 w-8 shrink-0" onClick={() => rm(f.id)}>
            <Trash2 className="h-3 w-3" />
          </Button>
        </div>
      ))}
    </Section>
  );
}

export function LivePreview({
  row,
  cfg,
  headerMap,
  rowIdx,
  setRowIdx,
  total,
  splitPhones,
}: {
  row: Record<string, CellValue> | null;
  cfg: MappingConfig;
  headerMap: Map<string, string>;
  rowIdx: number;
  setRowIdx: (n: number) => void;
  total: number;
  splitPhones?: boolean;
}) {
  const cell = (r: Record<string, CellValue>, key: string | null): string => {
    if (!key) return "";
    const v = r[key];
    if (v == null) return "";
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return String(v).trim();
  };
  const vcards = (() => {
    if (!row) return [""];
    if (!splitPhones) return [buildVCardText(row, cfg, headerMap)];
    const validPhones = cfg.phones.filter((e) => e.columnKey && cell(row, e.columnKey));
    if (validPhones.length <= 1) return [buildVCardText(row, cfg, headerMap)];
    return validPhones.map((pe) => {
      const header = headerMap.get(pe.columnKey!) ?? "";
      const splitCfg: MappingConfig = {
        ...cfg,
        phones: [pe],
      };
      const lastNameParts: NamePart[] = [];
      if (cfg.lastNameAssembly.length > 0) {
        lastNameParts.push(...cfg.lastNameAssembly);
      } else if (cfg.familyName) {
        lastNameParts.push({ kind: "col", columnKey: cfg.familyName });
      } else if (cfg.fullName) {
        lastNameParts.push({ kind: "col", columnKey: cfg.fullName });
      }
      if (lastNameParts.length > 0) {
        lastNameParts.push({ kind: "const", value: ` (${header})` });
        splitCfg.lastNameAssembly = lastNameParts;
        splitCfg.familyName = null;
        splitCfg.fullName = null;
      }
      return buildVCardText(row, splitCfg, headerMap);
    });
  })();

  return (
    <div className="lg:sticky lg:top-4 lg:h-fit">
      <div className="rounded-lg border border-border bg-card">
        <div className="flex items-center justify-between border-b border-border px-4 py-3">
          <div className="text-sm font-medium">Live preview</div>
          <div className="flex items-center gap-1">
            <Button
              size="icon"
              variant="ghost"
              className="h-7 w-7"
              disabled={rowIdx <= 0}
              onClick={() => setRowIdx(Math.max(0, rowIdx - 1))}
            >
              ‹
            </Button>
            <span className="text-xs text-muted-foreground">
              row {Math.min(rowIdx + 1, total)} / {total}
            </span>
            <Button
              size="icon"
              variant="ghost"
              className="h-7 w-7"
              disabled={rowIdx >= total - 1}
              onClick={() => setRowIdx(Math.min(total - 1, rowIdx + 1))}
            >
              ›
            </Button>
          </div>
        </div>
        <Tabs defaultValue="card">
          <TabsList className="mx-4 mt-3">
            <TabsTrigger value="card">Card</TabsTrigger>
            <TabsTrigger value="raw">.vcf</TabsTrigger>
          </TabsList>
          <TabsContent value="card" className="p-4">
            {row ? (
              <div className="space-y-3">
                {vcards.map((vcf, i) => (
                  <ContactCard key={i} vcf={vcf} />
                ))}
              </div>
            ) : (
              <p className="text-sm text-muted-foreground">No rows to preview.</p>
            )}
          </TabsContent>
          <TabsContent value="raw" className="p-4">
            <pre className="max-h-[300px] sm:max-h-[500px] overflow-auto rounded-md bg-muted p-3 font-mono text-xs">
              {vcards.join("\n") || "—"}
            </pre>
          </TabsContent>
        </Tabs>
      </div>
    </div>
  );
}

export function ContactCard({ vcf }: { vcf: string }) {
  const lines = vcf.split(/\r?\n/).filter(Boolean);

  const groupLabels = new Map<string, string>();
  for (const l of lines) {
    const dot = l.indexOf(".");
    if (dot > 0) {
      const rest = l.slice(dot + 1).toUpperCase();
      if (rest.startsWith("X-ABLABEL:")) {
        groupLabels.set(l.slice(0, dot), rest.slice(rest.indexOf(":") + 1));
      }
    }
  }

  const get = (key: string) => {
    const l = lines.find((l) => {
      const up = l.toUpperCase();
      return up.startsWith(key + ":") || up.startsWith(key + ";");
    });
    if (!l) return "";
    const idx = l.indexOf(":");
    return idx >= 0 ? l.slice(idx + 1) : "";
  };

  const getAll = (prefix: string): { label: string; value: string }[] => {
    const out: { label: string; value: string }[] = [];
    for (const l of lines) {
      const up = l.toUpperCase();
      if (up.startsWith(prefix + ":") || up.startsWith(prefix + ";")) {
        const idx = l.indexOf(":");
        const head = l.slice(0, idx);
        const val = l.slice(idx + 1);
        const typeMatch = head.match(/TYPE=([^;:]+)/i);
        out.push({ label: typeMatch?.[1] ?? "", value: val });
      } else {
        const dot = l.indexOf(".");
        if (dot > 0) {
          const group = l.slice(0, dot);
          const propUp = l.slice(dot + 1).toUpperCase();
          if (propUp.startsWith(prefix + ":") || propUp.startsWith(prefix + ";")) {
            const idx = l.indexOf(":", dot + 1);
            const val = idx >= 0 ? l.slice(idx + 1) : "";
            out.push({ label: groupLabels.get(group) ?? "", value: val });
          }
        }
      }
    }
    return out;
  };

  const getAllDates = (): { label: string; value: string }[] => {
    const out: { label: string; value: string }[] = [];
    for (const l of lines) {
      const up = l.toUpperCase();
      if (up.startsWith("BDAY:") || up.startsWith("BDAY;")) {
        const idx = l.indexOf(":");
        if (idx >= 0) out.push({ label: "BIRTHDAY", value: l.slice(idx + 1) });
      } else if (up.includes("X-ABDATE")) {
        const dot = l.indexOf(".");
        let value = "";
        let label = "";
        if (dot > 0) {
          const group = l.slice(0, dot);
          const idx = l.indexOf(":", dot + 1);
          value = idx >= 0 ? l.slice(idx + 1) : "";
          label = groupLabels.get(group) ?? "";
        } else {
          const idx = l.indexOf(":");
          value = idx >= 0 ? l.slice(idx + 1) : "";
        }
        if (value) out.push({ label, value });
      }
    }
    return out;
  };

  const fn = get("FN") || "(no name)";
  const org = get("ORG");
  const title = get("TITLE");
  const role = get("ROLE");
  const phones = getAll("TEL");
  const emails = getAll("EMAIL");
  const urls = getAll("URL");
  const adrs = getAll("ADR");
  const dates = getAllDates();
  const categories = get("CATEGORIES");
  const note = get("NOTE");
  const initials = fn
    .split(/\s+/)
    .map((s) => s[0])
    .filter(Boolean)
    .slice(0, 2)
    .join("")
    .toUpperCase();

  function parseAdr(val: string) {
    const parts = val.split(";");
    const labels = ["", "", "street", "city", "region", "zip", "country"];
    return parts
      .map((p, i) => ({ label: labels[i], value: p }))
      .filter((p) => p.label && p.value);
  }

  return (
    <div className="space-y-3">
      <div className="flex items-center gap-3">
        <div className="flex h-12 w-12 items-center justify-center rounded-full bg-primary text-lg font-semibold text-primary-foreground">
          {initials || "?"}
        </div>
        <div className="min-w-0">
          <div className="truncate font-medium">{fn}</div>
          {(title || role || org) && (
            <div className="truncate text-xs text-muted-foreground">
              {[title, role, org].filter(Boolean).join(" · ")}
            </div>
          )}
        </div>
      </div>
      {phones.length > 0 && (
        <div className="space-y-1">
          {phones.map((p, i) => (
            <Row key={i} icon={<Phone className="h-3 w-3" />} label={p.label} value={p.value} />
          ))}
        </div>
      )}
      {emails.length > 0 && (
        <div className="space-y-1">
          {emails.map((e, i) => (
            <Row key={i} icon={<Mail className="h-3 w-3" />} label={e.label} value={e.value} />
          ))}
        </div>
      )}
      {urls.length > 0 && (
        <div className="space-y-1">
          {urls.map((u, i) => (
            <Row key={i} icon={<LinkIcon className="h-3 w-3" />} label={u.label} value={u.value} />
          ))}
        </div>
      )}
      {adrs.length > 0 && (
        <div className="space-y-1">
          {adrs.map((a, i) => {
            const parts = parseAdr(a.value);
            return (
              <Row
                key={i}
                icon={<MapPin className="h-3 w-3" />}
                label={a.label}
                value={parts.map((p) => p.value).filter(Boolean).join(", ")}
              />
            );
          })}
        </div>
      )}
      {dates.length > 0 && (
        <div className="space-y-1">
          {dates.map((d, i) => (
            <Row key={i} icon={<Calendar className="h-3 w-3" />} label={d.label} value={d.value} />
          ))}
        </div>
      )}
      {categories && (
        <div className="space-y-1">
          <Row icon={<Tag className="h-3 w-3" />} label="" value={categories} />
        </div>
      )}
      {note && (
        <div className="rounded-md bg-muted/50 px-2 py-1">
          <p className="text-xs text-muted-foreground">{note}</p>
        </div>
      )}
    </div>
  );
}

function Row({ icon, label, value }: { icon: React.ReactNode; label: string; value: string }) {
  return (
    <div className="flex items-center gap-2 rounded-md bg-muted/50 px-2 py-1 text-sm">
      <span className="text-muted-foreground">{icon}</span>
      {label && (
        <Badge variant="outline" className="text-[10px]">
          {label}
        </Badge>
      )}
      <span className="truncate">{value}</span>
    </div>
  );
}

export function ExportStep({
  count,
  fileName,
  onFileNameChange,
  onBack,
  onDownload,
  onReset,
  previewRow,
  cfg,
  headerMap,
  splitPhones,
  rowIdx,
  setRowIdx,
  total,
}: {
  count: number;
  fileName: string;
  onFileNameChange: (v: string) => void;
  onBack: () => void;
  onDownload: () => void;
  onReset: () => void;
  previewRow: Record<string, CellValue> | null;
  cfg: MappingConfig;
  headerMap: Map<string, string>;
  splitPhones: boolean;
  rowIdx: number;
  setRowIdx: (n: number) => void;
  total: number;
}) {
  const ext = ".vcf";
  const cell = (row: Record<string, CellValue>, key: string | null): string => {
    if (!key) return "";
    const v = row[key];
    if (v == null) return "";
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    return String(v).trim();
  };
  const vcards = (() => {
    if (!previewRow) return [""];
    if (!splitPhones) return [buildVCardText(previewRow, cfg, headerMap)];
    const validPhones = cfg.phones.filter((e) => e.columnKey && cell(previewRow, e.columnKey));
    if (validPhones.length <= 1) return [buildVCardText(previewRow, cfg, headerMap)];
    return validPhones.map((pe) => {
      const header = headerMap.get(pe.columnKey!) ?? "";
      const splitCfg: MappingConfig = {
        ...cfg,
        phones: [pe],
      };
      const lastNameParts: NamePart[] = [];
      if (cfg.lastNameAssembly.length > 0) {
        lastNameParts.push(...cfg.lastNameAssembly);
      } else if (cfg.familyName) {
        lastNameParts.push({ kind: "col", columnKey: cfg.familyName });
      } else if (cfg.fullName) {
        lastNameParts.push({ kind: "col", columnKey: cfg.fullName });
      }
      if (lastNameParts.length > 0) {
        lastNameParts.push({ kind: "const", value: ` (${header})` });
        splitCfg.lastNameAssembly = lastNameParts;
        splitCfg.familyName = null;
        splitCfg.fullName = null;
      }
      return buildVCardText(previewRow, splitCfg, headerMap);
    });
  })();
  return (
    <div className="grid gap-6 lg:grid-cols-[1fr_400px] py-8">
      <div className="flex flex-col items-center justify-center">
        <div className="w-full max-w-md rounded-lg border border-border bg-card p-8 text-center">
          <div className="mx-auto mb-4 flex h-12 w-12 items-center justify-center rounded-full bg-primary/10 text-primary">
            <Download className="h-6 w-6" />
          </div>
          <h2 className="text-lg font-semibold">Ready to export</h2>
          <p className="mt-1 text-sm text-muted-foreground">
            {count} contact{count === 1 ? "" : "s"}{splitPhones ? ` (split into ${vcards.length} per row)` : ""} in a single .vcf file
          </p>
          <div className="mt-4 text-left">
            <Label className="text-xs">File name</Label>
            <div className="mt-1 flex items-center gap-0">
              <Input
                className="h-9 flex-1 rounded-r-none text-sm"
                value={fileName}
                onChange={(e) => onFileNameChange(e.target.value)}
              />
              <div className="flex h-9 items-center rounded-md rounded-l-none border border-l-0 border-input bg-muted px-3 text-sm text-muted-foreground">
                {ext}
              </div>
            </div>
          </div>
          <Button className="mt-4 w-full" onClick={onDownload}>
            <Download className="h-4 w-4" /> Download{fileName ? ` ${fileName}${ext}` : ""}
          </Button>
          <div className="mt-3 flex gap-2">
            <Button variant="outline" className="flex-1" onClick={onBack}>
              <ArrowLeft className="h-4 w-4" /> Back
            </Button>
            <Button variant="ghost" className="flex-1" onClick={onReset}>
              New file
            </Button>
          </div>
        </div>
      </div>
      <div className="space-y-3">
        <div className="rounded-lg border border-border bg-card">
          <div className="flex items-center justify-between border-b border-border px-4 py-3">
            <div className="text-sm font-medium">Live preview</div>
            <div className="flex items-center gap-1">
              <Button
                size="icon"
                variant="ghost"
                className="h-7 w-7"
                disabled={rowIdx <= 0}
                onClick={() => setRowIdx(Math.max(0, rowIdx - 1))}
              >
                ‹
              </Button>
              <span className="text-xs text-muted-foreground">
                row {Math.min(rowIdx + 1, total)} / {total}
              </span>
              <Button
                size="icon"
                variant="ghost"
                className="h-7 w-7"
                disabled={rowIdx >= total - 1}
                onClick={() => setRowIdx(Math.min(total - 1, rowIdx + 1))}
              >
                ›
              </Button>
            </div>
          </div>
          <Tabs defaultValue="card">
            <TabsList className="mx-4 mt-3">
              <TabsTrigger value="card">Card</TabsTrigger>
              <TabsTrigger value="raw">.vcf</TabsTrigger>
            </TabsList>
            <TabsContent value="card" className="p-4">
              <div className="space-y-3">
                {vcards.map((vcf, i) => (
                  <ContactCard key={i} vcf={vcf} />
                ))}
              </div>
            </TabsContent>
            <TabsContent value="raw" className="p-4">
              <pre className="max-h-[300px] sm:max-h-[500px] overflow-auto rounded-md bg-muted p-3 font-mono text-xs">
                {vcards.join("\n") || "—"}
              </pre>
            </TabsContent>
          </Tabs>
        </div>
      </div>
    </div>
  );
}
