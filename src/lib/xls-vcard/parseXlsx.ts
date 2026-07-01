import * as XLSX from "xlsx";
import type { CellValue, ParsedSheet, WorkbookData } from "./types";

export async function readWorkbook(file: File): Promise<WorkbookData> {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array", cellDates: true });
  return {
    raw: wb,
    sheets: wb.SheetNames.map((name) => {
      const ws = wb.Sheets[name];
      const ref = ws["!ref"];
      const range = ref ? XLSX.utils.decode_range(ref) : { s: { r: 0 }, e: { r: 0 } };
      return { name, rowCount: range.e.r - range.s.r + 1 };
    }),
  };
}

function colLetter(i: number): string {
  let s = "";
  let n = i;
  while (n >= 0) {
    s = String.fromCharCode((n % 26) + 65) + s;
    n = Math.floor(n / 26) - 1;
  }
  return s;
}

export function parseSheet(
  wb: WorkbookData,
  sheetName: string,
  opts: { skipRows: number; firstRowIsHeader: boolean }
): ParsedSheet {
  const rawWb = wb.raw as XLSX.WorkBook;
  const ws = rawWb.Sheets[sheetName];
  if (!ws) return { columns: [], rows: [] };
  const arr = XLSX.utils.sheet_to_json<CellValue[]>(ws, {
    header: 1,
    defval: null,
    blankrows: false,
    raw: true,
  });
  const sliced = arr.slice(opts.skipRows);
  if (sliced.length === 0) return { columns: [], rows: [] };

  const width = Math.max(...sliced.map((r) => r.length));
  let headers: string[];
  let dataRows: CellValue[][];
  if (opts.firstRowIsHeader) {
    headers = Array.from({ length: width }, (_, i) => {
      const raw = sliced[0][i];
      const s = raw == null ? "" : String(raw).trim();
      return s || `Column ${colLetter(i)}`;
    });
    dataRows = sliced.slice(1);
  } else {
    headers = Array.from({ length: width }, (_, i) => `Column ${colLetter(i)}`);
    dataRows = sliced;
  }

  const columns = headers.map((header, index) => ({
    key: `c${index}`,
    header,
    index,
  }));

  const rows = dataRows.map((r) => {
    const obj: Record<string, CellValue> = {};
    for (const c of columns) {
      const v = r[c.index];
      obj[c.key] = v === undefined ? null : v;
    }
    return obj;
  });

  // Remove all-empty rows
  const nonEmpty = rows.filter((r) =>
    Object.values(r).some((v) => v !== null && v !== "" && v !== undefined)
  );

  return { columns, rows: nonEmpty };
}
