import type { CellValue, FilterRow } from "./types";

function s(v: CellValue): string {
  if (v == null) return "";
  if (v instanceof Date) return v.toISOString();
  return String(v);
}

export function applyFilters(
  rows: Record<string, CellValue>[],
  filters: FilterRow[]
): Record<string, CellValue>[] {
  const active = filters.filter((f) => f.columnKey);
  if (active.length === 0) return rows;
  return rows.filter((row) =>
    active.every((f) => {
      const cell = s(row[f.columnKey]).toLowerCase();
      const val = f.value.toLowerCase();
      switch (f.op) {
        case "equals":
          return cell === val;
        case "not-equals":
          return cell !== val;
        case "contains":
          return cell.includes(val);
        case "not-contains":
          return !cell.includes(val);
        case "blank":
          return cell === "";
        case "not-blank":
          return cell !== "";
        case "starts-with":
          return cell.startsWith(val);
        case "ends-with":
          return cell.endsWith(val);
      }
    })
  );
}
