import VCard from "vcard-creator";
import type {
  CellValue,
  MappingConfig,
  MultiFieldEntry,
  NamePart,
} from "./types";

function cell(row: Record<string, CellValue>, key: string | null): string {
  if (!key) return "";
  const v = row[key];
  if (v == null) return "";
  if (v instanceof Date) return v.toISOString().slice(0, 10);
  return String(v).trim();
}

function assemble(row: Record<string, CellValue>, parts: NamePart[]): string {
  return parts
    .map((p) => (p.kind === "const" ? p.value : cell(row, p.columnKey)))
    .join("")
    .trim();
}

function entryLabel(row: Record<string, CellValue>, e: MultiFieldEntry, headerFallback: string): string {
  if (e.labelStrategy === "column" && e.labelColumnKey) return cell(row, e.labelColumnKey).toUpperCase();
  if (e.labelStrategy === "suffix") {
    const m = headerFallback.match(/[_\-\s]+(.+)$/);
    if (m) return m[1].toUpperCase().replace(/[^A-Z0-9]/g, "");
  }
  return e.customLabel.toUpperCase();
}

function toDate(v: string): string {
  // best effort YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}/.test(v)) return v.slice(0, 10);
  const d = new Date(v);
  if (!isNaN(d.getTime())) return d.toISOString().slice(0, 10);
  return v;
}

export function buildVCard(
  row: Record<string, CellValue>,
  cfg: MappingConfig,
  headers: Map<string, string>
): VCard {
  const card = new VCard();

  const given = cell(row, cfg.givenName);
  const family = cell(row, cfg.familyName);
  const middle = cell(row, cfg.additionalNames);
  const pre = cell(row, cfg.honorificPrefix);
  const suf = cell(row, cfg.honorificSuffix);

  // FN must be set before addName() — addName auto-generates FN and will
  // skip it if FN is already set (checks hasProperty('FN') internally).
  const fn = assemble(row, cfg.nameAssembly);
  if (fn) card.addFullName(fn);
  else if (!given && !family) {
    const first = Object.values(row).find((v) => v != null && v !== "");
    if (first) card.addFullName(String(first));
  }

  if (given || family || middle || pre || suf) {
    card.addName({
      givenName: given || undefined,
      familyName: family || undefined,
      additionalNames: middle || undefined,
      honorificPrefix: pre || undefined,
      honorificSuffix: suf || undefined,
    });
  }

  // Org
  const company = cell(row, cfg.company);
  const dept = cell(row, cfg.department);
  if (company || dept) card.addCompany({ name: company || " ", department: dept || undefined });
  const jt = cell(row, cfg.jobTitle);
  if (jt) card.addJobtitle(jt);
  const role = cell(row, cfg.role);
  if (role) card.addRole(role);

  // Phones
  let itemCounter = 0;
  for (const e of cfg.phones) {
    const v = cell(row, e.columnKey);
    if (!v) continue;
    const label = entryLabel(row, e, headers.get(e.columnKey!) ?? "");
    const std = ["PREF", "WORK", "HOME", "VOICE", "FAX", "CELL", "MOBILE", "MAIN"];
    if (label && !std.includes(label)) {
      const g = `item${++itemCounter}`;
      card.addCustomProperty({ name: "TEL", value: v, group: g });
      card.addCustomProperty({ name: "X-ABLabel", value: label, group: g });
    } else {
      const type = label ? [label.toLowerCase() as "cell"] : undefined;
      card.addPhoneNumber({ number: v, type });
    }
  }

  // Emails
  for (const e of cfg.emails) {
    const v = cell(row, e.columnKey);
    if (!v) continue;
    const label = entryLabel(row, e, headers.get(e.columnKey!) ?? "");
    const std = ["PREF", "WORK", "HOME", "INTERNET"];
    if (label && !std.includes(label)) {
      const g = `item${++itemCounter}`;
      card.addCustomProperty({ name: "EMAIL", value: v, group: g });
      card.addCustomProperty({ name: "X-ABLabel", value: label, group: g });
    } else {
      const type = label ? [label.toLowerCase() as "work"] : undefined;
      card.addEmail({ address: v, type });
    }
  }

  // URLs
  for (const e of cfg.urls) {
    const v = cell(row, e.columnKey);
    if (!v) continue;
    card.addUrl({ url: v });
  }

  // Dates
  for (const e of cfg.dates) {
    const v = cell(row, e.columnKey);
    if (!v) continue;
    const label = entryLabel(row, e, headers.get(e.columnKey!) ?? "");
    const iso = toDate(v);
    if (/BIRTH|BDAY|DOB/.test(label) || label === "") {
      try {
        card.addBirthday(iso as `${number}-${number}-${number}`);
      } catch {
        /* ignore */
      }
    } else {
      const g = `item${++itemCounter}`;
      card.addCustomProperty({ name: "X-ABDATE", value: iso, group: g });
      card.addCustomProperty({ name: "X-ABLabel", value: label, group: g });
    }
  }

  // Addresses
  for (const a of cfg.addresses) {
    const parts = {
      street: cell(row, a.street),
      locality: cell(row, a.locality),
      region: cell(row, a.region),
      postalCode: cell(row, a.postalCode),
      country: cell(row, a.country),
    };
    if (Object.values(parts).some(Boolean)) {
      card.addAddress({
        street: parts.street || undefined,
        locality: parts.locality || undefined,
        region: parts.region || undefined,
        postalCode: parts.postalCode || undefined,
        country: parts.country || undefined,
      });
    }
  }

  // Categories / note
  const cats = cell(row, cfg.categories);
  if (cats) card.addCategories(cats.split(/[,;]/).map((s) => s.trim()).filter(Boolean));
  const note = cell(row, cfg.note);
  if (note) card.addNote(note);

  // Custom
  for (const cp of cfg.customProps) {
    const v = cp.columnKey ? cell(row, cp.columnKey) : cp.constantValue;
    if (!v) continue;
    if (cp.ablLabel) {
      const g = `item${++itemCounter}`;
      card.addCustomProperty({ name: cp.name, value: v, group: g });
      card.addCustomProperty({ name: "X-ABLabel", value: cp.ablLabel, group: g });
    } else {
      card.addCustomProperty({ name: cp.name, value: v });
    }
  }

  // Extra constants
  for (const ex of cfg.extraConstants) {
    if (!ex.value) continue;
    switch (ex.target) {
      case "note":
        card.addNote(ex.value);
        break;
      case "category":
        card.addCategories([ex.value]);
        break;
      case "url":
        card.addUrl({ url: ex.value });
        break;
      case "org":
        card.addCompany({ name: ex.value });
        break;
      case "title":
        card.addJobtitle(ex.value);
        break;
      case "custom":
        if (ex.customName) card.addCustomProperty({ name: ex.customName, value: ex.value });
        break;
    }
  }

  return card;
}

export function buildVCardText(
  row: Record<string, CellValue>,
  cfg: MappingConfig,
  headers: Map<string, string>
): string {
  try {
    return buildVCard(row, cfg, headers).toString();
  } catch (e) {
    return `# error: ${(e as Error).message}`;
  }
}

export function buildAllVCards(
  rows: Record<string, CellValue>[],
  cfg: MappingConfig,
  headers: Map<string, string>
): string {
  return rows.map((r) => buildVCardText(r, cfg, headers)).join("\n");
}
