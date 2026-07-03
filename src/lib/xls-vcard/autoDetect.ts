import type { ColumnMeta, MappingConfig, MultiFieldEntry } from "./types";

const norm = (s: string) => s.toLowerCase().replace(/[^a-z0-9]/g, "");

function findCol(cols: ColumnMeta[], patterns: RegExp[]): ColumnMeta | null {
  for (const p of patterns) {
    const hit = cols.find((c) => p.test(norm(c.header)));
    if (hit) return hit;
  }
  return null;
}

function findAll(cols: ColumnMeta[], patterns: RegExp[]): ColumnMeta[] {
  return cols.filter((c) => patterns.some((p) => p.test(norm(c.header))));
}

function labelFromHeader(header: string): string {
  return header.trim().replace(/[^\w\s-]/g, "");
}

let idc = 0;
const nid = () => `e${++idc}_${Date.now().toString(36)}`;

function toEntry(col: ColumnMeta): MultiFieldEntry {
  return {
    id: nid(),
    columnKey: col.key,
    labelStrategy: "custom",
    labelColumnKey: null,
    customLabel: labelFromHeader(col.header),
  };
}

export function autoDetect(cols: ColumnMeta[]): MappingConfig {
  const given = findCol(cols, [/^(firstname|fname|givenname|first)$/, /^first/]);
  const family = findCol(cols, [/^(lastname|lname|familyname|surname|last)$/, /^last/]);
  const full = findCol(cols, [/^(fullname|name|contactname|displayname)$/]);
  const middle = findCol(cols, [/middle/]);
  const prefix = findCol(cols, [/^(prefix|title|salutation)$/]);
  const suffix = findCol(cols, [/^suffix$/]);

  const phones = findAll(cols, [/phone|mobile|cell|tel|fax/]).map(toEntry);
  const emails = findAll(cols, [/email|mail/]).map(toEntry);
  const dates = findAll(cols, [/birthday|bday|dob|anniversary|date/]).map(toEntry);
  const urls = findAll(cols, [/website|url|site|homepage/]).map(toEntry);

  const company = findCol(cols, [/company|organization|org|employer/]);
  const department = findCol(cols, [/department|dept/]);
  const jobTitle = findCol(cols, [/jobtitle|title|position/]);
  const role = findCol(cols, [/^role$/]);
  const categories = findCol(cols, [/categor|tag|group/]);
  const note = findCol(cols, [/note|comment|remark/]);

  const firstNameAssembly: MappingConfig["firstNameAssembly"] = given
    ? [{ kind: "col" as const, columnKey: given.key }]
    : [];
  const lastNameAssembly: MappingConfig["lastNameAssembly"] = family
    ? [{ kind: "col" as const, columnKey: family.key }]
    : [];

  return {
    mode: "simple",
    nameAssembly: [],
    firstNameAssembly,
    lastNameAssembly,
    fullName: full?.key ?? null,
    givenName: given?.key ?? null,
    familyName: family?.key ?? null,
    additionalNames: middle?.key ?? null,
    honorificPrefix: prefix?.key ?? null,
    honorificSuffix: suffix?.key ?? null,
    phones,
    emails,
    dates,
    urls,
    company: company?.key ?? null,
    department: department?.key ?? null,
    jobTitle: jobTitle?.key ?? null,
    role: role?.key ?? null,
    addresses: [],
    categories: categories?.key ?? null,
    note: note?.key ?? null,
    customProps: [],
    extraConstants: [],
  };
}

export function emptyMultiEntry(): MultiFieldEntry {
  return {
    id: nid(),
    columnKey: null,
    labelStrategy: "custom",
    labelColumnKey: null,
    customLabel: "",
  };
}

export function newId() {
  return nid();
}
