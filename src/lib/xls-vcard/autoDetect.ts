import type { ColumnMeta, MappingConfig, MultiFieldEntry } from "./types";

const norm = (s: string) =>
  s.toLowerCase().replace(/[^a-z0-9]/g, " ").replace(/\s+/g, " ").trim();

const tokens = (s: string) => norm(s).split(/\s+/).filter(Boolean);

const hasWords = (header: string, ...words: string[]) => {
  const t = tokens(header);
  return words.every((w) => t.includes(w));
};

const hasAnyWord = (header: string, ...words: string[]) => {
  const t = tokens(header);
  return words.some((w) => t.includes(w));
};

const isExactly = (header: string, ...phrases: string[]) => {
  const n = norm(header);
  return phrases.includes(n);
};

const EXCLUDED = new Set([
  "company", "organization", "org", "employer",
  "email", "mail",
  "phone", "mobile", "cell", "tel", "telephone", "fax",
  "address", "city", "state", "zip", "postal", "postcode", "country",
  "website", "url", "site", "homepage",
  "note", "comment", "remark",
  "price", "cost", "amount", "total", "rate",
  "id", "code", "number", "no", "num",
  "date", "birthday", "bday", "dob",
  "category", "tag", "group", "label",
  "department", "dept",
  "job", "position", "role",
  "status", "type", "source", "created", "updated",
]);

const hasExcluded = (header: string) => hasAnyWord(header, ...EXCLUDED);

function findBest(
  cols: ColumnMeta[],
  matchers: ((h: string) => boolean)[],
): ColumnMeta | null {
  for (const m of matchers) {
    const hit = cols.find((c) => m(c.header));
    if (hit) return hit;
  }
  return null;
}

const givenMatchers: ((h: string) => boolean)[] = [
  (h) =>
    isExactly(
      h,
      "first name", "first", "fname", "firstname",
      "given name", "given", "givenname",
      "forename", "vorname", "prenom", "nombre",
    ),
  (h) => hasWords(h, "first", "name") && !hasExcluded(h),
  (h) => hasWords(h, "given", "name") && !hasExcluded(h),
  (h) => hasAnyWord(h, "first", "fname", "given", "forename", "vorname", "prenom") && !hasExcluded(h),
];

const familyMatchers: ((h: string) => boolean)[] = [
  (h) =>
    isExactly(
      h,
      "last name", "last", "lname", "lastname",
      "family name", "family", "familyname",
      "surname", "sur", "nachname", "nom", "apellido", "sobrenome",
    ),
  (h) => hasWords(h, "last", "name") && !hasExcluded(h),
  (h) => hasWords(h, "family", "name") && !hasExcluded(h),
  (h) => hasAnyWord(h, "last", "lname", "family", "surname", "sur", "nachname", "apellido") && !hasExcluded(h),
];

const fullMatchers: ((h: string) => boolean)[] = [
  (h) =>
    isExactly(
      h,
      "full name", "fullname", "name",
      "contact name", "contactname",
      "display name", "displayname",
      "customer name", "client name",
      "person name", "participant name",
      "member name", "employee name",
    ),
  (h) => hasWords(h, "full", "name") && !hasExcluded(h),
  (h) => hasWords(h, "contact", "name") && !hasExcluded(h),
  (h) => hasWords(h, "display", "name") && !hasExcluded(h),
  (h) => hasWords(h, "customer", "name") && !hasExcluded(h),
  (h) => hasWords(h, "client", "name") && !hasExcluded(h),
  (h) => hasWords(h, "person", "name") && !hasExcluded(h),
  (h) => hasWords(h, "participant", "name") && !hasExcluded(h),
  (h) => hasWords(h, "member", "name") && !hasExcluded(h),
  (h) => hasWords(h, "employee", "name") && !hasExcluded(h),
  (h) => isExactly(h, "name") && !hasExcluded(h),
];

const nameExcluded = (h: string) =>
  hasAnyWord(h, "company", "org", "email", "phone", "address", "website", "date", "note", "category", "department", "job", "role", "price", "cost", "id", "code", "status", "type");

const fallbackMatchers: ((h: string) => boolean)[] = [
  (h) => {
    const t = tokens(h);
    return t.length <= 3 && !nameExcluded(h);
  },
];

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
  const given = findBest(cols, givenMatchers);
  const family = findBest(cols, familyMatchers);
  const full = findBest(cols, fullMatchers);
  const middle = findCol(cols, [/middle/, /initial/i]);
  const prefix = findCol(cols, [/^(prefix|title|salutation)$/]);
  const suffix = findCol(cols, [/^suffix$/]);

  const phones = findAll(cols, [
    /phone|mobile|cell|tel|fax|contact/i,
  ]).filter(
    (c) =>
      !hasWords(c.header, "email", "mail") &&
      !hasWords(c.header, "address") &&
      !hasWords(c.header, "note"),
  ).map(toEntry);
  const emails = findAll(cols, [/email|mail/i]).map(toEntry);
  const dates = findAll(cols, [/birthday|bday|dob|anniversary|date/i]).map(toEntry);
  const urls = findAll(cols, [/website|url|site|homepage/i]).map(toEntry);

  const company = findCol(cols, [/company|organization|org|employer|business/i]);
  const department = findCol(cols, [/department|dept/i]);
  const jobTitle = findCol(cols, [/job\s?title|title|position|designation/i]);
  const role = findCol(cols, [/^role$/i]);
  const categories = findCol(cols, [/categor|tag|group/i]);
  const note = findCol(cols, [/note|comment|remark/i]);

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
