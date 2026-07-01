# XLS → vCard Web App

React SPA on the existing TanStack Start template. No backend, no auth, no history — everything runs in the browser. Drop a file, configure, download `.vcf`, leave.

## Flow (single-page, stepped)

```
1. Drop file (.xls / .xlsx / .csv)
        ↓
2. Sheet & Header Preview
   - Pick which sheet (if multiple)
   - Skip N top rows (handles merged title rows)
   - Toggle: "First row is header" (on = use row values as column names; off = generic Col A/B/C)
   - Live table preview with detected columns
        ↓
3. Mapping (Simple ⇄ Advanced toggle)
   - Simple: auto-detected fields shown as chips, one-click tweak
   - Advanced: full control (see below)
   - LIVE vCard preview panel on the right — updates as you map,
     showing the first row rendered as a real contact card + raw .vcf text
        ↓
4. Filters (optional, collapsible)
   - Per-column rows: equals / contains / blank / not-blank / starts-with / ends-with
   - AND logic across rows
   - Row count updates live
        ↓
5. Export
   - "Download contacts.vcf" — single multi-contact file
   - Shows count of contacts about to export
```

## Advanced Mapping Capabilities

- **Name Assembly** — build FN from an ordered list of parts. Each part is either a column reference or a typed constant word. Example: `[const "Dr. "] + [col FirstName] + [const " "] + [col LastName] + [const " — "] + [col Company]`. Structured N (givenName / familyName / additionalNames / prefix / suffix) mapped separately.
- **Phones / Emails / Dates** — unlimited entries. Two label strategies per group:
  - *Suffix parse*: `Phone_Work` → value with label `WORK`
  - *Label column*: pick a column that holds the type string
  - Custom label override per entry
- **Org**: company, department, title, role
- **Address**: multi-instance, structured (street/locality/region/postal/country) with type
- **URLs, Categories, Notes**
- **Custom vCard properties** for anything unmapped (`X-*` keys, with optional Apple-style grouped labels for custom-label fields)
- **+ New contact** section at bottom of mapping — add extra fields to every generated card (constants across the whole export)

## Live Preview Panel

Right side, sticky. Shows the currently-selected preview row (dropdown to pick which row: 1st, 2nd, random). Two tabs:
- **Card view** — rendered like a phone contact (name, avatar initial, phone/email/date rows with labels)
- **Raw .vcf** — actual generated vCard text, syntax-highlighted

Updates on every mapping change — no "generate preview" button.

## Tech

- Existing template (TanStack Start + React 19 + Tailwind v4 + shadcn)
- Add deps: `xlsx` (SheetJS), `vcard-creator`
- Single route `/` — replaces the placeholder in `src/routes/index.tsx`
- State: local component state + `useReducer` for mapping config; no global store, no persistence
- File parsing entirely client-side via SheetJS `read()`
- vCard generation client-side via `vcard-creator`, `Blob` + `URL.createObjectURL` for download

## File Structure

```
src/
  routes/index.tsx                    (mounts <XlsToVcardApp />)
  routes/__root.tsx                   (update title/meta)
  components/xls-vcard/
    XlsToVcardApp.tsx                 (top-level, step state)
    steps/
      DropStep.tsx
      SheetPreviewStep.tsx            (sheet select, skip rows, header toggle, table)
      MappingStep.tsx                 (Simple/Advanced toggle, live preview split)
      FilterStep.tsx
      ExportStep.tsx
    mapping/
      SimpleMapping.tsx
      AdvancedMapping.tsx
      NameAssembler.tsx               (drag-orderable parts, add column/const)
      MultiFieldMapper.tsx            (phones/emails/dates)
      OrgMapper.tsx
      AddressMapper.tsx
      CustomPropsMapper.tsx
      ExtraFieldsMapper.tsx           (+ New contact constants)
    preview/
      LivePreviewPanel.tsx
      ContactCard.tsx
      VcfRaw.tsx
  lib/xls-vcard/
    types.ts
    parseXlsx.ts                      (SheetJS wrapper: sheets, skipRows, headerToggle)
    autoDetect.ts                     (fuzzy match headers → vCard fields)
    buildVcard.ts                     (config + row → VCard instance)
    applyFilters.ts
    labelParse.ts                     (suffix → label helper)
```

## Design

Clean, minimal, utility-app aesthetic — think Linear / Raycast. Neutral grays, one accent color, generous spacing, monospace for the raw .vcf tab. Left = configure, right = live preview (desktop). Stacked on mobile with sticky preview toggle. All colors as semantic tokens in `src/styles.css`, no hardcoded classes.

## Out of Scope

- No dashboard, history, saved templates, or stats
- No user accounts, no cloud storage
- No CSV export back (vCard only)
- No editing individual generated contacts (edit the source xls if wrong)
