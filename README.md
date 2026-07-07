# XLS → vCard Magic

A client-side web app that converts spreadsheet files (XLSX, XLS, CSV) into vCard (.vcf) contact files — entirely in the browser, no uploads or servers.

Built with [TanStack Start](https://start.tanstack.com/), React 19, Tailwind CSS v4, and shadcn/ui.

## Features

- **Drag-and-drop** file upload (`.xlsx`, `.xls`, `.csv`)
- **Intelligent column auto-detection** — fuzzy matches headers to vCard fields (name, phone, email, address, org, etc.)
- **Full mapping control** — Simple (auto-detect & override) or Advanced (drag-and-drop name assembly, multi-field phones/emails/dates, custom vCard properties)
- **Live preview** — see the rendered contact card and raw .vcf text as you configure
- **Filters** — include/exclude rows by column conditions
- **vCard export** — single multi-contact `.vcf` download
- **100% client-side** — your data never leaves your machine

## Quick Start

```bash
bun install
bun dev
```

Open `http://localhost:3000` in your browser.

## Development

```bash
bun dev          # Start dev server
bun build        # Production build
bun preview      # Preview production build
bun lint         # Run ESLint
bun format       # Format with Prettier
```

## Tech Stack

| Tool | Purpose |
| --- | --- |
| [TanStack Start](https://start.tanstack.com/) | Meta-framework (Vite + React Router) |
| [React 19](https://react.dev/) | UI library |
| [Tailwind CSS v4](https://tailwindcss.com/) | Styling |
| [shadcn/ui](https://ui.shadcn.com/) | Component primitives |
| [SheetJS (xlsx)](https://sheetjs.com/) | Spreadsheet parsing |
| [vcard-creator](https://github.com/CSpadinger/node-vcard-creator) | vCard generation |

## How It Works

1. **Upload** a spreadsheet file
2. **Preview** — pick a sheet, skip header rows, confirm column detection
3. **Map** columns to vCard fields (auto-detected, or manually configure)
4. **Filter** rows if needed
5. **Export** — download a `.vcf` file ready to import into any contact manager

The entire pipeline — parsing, detection, mapping, and vCard generation — runs in the browser using Web Workers and the SheetJS library.

## Project Structure

```
src/
  routes/                    # TanStack Router routes
  components/
    ui/                      # shadcn/ui primitives
    xls-vcard/               # App-specific components
  lib/
    xls-vcard/
      autoDetect.ts          # Column header → vCard field matching
      buildVcard.ts          # Config + row data → vCard string
      parseXlsx.ts           # SheetJS wrapper
      applyFilters.ts        # Row filter evaluation
      types.ts               # TypeScript type definitions
```
