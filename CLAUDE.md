# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

AutoReport is a vanilla HTML/CSS/JS web app that processes door access system logs into monthly commuter incentive reports. It parses .xls/.xlsx exports from a Jawl-based access control system, deduplicates card swipes to one per employee per day, and exports a clean spreadsheet with alphabetized employee names, companies, and day counts.

## Architecture

- **index.html** — Single-page app shell with six screens (home, upload, opt-in prompt, prefix filter, results, opt-in manager), plus custom modal and loading overlay
- **styles.css** — Dark theme using CSS custom properties. Figtree (Google Fonts) is the system font
- **app.js** — All logic in a single IIFE. Section-based navigation via `showSection()`. Two main flows:
  1. **Report flow**: Upload access log → (optional) upload opt-in reference → prefix filter → results/export
  2. **Opt-in manager**: Import/manually build an employee-company list, export as "Opt-in Reference" spreadsheet

## Key Design Decisions

- **No build step** — Vanilla stack, no bundler/framework. Intended to be wrapped in Electron later
- **Custom modals/inputs only** — No `alert()`, `confirm()`, `prompt()`, or browser-default file inputs. Everything is a custom DOM element for Electron compatibility
- **SheetJS (xlsx)** — Loaded via CDN (`https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js`). Used for both reading uploaded files and writing exports
- **Figtree font** — Loaded from Google Fonts. When moving to Electron, bundle the font files locally

## App Flows

### Report Generation
1. Upload .xls/.xlsx access log from Jawl/SmartLink system
2. Opt-in reference prompt: upload an opt-in list (exported from the manager) or skip
3. Prefix filter: choose which employee groups (RB, JAN, JPL, etc.) plus optional Guest & Spare Cards
4. Results: table with Employee Name, Last/First, Company, Days in Office
5. Export: downloads a .xlsx with the same four columns

If opt-in reference data is provided, only employees matching the opt-in list appear in the final report, and company names from the reference are included.

### Opt-in Reference Manager
- Upload a spreadsheet to bulk-import names (looks for "Employee Name"/"Name" and "Company" columns)
- Manually add/remove employees and assign company names via inline editable inputs
- Export as "Opt-in Reference" spreadsheet (Employee Name, Company)

## Access Log Format

The app expects exports from a Jawl/SmartLink access control system. Key structure:
- Rows 1-5: metadata (report name, date range, operator)
- Row 6: column headers — the app dynamically finds `Date and Time` (col B) and `Description #2` (col H)
- Data rows: `Description #2` contains employee identifiers in format `PREFIX - Name` or `PREFIX - Group (Name)`
- Prefixes (e.g., `RB`, `JAN`, `JPL`) represent organizational groups
- Generic entries like "Guest", "Spare", "Master" are grouped under a toggleable "Guest & Spare Cards" filter chip (off by default)

## Running

Open `index.html` directly in a browser. No server required.
