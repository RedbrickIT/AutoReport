# AutoReport

A lightweight web app for generating monthly commuter incentive reports from door access system logs. Upload a raw `.xls` or `.xlsx` export from a Jawl/SmartLink access control system, and AutoReport will deduplicate card swipes, organize employees alphabetically, and export a clean spreadsheet ready for payroll.

## Getting Started

No install or build step required. Open `index.html` in any modern browser.

```
open index.html
```

## Features

### Generate Report

Processes a raw access log into a clean commuter incentive report.

**Steps:**

1. From the home screen, click **Generate Report**.
2. Upload your `.xls` or `.xlsx` access log file (drag & drop or click to browse).
3. **Opt-in Reference (optional):** You'll be prompted to upload an opt-in reference file. This filters the report to only include employees on the list and attaches their company names. Click **Skip** to include everyone, or upload a reference file and click **Continue with Opt-in**.
4. **Filter by Employee Group:** Toggle which card prefixes to include (e.g., `RB`, `JAN`, `JPL`). There is also a **Guest & Spare Cards** option (off by default) for generic entries.
5. **Review results** in the on-screen table, then click **Export Spreadsheet** to download.

**Exported columns:**

| Employee Name | Last, First | Company | Days in Office |
|---|---|---|---|
| John Doe | Doe, John | Redbrick | 18 |

- Each employee is counted once per day regardless of how many times they swiped.
- Employees are sorted alphabetically by last name.
- The Company column is populated from opt-in reference data. If no opt-in reference is used, the column will be empty.

### Manage Opt-in Reference

Build and maintain a list of employees and their company names. This list can be exported and later uploaded during report generation to filter results.

**Steps:**

1. From the home screen, click **Manage Opt-in Reference**.
2. **Import names** by uploading a spreadsheet. The app accepts two formats:
   - **Raw access log** (same file used for reports) — employee names are automatically extracted from the `Description #2` column, deduplicated, with generic entries like Guest/Spare excluded.
   - **Opt-in reference spreadsheet** (previously exported from this tool) — names and company assignments are imported as-is.
3. **Add employees manually** using the name and company input fields at the top of the list.
4. **Assign companies** by clicking the company field next to any employee and typing a company name.
5. **Remove employees** by clicking the **Remove** button on their row.
6. **Clear all data** by clicking the **Clear Data** button. A confirmation prompt will appear before anything is deleted.
7. Click **Export Opt-in Reference** to download the list as an `.xlsx` file.

## Access Log Format

AutoReport expects exports from a Jawl/SmartLink access control system with the following structure:

- **Rows 1-5:** Report metadata (system name, date range, operator info)
- **Row 6:** Column headers — the app looks for `Date and Time` and `Description #2`
- **Data rows:** Each row is a card swipe event

Employee names are parsed from the `Description #2` column in one of two formats:

- `PREFIX - Employee Name` (e.g., `RB - John Doe`)
- `PREFIX - Group (Employee Name)` (e.g., `JAN - GDI (John Doe)`)

Prefixes represent organizational groups and are used as filter categories during report generation.

## File Structure

```
autoreport/
  index.html    Single-page app shell
  styles.css    Dark theme, responsive layout
  app.js        All application logic
  CLAUDE.md     AI assistant context file
  README.md     This file
```

## Notes

- All processing happens client-side in the browser. No data is sent to any server.
- The app is designed with Electron compatibility in mind — all modals, inputs, and file pickers are custom DOM elements (no browser defaults like `alert()` or `<input type="file">` styling).
- To bundle for Electron, the Figtree font files and SheetJS library should be included locally rather than loaded from CDN.
