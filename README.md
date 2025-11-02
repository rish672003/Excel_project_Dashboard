# Excel Project Dashboard

A user-friendly Excel dashboard for exploring, cleaning, visualizing, and exporting spreadsheet data. This repository contains the workbook(s), sample data, documentation, and optional helper macros to turn Excel workbooks into interactive charts, KPI summaries, and exportable reports.

Note: This README is written for an Excel-based project (workbook + VBA/macros). Replace any TODO entries with concrete filenames or details present in the repository.

## Table of contents
- [Project overview](#project-overview)
- [Features](#features)
- [Requirements](#requirements)
- [Getting started](#getting-started)
- [How to use the dashboard](#how-to-use-the-dashboard)
- [Data format and tips](#data-format-and-tips)
- [Customizing and maintenance](#customizing-and-maintenance)
- [Security and macros](#security-and-macros)
- [Troubleshooting](#troubleshooting)
- [Project structure](#project-structure)
- [Contributing](#contributing)
- [License](#license)
- [Contact](#contact)

## Project overview
This project provides an Excel-based dashboard that lets non-technical users:
- Import or paste spreadsheet data
- Clean and transform data using built-in macros or Excel formulas
- Interactively filter and slice data (Slicers, tables, or form controls)
- Visualize trends and KPIs with Excel charts and conditional formatting
- Export reports (PDF, Excel snapshot, or images)

The dashboard is designed to be used directly inside Excel (recommended: .xlsm workbook if macros are used).

## Features
- Single-click refresh of pivot tables/charts
- Pre-built KPI tiles and charts for common metrics
- Macro to import and normalize CSV/Excel files (if included)
- Buttons to export the current dashboard view to PDF or a new workbook
- Optional scheduled export via Windows Task Scheduler + script (advanced)

## Requirements
- Microsoft Excel (Windows preferred for full macro support)
  - Excel 2016, 2019, 2021, or Microsoft 365 recommended
  - Mac Excel supports many features but VBA compatibility and certain ActiveX controls may be limited
- If the workbook uses external data connections: access to the data files or a configured ODBC/ODBCDSN
- No Python/Node.js runtime required (unless the repo includes companion scripts)

## Getting started
1. Clone or download the repository:
   - Download ZIP from GitHub or git clone https://github.com/rish672003/Excel_project_Dashboard.git

2. Open the workbook:
   - Open the dashboard workbook (likely a .xlsm file) in Excel.

3. Enable macros (see Security and macros).

4. If the workbook has a README or notes sheet inside, read it for workbook-specific instructions (data layout, named ranges, buttons).

5. Place your data files in the repository data/ folder (if the project expects a data/ location) or use the built-in Import button to load files.

## How to use the dashboard
Typical workflow (may vary by workbook):
1. Open the workbook and enable macros.
2. Go to the "Data" or "Import" sheet and click Import / Load to bring in your Excel/CSV data.
3. Use the "Refresh" button to update pivot tables and charts after data changes.
4. Use slicers or drop-downs on the Dashboard sheet to filter date ranges, categories, or regions.
5. Click "Export PDF" or "Save Snapshot" to export the current dashboard view.

Look for these UI elements in the workbook:
- Ribbon buttons or ActiveX/Form Controls (Import, Refresh, Export)
- A Dashboard sheet with KPI tiles, charts, and slicers
- An "Admin" or "Config" sheet for named ranges and settings

## Data format and tips
- Keep column headers consistent across files (e.g., Date, Category, Amount).
- Dates should be Excel-recognized dates (not text) for time-series charts and filters.
- Remove leading/trailing spaces from text columns to avoid mismatched categories.
- If your dataset is large (>100k rows), consider using Power Query or reducing to a summarized dataset before charting.

If the workbook uses a specific schema, update the following example to match your files:
- Required columns: Date, Item, Category, Region, Value

## Customizing and maintenance
- To edit or extend VBA macros: open the VBA editor (Alt+F11) and locate modules under ThisWorkbook / Modules.
- To add chart variations: duplicate an existing sheet or chart and update the source ranges.
- To change data ranges used by pivot tables: right-click the pivot table > Change Data Source.
- If you prefer no macros: you can convert macro-driven steps into Power Query steps (recommended for reproducibility).

Document any changes you make in the workbook’s Notes/Admin sheet so other users can follow them.

## Security and macros
- The dashboard may include VBA macros; Excel will show a security warning on open.
- Enable macros only if you trust the file source.
- To allow macros:
  - In Excel, go to File > Options > Trust Center > Trust Center Settings > Macro Settings.
  - Choose "Disable all macros with notification" (recommended) and enable macros for this file when prompted.
  - Optionally add the repo folder to Trusted Locations to avoid repeated prompts.
- If the workbook requests "Trust access to the VBA project object model", this is needed only for macros that manipulate VBA; enable it if you understand and trust the code.

## Troubleshooting
- Buttons not working after download:
  - Make sure macros are enabled and ActiveX controls are allowed (Windows).
  - If ActiveX controls are blocked, try replacing buttons with Form Controls or shapes assigned to macros.

- Data not updating:
  - Confirm data is loaded into the expected named range/table.
  - Refresh pivot tables (right-click > Refresh All) or use provided Refresh button.

- Charts show #REF or broken:
  - Check that named ranges or table references still exist after edits.
  - Open the Name Manager (Formulas > Name Manager) to inspect named ranges.

- Large workbook performance:
  - Turn off automatic calculation during heavy updates (Formulas > Calculation Options > Manual) and set back to Automatic after finishing.
  - Use Excel Tables and PivotTables instead of many volatile formulas.

## Project structure
Update this list with the actual files in the repository.

- Dashboard.xlsm or Dashboard.xlsx        # Main dashboard workbook (macros-enabled = .xlsm)
- data/                                  # Sample data files (do not commit sensitive data)
- docs/                                  # Optional documentation, screenshots, and notes
- assets/                                # Images or exported reports
- README.md                               # This file

If you'd like, I can inspect the repository and replace the above with exact filenames and sheet descriptions.

## Contributing
- If you add macros, include comments and a short description at the top of each module.
- Test changes on a copy of the workbook before committing.
- Keep sample data anonymized — do not commit PII.
- For larger changes, open an issue describing the change and create a branch with your updates.

## License
TODO: Add a license file (for example, MIT). If you want, I can add a LICENSE file to the repository.

## Contact
Maintainer / Owner: rish672003

For issues, questions, or feature requests, open an issue in this repository.
