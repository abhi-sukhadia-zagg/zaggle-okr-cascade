# OKR Cascade - Interactive Presentation Tool

## What exists today

A single-file React app (`okr_cascade_react.html`) that presents organizational OKRs in two views:
- **Table view**: Three tabs (North Star Vision, Organization OKRs, Department OKRs) with a sidebar listing objectives and a content panel showing KR tables
- **Hierarchy view**: A 3-level tree (North Star -> Org OKRs -> Department OKRs grouped by department) with SVG/CSS connectors, expandable org cards, and clickable KR links that navigate to the table view

An Excel input template (`OKR_Cascade_Data_Input.xlsx`) with 5 sheets defining the data structure.

Currently, all data is hardcoded in the HTML file as JS objects. The goal is to make it dynamic.

## What needs to be built

### Architecture
- **Data source**: Excel file (.xlsx) hosted on OneDrive with a public share link
- **App**: Single React HTML file that fetches and parses the Excel on page load
- **Hosting**: GitHub Pages (static, free, public)
- **No server, no backend, no database, no API keys**

### Data flow
1. User edits the Excel file on OneDrive
2. User (or viewer) opens/refreshes the GitHub Pages URL
3. The app fetches the .xlsx from the OneDrive direct download link
4. SheetJS (xlsx library, available via CDN) parses the Excel in-browser
5. Parsed data populates the React state and renders the UI

### OneDrive public link setup
- User shares the Excel on OneDrive with "Anyone with the link can view"
- The share link (e.g. `https://1drv.ms/x/s!AbcDef...`) needs to be converted to a direct download URL
- Conversion: take the share link, base64 encode it, prepend with `https://api.onedrive.com/v1.0/shares/u!{base64}/root/content`
- Or use the embed URL pattern: replace `?e=xxxxx` with `&download=1`
- The app should have a config variable at the top where the user pastes their OneDrive share link

### Excel structure (5 sheets)

**Sheet: North Star**
| Column | Description |
|--------|-------------|
| Vision Statement | The north star vision text |
| Description | Supporting description paragraph |
| Timeline | e.g. "FY 2026-27" |

**Sheet: Org OKRs** (one row per KR, objective fields repeated on first KR row only)
| Column | Description |
|--------|-------------|
| Org OKR ID | Unique ID like ORG-1, ORG-2. Repeat for each KR row |
| Objective | Objective text (only on first row of each ID group) |
| Timeline | e.g. "FY 2026-27" |
| Sr. No. | Sequential number: 1, 2, 3 within each objective |
| EPIC-G | Single letter: E, P, I, C, or G |
| Key Result | The KR text shown in the table |
| KR Tooltip | Detailed description shown on hover over the Sr. No. cell |
| Measure | How the KR is measured |
| Target | Target value |
| KR Owner | Person or team responsible |

**Sheet: Dept OKRs** (same flat structure with extra columns)
| Column | Description |
|--------|-------------|
| Department | Department name (must match Departments sheet) |
| Dept OKR ID | Unique ID like ES-1, PRD-1 |
| Objective | Objective text (only on first row of each ID group) |
| Timeline | e.g. "FY 2026-27" |
| Linked Org OKR ID | Must match an Org OKR ID exactly (e.g. ORG-1) |
| Sr. No. | Sequential number within each objective |
| EPIC-G | Single letter |
| Key Result | KR text |
| KR Tooltip | Detailed hover description |
| Measure | Measurement method |
| Target | Target value |
| KR Owner | Responsible person/team |

**Sheet: Departments**
| Column | Description |
|--------|-------------|
| Department Name | Exact name used in Dept OKRs sheet |
| Display Order | Number controlling tab order |

### Parsing logic
When the app loads:
1. Fetch the .xlsx file as arraybuffer
2. Parse with SheetJS: `XLSX.read(data, {type: 'array'})`
3. For each sheet, convert to JSON: `XLSX.utils.sheet_to_json(wb.Sheets[sheetName])`
4. Group Org OKRs rows by `Org OKR ID` - first row of each group has the objective text, all rows have KR data
5. Group Dept OKRs rows by `Dept OKR ID` - same pattern, plus Department and Linked Org OKR ID
6. Sort departments by Display Order from the Departments sheet
7. Set React state with the parsed data

### Key parsing edge case
In the flat Excel structure, Objective/Timeline/Department fields are only filled on the FIRST row of each OKR ID group. Subsequent rows for the same ID have those cells empty. The parser must carry forward the last non-empty value for these fields when grouping.

### UI features (all already built, just need data wiring)

**Table view:**
- 3 tabs: North Star Vision, Organization OKRs, Department OKRs
- Table/Hierarchy toggle button in the tab bar
- Department sub-tab bar (horizontal, scrollable) visible only on Dept tab
- Left sidebar with goal icons listing objectives
- Right content panel with:
  - Objective heading + metadata chips
  - Linked Org OKR banner (dept view only) - clickable, navigates to that org OKR
  - KR table with columns: Sr. No., EPIC-G, Key Result, Measure, Target, Owner
  - Sr. No. cell has a tooltip showing the KR Tooltip text on hover
- Color scheme: Zaggle red (#E8281A) and white
- EPIC-G column uses amber (#FEF3D7 bg, #7C4D00 text)

**Hierarchy view:**
- Full page takeover (hides tabs, sidebar, dept bar)
- "Back to Table View" button
- 3-level tree:
  - Level 1: North Star node (dark navy card, centered)
  - Level 2: Org OKR cards (horizontal row, red border)
  - Level 3: Department group columns (dashed border, department header) with OKRs stacked vertically inside
- Click org card to expand/collapse its department children
- SVG/CSS T-bar connectors between levels:
  - Horizontal line spans center of first child to center of last child (NOT edge to edge)
  - Vertical stubs drop from horizontal line to each child card
- Every card has a "X KRs ->" link that navigates to table view for that OKR
- Dept cards are clickable, navigate to dept table view

**Auto-refresh (nice to have):**
- Optional: poll the OneDrive link every 60 seconds and re-parse if data changed
- Show a small "Last synced: X min ago" indicator

### Tech stack
- React 18 (CDN: cdnjs.cloudflare.com)
- Babel Standalone (CDN: cdnjs.cloudflare.com) for in-browser JSX compilation
- Tailwind CSS (CDN: cdn.tailwindcss.com) - only core utility classes
- SheetJS (CDN: cdnjs.cloudflare.com/ajax/libs/xlsx/) for Excel parsing
- Single HTML file, no build step

### GitHub Pages deployment
1. Create a repo (e.g. `zaggle-okr-cascade`)
2. Push the HTML file as `index.html`
3. Enable GitHub Pages in repo settings (source: main branch, root)
4. Site goes live at `username.github.io/zaggle-okr-cascade`

### Config
At the top of the script, one variable the user sets:
```js
const ONEDRIVE_SHARE_URL = "PASTE_YOUR_ONEDRIVE_SHARE_LINK_HERE";
```

The app should show a friendly error state if:
- The URL is not set (still the placeholder)
- The fetch fails (network, permissions)
- The Excel structure doesn't match expected sheets/columns

## Files in this project
- `okr_cascade_react.html` - Current working app with hardcoded data (reference implementation)
- `OKR_Cascade_Data_Input.xlsx` - Excel template the user fills in and hosts on OneDrive
- `CLAUDE_md.md` - Background context on Zaggle, the founder's lens, and KPI design philosophy
- `Org_OKRs_FY_202627.xlsx` - Original source data (reference only)

## What NOT to change
- The visual design, color scheme, and layout are finalized
- The hierarchy view connector geometry (T-bar with center-to-center horizontal lines and vertical stubs) is correct
- The table column order: Sr. No., EPIC-G, Key Result, Measure, Target, Owner
- The Sr. No. tooltip behavior for KR descriptions
