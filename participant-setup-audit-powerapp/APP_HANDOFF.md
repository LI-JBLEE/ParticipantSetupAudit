# Participant Setup Audit - App Handoff

Last updated: 2026-04-15

## 1. App purpose

This app generates a `Participant Setup Audit` report entirely in the browser.

Primary goals:

- No server-side processing
- Suitable for later wrapping/deployment through Power Apps
- English UI
- Excel-based input and Excel-based output

This app was built by following the same overall operating model as the earlier `Participant Template Generator` app, while keeping all parsing and audit logic on the client side for security reasons.

## 2. Current tech stack

- React 19
- TypeScript
- Vite
- `@microsoft/power-apps-vite`
- `xlsx`
- `xlsx-js-style`

Key files:

- [App.tsx](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/App.tsx)
- [engine.ts](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/lib/engine.ts)
- [types.ts](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/lib/types.ts)
- [index.css](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/index.css)
- [power.config.json](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/power.config.json)

## 3. Project structure and ownership

- `src/App.tsx`
  - Main UI
  - Upload workflow
  - Global filter state
  - Audit result table rendering
  - Excel download trigger
- `src/lib/engine.ts`
  - File parsing
  - Reference mapping load
  - Filter option generation
  - Audit rule logic
  - Excel workbook generation
- `src/lib/types.ts`
  - Shared types for parsed source data and output rows
- `src/index.css`
  - Layout and styling
  - Upload card styling
  - Audit result scroll frame behavior
- `public/defaults/Country Region Mapping.xlsx`
  - Bundled reference workbook used by the app at runtime

## 4. Runtime model

1. User uploads all required files in the UI.
2. Files are parsed in the browser only.
3. Parsed data is stored in React state.
4. Audit rules run in the browser.
5. Results are shown on screen and can be exported to Excel.

There is no backend and no server-side data processing.

## 5. Power Apps wrapper notes

- The app is Power Apps wrapper-ready but currently runs as a normal Vite app.
- [power.config.json](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/power.config.json) still uses a placeholder `appId`.
- `buildPath` is `./dist`.
- The current local app URL in config is `http://localhost:5173`.

## 6. Upload files and current UI order

Current upload order is defined in [engine.ts](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/lib/engine.ts#L27).

The required files are:

1. `Sales Compensation Report (Current Month)`
2. `Sales Compensation Report (Previous Month)`
3. `People`
4. `Position`
5. `Quota Assignment`
6. `Payment Balance`
7. `LOA Report`
8. `Transfer to MSFT`

Notes:

- `Worker Change Report` is intentionally excluded from this app.
- `People` and `Position` are expected to be normal Excel files. Earlier encrypted versions were replaced with re-saved standard Excel files.

## 7. Employee ID rules

Important custom rule from the business discussion:

- Valid employee IDs are numeric only.
- Numeric IDs shorter than 6 digits are valid and are left-padded to 6 digits.
- IDs containing letters are treated as placeholders or non-real employees and are ignored.

Examples:

- `81` -> `000081`
- `217929` -> `217929`
- `TBH-EM-UK-TECH-AE5` -> ignored

## 8. Source parsing rules

### People

Parsed in [engine.ts](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/lib/engine.ts#L103).

Important extracted fields:

- `Region`
- `Business_Unit`
- `Country`
- `Analyst_Name`
- `Plan_Effective_Date`
- `Plan_Type`
- `Effective Start Date`
- `Upload_Date`

Important fix already applied:

- Header normalization removes `_`, so columns such as `Business_Unit`, `Analyst_Name`, `Plan_Effective_Date`, and `Upload_Date` are parsed correctly.

### Sales Compensation Report

Parsed in [engine.ts](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/lib/engine.ts#L319).

Important extracted fields:

- `Original Hire Date`
- `Hire Date`
- `Is Rehire`
- `Active Status`
- `On Leave`
- `Termination Date`
- `Job Title`
- `Supervisory Manager`
- `OTE (Base+Comm)`
- `Commission Amount`
- `Business Unit`
- `Country`
- `Currency`

Important fixes already applied:

- `Business Unit` is read from the actual `Business Unit` column, not `Business Unit Organization`.
- `Hire Date` is now read as the actual `Hire Date` column, not accidentally from `Original Hire Date`.

### Quota Assignment

Parsed in [engine.ts](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/lib/engine.ts#L239).

Important behavior:

- Only OKR quota rows are kept.
- Month headers are restricted to real `MMM-YYYY` month columns.
- `YEAR-2026`, `H1-2026`, `H2-2026`, `QTR-*` headers are ignored.

This fixed the earlier bug where `YEAR-2026` could be misread as `EAR-2026`.

## 9. Global filter logic

### Region

Region options are fixed to:

- `APAC`
- `EMEA`
- `LATAM`
- `NAMER`

Resolution rules:

1. If SCR country exists, map `SCR.Country` through the reference workbook:
   - [Country Region Mapping.xlsx](/c:/Codex/PowerApps/Participant%20Setup%20Audit/Reference/Country%20Region%20Mapping.xlsx)
2. If the employee is not available in SCR and only exists in People, use `People.Region`
3. Special case: `People.Region = CHINA` is treated as `APAC`

Relevant logic:

- [engine.ts](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/lib/engine.ts#L67)
- [engine.ts](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/lib/engine.ts#L729)

### LOB

LOB resolution order:

1. Current month SCR `Business Unit`
2. Previous month SCR `Business Unit`
3. People `Business_Unit`

Filter options are built from resolved employee context.

### Country

Country filter options come only from the current month SCR `Country` column.

Dynamic behavior:

- Country options update based on selected Region values.
- If Region selection changes, Country selection is automatically constrained to visible countries.

Relevant UI logic:

- [App.tsx](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/App.tsx#L79)

## 10. Audit items currently implemented

Implemented in [engine.ts](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/lib/engine.ts#L414).

### New Hire

Rule:

- `Active Status = Yes`
- `Hire Date` between previous month 16th and selected month 15th

Rehire rule:

- If `Termination Date` exists, the row is still treated as New Hire when:
  - `Is Rehire = Yes`
  - `Original Hire Date < Hire Date`

This logic exists in:

- [engine.ts](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/lib/engine.ts#L431)
- [engine.ts](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/lib/engine.ts#L792)

### Change to Existing Participant

Compared fields:

- `Job Title`
- `Supervisory Manager`
- `OTE (Base+Comm)`
- `Commission Amount`
- `Business Unit`
- `Country`
- `Currency`

Important note:

- The app still checks `OTE (Base+Comm)` changes for `changeSummary`.
- However, the `previousOteBaseComm` and `currentOteBaseComm` output columns were removed from the report.

### LOA Start / LOA Return

Rule:

- Triggered when `On Leave` changes between previous and current SCR.

### OKR Plan End

Rule:

- Latest non-zero OKR quota month ends in the month before the selected processing month.

Current `changeSummary` text:

- `Non-zero OKR quota month ended in the month before the selected processing month.`

### Transfer to Non-Sales

Rule:

- Active in previous month SCR
- Missing from current month SCR

### Transfer to Sales

Rule:

- Active in current month SCR
- Missing from previous month SCR
- `Hire Date` earlier than previous month 15th

### Termination

Rule:

- Active in previous month SCR
- Current SCR exists but no longer active
- Also flags whether employee exists in Transfer to MSFT file

## 11. Output report design

Current visible output columns are defined in [App.tsx](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/App.tsx#L23).

Important output decisions:

- `changeSummary` is placed immediately after `Country`
- `peoplePlanEffectiveDate` is placed immediately after `changeSummary`
- `Upload_Date` is the last column
- `Note` column was removed
- Any former note text is merged into `changeSummary`

People-driven fields:

- `peoplePlanEffectiveDate` comes from People `Plan_Effective_Date`
- `peopleBusinessUnit` comes from People `Business_Unit`
- `analystName` comes from People `Analyst_Name`
- `uploadDate` comes from People `Upload_Date`

If the employee does not exist in People:

- Those People-based fields remain blank

### Commission Amount cell type

`previousCommissionAmount` and `currentCommissionAmount` are intentionally numeric in the output model:

- [types.ts](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/lib/types.ts#L145)

Excel output also writes them as numeric cells, not text.

## 12. Excel export

Workbook generation is handled in [engine.ts](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/lib/engine.ts#L600).

Sheets:

1. `Audit Report`
2. `Summary`

`Summary` contains:

- Uploaded file names
- Audit counts by item
- Total row count

## 13. Current UI behaviors

### Header

- The old `POWER APPS WRAPPER READY` badge has been removed.
- Processing month selector remains in the top-right hero area.

### Upload cards

- Uploaded cards change background color to make completion visually obvious.
- Current styling is in [index.css](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/index.css)

### Audit result table scroll

- Horizontal and vertical scrolling are intentionally confined to the Audit Results frame.
- The entire page should no longer grow a global results scrollbar when the table becomes wide or tall.

## 14. Known sample-data observations

The current sample data includes at least two `Change to Existing Participant` rows without People records:

- `244406`
- `244407`

This appears to be a source-data condition, not an app parsing issue.

## 15. Verification performed during development

Common verification methods used so far:

- `npm run build`
- `npx tsx -` smoke tests with sample data
- workbook inspection through `xlsx`
- headless Edge screenshot checks

Examples already verified:

- App builds successfully
- Region/Country/LOB filter behavior works with sample data
- `EAR-2026` month parsing bug is fixed
- Rehire New Hire case `151518` now appears correctly
- People metadata fields populate the report when the People record exists
- `Commission Amount` exports as numeric
- `OTE (Base+Comm)` columns are removed from the report

## 16. Common change locations

When future changes are requested, use this guide:

- UI layout or text:
  - [App.tsx](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/App.tsx)
  - [index.css](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/index.css)
- Upload order:
  - [engine.ts](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/lib/engine.ts#L27)
- Parsing bugs:
  - [engine.ts](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/lib/engine.ts)
- Output schema:
  - [types.ts](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/lib/types.ts)
  - [App.tsx](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/App.tsx#L23)
- Audit logic:
  - [engine.ts](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/lib/engine.ts#L414)
- Excel formatting/export:
  - [engine.ts](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/lib/engine.ts#L600)

## 17. Recommended next-step maintenance practice

When making future changes:

1. Update code
2. Run `npm run build`
3. If logic changed, run a sample-data smoke test with `npx tsx -`
4. If UI changed, capture a fresh browser screenshot
5. Update this document if business rules or output columns changed

