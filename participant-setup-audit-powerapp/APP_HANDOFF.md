# Participant Setup Audit - App Handoff

Last updated: 2026-08-26

Current Power Apps deployment: Git commit `1b92e4a`, deployed on 2026-08-26

Current pre-SharePoint deployment: seven-day People verification, SCR-based manager dashboard, PDF/interactive HTML export, Audit Subcategory propagation, inferred analyst ownership, and CSV Quota Assignment support.

App UI version: `1.2`

## 1. App purpose

This app generates a `Participant Setup Audit` report, verifies later Xactly People updates, and provides a manager dashboard entirely in the browser.

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
6. The deployed app currently has no SharePoint source-file archive connection; all source processing remains client-side.

There is no backend and no server-side data processing.

## 5. Power Apps wrapper notes

- The app is deployed as a Power Apps code app and can also run locally through Vite.
- [power.config.json](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/power.config.json) contains the configured app and environment IDs.
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
- `First Day of Leave`
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

For active current-SCR employees, LOB is derived in this order:

1. Cost Center containing GCP -> `GCP`
2. Job Family beginning with Sales Development -> `SD`
3. Advertising Sales or Advertising Operations -> `LMS`
4. LCS Sales or LCS Operations -> `LTS`
5. Sales Solutions or Sales Solutions Operations -> `LSS`
6. Global Sales Operations, except SalesQ VP -> `SD`
7. Global Sales Operations with SalesQ VP -> `Global`
8. Otherwise use People `Business_Unit`, displaying `TS` as `LTS` and `MS` as `LMS`

The same derived LOB is used for filters, Dashboard population, and Analyst inference.

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
- If People or Position setup is missing for a new hire, the New Hire row shows `missingPeopleSetup` and/or `missingPositionSetup` instead of creating a separate `Missing Xactly Setup` row.

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
- `Job Level` from SCR `CF-CB-Career Band/Level - Worker`
- `Job Grade` from SCR `CF LRV Global Job grade`
- `Supervisory Manager`
- `OTE (Base+Comm)`
- `Commission Amount`
- `Business Unit`
- `Country`
- `Currency`

Important note:

- The app still checks `OTE (Base+Comm)` changes for `changeSummary`.
- However, the `previousOteBaseComm` and `currentOteBaseComm` output columns were removed from the report.
- `previousJobLevelGrade` and `currentJobLevelGrade` show the two SCR values as `Level (Grade)`, for example `MR2 (09.1)`. New Hires have no previous value.
- When Job Title, Job Level, or Job Grade changes, `auditSubcategory` uses a career-movement classification:
  - A higher normalized grade is `Promotion` using `04 < 05 < 06 < 07 < 08.1 < 08.2 < 09.1 < 09.2 < 09.3 < 10 < 11 < 12 < A`.
  - The same nonblank grade with an `IC` to `MR` level change is also `Promotion`.
  - Same-grade IC/SP changes, lower grades, and unknown or incomplete grade movements are `Job Change`.
  - The final value is `Promotion + Variable Change`, `Promotion - No Variable Change`, `Job Change + Variable Change`, or `Job Change - No Variable Change`.
- Changes without a career movement retain the existing single-field, `Variable + Other Changes`, or `Multiple Changes - No Variable` classification.
- The exact changed fields remain in `changeSummary`.

### Deferred Change While on LOA

Rule:

- Same compared fields as `Change to Existing Participant`
- Used instead of `Change to Existing Participant` when current month SCR `On Leave = Yes`
- `changeSummary` is still prefixed with `[Currently on LOA]`

Purpose:

- Changes found while the employee is on LOA may need to be held until the employee returns from LOA.

### Missing Xactly Setup

Rule:

- Current month SCR `Active Status = Yes`
- Employee is missing from either the People file or the Position file
- Employees inside the selected processing month's New Hire window are excluded from this audit item to avoid flagging normal pending new-hire setup.

Output behavior:

- `missingPeopleSetup` shows whether the People record is missing
- `missingPositionSetup` shows whether the Position record is missing

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

### Negative balance output

Behavior:

- Material negative balances are shown in the `negativeBalance` output column for these audit items:
  - `New Hire`
  - `Transfer to Non-Sales`
  - `Transfer to Sales`
  - `Termination`
- No separate `Negative Balance Risk` audit item is generated.

Current materiality threshold:

- Absolute negative balance amount must be at least `1`

### Unmapped Data Warning

Rule:

- Employee is in the active current SCR population, active previous SCR population, or an OKR Plan End candidate
- Resolved Region, LOB, or Country is `Unmapped`, or the SCR country is missing from the country-region map

Important behavior:

- These warning rows bypass the global filters so unmapped source data does not disappear silently.

## 11. Output report design

Current visible output columns are defined in [App.tsx](/c:/Codex/PowerApps/Participant%20Setup%20Audit/participant-setup-audit-powerapp/src/App.tsx#L23).

Important output decisions:

- `auditSubcategory` is placed immediately after `auditItem` in Audit and Verification tables.
- Current month SCR `Active Status`, `On Leave`, and `First Day of Leave` are placed immediately after `Country`
- `changeSummary` is placed immediately after the current month SCR LOA context columns
- If current month SCR `On Leave = Yes`, `changeSummary` is prefixed with `[Currently on LOA]`
- `missingPeopleSetup` and `missingPositionSetup` show People/Position setup gaps for `Missing Xactly Setup`
- `peoplePlanEffectiveDate` is placed immediately after `changeSummary`
- `peopleUploadDate` is the last column
- `Note` column was removed
- Any former note text is merged into `changeSummary`

People-driven fields:

- `peoplePlanEffectiveDate` comes from People `Plan_Effective_Date`
- `peopleBusinessUnit` comes from People `Business_Unit`
- `analystName` comes from People `Analyst_Name`
- `peopleUploadDate` comes from People `Upload_Date`

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
2. `Column Guide`
3. `Summary`
4. `Verification Baseline`
5. `SCR Population`

`Summary` contains:

- Uploaded file names
- Audit counts by item
- Audit counts by nonblank subcategory
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

### Audit Subcategory

- Audit Results and Follow-up Verification tables display `Audit Subcategory` next to `Audit Item`.
- `Variable Change Only`, `Variable + Other Changes`, `Promotion + Variable Change`, and `Job Change + Variable Change` identify every audit action that requires an Annual Variable update.
- Audit Results also displays Previous and Current Job Level (Grade). Job Level and Job Grade expectations are carried into Verification as `Not Verifiable` because the People-only follow-up does not contain those SCR fields.
- `Manager Change Only` continues to use the separate `Manager Mismatch Only` verification treatment and remains outside Setup Required and Completion Rate.

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
- Audit Subcategory is carried from Audit Report to Verification Baseline and Verification Report
- Older Verification Baselines without Audit Subcategory derive it from their field keys
- The July workbook's 1,108 Change to Existing Participant rows classify as 692 Manager Change Only, 216 Variable Change Only, 173 Variable + Other Changes, 13 Multiple Changes - No Variable, and 14 Job Title Change Only

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

## 18. Follow-up verification and dashboard

### Follow-up workflow

1. Run the initial audit with the eight source files.
2. Download the initial workbook. Its `Verification Baseline` sheet stores expected People field values and a due date seven calendar days after generation.
3. Later, upload the initial workbook and the latest People file in `Follow-up Verification`.
4. Generate a workbook containing:
   - `Verification Report`
   - `Column Guide`
   - `Field Details`
   - `Summary`

Progress values:

- `Completed`
- `Partially Completed`
- `Pending`
- `Manager Mismatch Only`
- `Deferred`
- `Not Verifiable`

SLA values:

- `On Time`
- `Overdue`
- `Not Due`
- `Not Applicable`

Direct People verification mappings:

- SCR Job Title -> People `HR_Job_Title`
- SCR Supervisory Manager -> People `Level_1_Manager`
- SCR Commission Amount -> People `Annual_Variable`
- SCR OTE less Commission Amount -> People `Salary`
- SCR Country -> People `Country`
- SCR Currency -> People `Salary Currency`

SCR Business Unit, Position setup, and OKR assignment are `Not Verifiable` until an approved source mapping or additional follow-up file is available.

### Dashboard

- Commissioned employee population: distinct current-SCR employees with `Active Status = Yes`
- A blank current-SCR Active Status is a Termination; absence from the current SCR is Transfer to Non-Sales
- Region filter: All Regions, APAC, EMEA, LATAM, or NAMER when present
- Headcount breakdowns: Region, derived LOB, and assigned/inferred Analyst
- Setup Required: Completed + Partially Completed + Pending
- Completion Rate: Completed / Setup Required
- Manager Mismatch Only is shown separately; Deferred and Not Verifiable are excluded from Setup Required and Completion Rate
- KPI tiles provide hover descriptions
- PDF export uses a compact portrait print layout; HTML export is self-contained and retains an offline Region filter using aggregated data only

### Analyst ownership inference

Actual People `Analyst_Name` remains authoritative except for Transfer to Sales, which is treated like New Hire. Missing ownership is inferred from active current-SCR employees with existing People analyst mappings:

1. Unique top Analyst for normalized SCR Country + derived LOB
2. Unique top Analyst for resolved Region + derived LOB
3. `Unassigned` when no candidates exist or the highest count is tied

The baseline and follow-up reports retain `analystSource`, `analystConfidence`, and `analystSampleSize`. Inference affects dashboard ownership only and never writes back to People.

### Audit Subcategory propagation

- `auditItem` remains the stable parent action.
- `auditSubcategory` is stored in Audit Report, Verification Baseline, Verification Report, and Field Details.
- `changeSummary` remains the detailed list of changed fields.
- Verification Summary includes counts by Audit Subcategory.
- Dashboard Setup Required and Completion Rate calculations remain status-based and are not changed by the new classification.

### LOB derivation

For active current-SCR employees, apply this priority: Cost Center containing GCP -> GCP; Job Family beginning with Sales Development -> SD; Advertising Sales/Operations -> LMS; LCS Sales/Operations -> LTS; Sales Solutions/Operations -> LSS; Global Sales Operations -> SD except SalesQ VP -> Global. Otherwise use People `Business_Unit`, displaying TS as LTS and MS as LMS.

## 19. Planned SharePoint source-file archive

Not implemented in this checkpoint. The recommended future design is a solution-aware Power Automate flow triggered by the code app after a successful audit. It should create a unique processing-month/run folder and save the eight original input files without overwriting prior runs. Implementation requires the SharePoint site URL, document library, parent folder, connection reference, tenant permissions, and an upgrade of `@microsoft/power-apps` from `1.0.3` to at least `1.1.1`.
