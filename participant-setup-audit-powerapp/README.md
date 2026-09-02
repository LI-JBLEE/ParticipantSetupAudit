# Participant Setup Audit

React + TypeScript Power Apps code app for generating the Participant Setup Audit report, verifying People updates after seven days, and monitoring setup execution entirely in the browser.

## Run

```bash
npm run dev
```

## Build

```bash
npm run build
```

## Runtime model

1. All uploaded files are parsed in the front end.
2. Audit logic runs in the browser without server-side processing.
3. The built app is deployed as the `Participant Setup Audit` Power Apps code app.

## Required uploads

1. People
2. Position
3. Payment Balance
4. Quota Assignment
5. LOA Report
6. Sales Compensation Report (Current Month)
7. Sales Compensation Report (Previous Month)
8. Transfer to MSFT

## Output

1. On-screen audit table
2. Downloadable Excel workbook with:
   - `Audit Report` sheet
   - `Column Guide` sheet
   - `Summary` sheet
   - `Verification Baseline` sheet
   - `SCR Population` sheet

`Audit Report` includes `auditSubcategory` immediately after `auditItem`, plus `previousJobLevelGrade` and `currentJobLevelGrade`. The two SCR values are displayed as `Level (Grade)`, for example `MR2 (09.1)`.

Participant changes use controlled operational subcategories:

- `Manager Change Only`
- `Variable Change Only`
- `Variable + Other Changes`
- `Promotion + Variable Change`
- `Promotion - No Variable Change`
- `Job Change + Variable Change`
- `Job Change - No Variable Change`
- `Job Title Change Only`
- `OTE Change Only`
- `Business Unit Change Only`
- `Country Change Only`
- `Currency Change Only`
- `Multiple Changes - No Variable`

The exact changed fields remain in `changeSummary`.

For an existing participant, a Job Title, Job Level, or Job Grade change is classified as a career movement. It is a `Promotion` when the normalized Job Grade rises in this order: `04 < 05 < 06 < 07 < 08.1 < 08.2 < 09.1 < 09.2 < 09.3 < 10 < 11 < 12 < A`. A same-grade change from an `IC` level to an `MR` level is also a Promotion. Other career movements, including same-grade IC/SP changes, lower grades, and unknown grades, are classified as `Job Change`. The suffix records whether Commission Amount also changed.

## Follow-up verification

1. Generate and download the initial audit workbook.
2. About seven days later, open `Follow-up Verification`.
3. Upload the initial audit workbook and the latest People file.
4. Generate the verification report and review `Completed`, `Partially Completed`, `Pending`, `Manager Mismatch Only`, and SLA status.

The verification workbook includes `Verification Report`, `Column Guide`, `Field Details`, and `Summary`. `auditSubcategory` is carried from the initial audit through the Verification Baseline and into the Verification Report. Older audit workbooks without this column derive it from their verification field set.

Job Level, Job Grade, Business Unit, Position, and OKR changes remain `Not Verifiable` when the People-only follow-up does not contain an approved direct mapping.

## Dashboard

The Dashboard uses distinct current-SCR employees with `Active Status = Yes` for commissioned employee counts. Region filtering updates the KPI, derived LOB, Analyst, and execution views. Setup Required includes Completed, Partially Completed, and Pending; Completion Rate is Completed divided by Setup Required. Manager Mismatch Only is separate, while Deferred and Not Verifiable are excluded.

Dashboard KPI tiles include hover descriptions. `PDF` opens the compact portrait print layout. `HTML` downloads a self-contained interactive dashboard whose Region filter works offline using aggregated data only.

When a setup target has no People `Analyst_Name`, the dashboard infers ownership from existing active employee mappings:

1. Unique top Analyst for SCR Country + derived LOB
2. Otherwise, unique top Analyst for Region + derived LOB
3. Ties remain `Unassigned`

New Hire and Transfer to Sales use inferred ownership instead of the employee's historical People Analyst. Inferred assignments are labeled with source, confidence, and sample size; they do not overwrite People data.

## Deployment checkpoint

- Power Apps app ID: `c57dffff-3ceb-4d92-893b-d91f2b67de6c`
- Current Power Apps deployment: Git commit `1b92e4a`, deployed on 2026-08-26
- The deployment includes Audit Subcategory propagation, inferred analyst ownership, and CSV Quota Assignment support, and intentionally excludes the proposed SharePoint source-file archive feature.
