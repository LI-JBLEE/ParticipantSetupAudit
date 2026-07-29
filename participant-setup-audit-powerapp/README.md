# Participant Setup Audit

React + TypeScript app for generating the Participant Setup Audit report, verifying People updates after seven days, and monitoring setup execution entirely in the browser.

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
3. Power Apps can wrap the built app later through the Code App flow.

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

## Follow-up verification

1. Generate and download the initial audit workbook.
2. About seven days later, open `Follow-up Verification`.
3. Upload the initial audit workbook and the latest People file.
4. Generate the verification report and review `Completed`, `Partially Completed`, `Pending`, `Manager Mismatch Only`, and SLA status.

The verification workbook includes `Verification Report`, `Column Guide`, `Field Details`, and `Summary`.

Business Unit, Position, and OKR changes remain `Not Verifiable` when the People-only follow-up does not contain an approved direct mapping.

## Dashboard

The Dashboard uses distinct current-SCR employees with `Active Status = Yes` for commissioned employee counts. Region filtering updates the KPI, derived LOB, Analyst, and execution views. Setup Required includes Completed, Partially Completed, and Pending; Completion Rate is Completed divided by Setup Required. Manager Mismatch Only is separate, while Deferred and Not Verifiable are excluded.

Dashboard KPI tiles include hover descriptions. `PDF` opens the compact portrait print layout. `HTML` downloads a self-contained interactive dashboard whose Region filter works offline using aggregated data only.

When a setup target has no People `Analyst_Name`, the dashboard infers ownership from existing active employee mappings:

1. Unique top Analyst for SCR Country + derived LOB
2. Otherwise, unique top Analyst for Region + derived LOB
3. Ties remain `Unassigned`

New Hire and Transfer to Sales use inferred ownership instead of the employee's historical People Analyst. Inferred assignments are labeled with source, confidence, and sample size; they do not overwrite People data.
