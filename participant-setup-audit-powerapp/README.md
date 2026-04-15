# Participant Setup Audit

React + TypeScript app for generating the Participant Setup Audit report entirely in the browser.

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
   - `Summary` sheet
