import XLSXImport from "xlsx-js-style";
import { buildDashboardHtml, buildDashboardHtmlFileName } from "./dashboardHtml";
import {
  buildAuditWorkbook,
  buildAuditReport,
  buildDashboardModel,
  buildFollowUpVerification,
  buildFollowUpWorkbook,
  createEmptyAppData,
  deriveAuditSubcategory,
  parseQuotaAssignmentFile,
  parseScrFile,
  parseVerificationBaselineFile,
  summarizeSetupExecution,
} from "./engine";
import type { PeopleRecord, ScrRecord, VerificationExpectation } from "./types";

const XLSX = ((XLSXImport as unknown as { default?: typeof XLSXImport }).default ?? XLSXImport) as typeof XLSXImport;

function assertColumnGuide(workbookBuffer: ArrayBuffer, reportName: string): void {
  const workbook = XLSX.read(workbookBuffer, { type: "array" });
  const reportSheet = workbook.Sheets[reportName];
  const guideSheet = workbook.Sheets["Column Guide"];
  if (!reportSheet || !guideSheet) throw new Error(`${reportName} workbook is missing the Column Guide sheet.`);
  const reportHeaders = (XLSX.utils.sheet_to_json<string[]>(reportSheet, { header: 1 })[0] ?? []).map(String);
  const guideRows = XLSX.utils.sheet_to_json<{ Column: string }>(guideSheet, { defval: "" });
  const guideColumns = new Set(guideRows.map((row) => row.Column));
  if (reportHeaders.length !== guideColumns.size || reportHeaders.some((header) => !guideColumns.has(header))) {
    throw new Error(`${reportName} Column Guide does not cover every report column.`);
  }
}

const people: PeopleRecord = {
  employeeId: "000081",
  fullName: "Test Employee",
  firstName: "Test",
  lastName: "Employee",
  region: "APAC",
  businessUnit: "TS",
  country: "Singapore",
  analystName: "Test Analyst",
  planEffectiveDate: new Date(2026, 6, 1),
  planType: "Sales",
  effectiveStartDate: new Date(2026, 6, 1),
  uploadDate: new Date(2026, 6, 22),
  employeeStatus: "Active",
  terminationDate: null,
  salary: 100,
  salaryCurrency: "SGD",
  annualVariable: 50,
  hrJobTitle: "Account Executive",
  level1Manager: "Current Manager (123456)",
};

const base = {
  processingMonth: "JUL-2026",
  generatedAt: "2026-07-15T00:00:00.000Z",
  dueDate: "2026-07-21",
  employeeId: people.employeeId,
  employeeName: people.fullName,
  region: people.region,
  lob: people.businessUnit,
  country: people.country,
  analystName: people.analystName,
  analystSource: "People",
  analystConfidence: "Confirmed",
  analystSampleSize: 0,
  auditItem: "Change to Existing Participant",
  auditSubcategory: "",
  baselineValue: "",
  deferred: "No",
  note: "",
} satisfies Omit<VerificationExpectation, "verificationId" | "fieldKey" | "fieldLabel" | "expectedValue" | "rule">;

const expectations: VerificationExpectation[] = [
  {
    ...base,
    verificationId: "completed",
    fieldKey: "hrJobTitle",
    fieldLabel: "HR Job Title",
    expectedValue: people.hrJobTitle,
    rule: "text",
  },
  {
    ...base,
    verificationId: "partial",
    fieldKey: "record",
    fieldLabel: "People Record",
    expectedValue: "Present",
    rule: "exists",
  },
  {
    ...base,
    verificationId: "partial",
    fieldKey: "level1Manager",
    fieldLabel: "Level 1 Manager",
    expectedValue: "Expected Manager (654321)",
    rule: "text",
  },
  {
    ...base,
    verificationId: "manager-only",
    fieldKey: "level1Manager",
    fieldLabel: "Level 1 Manager",
    expectedValue: "Expected Manager (654321)",
    rule: "text",
  },
];

const result = buildFollowUpVerification(expectations, { [people.employeeId]: people }, people.uploadDate ?? new Date());
const completed = result.rows.find((row) => row.verificationId === "completed");
const partial = result.rows.find((row) => row.verificationId === "partial");
const managerOnly = result.rows.find((row) => row.verificationId === "manager-only");
if (completed?.progressStatus !== "Completed") throw new Error("Expected a completed verification row.");
if (partial?.progressStatus !== "Partially Completed" || partial.slaStatus !== "Overdue") {
  throw new Error("Expected a partially completed overdue verification row.");
}
if (managerOnly?.progressStatus !== "Manager Mismatch Only" || managerOnly.slaStatus !== "Not Applicable") {
  throw new Error("Expected a manager-only mismatch outside setup execution and SLA counts.");
}
if (managerOnly.auditSubcategory !== "Manager Change Only") {
  throw new Error("Expected an old baseline without a subcategory to derive Manager Change Only.");
}
if (
  deriveAuditSubcategory(["Commission Amount"]) !== "Variable Change Only" ||
  deriveAuditSubcategory(["Job Title", "Commission Amount"]) !== "Variable + Other Changes" ||
  deriveAuditSubcategory(["Job Title", "Supervisory Manager"]) !== "Multiple Changes - No Variable"
) {
  throw new Error("Audit subcategory classification did not preserve Variable change visibility.");
}
const executionSummary = summarizeSetupExecution(result.rows);
if (
  executionSummary.setupRequired !== 2 ||
  executionSummary.completed !== 1 ||
  executionSummary.partiallyCompleted !== 1 ||
  executionSummary.pending !== 0 ||
  executionSummary.managerMismatchOnly !== 1 ||
  executionSummary.completionRate !== 0.5
) {
  throw new Error("Setup execution summary did not match the Dashboard status rules.");
}

const followUpWorkbook = buildFollowUpWorkbook(result, {});
if (followUpWorkbook.byteLength === 0) {
  throw new Error("Follow-up verification workbook was empty.");
}
assertColumnGuide(followUpWorkbook, "Verification Report");

const scr = (employeeId: string, hireDate: Date, fullName: string, country = "Singapore"): ScrRecord => ({
  employeeId,
  firstName: fullName.split(" ")[0] ?? fullName,
  lastName: fullName.split(" ")[1] ?? "",
  fullName,
  originalHireDate: hireDate,
  activeStatus: "Yes",
  onLeave: "",
  firstDayOfLeave: null,
  hireDate,
  isRehire: "No",
  terminationDate: null,
  jobTitle: "Account Executive",
  jobLevel: "IC3",
  jobGrade: "08.2",
  supervisoryManager: "Manager (123456)",
  oteBaseComm: 150,
  commissionAmount: 50,
  costCenter: "",
  jobFamily: "",
  businessUnit: "Sales Solutions",
  country,
  currency: "SGD",
});

const parsedScr = await parseScrFile(
  new File(
    [
      [
        "Employee ID,Active Status,Hire Date,Business Unit,Country,Job Family Group,Job Family,Cost Center - ID,Cost Center,CF-CB-Career Band/Level - Worker,CF LRV Global Job grade",
        "000100,Yes,2020-01-01,Other,Singapore,Wrong Group,Sales Development Representative,12345,NAMER GCP Enterprise,MR2,09.1",
      ].join("\n"),
    ],
    "scr.csv",
    { type: "text/csv" },
  ),
);
if (
  parsedScr.data["000100"]?.jobFamily !== "Sales Development Representative" ||
  parsedScr.data["000100"]?.costCenter !== "NAMER GCP Enterprise" ||
  parsedScr.data["000100"]?.jobLevel !== "MR2" ||
  parsedScr.data["000100"]?.jobGrade !== "09.1"
) {
  throw new Error("SCR parser did not select the exact Job Family, Cost Center, Job Level, and Job Grade columns.");
}

const quotaMatrix = [
  ["# Quota Name", "Type", "Name", "Person Name (Employee ID)", "Effective Start Date", "JUL-2026", "AUG-2026"],
  ["OKR Quota", "Position", "OKR Quota", "Nitya Rao (245203)", "2026-07-01", 7160.3775, 0],
];
const quotaCsv = quotaMatrix.map((row) => row.join(",")).join("\n");
const quotaWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(quotaWorkbook, XLSX.utils.aoa_to_sheet(quotaMatrix), "Quota Assignment");
const [parsedQuotaCsv, parsedQuotaXlsx] = await Promise.all([
  parseQuotaAssignmentFile(new File([quotaCsv], "quota.csv", { type: "text/csv" })),
  parseQuotaAssignmentFile(new File([XLSX.write(quotaWorkbook, { type: "array", bookType: "xlsx" })], "quota.xlsx")),
]);
const csvQuotaRow = parsedQuotaCsv.data[0];
const xlsxQuotaRow = parsedQuotaXlsx.data[0];
if (
  csvQuotaRow?.monthValues["JUL-2026"] !== 7160.3775 ||
  csvQuotaRow.monthValues["AUG-2026"] !== 0 ||
  JSON.stringify(csvQuotaRow.monthValues) !== JSON.stringify(xlsxQuotaRow?.monthValues)
) {
  throw new Error("Quota CSV month columns did not parse like XLSX month columns.");
}
const okrData = createEmptyAppData();
okrData.quotaRows = parsedQuotaCsv.data;
okrData.peopleById["245203"] = { ...people, employeeId: "245203", fullName: "Nitya Rao" };
const okrAudit = buildAuditReport(
  "AUG-2026",
  { regions: ["APAC"], lobs: ["LTS"], countries: ["Singapore"] },
  okrData,
  { singapore: "APAC" },
  new Date(2026, 7, 20),
);
if (!okrAudit.rows.some((row) => row.employeeId === "245203" && row.auditItem === "OKR Plan End")) {
  throw new Error("Quota CSV did not produce the expected August OKR Plan End audit row.");
}

const dashboardActiveScr = {
  ...scr(people.employeeId, new Date(2020, 0, 1), people.fullName),
  costCenter: "Standard Cost Center",
  jobFamily: "SalesQ IC",
};
const dashboardTerminatedScr = {
  ...scr("000089", new Date(2020, 0, 1), "Terminated Employee"),
  activeStatus: "",
};
const baselineWorkbook = buildAuditWorkbook(
  [],
  {},
  expectations,
  { [people.employeeId]: dashboardActiveScr, [dashboardTerminatedScr.employeeId]: dashboardTerminatedScr },
);
const parsedBaseline = await parseVerificationBaselineFile(new File([baselineWorkbook], "initial-output.xlsx"));
if (
  parsedBaseline.data.expectations.length !== expectations.length ||
  parsedBaseline.data.expectations[0]?.employeeId !== people.employeeId
) {
  throw new Error("Verification Baseline workbook round trip failed.");
}
if (
  Object.keys(parsedBaseline.data.currentScrById).length !== 1 ||
  !parsedBaseline.data.currentScrById[people.employeeId] ||
  parsedBaseline.data.currentScrById[dashboardTerminatedScr.employeeId] ||
  parsedBaseline.data.currentScrById[people.employeeId]?.costCenter !== dashboardActiveScr.costCenter ||
  parsedBaseline.data.currentScrById[people.employeeId]?.jobFamily !== dashboardActiveScr.jobFamily
) {
  throw new Error("Active SCR Population workbook round trip failed.");
}
const dashboardPeopleOnly = { ...people, employeeId: "000090", fullName: "People Only Employee" };
const dashboard = buildDashboardModel(
  { [people.employeeId]: dashboardActiveScr, [dashboardTerminatedScr.employeeId]: dashboardTerminatedScr },
  {
    [people.employeeId]: { ...people, region: "EMEA", businessUnit: "Incorrect People LOB" },
    [dashboardTerminatedScr.employeeId]: { ...people, employeeId: dashboardTerminatedScr.employeeId },
    [dashboardPeopleOnly.employeeId]: dashboardPeopleOnly,
  },
  result.rows,
  { singapore: "APAC" },
);
const dashboardHtml = buildDashboardHtml(
  { "All Regions": dashboard, APAC: { ...dashboard, commissionedEmployees: 7 } },
  "APAC",
  "JUL-2026",
);
const dashboardScript = dashboardHtml.slice(
  dashboardHtml.indexOf("<script>") + "<script>".length,
  dashboardHtml.lastIndexOf("</script>"),
);
const dashboardStaticHtml = dashboardHtml.slice(0, dashboardHtml.indexOf("<script>"));
if (
  !dashboardStaticHtml.includes('<option value="APAC" selected>APAC</option>') ||
  !dashboardStaticHtml.includes(">7</strong>") ||
  !dashboardStaticHtml.includes("Analyst setup ownership") ||
  !dashboardStaticHtml.includes('id="dashboard-region-0"') ||
  !dashboardStaticHtml.includes('for="dashboard-region-1"') ||
  !dashboardStaticHtml.includes("#dashboard-region-1:checked ~ .shell .dashboard-fallback-view-1") ||
  !dashboardStaticHtml.includes(">1</strong>")
) {
  throw new Error("Dashboard HTML did not include usable CSS-only Region fallback views.");
}
if (/(?:=>|\?\?|\bconst\b|\blet\b|\.\.\.)/.test(dashboardScript)) {
  throw new Error("Dashboard HTML runtime includes JavaScript syntax that can fail in older mobile WebViews.");
}
const fakeElements = new Map(
  ["region-filter", "kpis", "snapshot", "print-region", "analyst-table", "region-bars", "lob-bars", "lob-table"].map(
    (id) => [
      id,
      {
        value: "",
        innerHTML: "",
        textContent: "",
        className: "",
        options: [] as Array<{ text: string; value: string }>,
        listeners: {} as Record<string, () => void>,
        add(option: { text: string; value: string }) {
          this.options.push(option);
        },
        addEventListener(event: string, listener: () => void) {
          this.listeners[event] = listener;
        },
      },
    ],
  ),
);
fakeElements.get("region-filter")?.options.push(
  { text: "All Regions", value: "All Regions" },
  { text: "APAC", value: "APAC" },
);
const fakeDocument = {
  title: "",
  documentElement: { className: "" },
  getElementById: (id: string) => fakeElements.get(id),
};
Object.defineProperty(globalThis, "document", { configurable: true, value: fakeDocument });
try {
  new Function(dashboardScript)();
  const regionFilter = fakeElements.get("region-filter");
  const kpis = fakeElements.get("kpis");
  if (regionFilter?.options.length !== 2 || regionFilter.value !== "APAC" || !kpis?.innerHTML.includes(">7</strong>")) {
    throw new Error("Dashboard HTML did not render the default Region model.");
  }
  if (!fakeDocument.documentElement.className.includes("dashboard-js")) {
    throw new Error("Dashboard HTML did not switch from the CSS fallback to the scripted Region filter.");
  }
  regionFilter.value = "All Regions";
  regionFilter.listeners.change?.();
  if (!kpis.innerHTML.includes(">1</strong>")) {
    throw new Error("Dashboard HTML Region filter did not update the KPI values.");
  }
} finally {
  Reflect.deleteProperty(globalThis, "document");
}
if (
  !dashboardHtml.includes('id="region-filter"') ||
  !dashboardHtml.includes('regionFilter.addEventListener("change", renderDashboard)') ||
  !dashboardHtml.includes('"All Regions"') ||
  !dashboardHtml.includes('"APAC"')
) {
  throw new Error("Interactive Dashboard HTML export was missing its Region models or filter behavior.");
}
if (
  buildDashboardHtmlFileName("JUL-2026", new Date(2026, 6, 29, 12, 18, 15)) !==
  "Participant_Setup_Dashboard_JUL-2026_20260729_121815.html"
) {
  throw new Error("Dashboard HTML download filename was incorrect.");
}
if (dashboard.commissionedEmployees !== 1 || dashboard.setupRequired !== 2 || dashboard.managerMismatchOnly !== 1) {
  throw new Error("Dashboard did not use current SCR Active Status for the commissioned population.");
}
if (dashboard.byRegion.find((row) => row.label === "APAC")?.employees !== 1) {
  throw new Error("Dashboard Region breakdown did not use the current SCR country mapping.");
}
if (dashboard.byLob.find((row) => row.label === "LSS")?.employees !== 1) {
  throw new Error("Dashboard LOB breakdown did not apply the SCR LOB mapping.");
}
if (dashboard.completed !== 1 || dashboard.partiallyCompleted !== 1) {
  throw new Error("Dashboard verification metrics did not reconcile.");
}
if (dashboard.setupRequiredRate !== 2) {
  throw new Error("Dashboard Setup Required rate did not use Setup Required divided by Commissioned Employees.");
}
if (dashboard.completionRate !== 0.5) {
  throw new Error("Dashboard Completion Rate included Manager Mismatch Only in its denominator.");
}
if (buildDashboardModel({}, {}, [], {}).setupRequiredRate !== 0) {
  throw new Error("Dashboard Setup Required rate did not handle an empty commissioned population.");
}

const lobMappingCases = [
  {
    employeeId: "000091",
    costCenter: "NAMER GCP Enterprise",
    jobFamily: "Sales Development Representative",
    businessUnit: "Advertising Sales",
    peopleBusinessUnit: "TS",
  },
  {
    employeeId: "000092",
    costCenter: "",
    jobFamily: "Sales Development Manager",
    businessUnit: "Advertising Sales",
    peopleBusinessUnit: "MS",
  },
  {
    employeeId: "000093",
    costCenter: "",
    jobFamily: "SalesQ IC",
    businessUnit: "Advertising Operations",
    peopleBusinessUnit: "TS",
  },
  {
    employeeId: "000094",
    costCenter: "",
    jobFamily: "SalesQ IC",
    businessUnit: "LCS Operations",
    peopleBusinessUnit: "MS",
  },
  {
    employeeId: "000095",
    costCenter: "",
    jobFamily: "SalesQ IC",
    businessUnit: "Sales Solutions Operations",
    peopleBusinessUnit: "MS",
  },
  {
    employeeId: "000096",
    costCenter: "",
    jobFamily: "SalesQ IC",
    businessUnit: "Global Sales Operations",
    peopleBusinessUnit: "MS",
  },
  {
    employeeId: "000097",
    costCenter: "",
    jobFamily: "SalesQ VP",
    businessUnit: "Global Sales Operations",
    peopleBusinessUnit: "MS",
  },
  {
    employeeId: "000098",
    costCenter: "",
    jobFamily: "Engineering",
    businessUnit: "Other",
    peopleBusinessUnit: "TS",
  },
  {
    employeeId: "000099",
    costCenter: "",
    jobFamily: "Engineering",
    businessUnit: "Other",
    peopleBusinessUnit: "MS",
  },
];
const lobScrById: Record<string, ScrRecord> = {};
const lobPeopleById: Record<string, PeopleRecord> = {};
for (const item of lobMappingCases) {
  lobScrById[item.employeeId] = {
    ...scr(item.employeeId, new Date(2020, 0, 1), `LOB Employee ${item.employeeId}`),
    costCenter: item.costCenter,
    jobFamily: item.jobFamily,
    businessUnit: item.businessUnit,
  };
  lobPeopleById[item.employeeId] = {
    ...people,
    employeeId: item.employeeId,
    fullName: `LOB Employee ${item.employeeId}`,
    businessUnit: item.peopleBusinessUnit,
  };
}
const lobDashboard = buildDashboardModel(lobScrById, lobPeopleById, [], { singapore: "APAC" });
const expectedLobCounts = { GCP: 1, SD: 2, LMS: 2, LTS: 2, LSS: 1, Global: 1 };
for (const [lob, expected] of Object.entries(expectedLobCounts)) {
  if (lobDashboard.byLob.find((row) => row.label === lob)?.employees !== expected) {
    throw new Error(`LOB mapping failed for ${lob}.`);
  }
}

const inferenceData = createEmptyAppData();
const secondPeople = { ...people, employeeId: "000082", fullName: "Second Employee" };
const historicalNewHirePeople = {
  ...people,
  employeeId: "000083",
  fullName: "New Employee",
  analystName: "Historical Analyst",
};
inferenceData.peopleById = {
  [people.employeeId]: people,
  [secondPeople.employeeId]: secondPeople,
  [historicalNewHirePeople.employeeId]: historicalNewHirePeople,
};
inferenceData.currentScrById = {
  [people.employeeId]: scr(people.employeeId, new Date(2020, 0, 1), people.fullName),
  [secondPeople.employeeId]: scr(secondPeople.employeeId, new Date(2020, 0, 1), secondPeople.fullName),
  "000083": scr("000083", new Date(2026, 6, 10), "New Employee"),
  "000084": scr("000084", new Date(2026, 6, 10), "Regional Employee", "Australia"),
};
inferenceData.previousScrById = {
  [people.employeeId]: inferenceData.currentScrById[people.employeeId],
  [secondPeople.employeeId]: inferenceData.currentScrById[secondPeople.employeeId],
};
inferenceData.positionById["000083"] = {
  employeeId: "000083",
  positionName: "New Employee (000083)",
  personName: "New Employee (000083)",
  title: "Account Executive",
  businessGroup: "Sales",
  effectiveStartDate: new Date(2026, 6, 10),
};
inferenceData.positionById["000084"] = {
  ...inferenceData.positionById["000083"],
  employeeId: "000084",
  positionName: "Regional Employee (000084)",
  personName: "Regional Employee (000084)",
};
const inferredAudit = buildAuditReport(
  "JUL-2026",
  { regions: ["APAC"], lobs: ["LSS"], countries: ["Australia", "Singapore"] },
  inferenceData,
  { australia: "APAC", singapore: "APAC" },
  new Date(2026, 6, 16),
);
const inferredExpectation = inferredAudit.expectations.find((item) => item.employeeId === "000083");
const inferredAuditRow = inferredAudit.rows.find((item) => item.employeeId === "000083");
if (
  inferredAuditRow?.analystName !== people.analystName ||
  inferredAuditRow.previousJobLevelGrade !== "" ||
  inferredAuditRow.currentJobLevelGrade !== "IC3 (08.2)" ||
  inferredExpectation?.analystName !== people.analystName ||
  inferredExpectation.analystSource !== "Inferred: Country + LOB" ||
  inferredExpectation.analystConfidence !== "100%" ||
  inferredExpectation.analystSampleSize !== 2
) {
  throw new Error("New Hire analyst inference did not replace the historical assignment with the expected unique leader.");
}
const regionalExpectation = inferredAudit.expectations.find((item) => item.employeeId === "000084");
if (
  regionalExpectation?.analystName !== people.analystName ||
  regionalExpectation.analystSource !== "Inferred: Region + LOB"
) {
  throw new Error("Region + LOB analyst fallback did not select the expected unique leader.");
}
const inferenceDashboard = buildDashboardModel(
  inferenceData.currentScrById,
  inferenceData.peopleById,
  buildFollowUpVerification(inferredAudit.expectations, inferenceData.peopleById).rows,
  { australia: "APAC", singapore: "APAC" },
  "APAC",
);
if (
  inferenceDashboard.commissionedEmployees !== 4 ||
  inferenceDashboard.byAnalyst.find((row) => row.label === people.analystName)?.employees !== 4 ||
  inferenceDashboard.byAnalyst.some((row) => row.label === historicalNewHirePeople.analystName && row.employees > 0)
) {
  throw new Error("SCR-based Dashboard did not retain inferred Analyst ownership for commissioned employees.");
}
const inferredWorkbook = XLSX.read(
  buildAuditWorkbook(inferredAudit.rows, {}, inferredAudit.expectations, inferenceData.currentScrById),
  { type: "array", cellStyles: true },
);
const inferredReportSheet = inferredWorkbook.Sheets["Audit Report"];
const inferredReportRows = XLSX.utils.sheet_to_json<string[]>(inferredReportSheet, { header: 1, defval: "" });
const inferredHeaders = inferredReportRows[0] ?? [];
const inferredRowIndex = inferredReportRows.findIndex(
  (row) => String(row[inferredHeaders.indexOf("employeeId")]) === "000083",
);
const inferredAnalystCell = inferredReportSheet[
  XLSX.utils.encode_cell({ r: inferredRowIndex, c: inferredHeaders.indexOf("analystName") })
] as { s?: { patternType?: string; fgColor?: { rgb?: string } } } | undefined;
if (inferredAnalystCell?.s?.patternType !== "solid" || inferredAnalystCell.s.fgColor?.rgb !== "FFF2CC") {
  throw new Error("Inferred Analyst Name cells are not highlighted light yellow in Audit Excel.");
}

const filters = { regions: ["APAC"], lobs: ["LSS"], countries: ["Singapore"] };
const countryToRegion = { singapore: "APAC" };
const transferOutEmployeeId = "000087";
const terminatedEmployeeId = "000088";
const statusData = createEmptyAppData();
statusData.previousScrById[transferOutEmployeeId] = scr(
  transferOutEmployeeId,
  new Date(2020, 0, 1),
  "Transfer Out Employee",
);
statusData.previousScrById[terminatedEmployeeId] = scr(
  terminatedEmployeeId,
  new Date(2020, 0, 1),
  "Terminated Employee",
);
statusData.currentScrById[terminatedEmployeeId] = {
  ...statusData.previousScrById[terminatedEmployeeId],
  activeStatus: "",
  terminationDate: new Date(2026, 6, 10),
};
statusData.peopleById[transferOutEmployeeId] = {
  ...people,
  employeeId: transferOutEmployeeId,
  fullName: "Transfer Out Employee",
};
statusData.peopleById[terminatedEmployeeId] = {
  ...people,
  employeeId: terminatedEmployeeId,
  fullName: "Terminated Employee",
};
const statusAudit = buildAuditReport("JUL-2026", filters, statusData, countryToRegion, new Date(2026, 6, 16));
if (statusAudit.rows.find((row) => row.employeeId === transferOutEmployeeId)?.auditItem !== "Transfer to Non-Sales") {
  throw new Error("An employee missing from the current SCR was not classified as Transfer to Non-Sales.");
}
if (statusAudit.rows.find((row) => row.employeeId === terminatedEmployeeId)?.auditItem !== "Termination") {
  throw new Error("A blank current SCR Active Status was not classified as Termination.");
}

const changedEmployeeId = "000111";
const changeData = createEmptyAppData();
changeData.previousScrById[changedEmployeeId] = scr(changedEmployeeId, new Date(2020, 0, 1), "Changed Employee");
changeData.currentScrById[changedEmployeeId] = {
  ...changeData.previousScrById[changedEmployeeId],
  jobTitle: "Senior Account Executive",
  jobLevel: "IC4",
  jobGrade: "09.1",
  commissionAmount: 75,
};
changeData.peopleById[changedEmployeeId] = { ...people, employeeId: changedEmployeeId, fullName: "Changed Employee" };
changeData.positionById[changedEmployeeId] = {
  employeeId: changedEmployeeId,
  positionName: "Changed Employee (000111)",
  personName: "Changed Employee (000111)",
  title: "Senior Account Executive",
  businessGroup: "Sales",
  effectiveStartDate: new Date(2020, 0, 1),
};
const changeAudit = buildAuditReport("JUL-2026", filters, changeData, countryToRegion, new Date(2026, 6, 16));
const changeRow = changeAudit.rows.find((row) => row.employeeId === changedEmployeeId);
if (
  changeRow?.auditItem !== "Change to Existing Participant" ||
  changeRow.auditSubcategory !== "Promotion + Variable Change" ||
  changeRow.previousJobLevelGrade !== "IC3 (08.2)" ||
  changeRow.currentJobLevelGrade !== "IC4 (09.1)" ||
  changeAudit.expectations.some((item) => item.auditSubcategory !== "Promotion + Variable Change")
) {
  throw new Error("Audit and Verification Baseline did not carry the promotion classification or Job Level (Grade) values.");
}
const changeWorkbook = XLSX.read(
  buildAuditWorkbook(changeAudit.rows, {}, changeAudit.expectations, changeData.currentScrById),
  { type: "array" },
);
const changeWorkbookRows = XLSX.utils.sheet_to_json<Record<string, string>>(changeWorkbook.Sheets["Audit Report"], {
  defval: "",
});
if (
  changeWorkbookRows[0]?.previousJobLevelGrade !== "IC3 (08.2)" ||
  changeWorkbookRows[0]?.currentJobLevelGrade !== "IC4 (09.1)"
) {
  throw new Error("Audit Excel did not retain the previous and current Job Level (Grade) values.");
}
const changeVerification = buildFollowUpVerification(
  changeAudit.expectations,
  {
    [changedEmployeeId]: {
      ...people,
      employeeId: changedEmployeeId,
      fullName: "Changed Employee",
      hrJobTitle: "Senior Account Executive",
      annualVariable: 75,
    },
  },
  new Date(2026, 6, 22),
);
if (changeVerification.rows[0]?.auditSubcategory !== "Promotion + Variable Change") {
  throw new Error("Verification Report did not retain the audit subcategory.");
}

const careerMovementCases: Array<{
  name: string;
  previous: Partial<ScrRecord>;
  current: Partial<ScrRecord>;
  expectedSubcategory: string;
}> = [
  {
    name: "higher grade with variable change",
    previous: { jobLevel: "IC3", jobGrade: "08.2", commissionAmount: 50 },
    current: { jobLevel: "IC4", jobGrade: "09.1", commissionAmount: 75 },
    expectedSubcategory: "Promotion + Variable Change",
  },
  {
    name: "same-grade IC to manager",
    previous: { jobLevel: "IC3", jobGrade: "08.2" },
    current: { jobLevel: "MR2", jobGrade: "08.2" },
    expectedSubcategory: "Promotion - No Variable Change",
  },
  {
    name: "same-grade IC to SP",
    previous: { jobLevel: "IC1", jobGrade: "06" },
    current: { jobLevel: "SP4", jobGrade: "06" },
    expectedSubcategory: "Job Change - No Variable Change",
  },
  {
    name: "lower grade",
    previous: { jobLevel: "IC4", jobGrade: "09.1" },
    current: { jobLevel: "IC3", jobGrade: "08.2" },
    expectedSubcategory: "Job Change - No Variable Change",
  },
  {
    name: "unknown grade",
    previous: { jobLevel: "IC3", jobGrade: "B" },
    current: { jobLevel: "IC4", jobGrade: "C" },
    expectedSubcategory: "Job Change - No Variable Change",
  },
];
for (const [index, testCase] of careerMovementCases.entries()) {
  const employeeId = `0002${index + 10}`;
  const movementData = createEmptyAppData();
  const previousRecord = {
    ...scr(employeeId, new Date(2020, 0, 1), `Career Case ${index}`),
    ...testCase.previous,
  };
  movementData.previousScrById[employeeId] = previousRecord;
  movementData.currentScrById[employeeId] = { ...previousRecord, ...testCase.current };
  movementData.peopleById[employeeId] = { ...people, employeeId, fullName: `Career Case ${index}` };
  movementData.positionById[employeeId] = {
    employeeId,
    positionName: `Career Case ${index} (${employeeId})`,
    personName: `Career Case ${index} (${employeeId})`,
    title: movementData.currentScrById[employeeId].jobTitle,
    businessGroup: "Sales",
    effectiveStartDate: new Date(2020, 0, 1),
  };
  const movementAudit = buildAuditReport("JUL-2026", filters, movementData, countryToRegion, new Date(2026, 6, 16));
  const movementRow = movementAudit.rows.find((row) => row.employeeId === employeeId);
  if (movementRow?.auditSubcategory !== testCase.expectedSubcategory) {
    throw new Error(
      `${testCase.name} classified as ${movementRow?.auditSubcategory || "no row"}; expected ${testCase.expectedSubcategory}.`,
    );
  }
}

const gradeOnlyEmployeeId = "000299";
const gradeOnlyData = createEmptyAppData();
gradeOnlyData.previousScrById[gradeOnlyEmployeeId] = scr(
  gradeOnlyEmployeeId,
  new Date(2020, 0, 1),
  "Grade Only Employee",
);
gradeOnlyData.currentScrById[gradeOnlyEmployeeId] = {
  ...gradeOnlyData.previousScrById[gradeOnlyEmployeeId],
  jobGrade: "09.1",
};
gradeOnlyData.peopleById[gradeOnlyEmployeeId] = {
  ...people,
  employeeId: gradeOnlyEmployeeId,
  fullName: "Grade Only Employee",
};
gradeOnlyData.positionById[gradeOnlyEmployeeId] = {
  employeeId: gradeOnlyEmployeeId,
  positionName: "Grade Only Employee (000299)",
  personName: "Grade Only Employee (000299)",
  title: "Account Executive",
  businessGroup: "Sales",
  effectiveStartDate: new Date(2020, 0, 1),
};
const gradeOnlyAudit = buildAuditReport("JUL-2026", filters, gradeOnlyData, countryToRegion, new Date(2026, 6, 16));
if (
  gradeOnlyAudit.rows[0]?.auditSubcategory !== "Promotion - No Variable Change" ||
  gradeOnlyAudit.expectations.length !== 1 ||
  gradeOnlyAudit.expectations[0]?.fieldKey !== "jobGrade" ||
  gradeOnlyAudit.expectations[0]?.rule !== "unverifiable"
) {
  throw new Error("A grade-only promotion was not audited with an unverifiable People follow-up expectation.");
}
const gradeOnlyVerification = buildFollowUpVerification(
  gradeOnlyAudit.expectations,
  gradeOnlyData.peopleById,
  new Date(2026, 6, 22),
);
if (gradeOnlyVerification.rows[0]?.progressStatus !== "Not Verifiable") {
  throw new Error("A grade-only promotion did not remain Not Verifiable in the People-only follow-up.");
}

const transferEmployeeId = "000085";
const transferData = createEmptyAppData();
transferData.currentScrById[transferEmployeeId] = scr(
  transferEmployeeId,
  new Date(2020, 0, 1),
  "Transfer Employee",
);
const transferAudit = buildAuditReport("JUL-2026", filters, transferData, countryToRegion, new Date(2026, 6, 16));
assertColumnGuide(buildAuditWorkbook(transferAudit.rows, {}), "Audit Report");
const transferRows = transferAudit.rows.filter((row) => row.employeeId === transferEmployeeId);
if (
  transferRows.length !== 1 ||
  transferRows[0]?.auditItem !== "Transfer to Sales - Xactly Setup Required" ||
  transferRows[0]?.missingPeopleSetup !== "Yes" ||
  transferRows[0]?.missingPositionSetup !== "Yes"
) {
  throw new Error("Transfer to Sales and Missing Xactly Setup were not merged into one audit action.");
}
if (new Set(transferAudit.expectations.map((item) => item.verificationId)).size !== 1) {
  throw new Error("Merged Transfer to Sales audit did not produce one verification action.");
}

const historicalTransferEmployeeId = "000101";
const historicalTransferData = createEmptyAppData();
historicalTransferData.currentScrById[historicalTransferEmployeeId] = scr(
  historicalTransferEmployeeId,
  new Date(2020, 0, 1),
  "Historical Transfer Employee",
);
for (const peerEmployeeId of ["000102", "000103"]) {
  const peerScr = scr(peerEmployeeId, new Date(2020, 0, 1), `Peer Employee ${peerEmployeeId}`);
  historicalTransferData.currentScrById[peerEmployeeId] = peerScr;
  historicalTransferData.previousScrById[peerEmployeeId] = peerScr;
  historicalTransferData.peopleById[peerEmployeeId] = {
    ...people,
    employeeId: peerEmployeeId,
    fullName: `Peer Employee ${peerEmployeeId}`,
    analystName: "New Analyst",
  };
}
historicalTransferData.peopleById[historicalTransferEmployeeId] = {
  ...people,
  employeeId: historicalTransferEmployeeId,
  fullName: "Historical Transfer Employee",
  analystName: "Old Analyst",
};
historicalTransferData.positionById[historicalTransferEmployeeId] = {
  employeeId: historicalTransferEmployeeId,
  positionName: "Historical Transfer Employee (000101)",
  personName: "Historical Transfer Employee (000101)",
  title: "Account Executive",
  businessGroup: "Sales",
  effectiveStartDate: new Date(2020, 0, 1),
};
const historicalTransferAudit = buildAuditReport(
  "JUL-2026",
  filters,
  historicalTransferData,
  countryToRegion,
  new Date(2026, 6, 16),
);
const historicalTransferRow = historicalTransferAudit.rows.find(
  (row) => row.employeeId === historicalTransferEmployeeId,
);
const historicalTransferExpectation = historicalTransferAudit.expectations.find(
  (item) => item.employeeId === historicalTransferEmployeeId,
);
if (
  historicalTransferRow?.auditItem !== "Transfer to Sales" ||
  historicalTransferRow.analystName !== "New Analyst" ||
  historicalTransferExpectation?.analystName !== "New Analyst" ||
  historicalTransferExpectation.analystSource !== "Inferred: Country + LOB" ||
  historicalTransferExpectation.analystSampleSize !== 2
) {
  throw new Error("Transfer to Sales reused the employee's historical People Analyst instead of inferring ownership.");
}
const historicalTransferVerification = buildFollowUpVerification(
  historicalTransferAudit.expectations,
  historicalTransferData.peopleById,
  new Date(2026, 6, 29),
);
const historicalTransferDashboard = buildDashboardModel(
  historicalTransferData.currentScrById,
  historicalTransferData.peopleById,
  historicalTransferVerification.rows,
  countryToRegion,
);
if (
  historicalTransferDashboard.byAnalyst.find((row) => row.label === "New Analyst")?.employees !== 3 ||
  historicalTransferDashboard.byAnalyst.some((row) => row.label === "Old Analyst" && row.employees > 0)
) {
  throw new Error("Dashboard reused the historical People Analyst for a Transfer to Sales employee.");
}

const loaEmployeeId = "000086";
const loaData = createEmptyAppData();
const loaCurrent = {
  ...scr(loaEmployeeId, new Date(2020, 0, 1), "LOA Employee"),
  supervisoryManager: "New Manager (222222)",
};
const loaPrevious = {
  ...loaCurrent,
  onLeave: "Yes",
  supervisoryManager: "Old Manager (111111)",
};
loaData.currentScrById[loaEmployeeId] = loaCurrent;
loaData.previousScrById[loaEmployeeId] = loaPrevious;
loaData.peopleById[loaEmployeeId] = {
  ...people,
  employeeId: loaEmployeeId,
  fullName: "LOA Employee",
  firstName: "LOA",
  lastName: "Employee",
  businessUnit: "Sales Solutions",
  employeeStatus: "LOA",
  level1Manager: loaPrevious.supervisoryManager,
};
loaData.positionById[loaEmployeeId] = {
  employeeId: loaEmployeeId,
  positionName: "LOA Employee (000086)",
  personName: "LOA Employee (000086)",
  title: "Account Executive",
  businessGroup: "Sales",
  effectiveStartDate: new Date(2020, 0, 1),
};
loaData.loaById[loaEmployeeId] = {
  employeeId: loaEmployeeId,
  region: "APAC",
  firstDayOfLeave: new Date(2026, 5, 1),
  estimatedLastDayOfLeave: new Date(2026, 6, 15),
  totalDaysOnLeave: "45",
  dateTimeCompleted: null,
  latestCorrection: null,
};
const loaAudit = buildAuditReport("JUL-2026", filters, loaData, countryToRegion, new Date(2026, 6, 16));
const loaRows = loaAudit.rows.filter((row) => row.employeeId === loaEmployeeId);
if (loaRows.length !== 1 || loaRows[0]?.auditItem !== "LOA Return with Participant Changes") {
  throw new Error("LOA Return and participant changes were not merged into one audit action.");
}
const loaFieldKeys = new Set(loaAudit.expectations.map((item) => item.fieldKey));
if (!loaFieldKeys.has("level1Manager") || !loaFieldKeys.has("employeeStatus")) {
  throw new Error("Merged LOA Return verification did not retain change and status expectations.");
}

const transferFollowUpPeople: PeopleRecord = {
  ...people,
  employeeId: transferEmployeeId,
  fullName: "Transfer Employee",
  firstName: "Transfer",
  lastName: "Employee",
  businessUnit: "Sales Solutions",
  level1Manager: "Manager (123456)",
};
const loaFollowUpPeople: PeopleRecord = {
  ...loaData.peopleById[loaEmployeeId],
  employeeStatus: "Active",
  level1Manager: loaCurrent.supervisoryManager,
  uploadDate: new Date(2026, 6, 29),
};
const mergedVerification = buildFollowUpVerification(
  [...transferAudit.expectations, ...loaAudit.expectations],
  { [transferEmployeeId]: transferFollowUpPeople, [loaEmployeeId]: loaFollowUpPeople },
  new Date(2026, 6, 29),
);
const mergedDashboard = buildDashboardModel(
  {
    [transferEmployeeId]: transferData.currentScrById[transferEmployeeId],
    [loaEmployeeId]: loaData.currentScrById[loaEmployeeId],
  },
  { [transferEmployeeId]: transferFollowUpPeople, [loaEmployeeId]: loaFollowUpPeople },
  mergedVerification.rows,
  countryToRegion,
  "APAC",
);
if (mergedVerification.rows.length !== 2 || mergedDashboard.setupRequired !== 2) {
  throw new Error("Dashboard did not count merged audit combinations as one setup action per employee.");
}

console.log("Follow-up verification smoke test passed.");
