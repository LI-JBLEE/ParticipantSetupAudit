import XLSXImport from "xlsx-js-style";
import type { ColInfo, WorkBook, WorkSheet } from "xlsx-js-style";
import { DEFAULT_COUNTRY_REGION_MAP } from "./countryRegionMap";
import type {
  AppData,
  AuditBuildResult,
  AuditRow,
  BalanceRow,
  BalanceSummary,
  DashboardBreakdownRow,
  DashboardModel,
  FileParseResult,
  FilterOptions,
  Filters,
  FollowUpBuildResult,
  LoaRecord,
  MsftTransferRecord,
  PeopleRecord,
  PositionRecord,
  ProcessingMonthOption,
  QuotaAssignmentRow,
  ScrRecord,
  UploadDefinition,
  VerificationExpectation,
  VerificationFieldResult,
  VerificationProgressStatus,
  VerificationResultRow,
  VerificationRule,
  VerificationSlaStatus,
  WorkerChangeRecord,
  WorkerChangeSignal,
} from "./types";

const XLSX = ((XLSXImport as unknown as { default?: typeof XLSXImport }).default ?? XLSXImport) as typeof XLSXImport;

const MONTH_NAMES = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
const REGION_OPTIONS = ["APAC", "EMEA", "LATAM", "NAMER"];
const CURRENTLY_ON_LOA_PREFIX = "[Currently on LOA]";
const NEGATIVE_BALANCE_MATERIALITY_THRESHOLD = 1;
const VARIABLE_COMPENSATION_MATERIALITY_THRESHOLD = 10;
const JOB_GRADE_ORDER = ["4", "5", "6", "7", "8.1", "8.2", "9.1", "9.2", "9.3", "10", "11", "12", "a"];

const AUDIT_COLUMN_DESCRIPTIONS: Record<keyof AuditRow, string> = {
  auditItem: "Consolidated audit action or issue identified for the employee.",
  auditSubcategory: "Operational subtype of a participant change, separating variable-related and other change patterns.",
  processingMonth: "Month selected when the audit was generated.",
  employeeId: "Unique employee identifier.",
  employeeName: "Employee's full name.",
  region: "Employee region derived from the configured country-to-region mapping.",
  lob: "Employee line of business derived from the configured SCR and People mapping rules.",
  country: "Employee country used for audit filtering and ownership mapping.",
  currentActiveStatus: "Active Status from the current-month SCR.",
  currentOnLeave: "Current leave-of-absence indicator from the SCR.",
  currentFirstDayOfLeave: "First day of leave recorded in the current SCR.",
  changeSummary: "Brief explanation of the setup action, change, or data issue detected.",
  wcrEffectiveDate: "Relevant effective date from the Worker Change Report. Multiple matching dates are separated by semicolons.",
  peoplePlanEffectiveDate: "Current plan effective date in the People record.",
  peopleBusinessUnit: "Current Business_Unit value in the People record.",
  analystName: "Analyst assigned or inferred to own the setup action. Light yellow cells indicate an inferred assignment.",
  inferredAnalystName: "Analyst suggested from the employee's current Country and derived LOB. Light yellow indicates an app inference.",
  analystReview: "Review outcome comparing the current People Analyst with the inferred Analyst.",
  inferenceBasis: "Inference level, confidence, and supporting employee count, or the reason no routing recommendation was made.",
  planType: "Current plan type in the People record.",
  hireDate: "Hire date relevant to a new-hire or rehire audit item.",
  terminationDate: "Termination date relevant to a termination audit item.",
  rehireInPeople: "Indicates whether the People record reflects the employee as a rehire.",
  negativeBalance: "Material negative payment balance identified for the employee.",
  missingPeopleSetup: "Indicates whether the employee is missing from the People setup.",
  missingPositionSetup: "Indicates whether the employee is missing from the Position setup.",
  previousJobTitle: "Job title in the previous-month SCR.",
  currentJobTitle: "Job title in the current-month SCR.",
  previousJobLevelGrade: "Job level and global job grade in the previous-month SCR, displayed as Level (Grade).",
  currentJobLevelGrade: "Job level and global job grade in the current-month SCR, displayed as Level (Grade).",
  previousSupervisoryManager: "Supervisory manager in the previous-month SCR.",
  currentSupervisoryManager: "Supervisory manager in the current-month SCR.",
  previousCommissionAmount: "Commission amount in the previous-month SCR.",
  currentCommissionAmount: "Commission amount in the current-month SCR.",
  peopleAnnualVariable: "Annual_Variable value in the current People record when the SCR-to-People gap is 10 or more.",
  variableCompensationGap: "Current SCR Commission Amount minus People Annual_Variable when the absolute gap is 10 or more.",
  previousBusinessUnit: "Business unit in the previous-month SCR.",
  currentBusinessUnit: "Business unit in the current-month SCR.",
  previousCountry: "Country in the previous-month SCR.",
  currentCountry: "Country in the current-month SCR.",
  previousCurrency: "Currency in the previous-month SCR.",
  currentCurrency: "Currency in the current-month SCR.",
  loaFirstDayOfLeave: "First day of leave from the LOA report.",
  loaEstimatedLastDay: "Estimated last day of leave from the LOA report.",
  loaTotalDays: "Total leave duration from the LOA report.",
  okrStartMonth: "Start month of the relevant OKR assignment.",
  okrEndMonth: "End month of the relevant OKR assignment.",
  transferDirection: "Detected direction of the employee's sales-role transfer.",
  microsoftTransfer: "Indicates whether the employee appears in the Transfer to MSFT file.",
  peopleUploadDate: "Upload date of the People record used in the audit.",
};

const VERIFICATION_COLUMN_DESCRIPTIONS: Record<keyof VerificationResultRow, string> = {
  verificationId: "Unique identifier for the employee audit action being verified.",
  processingMonth: "Processing month carried forward from the initial audit.",
  employeeId: "Unique employee identifier.",
  employeeName: "Employee's full name.",
  region: "Employee region recorded in the initial audit baseline.",
  lob: "Employee line of business recorded in the initial audit baseline.",
  country: "Employee country recorded in the initial audit baseline.",
  analystName: "Analyst assigned or inferred to own the setup action.",
  inferredAnalystName: "Analyst suggested during the initial audit for an existing participant whose routing key changed.",
  analystReview: "Initial audit comparison between the current and inferred Analyst.",
  inferenceBasis: "Inference level, confidence, and supporting employee count carried from the initial audit.",
  analystSource: "Source of the analyst assignment, such as People or an inferred mapping.",
  analystConfidence: "Confidence percentage for an inferred analyst assignment.",
  analystSampleSize: "Number of existing employees supporting the inferred analyst assignment.",
  auditItem: "Audit action or issue carried forward from the initial audit.",
  auditSubcategory: "Operational change subtype carried forward from the initial audit baseline.",
  wcrEffectiveDate: "Worker Change Report effective date carried forward from the initial audit.",
  progressStatus: "Overall result: Completed, Partially Completed, Pending, Manager Mismatch Only, Deferred, or Not Verifiable.",
  slaStatus: "Timeliness result based on the due date and follow-up People snapshot.",
  baselineGeneratedAt: "Date and time when the initial verification baseline was generated.",
  dueDate: "Expected completion date, seven days after the initial audit was generated.",
  followUpPeopleDate: "Snapshot or upload date of the follow-up People file.",
  completedDate: "People upload date recorded when all verifiable fields are completed.",
  timely: "Yes when a completed setup was reflected by the due date; otherwise No.",
  completedFields: "Fields whose follow-up People values match the expected values.",
  pendingFields: "Verifiable fields that do not yet match the expected values.",
  notVerifiableFields: "Fields that cannot be confirmed from a People-only follow-up.",
  verificationNotes: "Notes explaining verification limitations or special handling.",
};

const UPLOAD_DEFINITIONS: UploadDefinition[] = [
  { key: "currentScr", label: "Sales Compensation Report (Current Month)", accept: ".xlsx,.xls" },
  { key: "previousScr", label: "Sales Compensation Report (Previous Month)", accept: ".xlsx,.xls" },
  { key: "people", label: "People", accept: ".xlsx,.xls" },
  { key: "position", label: "Position", accept: ".xlsx,.xls" },
  {
    key: "workerChangeReport",
    label: "Worker Change Report",
    accept: ".xlsx,.xls,.csv",
  },
  { key: "quota", label: "Quota Assignment", accept: ".xlsx,.xls,.csv" },
  { key: "balance", label: "Payment Balance", accept: ".xls,.xlsx" },
  { key: "loa", label: "LOA Report", accept: ".xlsx,.xls" },
  { key: "msftTransfer", label: "Transfer to MSFT", accept: ".xlsx,.xls" },
];

export function getUploadDefinitions(): UploadDefinition[] {
  return UPLOAD_DEFINITIONS;
}

export function createEmptyAppData(): AppData {
  return {
    peopleById: {},
    peopleHistoryById: {},
    positionById: {},
    balanceById: {},
    quotaRows: [],
    loaById: {},
    currentScrById: {},
    previousScrById: {},
    msftTransferById: {},
    workerChangesById: {},
  };
}

export function buildProcessingMonthOptions(today = new Date()): ProcessingMonthOption[] {
  const current = new Date(today.getFullYear(), today.getMonth(), 1);
  const fiscalStartYear = current.getMonth() >= 6 ? current.getFullYear() : current.getFullYear() - 1;
  const fiscalStart = new Date(fiscalStartYear, 6, 1);
  const options: ProcessingMonthOption[] = [];
  for (let cursor = current; cursor >= fiscalStart; cursor = new Date(cursor.getFullYear(), cursor.getMonth() - 1, 1)) {
    options.push({ label: formatMonthKey(cursor), date: new Date(cursor) });
  }
  return options;
}

export async function loadCountryRegionReferenceMap(): Promise<Record<string, string>> {
  return { ...DEFAULT_COUNTRY_REGION_MAP };
}

export function buildFilterOptions(data: AppData, countryToRegion: Record<string, string>): FilterOptions {
  const regionSet = new Set<string>();
  const countrySet = new Set<string>();

  const ids = new Set<string>([
    ...Object.keys(data.peopleById),
    ...Object.keys(data.currentScrById),
    ...Object.keys(data.previousScrById),
    ...Object.keys(data.loaById),
  ]);

  for (const employeeId of ids) {
    const context = resolveEmployeeContext(employeeId, data, countryToRegion);
    regionSet.add(context.region);
  }

  for (const record of Object.values(data.currentScrById)) {
    if (record.country && record.country !== "Unmapped") {
      countrySet.add(record.country);
    }
  }

  return {
    regions: REGION_OPTIONS,
    lobs: collectScrLobOptions(data),
    countries: sortDisplayValues(countrySet),
  };
}

export async function parsePeopleFile(file: File): Promise<FileParseResult<{ byId: Record<string, PeopleRecord>; historyById: Record<string, PeopleRecord[]> }>> {
  const matrix = await readMatrixFromFile(file, 0);
  const headerIndex = findHeaderRow(matrix, ["employeeid", "firstname", "lastname"]);
  if (headerIndex < 0) throw new Error("Could not find the People header row.");
  const header = matrix[headerIndex].map(normalizeHeader);
  const cols = {
    id: findColumn(header, ["employeeid"]),
    firstName: findColumn(header, ["firstname"]),
    lastName: findColumn(header, ["lastname"]),
    region: findColumn(header, ["region"]),
    businessUnit: findColumn(header, ["businessunit"]),
    country: findColumn(header, ["country"]),
    analystName: findColumn(header, ["analystname"]),
    planEffectiveDate: findColumn(header, ["planeffectivedate"]),
    planType: findColumn(header, ["plantype"]),
    effectiveStartDate: findColumn(header, ["effectivestartdate"]),
    uploadDate: findColumn(header, ["uploaddate"]),
    employeeStatus: findColumn(header, ["employeestatus"]),
    terminationDate: findColumn(header, ["terminationdate"]),
    salary: findExactColumn(header, ["salary"], ["salary"]),
    salaryCurrency: findColumn(header, ["salarycurrency"]),
    annualVariable: findExactColumn(header, ["annualvariable"], ["annualvariable"]),
    hrJobTitle: findColumn(header, ["hrjobtitle"]),
    level1Manager: findColumn(header, ["level1manager"]),
  };

  const historyById: Record<string, PeopleRecord[]> = {};
  const byId: Record<string, PeopleRecord> = {};
  let rows = 0;

  for (let rowIndex = headerIndex + 1; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] ?? [];
    const employeeId = normalizeEmployeeIdFromCell(cell(row, cols.id));
    if (!employeeId) continue;

    const firstName = text(cell(row, cols.firstName));
    const lastName = text(cell(row, cols.lastName));
    const record: PeopleRecord = {
      employeeId,
      fullName: `${firstName} ${lastName}`.trim() || employeeId,
      firstName,
      lastName,
      region: displayOrUnmapped(text(cell(row, cols.region))),
      businessUnit: text(cell(row, cols.businessUnit)),
      country: displayOrUnmapped(text(cell(row, cols.country))),
      analystName: text(cell(row, cols.analystName)),
      planEffectiveDate: toDate(cell(row, cols.planEffectiveDate)),
      planType: text(cell(row, cols.planType)),
      effectiveStartDate: toDate(cell(row, cols.effectiveStartDate)),
      uploadDate: toDate(cell(row, cols.uploadDate)),
      employeeStatus: text(cell(row, cols.employeeStatus)),
      terminationDate: toDate(cell(row, cols.terminationDate)),
      salary: toNumber(cell(row, cols.salary)),
      salaryCurrency: text(cell(row, cols.salaryCurrency)),
      annualVariable: toNumber(cell(row, cols.annualVariable)),
      hrJobTitle: text(cell(row, cols.hrJobTitle)),
      level1Manager: text(cell(row, cols.level1Manager)),
    };

    rows += 1;
    if (!historyById[employeeId]) historyById[employeeId] = [];
    historyById[employeeId].push(record);
    byId[employeeId] = pickLatestPeopleRecord(byId[employeeId], record);
  }

  return { fileName: file.name, rows, data: { byId, historyById } };
}

export async function parsePositionFile(file: File): Promise<FileParseResult<Record<string, PositionRecord>>> {
  const matrix = await readMatrixFromFile(file, 0);
  const headerIndex = findHeaderRow(matrix, ["positionname", "employeeid"]);
  if (headerIndex < 0) throw new Error("Could not find the Position header row.");
  const header = matrix[headerIndex].map(normalizeHeader);
  const cols = {
    positionName: findColumn(header, ["positionname"]),
    effectiveStartDate: findColumn(header, ["effectivestartdate"]),
    employeeId: findColumn(header, ["employeeid"]),
    title: findColumn(header, ["title"]),
    personName: findColumn(header, ["personname"]),
    businessGroup: findColumn(header, ["businessgroup"]),
  };

  const byId: Record<string, PositionRecord> = {};
  let rows = 0;

  for (let rowIndex = headerIndex + 1; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] ?? [];
    const employeeId =
      normalizeEmployeeIdFromCell(cell(row, cols.employeeId)) ?? normalizeEmployeeIdFromText(cell(row, cols.personName));
    if (!employeeId) continue;
    const record: PositionRecord = {
      employeeId,
      positionName: text(cell(row, cols.positionName)),
      personName: text(cell(row, cols.personName)),
      title: text(cell(row, cols.title)),
      businessGroup: text(cell(row, cols.businessGroup)),
      effectiveStartDate: toDate(cell(row, cols.effectiveStartDate)),
    };
    rows += 1;
    byId[employeeId] = pickLatestPositionRecord(byId[employeeId], record);
  }

  return { fileName: file.name, rows, data: byId };
}

export async function parseBalanceFile(file: File): Promise<FileParseResult<Record<string, BalanceSummary>>> {
  const matrix = await readMatrixFromFile(file, 0);
  const headerIndex = findHeaderRow(matrix, ["personname", "remainingbalance"]);
  if (headerIndex < 0) throw new Error("Could not find the Payment Balance header row.");
  const header = matrix[headerIndex].map(normalizeHeader);
  const cols = {
    personName: findColumn(header, ["personname"]),
    positionName: findColumn(header, ["positionname"]),
    remainingBalance: findColumn(header, ["remainingbalance"]),
    currency: findColumn(header, ["remainingbalancecurrency"]),
    createdDate: findColumn(header, ["createddate"]),
  };

  const byId: Record<string, BalanceSummary> = {};
  let rows = 0;

  for (let rowIndex = headerIndex + 1; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] ?? [];
    const employeeId = normalizeEmployeeIdFromText(cell(row, cols.personName));
    if (!employeeId) continue;
    const remainingBalance = toNumber(cell(row, cols.remainingBalance));
    if (remainingBalance === null || remainingBalance >= 0) continue;

    const balanceRow: BalanceRow = {
      employeeId,
      personName: text(cell(row, cols.personName)),
      positionName: text(cell(row, cols.positionName)),
      remainingBalance,
      currency: text(cell(row, cols.currency)),
      createdDate: toDate(cell(row, cols.createdDate)),
    };

    if (!byId[employeeId]) {
      byId[employeeId] = { employeeId, negativeTotalByCurrency: {}, rows: [] };
    }
    const currency = balanceRow.currency || "N/A";
    byId[employeeId].negativeTotalByCurrency[currency] =
      (byId[employeeId].negativeTotalByCurrency[currency] ?? 0) + remainingBalance;
    byId[employeeId].rows.push(balanceRow);
    rows += 1;
  }

  return { fileName: file.name, rows, data: byId };
}

export async function parseQuotaAssignmentFile(file: File): Promise<FileParseResult<QuotaAssignmentRow[]>> {
  const matrix = await readMatrixFromFile(file, 0);
  const headerIndex = findHeaderRow(matrix, ["quotaname", "personnameemployeeid"]);
  if (headerIndex < 0) throw new Error("Could not find the Quota Assignment header row.");
  const headerRow = matrix[headerIndex].map(text);
  const header = headerRow.map(normalizeHeader);
  const monthColumns = extractMonthColumns(headerRow);
  const cols = {
    quotaName: findColumn(header, ["quotaname"]),
    type: findColumn(header, ["type"]),
    name: findColumn(header, ["name"]),
    personName: findColumn(header, ["personnameemployeeid"]),
    effectiveStartDate: findColumn(header, ["effectivestartdate"]),
  };

  const rows: QuotaAssignmentRow[] = [];
  for (let rowIndex = headerIndex + 1; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] ?? [];
    const quotaName = text(cell(row, cols.quotaName));
    if (!quotaName.toLowerCase().includes("okr")) continue;
    const personName = text(cell(row, cols.personName));
    const employeeId = normalizeEmployeeIdFromText(personName);
    if (!employeeId) continue;

    const monthValues: Record<string, number> = {};
    for (const [label, columnIndex] of Object.entries(monthColumns)) {
      monthValues[label] = toNumber(cell(row, columnIndex)) ?? 0;
    }
    rows.push({
      employeeId,
      quotaName,
      type: text(cell(row, cols.type)),
      name: text(cell(row, cols.name)),
      personName,
      effectiveStartDate: toDate(cell(row, cols.effectiveStartDate)),
      monthValues,
    });
  }

  return { fileName: file.name, rows: rows.length, data: rows };
}

export async function parseLoaFile(file: File): Promise<FileParseResult<Record<string, LoaRecord>>> {
  const matrix = await readMatrixFromFile(file, 0);
  const headerIndex = findHeaderRow(matrix, ["employeeid", "firstdayofleave", "estimatedlastdayofleave"]);
  if (headerIndex < 0) throw new Error("Could not find the LOA Report header row.");
  const header = matrix[headerIndex].map(normalizeHeader);
  const cols = {
    employeeId: findColumn(header, ["employeeid"]),
    region: findColumn(header, ["region"]),
    firstDayOfLeave: findColumn(header, ["firstdayofleave"]),
    estimatedLastDayOfLeave: findColumn(header, ["estimatedlastdayofleave"]),
    totalDaysOnLeave: findColumn(header, ["totaldaysonleave"]),
    dateTimeCompleted: findColumn(header, ["datetimecompleted"]),
    latestCorrection: findColumn(header, ["datetimeoflatestloacorrection"]),
  };

  const byId: Record<string, LoaRecord> = {};
  let rows = 0;

  for (let rowIndex = headerIndex + 1; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] ?? [];
    const employeeId = normalizeEmployeeIdFromCell(cell(row, cols.employeeId));
    if (!employeeId) continue;
    const record: LoaRecord = {
      employeeId,
      region: displayOrUnmapped(text(cell(row, cols.region))),
      firstDayOfLeave: toDate(cell(row, cols.firstDayOfLeave)),
      estimatedLastDayOfLeave: toDate(cell(row, cols.estimatedLastDayOfLeave)),
      totalDaysOnLeave: text(cell(row, cols.totalDaysOnLeave)),
      dateTimeCompleted: toDate(cell(row, cols.dateTimeCompleted)),
      latestCorrection: toDate(cell(row, cols.latestCorrection)),
    };
    byId[employeeId] = pickPreferredLoaRecord(byId[employeeId], record);
    rows += 1;
  }

  return { fileName: file.name, rows, data: byId };
}

export async function parseScrFile(file: File): Promise<FileParseResult<Record<string, ScrRecord>>> {
  const matrix = await readMatrixFromFile(file, 0);
  const headerIndex = findHeaderRow(matrix, ["employeeid", "activestatus", "hiredate", "businessunit"]);
  if (headerIndex < 0) throw new Error("Could not find the Sales Compensation Report header row.");
  const header = matrix[headerIndex].map(normalizeHeader);
  const cols = {
    employeeId: findColumn(header, ["employeeid"]),
    firstName: findColumn(header, ["firstname"]),
    lastName: findColumn(header, ["lastname"]),
    fullName: findColumn(header, ["fulllegalname"]),
    originalHireDate: findExactColumn(header, ["originalhiredate"], ["originalhiredate"]),
    activeStatus: findColumn(header, ["activestatus"]),
    onLeave: findColumn(header, ["onleave"]),
    firstDayOfLeave: findColumn(header, ["firstdayofleave"]),
    hireDate: findExactColumn(header, ["hiredate"], ["hiredate"]),
    isRehire: findExactColumn(header, ["isrehire"], ["isrehire"]),
    terminationDate: findColumn(header, ["terminationdate"]),
    jobTitle: findColumn(header, ["jobtitle"]),
    jobLevel: findExactColumn(header, ["cfcbcareerbandlevelworker"], ["careerbandlevelworker"]),
    jobGrade: findExactColumn(header, ["cflrvglobaljobgrade"], ["globaljobgrade"]),
    supervisoryManager: findColumn(header, ["supervisorymanager"]),
    oteBaseComm: findColumn(header, ["otebasecomm"]),
    commissionAmount: findColumn(header, ["commissionamount"]),
    costCenter: findExactColumn(header, ["costcenter"], ["costcenter"]),
    jobFamily: findExactColumn(header, ["jobfamily"], ["jobfamily"]),
    businessUnit: findExactColumn(header, ["businessunit"], ["businessunit"]),
    country: findColumn(header, ["country"]),
    currency: findColumn(header, ["currency"]),
  };

  const byId: Record<string, ScrRecord> = {};
  let rows = 0;

  for (let rowIndex = headerIndex + 1; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] ?? [];
    const employeeId = normalizeEmployeeIdFromCell(cell(row, cols.employeeId));
    if (!employeeId) continue;
    const firstName = text(cell(row, cols.firstName));
    const lastName = text(cell(row, cols.lastName));
    const fullName = text(cell(row, cols.fullName)) || `${firstName} ${lastName}`.trim();
    byId[employeeId] = {
      employeeId,
      firstName,
      lastName,
      fullName,
      originalHireDate: toDate(cell(row, cols.originalHireDate)),
      activeStatus: text(cell(row, cols.activeStatus)),
      onLeave: text(cell(row, cols.onLeave)),
      firstDayOfLeave: toDate(cell(row, cols.firstDayOfLeave)),
      hireDate: toDate(cell(row, cols.hireDate)),
      isRehire: text(cell(row, cols.isRehire)),
      terminationDate: toDate(cell(row, cols.terminationDate)),
      jobTitle: text(cell(row, cols.jobTitle)),
      jobLevel: text(cell(row, cols.jobLevel)),
      jobGrade: text(cell(row, cols.jobGrade)),
      supervisoryManager: text(cell(row, cols.supervisoryManager)),
      oteBaseComm: toNumber(cell(row, cols.oteBaseComm)),
      commissionAmount: toNumber(cell(row, cols.commissionAmount)),
      costCenter: text(cell(row, cols.costCenter)),
      jobFamily: text(cell(row, cols.jobFamily)),
      businessUnit: displayOrUnmapped(text(cell(row, cols.businessUnit))),
      country: displayOrUnmapped(text(cell(row, cols.country))),
      currency: text(cell(row, cols.currency)),
    };
    rows += 1;
  }

  return { fileName: file.name, rows, data: byId };
}

export async function parseMsftTransferFile(file: File): Promise<FileParseResult<Record<string, MsftTransferRecord>>> {
  const matrix = await readMatrixFromFile(file, 0);
  const headerIndex = findHeaderRow(matrix, ["subject", "effectivedate"]);
  if (headerIndex < 0) throw new Error("Could not find the Transfer to MSFT header row.");
  const header = matrix[headerIndex].map(normalizeHeader);
  const cols = {
    subject: findColumn(header, ["subject"]),
    effectiveDate: findColumn(header, ["effectivedate"]),
    businessProcessReason: findColumn(header, ["businessprocessreason"]),
    businessUnitOrganization: findColumn(header, ["businessunitorganization"]),
    region: findColumn(header, ["region"]),
  };

  const byId: Record<string, MsftTransferRecord> = {};
  let rows = 0;

  for (let rowIndex = headerIndex + 1; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] ?? [];
    const subject = text(cell(row, cols.subject));
    const employeeId = normalizeEmployeeIdFromText(subject);
    if (!employeeId) continue;
    byId[employeeId] = {
      employeeId,
      subject,
      effectiveDate: toDate(cell(row, cols.effectiveDate)),
      businessProcessReason: text(cell(row, cols.businessProcessReason)),
      businessUnitOrganization: text(cell(row, cols.businessUnitOrganization)),
      region: displayOrUnmapped(text(cell(row, cols.region))),
    };
    rows += 1;
  }

  return { fileName: file.name, rows, data: byId };
}

export async function parseWorkerChangeReportFile(
  file: File,
): Promise<FileParseResult<Record<string, WorkerChangeRecord[]>>> {
  const matrix = await readMatrixFromFile(file, 0);
  const headerIndex = findHeaderRow(matrix, ["employeeid", "effectivedate", "businessprocesstype"]);
  if (headerIndex < 0) throw new Error("Could not find the Worker Change Report header row.");

  const header = matrix[headerIndex].map(normalizeHeader);
  const cols = {
    employeeId: findColumn(header, ["employeeid"]),
    effectiveDate: findColumn(header, ["effectivedate"]),
    businessProcessType: findColumn(header, ["businessprocesstype"]),
    businessProcessReason: findColumn(header, ["businessprocessreason"]),
    jobCurrent: findColumn(header, ["jobprofilecurrent"]),
    jobProposed: findColumn(header, ["jobprofileproposed"]),
    managerCurrent: findColumn(header, ["managercurrent"]),
    managerProposed: findColumn(header, ["managerproposed"]),
    costCenterCurrent: findColumn(header, ["costcentercurrent"]),
    costCenterProposed: findColumn(header, ["costcenterproposed"]),
    companyCurrent: findColumn(header, ["companiescurrent", "companycurrent"]),
    companyProposed: findColumn(header, ["companyproposed", "companiesproposed"]),
    locationCurrent: findColumn(header, ["locationcurrent"]),
    locationProposed: findColumn(header, ["locationproposed"]),
    basePayCurrent: findColumn(header, ["basepaycurrent"]),
    basePayProposed: findColumn(header, ["basepayproposed"]),
    commissionCurrent: findColumn(header, ["commissionamountcurrent"]),
    commissionProposed: findColumn(header, ["commissionamountproposed"]),
  };
  const currencyColumns = header
    .map((value, index) => (value === "currency" ? index : -1))
    .filter((index) => index >= 0);
  const byId: Record<string, WorkerChangeRecord[]> = {};
  let rows = 0;

  for (let rowIndex = headerIndex + 1; rowIndex < matrix.length; rowIndex += 1) {
    const row = matrix[rowIndex] ?? [];
    const employeeId = normalizeEmployeeIdFromCell(cell(row, cols.employeeId));
    if (!employeeId) continue;
    const signals = new Set<WorkerChangeSignal>();
    if (wcrValuesChanged(row, cols.jobCurrent, cols.jobProposed)) signals.add("job");
    if (wcrValuesChanged(row, cols.managerCurrent, cols.managerProposed)) signals.add("manager");
    if (wcrValuesChanged(row, cols.costCenterCurrent, cols.costCenterProposed)) signals.add("businessUnit");
    if (
      wcrValuesChanged(row, cols.companyCurrent, cols.companyProposed) ||
      wcrValuesChanged(row, cols.locationCurrent, cols.locationProposed)
    ) {
      signals.add("country");
    }
    if (wcrValuesChanged(row, cols.basePayCurrent, cols.basePayProposed)) signals.add("ote");
    if (wcrValuesChanged(row, cols.commissionCurrent, cols.commissionProposed)) signals.add("commission");
    if (
      (currencyColumns.length >= 2 && wcrValuesChanged(row, currencyColumns[0] ?? -1, currencyColumns[1] ?? -1)) ||
      (currencyColumns.length >= 4 && wcrValuesChanged(row, currencyColumns[2] ?? -1, currencyColumns[3] ?? -1))
    ) {
      signals.add("currency");
    }
    (byId[employeeId] ??= []).push({
      employeeId,
      effectiveDate: toDate(cell(row, cols.effectiveDate)),
      businessProcessType: text(cell(row, cols.businessProcessType)),
      businessProcessReason: text(cell(row, cols.businessProcessReason)),
      signals: [...signals],
    });
    rows += 1;
  }

  if (rows === 0) throw new Error("The Worker Change Report does not contain any valid employee records.");
  return { fileName: file.name, rows, data: byId };
}

export function buildAuditReport(
  processingMonth: string,
  filters: Filters,
  data: AppData,
  countryToRegion: Record<string, string>,
  generatedAt = new Date(),
): AuditBuildResult {
  const processingMonthDate = parseMonthKey(processingMonth);
  if (!processingMonthDate) throw new Error("Invalid processing month.");

  const previousMonthDate = new Date(processingMonthDate.getFullYear(), processingMonthDate.getMonth() - 1, 1);
  const previousMonthKey = formatMonthKey(previousMonthDate);
  const newHireStart = new Date(previousMonthDate.getFullYear(), previousMonthDate.getMonth(), 16);
  const newHireEnd = new Date(processingMonthDate.getFullYear(), processingMonthDate.getMonth(), 15);
  const transferInCutoff = new Date(previousMonthDate.getFullYear(), previousMonthDate.getMonth(), 15);
  const warnings: string[] = [];
  const rows: AuditRow[] = [];
  const transferToSalesIds = new Set(
    Object.entries(data.currentScrById)
      .filter(
        ([employeeId, current]) =>
          isYes(current.activeStatus) &&
          !data.previousScrById[employeeId] &&
          current.hireDate !== null &&
          current.hireDate < transferInCutoff,
      )
      .map(([employeeId]) => employeeId),
  );
  const newHireIds = new Set(
    Object.values(data.currentScrById)
      .filter(
        (current) =>
          isYes(current.activeStatus) &&
          (!current.terminationDate || isRehireNewHire(current)) &&
          isNewHireInProcessingWindow(current, newHireStart, newHireEnd),
      )
      .map((current) => current.employeeId),
  );
  const sharedActiveIds = intersectKeys(data.currentScrById, data.previousScrById).filter((employeeId) => {
    const current = data.currentScrById[employeeId];
    const previous = data.previousScrById[employeeId];
    return Boolean(current) && Boolean(previous) && isYes(current.activeStatus) && isYes(previous.activeStatus);
  });
  const analystRecommendationIds = new Set(
    sharedActiveIds.filter((employeeId) => {
      const current = data.currentScrById[employeeId];
      const previous = data.previousScrById[employeeId];
      if (!current || !previous) return false;
      const routingFieldChanged =
        compareField("Business Unit", previous.businessUnit, current.businessUnit).changed ||
        compareField("Country", previous.country, current.country).changed;
      return routingFieldChanged && hasAnalystRoutingChange(previous, current, data.peopleById[employeeId]);
    }),
  );
  const inferredOwnershipIds = new Set([...newHireIds, ...transferToSalesIds]);
  const analystInferenceIndex = buildAnalystInferenceIndex(
    data,
    countryToRegion,
    new Set([...inferredOwnershipIds, ...analystRecommendationIds]),
  );

  for (const current of Object.values(data.currentScrById)) {
    if (!isYes(current.activeStatus)) continue;
    if (current.terminationDate && !isRehireNewHire(current)) continue;
    if (!isNewHireInProcessingWindow(current, newHireStart, newHireEnd)) continue;
    const context = resolveEmployeeContext(current.employeeId, data, countryToRegion);
    if (!matchesFilters(context, filters)) continue;
    const people = data.peopleById[current.employeeId];
    const position = data.positionById[current.employeeId];
    const balance = data.balanceById[current.employeeId];
    const missingPeople = !people;
    const missingPosition = !position;
    const materialNegativeBalance = people && balance ? formatMaterialNegativeBalance(balance) : "";
    const analystAssignment = resolveAnalystAssignment(
      current.employeeId,
      data,
      countryToRegion,
      analystInferenceIndex,
      true,
    );
    rows.push(
      createAuditRow("New Hire", processingMonth, current.employeeId, current.fullName || context.name, context, {
        analystName: analystAssignment.name,
        hireDate: formatDate(current.hireDate),
        currentJobTitle: current.jobTitle,
        currentBusinessUnit: current.businessUnit,
        currentCountry: current.country,
        currentCurrency: current.currency,
        currentCommissionAmount: numericCell(current.commissionAmount),
        rehireInPeople: people ? "Yes" : "No",
        missingPeopleSetup: missingPeople ? "Yes" : "No",
        missingPositionSetup: missingPosition ? "Yes" : "No",
        negativeBalance: materialNegativeBalance,
        changeSummary: buildNewHireSummary(Boolean(people), Boolean(materialNegativeBalance), missingPeople, missingPosition),
      }),
    );
  }

  for (const current of Object.values(data.currentScrById)) {
    if (!isYes(current.activeStatus)) continue;
    if (isNewHireInProcessingWindow(current, newHireStart, newHireEnd)) continue;
    if (transferToSalesIds.has(current.employeeId)) continue;
    const missingPeople = !data.peopleById[current.employeeId];
    const missingPosition = !data.positionById[current.employeeId];
    if (!missingPeople && !missingPosition) continue;
    const context = resolveEmployeeContext(current.employeeId, data, countryToRegion);
    if (!matchesFilters(context, filters)) continue;
    rows.push(
      createAuditRow("Missing Xactly Setup", processingMonth, current.employeeId, current.fullName || context.name, context, {
        currentActiveStatus: current.activeStatus,
        currentJobTitle: current.jobTitle,
        currentBusinessUnit: current.businessUnit,
        currentCountry: current.country,
        currentCurrency: current.currency,
        hireDate: formatDate(current.hireDate),
        missingPeopleSetup: missingPeople ? "Yes" : "No",
        missingPositionSetup: missingPosition ? "Yes" : "No",
        changeSummary: buildMissingXactlySetupSummary(missingPeople, missingPosition),
      }),
    );
  }

  for (const employeeId of sharedActiveIds) {
    const current = data.currentScrById[employeeId];
    const previous = data.previousScrById[employeeId];
    if (!current || !previous) continue;
    const context = resolveEmployeeContext(employeeId, data, countryToRegion);
    if (!matchesFilters(context, filters)) continue;

    const changes = [
      compareField("Job Title", previous.jobTitle, current.jobTitle),
      compareField("Job Level", previous.jobLevel, current.jobLevel),
      compareField("Job Grade", previous.jobGrade, current.jobGrade),
      compareField("Supervisory Manager", previous.supervisoryManager, current.supervisoryManager),
      compareField("OTE (Base+Comm)", previous.oteBaseComm, current.oteBaseComm),
      compareField("Commission Amount", previous.commissionAmount, current.commissionAmount),
      compareField("Business Unit", previous.businessUnit, current.businessUnit),
      compareField("Country", previous.country, current.country),
      compareField("Currency", previous.currency, current.currency),
    ].filter((item) => item.changed);
    const previousOnLeave = isYes(previous.onLeave);
    const currentOnLeave = isYes(current.onLeave);
    const isLoaReturn = previousOnLeave && !currentOnLeave;
    const loa = previousOnLeave !== currentOnLeave ? data.loaById[employeeId] : undefined;
    if (previousOnLeave !== currentOnLeave && !loa) warnings.push(`LOA detail not found for employee ${employeeId}.`);

    if (changes.length > 0) {
      const analystRecommendation = deriveExistingAnalystRecommendation(
        employeeId,
        changes,
        previous,
        current,
        data,
        countryToRegion,
        analystInferenceIndex,
      );
      rows.push(
        createAuditRow(
          isLoaReturn
            ? "LOA Return with Participant Changes"
            : currentOnLeave
              ? "Deferred Change While on LOA"
              : "Change to Existing Participant",
          processingMonth,
          employeeId,
          current.fullName || context.name,
          context,
          {
            ...analystRecommendation,
            auditSubcategory: deriveExistingParticipantSubcategory(changes, previous, current),
            previousJobTitle: hasChanged(changes, "Job Title") ? previous.jobTitle : "",
            currentJobTitle: hasChanged(changes, "Job Title") ? current.jobTitle : "",
            previousSupervisoryManager: hasChanged(changes, "Supervisory Manager") ? previous.supervisoryManager : "",
            currentSupervisoryManager: hasChanged(changes, "Supervisory Manager") ? current.supervisoryManager : "",
            previousCommissionAmount: hasChanged(changes, "Commission Amount") ? numericCell(previous.commissionAmount) : "",
            currentCommissionAmount: hasChanged(changes, "Commission Amount") ? numericCell(current.commissionAmount) : "",
            previousBusinessUnit: hasChanged(changes, "Business Unit") ? previous.businessUnit : "",
            currentBusinessUnit: hasChanged(changes, "Business Unit") ? current.businessUnit : "",
            previousCountry: hasChanged(changes, "Country") ? previous.country : "",
            currentCountry: hasChanged(changes, "Country") ? current.country : "",
            previousCurrency: hasChanged(changes, "Currency") ? previous.currency : "",
            currentCurrency: hasChanged(changes, "Currency") ? current.currency : "",
            loaFirstDayOfLeave: isLoaReturn ? formatDate(loa?.firstDayOfLeave ?? null) : "",
            loaEstimatedLastDay: isLoaReturn ? formatDate(loa?.estimatedLastDayOfLeave ?? null) : "",
            loaTotalDays: isLoaReturn ? (loa?.totalDaysOnLeave ?? "") : "",
            changeSummary: isLoaReturn
              ? `LOA Return; ${changes.map((item) => item.label).join(", ")}`
              : changes.map((item) => item.label).join(", "),
          },
        ),
      );
    }

    if (previousOnLeave !== currentOnLeave && !(isLoaReturn && changes.length > 0)) {
      rows.push(
        createAuditRow(
          currentOnLeave ? "LOA Start" : "LOA Return",
          processingMonth,
          employeeId,
          current.fullName || context.name,
          context,
          {
            loaFirstDayOfLeave: formatDate(loa?.firstDayOfLeave ?? null),
            loaEstimatedLastDay: formatDate(loa?.estimatedLastDayOfLeave ?? null),
            loaTotalDays: loa?.totalDaysOnLeave ?? "",
            changeSummary: currentOnLeave ? "On Leave changed from blank to Yes." : "On Leave changed from Yes to blank.",
          },
        ),
      );
    }
  }

  const okrByEmployee = aggregateOkrAssignments(data.quotaRows);
  for (const [employeeId, okrSummary] of Object.entries(okrByEmployee)) {
    if (okrSummary.endMonth !== previousMonthKey) continue;
    const context = resolveEmployeeContext(employeeId, data, countryToRegion);
    if (!matchesFilters(context, filters)) continue;
    rows.push(
      createAuditRow("OKR Plan End", processingMonth, employeeId, context.name, context, {
        okrStartMonth: okrSummary.startMonth,
        okrEndMonth: okrSummary.endMonth,
        changeSummary: "Non-zero OKR quota month ended in the month before the selected processing month.",
      }),
    );
  }

  for (const [employeeId, previous] of Object.entries(data.previousScrById)) {
    if (!isYes(previous.activeStatus)) continue;
    const current = data.currentScrById[employeeId];
    if (!current) {
      const context = resolveEmployeeContext(employeeId, data, countryToRegion);
      if (!matchesFilters(context, filters)) continue;
      const balance = data.balanceById[employeeId];
      rows.push(
        createAuditRow("Transfer to Non-Sales", processingMonth, employeeId, previous.fullName || context.name, context, {
          transferDirection: "Sales to Non-Sales",
          previousJobTitle: previous.jobTitle,
          previousBusinessUnit: previous.businessUnit,
          previousCountry: previous.country,
          negativeBalance: balance ? formatMaterialNegativeBalance(balance) : "",
          changeSummary: "Active in the previous month SCR but missing from the current month SCR.",
        }),
      );
    }
  }

  for (const employeeId of transferToSalesIds) {
    const current = data.currentScrById[employeeId];
    if (!current) continue;
    const context = resolveEmployeeContext(employeeId, data, countryToRegion);
    if (!matchesFilters(context, filters)) continue;
    const balance = data.balanceById[employeeId];
    const missingPeople = !data.peopleById[employeeId];
    const missingPosition = !data.positionById[employeeId];
    const setupRequired = missingPeople || missingPosition;
    const analystAssignment = resolveAnalystAssignment(
      employeeId,
      data,
      countryToRegion,
      analystInferenceIndex,
      true,
    );
    rows.push(
      createAuditRow(setupRequired ? "Transfer to Sales - Xactly Setup Required" : "Transfer to Sales", processingMonth, employeeId, current.fullName || context.name, context, {
        analystName: analystAssignment.name,
        transferDirection: "Non-Sales to Sales",
        hireDate: formatDate(current.hireDate),
        currentJobTitle: current.jobTitle,
        currentBusinessUnit: current.businessUnit,
        currentCountry: current.country,
        currentCurrency: current.currency,
        missingPeopleSetup: missingPeople ? "Yes" : "No",
        missingPositionSetup: missingPosition ? "Yes" : "No",
        negativeBalance: balance ? formatMaterialNegativeBalance(balance) : "",
        changeSummary: setupRequired
          ? `Transfer to Sales requires Xactly setup. ${buildMissingXactlySetupSummary(missingPeople, missingPosition)}`
          : "Active in the current month SCR, not present in the previous month SCR, and hire date is earlier than the previous month 15th.",
      }),
    );
  }

  for (const [employeeId, previous] of Object.entries(data.previousScrById)) {
    if (!isYes(previous.activeStatus)) continue;
    const current = data.currentScrById[employeeId];
    if (!current || isYes(current.activeStatus) || text(current.activeStatus)) continue;
    const context = resolveEmployeeContext(employeeId, data, countryToRegion);
    if (!matchesFilters(context, filters)) continue;
    const transfer = data.msftTransferById[employeeId];
    const balance = data.balanceById[employeeId];
    rows.push(
      createAuditRow("Termination", processingMonth, employeeId, current.fullName || previous.fullName || context.name, context, {
        terminationDate: formatDate(current.terminationDate ?? previous.terminationDate),
        microsoftTransfer: transfer ? "Yes" : "No",
        negativeBalance: balance ? formatMaterialNegativeBalance(balance) : "",
        changeSummary: transfer ? "Employee also appears in the Transfer to MSFT file." : "",
      }),
    );
  }

  for (const employeeId of collectUnmappedWarningCandidateIds(data, okrByEmployee, previousMonthKey)) {
    const context = resolveEmployeeContext(employeeId, data, countryToRegion);
    const issues = getUnmappedDataIssues(employeeId, data, countryToRegion, context);
    if (issues.length === 0) continue;
    const current = data.currentScrById[employeeId];
    const previous = data.previousScrById[employeeId];
    rows.push(
      createAuditRow("Unmapped Data Warning", processingMonth, employeeId, context.name, context, {
        previousBusinessUnit: previous?.businessUnit ?? "",
        currentBusinessUnit: current?.businessUnit ?? "",
        previousCountry: previous?.country ?? "",
        currentCountry: current?.country ?? "",
        changeSummary: `Unmapped data detected: ${issues.join(", ")}.`,
      }),
    );
  }

  appendVariableCompensationMismatches(rows, processingMonth, filters, data, countryToRegion);

  for (const row of rows) {
    row.previousJobLevelGrade = formatJobLevelGrade(data.previousScrById[row.employeeId]);
    row.currentJobLevelGrade = formatJobLevelGrade(data.currentScrById[row.employeeId]);
    row.wcrEffectiveDate = resolveWcrEffectiveDate(row, data.workerChangesById[row.employeeId] ?? []);
  }

  rows.sort((left, right) => {
    const itemCompare = left.auditItem.localeCompare(right.auditItem);
    if (itemCompare !== 0) return itemCompare;
    return left.employeeId.localeCompare(right.employeeId);
  });

  return {
    rows,
    warnings: dedupeStrings(warnings),
    expectations: buildVerificationExpectations(rows, data, countryToRegion, generatedAt),
  };
}

export function buildAuditWorkbook(
  rows: AuditRow[],
  fileNames: Record<string, string>,
  expectations: VerificationExpectation[] = [],
  currentScrById: Record<string, ScrRecord> = {},
): ArrayBuffer {
  const wb = XLSX.utils.book_new();
  const reportRows = rows.map((row) => ({ ...row }));
  const reportSheet = XLSX.utils.json_to_sheet(reportRows);
  applyHeaderStyle(reportSheet);
  if (reportRows.length > 0) {
    const analystColumn = Object.keys(reportRows[0]).indexOf("analystName");
    const inferredAnalystColumn = Object.keys(reportRows[0]).indexOf("inferredAnalystName");
    const inferredVerificationIds = new Set(
      expectations
        .filter((expectation) => expectation.analystSource.startsWith("Inferred:"))
        .map((expectation) => expectation.verificationId),
    );
    if (analystColumn >= 0) {
      reportRows.forEach((row, rowIndex) => {
        const verificationId = `${row.processingMonth}|${row.auditItem}|${row.employeeId}`;
        if (!inferredVerificationIds.has(verificationId)) return;
        const cell = reportSheet[XLSX.utils.encode_cell({ r: rowIndex + 1, c: analystColumn })];
        if (cell) cell.s = { ...(cell.s ?? {}), fill: { patternType: "solid", fgColor: { rgb: "FFF2CC" } } };
      });
    }
    if (inferredAnalystColumn >= 0) {
      reportRows.forEach((row, rowIndex) => {
        if (!row.inferredAnalystName) return;
        const cell = reportSheet[XLSX.utils.encode_cell({ r: rowIndex + 1, c: inferredAnalystColumn })];
        if (cell) cell.s = { ...(cell.s ?? {}), fill: { patternType: "solid", fgColor: { rgb: "FFF2CC" } } };
      });
    }
  }
  reportSheet["!autofilter"] = {
    ref: XLSX.utils.encode_range(
      reportSheet["!ref"] ? XLSX.utils.decode_range(reportSheet["!ref"]) : { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } },
    ),
  };
  reportSheet["!cols"] = buildColumnWidths(reportRows);
  XLSX.utils.book_append_sheet(wb, reportSheet, "Audit Report");
  appendColumnGuide(wb, AUDIT_COLUMN_DESCRIPTIONS);

  const summaryRows = [
    ...Object.entries(fileNames).map(([key, value]) => ({ Section: "Uploaded File", Name: key, Value: value })),
    ...Object.entries(countByAuditItem(rows)).map(([key, value]) => ({ Section: "Audit Count", Name: key, Value: value })),
    { Section: "Audit Count", Name: "Total Rows", Value: rows.length },
    ...Object.entries(countByAuditSubcategory(rows)).map(([key, value]) => ({
      Section: "Audit Subcategory Count",
      Name: key,
      Value: value,
    })),
  ];
  const summarySheet = XLSX.utils.json_to_sheet(summaryRows);
  applyHeaderStyle(summarySheet);
  summarySheet["!cols"] = buildColumnWidths(summaryRows);
  XLSX.utils.book_append_sheet(wb, summarySheet, "Summary");

  const baselineSheet = XLSX.utils.json_to_sheet(expectations);
  applyHeaderStyle(baselineSheet);
  baselineSheet["!autofilter"] = {
    ref: XLSX.utils.encode_range(
      baselineSheet["!ref"] ? XLSX.utils.decode_range(baselineSheet["!ref"]) : { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } },
    ),
  };
  baselineSheet["!cols"] = buildColumnWidths(expectations);
  XLSX.utils.book_append_sheet(wb, baselineSheet, "Verification Baseline");

  const populationRows = Object.values(currentScrById)
    .filter((current) => isYes(current.activeStatus))
    .map((current) => ({
      employeeId: current.employeeId,
      fullName: current.fullName,
      activeStatus: current.activeStatus,
      onLeave: current.onLeave,
      jobLevel: current.jobLevel,
      jobGrade: current.jobGrade,
      costCenter: current.costCenter,
      jobFamily: current.jobFamily,
      businessUnit: current.businessUnit,
      country: current.country,
    }));
  const populationSheet = XLSX.utils.json_to_sheet(populationRows);
  applyHeaderStyle(populationSheet);
  populationSheet["!cols"] = buildColumnWidths(populationRows);
  XLSX.utils.book_append_sheet(wb, populationSheet, "SCR Population");

  return XLSX.write(wb, { bookType: "xlsx", type: "array", cellStyles: true });
}

export function buildDownloadFileName(now = new Date()): string {
  const yyyy = now.getFullYear();
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  const hh = String(now.getHours()).padStart(2, "0");
  const mi = String(now.getMinutes()).padStart(2, "0");
  const ss = String(now.getSeconds()).padStart(2, "0");
  return `Participant_Setup_Audit_${yyyy}${mm}${dd}_${hh}${mi}${ss}.xlsx`;
}

export async function parseVerificationBaselineFile(
  file: File,
): Promise<FileParseResult<{ expectations: VerificationExpectation[]; currentScrById: Record<string, ScrRecord> }>> {
  const workbook = normalizeWorkbookRanges(XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false }));
  const sheet = workbook.Sheets["Verification Baseline"];
  if (!sheet) throw new Error("The selected workbook does not contain a Verification Baseline sheet.");

  const rawRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: "", raw: true });
  const requiredColumns = ["verificationId", "employeeId", "auditItem", "fieldKey", "expectedValue", "rule"];
  if (rawRows.length === 0 || requiredColumns.some((column) => !(column in rawRows[0]))) {
    throw new Error("The Verification Baseline sheet is missing required columns.");
  }

  const validRules = new Set<VerificationRule>(["exists", "text", "number", "date", "oneOf", "unverifiable"]);
  const expectations = rawRows.map((row) => {
    const rule = text(row.rule) as VerificationRule;
    if (!validRules.has(rule)) throw new Error(`Unknown verification rule: ${rule || "(blank)"}.`);
    return {
      verificationId: text(row.verificationId),
      processingMonth: text(row.processingMonth),
      generatedAt: text(row.generatedAt),
      dueDate: text(row.dueDate),
      employeeId: normalizeEmployeeIdFromCell(row.employeeId) ?? text(row.employeeId),
      employeeName: text(row.employeeName),
      region: text(row.region),
      lob: text(row.lob),
      country: text(row.country),
      analystName: text(row.analystName),
      inferredAnalystName: text(row.inferredAnalystName),
      analystReview: text(row.analystReview),
      inferenceBasis: text(row.inferenceBasis),
      analystSource: text(row.analystSource) || (text(row.analystName) ? "People" : "Unassigned"),
      analystConfidence: text(row.analystConfidence),
      analystSampleSize: toNumber(row.analystSampleSize) ?? 0,
      auditItem: text(row.auditItem),
      auditSubcategory: text(row.auditSubcategory),
      wcrEffectiveDate: text(row.wcrEffectiveDate),
      fieldKey: text(row.fieldKey),
      fieldLabel: text(row.fieldLabel),
      baselineValue: text(row.baselineValue),
      expectedValue: text(row.expectedValue),
      rule,
      deferred: text(row.deferred),
      note: text(row.note),
    } satisfies VerificationExpectation;
  });

  const currentScrById: Record<string, ScrRecord> = {};
  const populationSheet = workbook.Sheets["SCR Population"];
  if (populationSheet) {
    const populationRows = XLSX.utils.sheet_to_json<Record<string, unknown>>(populationSheet, { defval: "", raw: true });
    for (const row of populationRows) {
      const employeeId = normalizeEmployeeIdFromCell(row.employeeId) ?? text(row.employeeId);
      if (!employeeId) continue;
      currentScrById[employeeId] = {
        employeeId,
        firstName: "",
        lastName: "",
        fullName: text(row.fullName),
        originalHireDate: null,
        activeStatus: text(row.activeStatus),
        onLeave: text(row.onLeave),
        firstDayOfLeave: null,
        hireDate: null,
        isRehire: "",
        terminationDate: null,
        jobTitle: "",
        jobLevel: text(row.jobLevel),
        jobGrade: text(row.jobGrade),
        supervisoryManager: "",
        oteBaseComm: null,
        commissionAmount: null,
        costCenter: text(row.costCenter),
        jobFamily: text(row.jobFamily),
        businessUnit: text(row.businessUnit),
        country: text(row.country),
        currency: "",
      };
    }
  }

  return { fileName: file.name, rows: expectations.length, data: { expectations, currentScrById } };
}

export function buildFollowUpVerification(
  expectations: VerificationExpectation[],
  peopleById: Record<string, PeopleRecord>,
  snapshotDate = findLatestPeopleDate(peopleById) ?? new Date(),
): FollowUpBuildResult {
  const grouped = new Map<string, VerificationExpectation[]>();
  for (const expectation of expectations) {
    const group = grouped.get(expectation.verificationId);
    if (group) group.push(expectation);
    else grouped.set(expectation.verificationId, [expectation]);
  }

  const fieldResults: VerificationFieldResult[] = [];
  const rows: VerificationResultRow[] = [];
  const followUpPeopleDate = formatDate(snapshotDate);

  for (const group of grouped.values()) {
    const first = group[0];
    if (!first) continue;
    const people = peopleById[first.employeeId];
    const results = group.map((expectation) => {
      const actualValue = getPeopleVerificationValue(people, expectation.fieldKey);
      const matched = matchesExpectation(expectation, actualValue);
      return { ...expectation, actualValue, matched: matched ? "Yes" : "No" } satisfies VerificationFieldResult;
    });
    fieldResults.push(...results);

    const verifiable = results.filter((result) => result.rule !== "unverifiable");
    const completed = verifiable.filter((result) => result.matched === "Yes");
    const deferred = results.some((result) => isYes(result.deferred));
    const auditSubcategory = first.auditSubcategory || deriveVerificationSubcategory(results);
    const managerMismatchOnly =
      first.auditItem === "Change to Existing Participant" &&
      results.length === 1 &&
      results[0]?.fieldKey === "level1Manager";
    let progressStatus: VerificationProgressStatus;
    if (managerMismatchOnly) progressStatus = "Manager Mismatch Only";
    else if (verifiable.length === 0) progressStatus = "Not Verifiable";
    else if (completed.length === verifiable.length) progressStatus = "Completed";
    else if (deferred) progressStatus = "Deferred";
    else if (completed.length > 0) progressStatus = "Partially Completed";
    else progressStatus = "Pending";

    const completedDate = progressStatus === "Completed" ? formatDate(people?.uploadDate ?? snapshotDate) : "";
    const dueDate = parseIsoDate(first.dueDate);
    const comparisonDate = parseIsoDate(completedDate) ?? snapshotDate;
    let slaStatus: VerificationSlaStatus = "Not Applicable";
    if (
      progressStatus !== "Manager Mismatch Only" &&
      progressStatus !== "Deferred" &&
      progressStatus !== "Not Verifiable" &&
      dueDate
    ) {
      slaStatus = comparisonDate.getTime() <= dueDate.getTime() ? "On Time" : "Overdue";
      if (progressStatus !== "Completed" && snapshotDate.getTime() <= dueDate.getTime()) slaStatus = "Not Due";
    }

    rows.push({
      verificationId: first.verificationId,
      processingMonth: first.processingMonth,
      employeeId: first.employeeId,
      employeeName: first.employeeName,
      region: first.region,
      lob: first.lob,
      country: first.country,
      analystName: first.analystName,
      inferredAnalystName: first.inferredAnalystName,
      analystReview: first.analystReview,
      inferenceBasis: first.inferenceBasis,
      analystSource: first.analystSource,
      analystConfidence: first.analystConfidence,
      analystSampleSize: first.analystSampleSize,
      auditItem: first.auditItem,
      auditSubcategory,
      wcrEffectiveDate: first.wcrEffectiveDate,
      progressStatus,
      slaStatus,
      baselineGeneratedAt: first.generatedAt,
      dueDate: first.dueDate,
      followUpPeopleDate,
      completedDate,
      timely: progressStatus === "Completed" ? (slaStatus === "On Time" ? "Yes" : "No") : "",
      completedFields: results.filter((result) => result.matched === "Yes").map((result) => result.fieldLabel).join(", "),
      pendingFields: results
        .filter((result) => result.rule !== "unverifiable" && result.matched !== "Yes")
        .map((result) => result.fieldLabel)
        .join(", "),
      notVerifiableFields: results
        .filter((result) => result.rule === "unverifiable")
        .map((result) => result.fieldLabel)
        .join(", "),
      verificationNotes: dedupeStrings(results.map((result) => result.note).filter(Boolean)).join(" | "),
    });
  }

  rows.sort((left, right) => {
    const statusCompare = left.progressStatus.localeCompare(right.progressStatus);
    if (statusCompare !== 0) return statusCompare;
    return left.employeeId.localeCompare(right.employeeId);
  });

  const warnings = followUpPeopleDate ? [] : ["No People Upload_Date was available; the current date was used for SLA comparison."];
  return { rows, fieldResults, warnings };
}

export function buildFollowUpWorkbook(
  result: FollowUpBuildResult,
  fileNames: Record<string, string>,
): ArrayBuffer {
  const wb = XLSX.utils.book_new();
  const reportSheet = XLSX.utils.json_to_sheet(result.rows);
  applyHeaderStyle(reportSheet);
  if (result.rows.length > 0) {
    const inferredAnalystColumn = Object.keys(result.rows[0]).indexOf("inferredAnalystName");
    if (inferredAnalystColumn >= 0) {
      result.rows.forEach((row, rowIndex) => {
        if (!row.inferredAnalystName) return;
        const cell = reportSheet[XLSX.utils.encode_cell({ r: rowIndex + 1, c: inferredAnalystColumn })];
        if (cell) cell.s = { ...(cell.s ?? {}), fill: { patternType: "solid", fgColor: { rgb: "FFF2CC" } } };
      });
    }
  }
  reportSheet["!cols"] = buildColumnWidths(result.rows);
  XLSX.utils.book_append_sheet(wb, reportSheet, "Verification Report");
  appendColumnGuide(wb, VERIFICATION_COLUMN_DESCRIPTIONS);

  const detailSheet = XLSX.utils.json_to_sheet(result.fieldResults);
  applyHeaderStyle(detailSheet);
  detailSheet["!cols"] = buildColumnWidths(result.fieldResults);
  XLSX.utils.book_append_sheet(wb, detailSheet, "Field Details");

  const summaryRows = [
    ...Object.entries(fileNames).map(([key, value]) => ({ Section: "Uploaded File", Name: key, Value: value })),
    ...Object.entries(countVerificationStatuses(result.rows)).map(([key, value]) => ({
      Section: "Verification Status",
      Name: key,
      Value: value,
    })),
    { Section: "Verification Status", Name: "Total Rows", Value: result.rows.length },
    ...Object.entries(countByVerificationSubcategory(result.rows)).map(([key, value]) => ({
      Section: "Audit Subcategory Count",
      Name: key,
      Value: value,
    })),
  ];
  const summarySheet = XLSX.utils.json_to_sheet(summaryRows);
  applyHeaderStyle(summarySheet);
  summarySheet["!cols"] = buildColumnWidths(summaryRows);
  XLSX.utils.book_append_sheet(wb, summarySheet, "Summary");
  return XLSX.write(wb, { bookType: "xlsx", type: "array", cellStyles: true });
}

export function buildFollowUpDownloadFileName(now = new Date()): string {
  return buildDownloadFileName(now).replace("Participant_Setup_Audit_", "Participant_Setup_Verification_");
}

export function summarizeSetupExecution(rows: VerificationResultRow[]) {
  const setupRows = rows.filter(
    (row) =>
      row.progressStatus === "Completed" ||
      row.progressStatus === "Partially Completed" ||
      row.progressStatus === "Pending",
  );
  const completed = setupRows.filter((row) => row.progressStatus === "Completed").length;
  return {
    setupRows,
    setupRequired: setupRows.length,
    completed,
    partiallyCompleted: setupRows.filter((row) => row.progressStatus === "Partially Completed").length,
    pending: setupRows.filter((row) => row.progressStatus === "Pending").length,
    managerMismatchOnly: rows.filter((row) => row.progressStatus === "Manager Mismatch Only").length,
    completionRate: setupRows.length > 0 ? completed / setupRows.length : 0,
  };
}

export function buildDashboardModel(
  currentScrById: Record<string, ScrRecord>,
  peopleById: Record<string, PeopleRecord>,
  verificationRows: VerificationResultRow[],
  countryToRegion: Record<string, string>,
  selectedRegion = "All Regions",
): DashboardModel {
  const analystData: AnalystData = { peopleById, currentScrById, previousScrById: {} };
  const inferredOwnershipIds = new Set(
    verificationRows.filter((row) => isInferredOwnershipAuditItem(row.auditItem)).map((row) => row.employeeId),
  );
  const analystIndex = buildAnalystInferenceIndex(analystData, countryToRegion, inferredOwnershipIds);
  const commissioned = Object.values(currentScrById)
    .filter((current) => isYes(current.activeStatus))
    .map((current) => ({
      region: dashboardRegion(countryToRegion[normalizeText(current.country)] ?? ""),
      lob: deriveLob(current, peopleById[current.employeeId]),
      analystName: resolveAnalystAssignment(
        current.employeeId,
        analystData,
        countryToRegion,
        analystIndex,
        inferredOwnershipIds.has(current.employeeId),
      ).name,
    }));
  const regionOptions = sortDisplayValues(new Set(commissioned.map((person) => person.region)));
  const visiblePeople = selectedRegion === "All Regions"
    ? commissioned
    : commissioned.filter((person) => person.region === selectedRegion);
  const regionVerificationRows = selectedRegion === "All Regions"
    ? verificationRows
    : verificationRows.filter((row) => row.region === selectedRegion);
  const execution = summarizeSetupExecution(regionVerificationRows);

  return {
    regionOptions,
    commissionedEmployees: visiblePeople.length,
    setupRequired: execution.setupRequired,
    setupRequiredRate: visiblePeople.length > 0 ? execution.setupRequired / visiblePeople.length : 0,
    completed: execution.completed,
    partiallyCompleted: execution.partiallyCompleted,
    pending: execution.pending,
    managerMismatchOnly: execution.managerMismatchOnly,
    completionRate: execution.completionRate,
    latestPeopleDate: formatDate(findLatestPeopleDate(peopleById)),
    byRegion: buildDashboardBreakdown(
      visiblePeople,
      execution.setupRows,
      (person) => person.region,
      (row) => row.region,
    ),
    byLob: buildDashboardBreakdown(visiblePeople, execution.setupRows, (person) => person.lob, (row) => row.lob),
    byAnalyst: buildDashboardBreakdown(
      visiblePeople,
      execution.setupRows,
      (person) => person.analystName || "Unassigned",
      (row) => row.analystName || "Unassigned",
    ),
  };
}

function buildVerificationExpectations(
  rows: AuditRow[],
  data: AppData,
  countryToRegion: Record<string, string>,
  generatedAt: Date,
): VerificationExpectation[] {
  const expectations: VerificationExpectation[] = [];
  const inferredOwnershipIds = new Set(
    rows.filter((row) => isInferredOwnershipAuditItem(row.auditItem)).map((row) => row.employeeId),
  );
  const analystIndex = buildAnalystInferenceIndex(data, countryToRegion, inferredOwnershipIds);
  const generatedAtText = generatedAt.toISOString();
  const dueDateValue = new Date(generatedAt);
  dueDateValue.setDate(dueDateValue.getDate() + 7);
  const dueDate = formatDate(dueDateValue);

  for (const row of rows) {
    if (row.auditItem === "Unmapped Data Warning") continue;
    const current = data.currentScrById[row.employeeId];
    const previous = data.previousScrById[row.employeeId];
    const people = data.peopleById[row.employeeId];
    const analystAssignment = resolveAnalystAssignment(
      row.employeeId,
      data,
      countryToRegion,
      analystIndex,
      isInferredOwnershipAuditItem(row.auditItem),
    );
    const verificationId = `${row.processingMonth}|${row.auditItem}|${row.employeeId}`;
    const deferred = row.auditItem === "Deferred Change While on LOA" ? "Yes" : "No";
    const add = (
      fieldKey: string,
      fieldLabel: string,
      expectedValue: string,
      rule: VerificationRule,
      note = "",
    ): void => {
      expectations.push({
        verificationId,
        processingMonth: row.processingMonth,
        generatedAt: generatedAtText,
        dueDate,
        employeeId: row.employeeId,
        employeeName: row.employeeName,
        region: row.region,
        lob: row.lob,
        country: row.country,
        analystName: analystAssignment.name,
        inferredAnalystName: row.inferredAnalystName,
        analystReview: row.analystReview,
        inferenceBasis: row.inferenceBasis,
        analystSource: analystAssignment.source,
        analystConfidence: analystAssignment.confidence,
        analystSampleSize: analystAssignment.sampleSize,
        auditItem: row.auditItem,
        auditSubcategory: row.auditSubcategory,
        wcrEffectiveDate: row.wcrEffectiveDate,
        fieldKey,
        fieldLabel,
        baselineValue: getPeopleVerificationValue(people, fieldKey),
        expectedValue,
        rule,
        deferred,
        note,
      });
    };
    const addCoreEmployeeExpectations = (): void => {
      if (!current) return;
      add("record", "People Record", "Present", "exists");
      add("employeeStatus", "Employee Status", isYes(current.onLeave) ? "LOA" : "Active", "text");
      if (current.jobTitle) add("hrJobTitle", "HR Job Title", current.jobTitle, "text");
      if (current.supervisoryManager) add("level1Manager", "Level 1 Manager", current.supervisoryManager, "text");
      if (current.commissionAmount !== null) {
        add("annualVariable", "Annual Variable", String(current.commissionAmount), "number");
      }
      if (current.oteBaseComm !== null && current.commissionAmount !== null) {
        add("salary", "Salary", String(current.oteBaseComm - current.commissionAmount), "number");
      }
      if (current.country && current.country !== "Unmapped") add("country", "Country", current.country, "text");
      if (current.currency) add("salaryCurrency", "Salary Currency", current.currency, "text");
      if (current.businessUnit && current.businessUnit !== "Unmapped") {
        add(
          "businessUnit",
          "Business Unit",
          current.businessUnit,
          "unverifiable",
          "SCR Business Unit requires an approved mapping to Xactly Business_Unit.",
        );
      }
    };

    if (row.auditItem === "Variable Compensation Mismatch") {
      if (current?.commissionAmount !== null && current?.commissionAmount !== undefined) {
        add("annualVariable", "Annual Variable", String(current.commissionAmount), "number");
      }
      continue;
    }

    if (
      row.auditItem === "New Hire" ||
      row.auditItem === "Transfer to Sales" ||
      row.auditItem === "Transfer to Sales - Xactly Setup Required"
    ) {
      addCoreEmployeeExpectations();
      if (row.missingPositionSetup === "Yes") {
        add("positionRecord", "Position Record", "Present", "unverifiable", "Position cannot be verified from a People-only follow-up.");
      }
      continue;
    }

    if (row.auditItem === "Missing Xactly Setup") {
      if (row.missingPeopleSetup === "Yes") addCoreEmployeeExpectations();
      if (row.variableCompensationGap !== "" && current?.commissionAmount !== null && current?.commissionAmount !== undefined) {
        add("annualVariable", "Annual Variable", String(current.commissionAmount), "number");
      }
      if (row.missingPositionSetup === "Yes") {
        add("positionRecord", "Position Record", "Present", "unverifiable", "Position cannot be verified from a People-only follow-up.");
      }
      continue;
    }

    if (
      (row.auditItem === "Change to Existing Participant" ||
        row.auditItem === "Deferred Change While on LOA" ||
        row.auditItem === "LOA Return with Participant Changes") &&
      current &&
      previous
    ) {
      if (compareField("Job Title", previous.jobTitle, current.jobTitle).changed) {
        add("hrJobTitle", "HR Job Title", current.jobTitle, "text");
      }
      if (compareField("Job Level", previous.jobLevel, current.jobLevel).changed) {
        add("jobLevel", "Job Level", current.jobLevel, "unverifiable", "Job Level is not available in the People-only follow-up.");
      }
      if (compareField("Job Grade", previous.jobGrade, current.jobGrade).changed) {
        add("jobGrade", "Job Grade", current.jobGrade, "unverifiable", "Job Grade is not available in the People-only follow-up.");
      }
      if (compareField("Supervisory Manager", previous.supervisoryManager, current.supervisoryManager).changed) {
        add("level1Manager", "Level 1 Manager", current.supervisoryManager, "text");
      }
      if (compareField("OTE", previous.oteBaseComm, current.oteBaseComm).changed) {
        if (current.oteBaseComm !== null && current.commissionAmount !== null) {
          add("salary", "Salary", String(current.oteBaseComm - current.commissionAmount), "number");
        } else {
          add("salary", "Salary", "", "unverifiable", "Salary cannot be derived when OTE or Commission Amount is blank.");
        }
      }
      if (
        compareField("Commission Amount", previous.commissionAmount, current.commissionAmount).changed ||
        row.variableCompensationGap !== ""
      ) {
        if (current.commissionAmount !== null) add("annualVariable", "Annual Variable", String(current.commissionAmount), "number");
      }
      if (compareField("Business Unit", previous.businessUnit, current.businessUnit).changed) {
        add(
          "businessUnit",
          "Business Unit",
          current.businessUnit,
          "unverifiable",
          "SCR Business Unit requires an approved mapping to Xactly Business_Unit.",
        );
      }
      if (compareField("Country", previous.country, current.country).changed) add("country", "Country", current.country, "text");
      if (compareField("Currency", previous.currency, current.currency).changed) {
        add("salaryCurrency", "Salary Currency", current.currency, "text");
      }
      if (row.auditItem === "LOA Return with Participant Changes") {
        add("employeeStatus", "Employee Status", "Active", "text");
      }
      continue;
    }

    if (row.auditItem === "LOA Start") {
      add("employeeStatus", "Employee Status", "LOA", "text");
      if (row.variableCompensationGap !== "" && current?.commissionAmount !== null && current?.commissionAmount !== undefined) {
        add("annualVariable", "Annual Variable", String(current.commissionAmount), "number");
      }
      continue;
    }
    if (row.auditItem === "LOA Return") {
      add("employeeStatus", "Employee Status", "Active", "text");
      if (row.variableCompensationGap !== "" && current?.commissionAmount !== null && current?.commissionAmount !== undefined) {
        add("annualVariable", "Annual Variable", String(current.commissionAmount), "number");
      }
      continue;
    }
    if (row.auditItem === "Transfer to Non-Sales") {
      add("employeeStatus", "Employee Status", "Transfer Out|Non-Commissionable", "oneOf");
      continue;
    }
    if (row.auditItem === "Termination") {
      add("employeeStatus", "Employee Status", "Terminated", "text");
      const expectedTerminationDate = current?.terminationDate ?? previous?.terminationDate ?? null;
      if (expectedTerminationDate) add("terminationDate", "Termination Date", formatDate(expectedTerminationDate), "date");
      continue;
    }
    if (row.auditItem === "OKR Plan End") {
      add("okrAssignment", "OKR Assignment", row.okrEndMonth, "unverifiable", "OKR assignment cannot be verified from a People-only follow-up.");
    }
  }

  return expectations;
}

function getPeopleVerificationValue(people: PeopleRecord | undefined, fieldKey: string): string {
  if (fieldKey === "record") return people ? "Present" : "Missing";
  if (!people) return "";
  if (fieldKey === "employeeStatus") return people.employeeStatus;
  if (fieldKey === "terminationDate") return formatDate(people.terminationDate);
  if (fieldKey === "salary") return people.salary === null ? "" : String(people.salary);
  if (fieldKey === "salaryCurrency") return people.salaryCurrency;
  if (fieldKey === "annualVariable") return people.annualVariable === null ? "" : String(people.annualVariable);
  if (fieldKey === "hrJobTitle") return people.hrJobTitle;
  if (fieldKey === "level1Manager") return people.level1Manager;
  if (fieldKey === "businessUnit") return people.businessUnit;
  if (fieldKey === "country") return people.country;
  return "";
}

function matchesExpectation(expectation: VerificationExpectation, actualValue: string): boolean {
  if (expectation.rule === "unverifiable") return false;
  if (expectation.rule === "exists") return actualValue === "Present";
  if (expectation.rule === "number") {
    const expected = toNumber(expectation.expectedValue);
    const actual = toNumber(actualValue);
    return (
      expected !== null &&
      actual !== null &&
      (expectation.fieldKey === "annualVariable"
        ? Math.abs(expected - actual) < VARIABLE_COMPENSATION_MATERIALITY_THRESHOLD
        : numbersEqual(expected, actual))
    );
  }
  if (expectation.rule === "oneOf") {
    const actual = normalizeText(actualValue);
    return expectation.expectedValue.split("|").some((value) => normalizeText(value) === actual);
  }
  if (expectation.rule === "date") return expectation.expectedValue === actualValue;
  return normalizeText(expectation.expectedValue) === normalizeText(actualValue);
}

function parseIsoDate(value: string): Date | null {
  if (!value) return null;
  const parsed = new Date(`${value.slice(0, 10)}T00:00:00`);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function findLatestPeopleDate(peopleById: Record<string, PeopleRecord>): Date | null {
  let latest: Date | null = null;
  for (const people of Object.values(peopleById)) {
    if (people.uploadDate && (!latest || people.uploadDate > latest)) latest = people.uploadDate;
  }
  return latest;
}

type AnalystCountMap = Map<string, number>;

interface AnalystInferenceIndex {
  countryLob: Map<string, AnalystCountMap>;
  regionLob: Map<string, AnalystCountMap>;
}

interface AnalystAssignment {
  name: string;
  source: string;
  confidence: string;
  sampleSize: number;
}

type AnalystData = Pick<AppData, "peopleById" | "currentScrById" | "previousScrById">;

function hasAnalystRoutingChange(previous: ScrRecord, current: ScrRecord, people: PeopleRecord | undefined): boolean {
  return (
    compareField("Country", previous.country, current.country).changed ||
    normalizeText(deriveLob(previous, people)) !== normalizeText(deriveLob(current, people))
  );
}

function deriveExistingAnalystRecommendation(
  employeeId: string,
  changes: { label: string; changed: boolean }[],
  previous: ScrRecord,
  current: ScrRecord,
  data: AnalystData,
  countryToRegion: Record<string, string>,
  index: AnalystInferenceIndex,
): Pick<AuditRow, "inferredAnalystName" | "analystReview" | "inferenceBasis"> | Record<string, never> {
  const routingFieldChanged = hasChanged(changes, "Business Unit") || hasChanged(changes, "Country");
  if (!routingFieldChanged) return {};
  const people = data.peopleById[employeeId];
  if (!hasAnalystRoutingChange(previous, current, people)) {
    return {
      inferredAnalystName: "",
      analystReview: "No Routing Change",
      inferenceBasis: "Country + derived LOB unchanged",
    };
  }

  const inferred = resolveAnalystAssignment(employeeId, data, countryToRegion, index, true);
  if (inferred.name === "Unassigned") {
    return {
      inferredAnalystName: "Unassigned",
      analystReview: "Ambiguous / Unassigned",
      inferenceBasis: "No unique Country + LOB or Region + LOB match",
    };
  }
  return {
    inferredAnalystName: inferred.name,
    analystReview:
      normalizeText(people?.analystName) === normalizeText(inferred.name) ? "No Change Suggested" : "Change Suggested",
    inferenceBasis: `${inferred.source.replace(/^Inferred:\s*/, "")} | ${inferred.confidence} | n=${inferred.sampleSize}`,
  };
}

function buildAnalystInferenceIndex(
  data: Pick<AnalystData, "peopleById" | "currentScrById">,
  countryToRegion: Record<string, string>,
  excludedEmployeeIds: ReadonlySet<string> = new Set(),
): AnalystInferenceIndex {
  const index: AnalystInferenceIndex = { countryLob: new Map(), regionLob: new Map() };
  for (const current of Object.values(data.currentScrById)) {
    if (!isYes(current.activeStatus)) continue;
    if (excludedEmployeeIds.has(current.employeeId)) continue;
    const people = data.peopleById[current.employeeId];
    const analyst = people?.analystName.trim();
    const country = normalizeText(current.country);
    const lob = normalizeText(deriveLob(current, people));
    if (!analyst || !country || country === "unmapped" || !lob || lob === "unmapped") continue;
    addAnalystCount(index.countryLob, analystKey(country, lob), analyst);
    const region = normalizeRegionValue(countryToRegion[country] ?? people?.region ?? "");
    if (region) addAnalystCount(index.regionLob, analystKey(region, lob), analyst);
  }
  return index;
}

function resolveAnalystAssignment(
  employeeId: string,
  data: AnalystData,
  countryToRegion: Record<string, string>,
  index: AnalystInferenceIndex,
  ignorePeopleAnalyst = false,
): AnalystAssignment {
  const people = data.peopleById[employeeId];
  if (!ignorePeopleAnalyst && people?.analystName.trim()) {
    return { name: people.analystName.trim(), source: "People", confidence: "Confirmed", sampleSize: 0 };
  }

  const current = data.currentScrById[employeeId];
  const previous = data.previousScrById[employeeId];
  const scr = isYes(current?.activeStatus) ? current : isYes(previous?.activeStatus) ? previous : current ?? previous;
  const country = normalizeText(scr?.country ?? "");
  const lob = normalizeText(deriveLob(scr, people));
  if (!country || country === "unmapped" || !lob || lob === "unmapped") {
    return { name: "Unassigned", source: "Unassigned", confidence: "", sampleSize: 0 };
  }

  const countryMatch = chooseAnalyst(index.countryLob.get(analystKey(country, lob)), "Inferred: Country + LOB");
  if (countryMatch) return countryMatch;

  const region = normalizeRegionValue(countryToRegion[country] ?? people?.region ?? "");
  const regionMatch = region
    ? chooseAnalyst(index.regionLob.get(analystKey(region, lob)), "Inferred: Region + LOB")
    : null;
  return regionMatch ?? { name: "Unassigned", source: "Unassigned", confidence: "", sampleSize: 0 };
}

function addAnalystCount(index: Map<string, AnalystCountMap>, key: string, analyst: string): void {
  let counts = index.get(key);
  if (!counts) {
    counts = new Map();
    index.set(key, counts);
  }
  counts.set(analyst, (counts.get(analyst) ?? 0) + 1);
}

function chooseAnalyst(counts: AnalystCountMap | undefined, source: string): AnalystAssignment | null {
  if (!counts || counts.size === 0) return null;
  const ranked = [...counts.entries()].sort((left, right) => {
    const countCompare = right[1] - left[1];
    return countCompare !== 0 ? countCompare : left[0].localeCompare(right[0]);
  });
  const top = ranked[0];
  if (!top || (ranked[1]?.[1] ?? -1) === top[1]) return null;
  const sampleSize = ranked.reduce((total, [, count]) => total + count, 0);
  return {
    name: top[0],
    source,
    confidence: `${Math.round((top[1] / sampleSize) * 100)}%`,
    sampleSize,
  };
}

function analystKey(first: string, lob: string): string {
  return `${normalizeText(first)}|${normalizeText(lob)}`;
}

function isTransferToSalesAuditItem(auditItem: string): boolean {
  return auditItem === "Transfer to Sales" || auditItem === "Transfer to Sales - Xactly Setup Required";
}

function isInferredOwnershipAuditItem(auditItem: string): boolean {
  return auditItem === "New Hire" || isTransferToSalesAuditItem(auditItem);
}

function dashboardRegion(value: string): string {
  return normalizeRegionValue(value) || "Unmapped";
}

function deriveLob(scr: ScrRecord | undefined, people: PeopleRecord | undefined): string {
  if (scr && isYes(scr.activeStatus)) {
    const costCenter = normalizeText(scr.costCenter);
    const jobFamily = normalizeText(scr.jobFamily);
    const businessUnit = normalizeText(scr.businessUnit);
    if (costCenter.includes("gcp")) return "GCP";
    if (jobFamily.startsWith("sales development")) return "SD";
    if (businessUnit === "advertising sales" || businessUnit === "advertising operations") return "LMS";
    if (businessUnit === "lcs sales" || businessUnit === "lcs operations") return "LTS";
    if (businessUnit === "sales solutions" || businessUnit === "sales solutions operations") return "LSS";
    if (businessUnit === "global sales operations") return jobFamily === "salesq vp" ? "Global" : "SD";
  }

  const peopleLob = text(people?.businessUnit);
  const normalizedPeopleLob = normalizeText(peopleLob);
  if (normalizedPeopleLob === "ts") return "LTS";
  if (normalizedPeopleLob === "ms") return "LMS";
  return displayOrUnmapped(peopleLob);
}

function buildDashboardBreakdown<TPopulation>(
  people: TPopulation[],
  verificationRows: VerificationResultRow[],
  peopleLabel: (people: TPopulation) => string,
  verificationLabel: (row: VerificationResultRow) => string,
): DashboardBreakdownRow[] {
  const byLabel = new Map<string, DashboardBreakdownRow>();
  const getRow = (rawLabel: string): DashboardBreakdownRow => {
    const label = rawLabel || "Unmapped";
    const existing = byLabel.get(label);
    if (existing) return existing;
    const created = { label, employees: 0, required: 0, completed: 0, pending: 0, overdue: 0, inferred: 0 };
    byLabel.set(label, created);
    return created;
  };

  for (const person of people) getRow(peopleLabel(person)).employees += 1;
  for (const verification of verificationRows) {
    const row = getRow(verificationLabel(verification));
    row.required += 1;
    if (verification.progressStatus === "Completed") row.completed += 1;
    if (verification.progressStatus === "Pending" || verification.progressStatus === "Partially Completed") row.pending += 1;
    if (verification.slaStatus === "Overdue") row.overdue += 1;
    if (verification.analystSource.startsWith("Inferred:")) row.inferred += 1;
  }

  return [...byLabel.values()].sort((left, right) => {
    const employeeCompare = right.employees - left.employees;
    if (employeeCompare !== 0) return employeeCompare;
    const requiredCompare = right.required - left.required;
    return requiredCompare !== 0 ? requiredCompare : left.label.localeCompare(right.label);
  });
}

function countVerificationStatuses(rows: VerificationResultRow[]): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const row of rows) {
    counts[row.progressStatus] = (counts[row.progressStatus] ?? 0) + 1;
    if (row.slaStatus === "Overdue") counts.Overdue = (counts.Overdue ?? 0) + 1;
  }
  return counts;
}

function applyHeaderStyle(sheet: WorkSheet): void {
  const range = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : null;
  if (!range) return;
  for (let column = range.s.c; column <= range.e.c; column += 1) {
    const address = XLSX.utils.encode_cell({ c: column, r: 0 });
    const cellRef = sheet[address];
    if (!cellRef) continue;
    cellRef.s = {
      font: { bold: true, color: { rgb: "F7F7F2" } },
      fill: { fgColor: { rgb: "16324F" } },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
    };
  }
}

function appendColumnGuide(wb: WorkBook, descriptions: Readonly<Record<string, string>>): void {
  const rows = Object.entries(descriptions).map(([Column, Description]) => ({ Column, Description }));
  const sheet = XLSX.utils.json_to_sheet(rows);
  applyHeaderStyle(sheet);
  sheet["!cols"] = [{ wch: 32 }, { wch: 80 }];
  for (let row = 2; row <= rows.length + 1; row += 1) {
    const descriptionCell = sheet[`B${row}`];
    if (descriptionCell) descriptionCell.s = { alignment: { vertical: "top", wrapText: true } };
  }
  XLSX.utils.book_append_sheet(wb, sheet, "Column Guide");
}

function buildColumnWidths(rows: object[]): ColInfo[] {
  if (rows.length === 0) return [];
  const records = rows as Record<string, unknown>[];
  const keys = Object.keys(records[0]);
  return keys.map((key) => {
    const maxValueLength = Math.max(key.length, ...records.map((row) => String(row[key] ?? "").length));
    return { wch: Math.min(Math.max(maxValueLength + 2, 14), 36) };
  });
}

function countByAuditItem(rows: AuditRow[]): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const row of rows) counts[row.auditItem] = (counts[row.auditItem] ?? 0) + 1;
  return counts;
}

function countByAuditSubcategory(rows: AuditRow[]): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const row of rows) {
    if (row.auditSubcategory) counts[row.auditSubcategory] = (counts[row.auditSubcategory] ?? 0) + 1;
  }
  return counts;
}

function countByVerificationSubcategory(rows: VerificationResultRow[]): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const row of rows) {
    if (row.auditSubcategory) counts[row.auditSubcategory] = (counts[row.auditSubcategory] ?? 0) + 1;
  }
  return counts;
}

function createAuditRow(
  auditItem: string,
  processingMonth: string,
  employeeId: string,
  employeeName: string,
  context: ResolvedContext,
  values: Partial<AuditRow> & { notes?: string },
): AuditRow {
  const { changeSummary, notes, ...restValues } = values;
  const mergedChangeSummary = combineSummaryParts(changeSummary ?? "", notes ?? "");
  const displayedChangeSummary = isYes(context.currentOnLeave)
    ? prefixChangeSummary(CURRENTLY_ON_LOA_PREFIX, mergedChangeSummary)
    : mergedChangeSummary;
  return {
    auditItem,
    auditSubcategory: "",
    processingMonth,
    employeeId,
    employeeName,
    region: context.region,
    lob: context.lob,
    country: context.country,
    currentActiveStatus: context.currentActiveStatus,
    currentOnLeave: context.currentOnLeave,
    currentFirstDayOfLeave: context.currentFirstDayOfLeave,
    changeSummary: displayedChangeSummary,
    wcrEffectiveDate: "",
    peoplePlanEffectiveDate: context.peoplePlanEffectiveDate,
    peopleBusinessUnit: context.peopleBusinessUnit,
    analystName: context.analystName,
    inferredAnalystName: "",
    analystReview: "",
    inferenceBasis: "",
    planType: context.planType,
    hireDate: "",
    terminationDate: "",
    rehireInPeople: "",
    negativeBalance: "",
    missingPeopleSetup: "",
    missingPositionSetup: "",
    previousJobTitle: "",
    currentJobTitle: "",
    previousJobLevelGrade: "",
    currentJobLevelGrade: "",
    previousSupervisoryManager: "",
    currentSupervisoryManager: "",
    previousCommissionAmount: "",
    currentCommissionAmount: "",
    peopleAnnualVariable: "",
    variableCompensationGap: "",
    previousBusinessUnit: "",
    currentBusinessUnit: "",
    previousCountry: "",
    currentCountry: "",
    previousCurrency: "",
    currentCurrency: "",
    loaFirstDayOfLeave: "",
    loaEstimatedLastDay: "",
    loaTotalDays: "",
    okrStartMonth: "",
    okrEndMonth: "",
    transferDirection: "",
    microsoftTransfer: "",
    peopleUploadDate: context.peopleUploadDate,
    ...restValues,
  };
}

function appendVariableCompensationMismatches(
  rows: AuditRow[],
  processingMonth: string,
  filters: Filters,
  data: AppData,
  countryToRegion: Record<string, string>,
): void {
  const nonActionItems = new Set(["OKR Plan End", "Unmapped Data Warning"]);
  const existingChangeItems = new Set([
    "Change to Existing Participant",
    "Deferred Change While on LOA",
    "LOA Return with Participant Changes",
  ]);

  for (const current of Object.values(data.currentScrById)) {
    if (!isYes(current.activeStatus)) continue;
    const people = data.peopleById[current.employeeId];
    const gap = materialVariableCompensationGap(current.commissionAmount, people?.annualVariable ?? null);
    if (gap === null || !people) continue;
    const context = resolveEmployeeContext(current.employeeId, data, countryToRegion);
    if (!matchesFilters(context, filters)) continue;

    const summary = `SCR Commission Amount ${formatNumber(current.commissionAmount)} vs People Annual Variable ${formatNumber(people.annualVariable)} (Gap ${gap >= 0 ? "+" : ""}${formatNumber(gap)})`;
    const candidates = rows.filter(
      (row) => row.employeeId === current.employeeId && !nonActionItems.has(row.auditItem),
    );
    const existing = candidates.find((row) => existingChangeItems.has(row.auditItem)) ?? candidates[0];
    if (existing) {
      existing.currentCommissionAmount = numericCell(current.commissionAmount);
      existing.peopleAnnualVariable = numericCell(people.annualVariable);
      existing.variableCompensationGap = gap;
      existing.changeSummary = combineSummaryParts(existing.changeSummary, summary);

      const previous = data.previousScrById[current.employeeId];
      const monthlyVariableChanged = Boolean(
        previous && compareField("Commission Amount", previous.commissionAmount, current.commissionAmount).changed,
      );
      if (existingChangeItems.has(existing.auditItem) && !monthlyVariableChanged) {
        const baseSubcategory = existing.auditSubcategory.replace(/ Only$/, "");
        existing.auditSubcategory = baseSubcategory
          ? `${baseSubcategory} + Variable Mismatch`
          : "Variable Mismatch Only";
      }
      continue;
    }

    rows.push(
      createAuditRow(
        "Variable Compensation Mismatch",
        processingMonth,
        current.employeeId,
        current.fullName || context.name,
        context,
        {
          auditSubcategory: "Variable Mismatch Only",
          currentCommissionAmount: numericCell(current.commissionAmount),
          peopleAnnualVariable: numericCell(people.annualVariable),
          variableCompensationGap: gap,
          changeSummary: summary,
        },
      ),
    );
  }
}

function materialVariableCompensationGap(currentCommissionAmount: number | null, peopleAnnualVariable: number | null): number | null {
  if (currentCommissionAmount === null || peopleAnnualVariable === null) return null;
  const gap = currentCommissionAmount - peopleAnnualVariable;
  return Math.abs(gap) >= VARIABLE_COMPENSATION_MATERIALITY_THRESHOLD ? gap : null;
}

function wcrValuesChanged(row: unknown[], currentColumn: number, proposedColumn: number): boolean {
  const current = normalizeText(text(cell(row, currentColumn)));
  const proposed = normalizeText(text(cell(row, proposedColumn)));
  return current !== proposed && Boolean(current || proposed);
}

function resolveWcrEffectiveDate(row: AuditRow, records: WorkerChangeRecord[]): string {
  const groups = new Map<string, { signals: Set<WorkerChangeSignal>; processText: string }>();
  for (const record of records) {
    const effectiveDate = formatDate(record.effectiveDate);
    if (!effectiveDate) continue;
    const group = groups.get(effectiveDate) ?? { signals: new Set<WorkerChangeSignal>(), processText: "" };
    record.signals.forEach((signal) => group.signals.add(signal));
    group.processText += " " + record.businessProcessType.toLowerCase() + " " + record.businessProcessReason.toLowerCase();
    groups.set(effectiveDate, group);
  }
  if (groups.size === 0) return "";

  const auditSignals = new Set<WorkerChangeSignal>();
  if (
    normalizeText(row.previousJobTitle) !== normalizeText(row.currentJobTitle) ||
    normalizeText(row.previousJobLevelGrade) !== normalizeText(row.currentJobLevelGrade)
  ) {
    auditSignals.add("job");
  }
  if (normalizeText(row.previousSupervisoryManager) !== normalizeText(row.currentSupervisoryManager)) {
    auditSignals.add("manager");
  }
  if (row.previousCommissionAmount !== "" || row.currentCommissionAmount !== "") auditSignals.add("commission");
  if (row.previousBusinessUnit || row.currentBusinessUnit) auditSignals.add("businessUnit");
  if (row.previousCountry || row.currentCountry) auditSignals.add("country");
  if (row.previousCurrency || row.currentCurrency) auditSignals.add("currency");
  if (row.changeSummary.includes("OTE (Base+Comm)")) auditSignals.add("ote");

  const existingChangeItems = new Set([
    "Change to Existing Participant",
    "Deferred Change While on LOA",
    "LOA Return with Participant Changes",
  ]);
  const eventPattern =
    row.auditItem === "New Hire"
      ? /\bhire\b|\brehire\b/
      : row.auditItem.startsWith("Transfer to Sales") || row.auditItem === "Transfer to Non-Sales"
        ? /\btransfer\b/
        : row.auditItem === "Termination"
          ? /\bterminat/
          : row.auditItem === "LOA Start" || row.auditItem === "LOA Return"
            ? /\bleave\b|\babsence\b/
            : null;

  const matches = [...groups.entries()]
    .filter(([, group]) =>
      existingChangeItems.has(row.auditItem)
        ? [...auditSignals].some((signal) => group.signals.has(signal))
        : Boolean(eventPattern?.test(group.processText)),
    )
    .map(([date]) => date)
    .sort();
  if (matches.length > 0) return matches.join("; ");

  const supportsSingleDateFallback = existingChangeItems.has(row.auditItem) || eventPattern !== null;
  return supportsSingleDateFallback && groups.size === 1 ? [...groups.keys()][0] ?? "" : "";
}

function formatJobLevelGrade(scr: ScrRecord | undefined): string {
  const level = text(scr?.jobLevel);
  const grade = text(scr?.jobGrade);
  return level && grade ? `${level} (${grade})` : level || grade;
}

interface ResolvedContext {
  name: string;
  region: string;
  lob: string;
  country: string;
  peoplePlanEffectiveDate: string;
  peopleBusinessUnit: string;
  analystName: string;
  planType: string;
  peopleUploadDate: string;
  currentActiveStatus: string;
  currentOnLeave: string;
  currentFirstDayOfLeave: string;
}

function resolveEmployeeContext(employeeId: string, data: AppData, countryToRegion: Record<string, string>): ResolvedContext {
  const people = data.peopleById[employeeId];
  const current = data.currentScrById[employeeId];
  const previous = data.previousScrById[employeeId];
  const loa = data.loaById[employeeId];
  const msft = data.msftTransferById[employeeId];

  const country = displayOrUnmapped(current?.country || previous?.country || people?.country || "");
  const scrCountry = current?.country || previous?.country || "";
  const mappedScrRegion = scrCountry ? normalizeRegionValue(countryToRegion[scrCountry.toLowerCase()] ?? "") : "";
  const peopleRegion = normalizeRegionValue(people?.region ?? "");
  const region = displayOrUnmapped(mappedScrRegion || peopleRegion || normalizeRegionValue(loa?.region ?? "") || normalizeRegionValue(msft?.region ?? ""));
  const lobScr = isYes(current?.activeStatus) ? current : isYes(previous?.activeStatus) ? previous : undefined;
  const lob = deriveLob(lobScr, people);

  return {
    name: current?.fullName || previous?.fullName || people?.fullName || employeeId,
    region,
    lob,
    country,
    peoplePlanEffectiveDate: formatDate(people?.planEffectiveDate ?? null),
    peopleBusinessUnit: people?.businessUnit ?? "",
    analystName: people?.analystName ?? "",
    planType: people?.planType ?? "",
    peopleUploadDate: formatDate(people?.uploadDate ?? null),
    currentActiveStatus: current?.activeStatus ?? "",
    currentOnLeave: current?.onLeave ?? "",
    currentFirstDayOfLeave: formatDate(current?.firstDayOfLeave ?? null),
  };
}

function matchesFilters(context: ResolvedContext, filters: Filters): boolean {
  return filters.regions.includes(context.region) && filters.lobs.includes(context.lob) && filters.countries.includes(context.country);
}

function collectUnmappedWarningCandidateIds(
  data: AppData,
  okrByEmployee: Record<string, { startMonth: string; endMonth: string }>,
  previousMonthKey: string,
): string[] {
  const ids = new Set<string>();
  for (const record of Object.values(data.currentScrById)) {
    if (isYes(record.activeStatus)) ids.add(record.employeeId);
  }
  for (const record of Object.values(data.previousScrById)) {
    if (isYes(record.activeStatus)) ids.add(record.employeeId);
  }
  for (const [employeeId, okrSummary] of Object.entries(okrByEmployee)) {
    if (okrSummary.endMonth === previousMonthKey) ids.add(employeeId);
  }
  return [...ids].sort((left, right) => left.localeCompare(right));
}

function getUnmappedDataIssues(
  employeeId: string,
  data: AppData,
  countryToRegion: Record<string, string>,
  context: ResolvedContext,
): string[] {
  const current = data.currentScrById[employeeId];
  const previous = data.previousScrById[employeeId];
  const scrCountry = current?.country || previous?.country || "";
  const issues: string[] = [];
  if (isUnmappedValue(context.region)) issues.push("Region");
  if (isUnmappedValue(context.lob)) issues.push("LOB");
  if (isUnmappedValue(context.country)) issues.push("Country");
  if (scrCountry && !isUnmappedValue(scrCountry) && !countryToRegion[scrCountry.toLowerCase()]) {
    issues.push(`SCR country not in region map (${scrCountry})`);
  }
  return dedupeStrings(issues);
}

function isUnmappedValue(value: string): boolean {
  return value.trim().toLowerCase() === "unmapped";
}

function aggregateOkrAssignments(rows: QuotaAssignmentRow[]): Record<string, { startMonth: string; endMonth: string }> {
  const byId: Record<string, { startMonth: string; endMonth: string }> = {};
  for (const row of rows) {
    const nonZeroMonths = Object.entries(row.monthValues)
      .filter(([, value]) => Math.abs(value) > 0.0000001)
      .map(([label]) => label)
      .sort(compareMonthKeys);
    if (nonZeroMonths.length === 0) continue;
    const startMonth = nonZeroMonths[0];
    const endMonth = nonZeroMonths[nonZeroMonths.length - 1];
    if (!byId[row.employeeId]) {
      byId[row.employeeId] = { startMonth, endMonth };
      continue;
    }
    if (compareMonthKeys(startMonth, byId[row.employeeId].startMonth) < 0) byId[row.employeeId].startMonth = startMonth;
    if (compareMonthKeys(endMonth, byId[row.employeeId].endMonth) > 0) byId[row.employeeId].endMonth = endMonth;
  }
  return byId;
}

function compareField(label: string, previous: unknown, current: unknown): { label: string; changed: boolean } {
  if (typeof previous === "number" || typeof current === "number") {
    return { label, changed: !numbersEqual(toNumber(previous), toNumber(current)) };
  }
  return { label, changed: normalizeText(previous) !== normalizeText(current) };
}

function deriveExistingParticipantSubcategory(
  changes: { label: string; changed: boolean }[],
  previous: ScrRecord,
  current: ScrRecord,
): string {
  const movement = classifyCareerMovement(previous, current);
  if (!movement) return deriveAuditSubcategory(changes.map((item) => item.label));
  return hasChanged(changes, "Commission Amount")
    ? `${movement} + Variable Change`
    : `${movement} - No Variable Change`;
}

function classifyCareerMovement(previous: ScrRecord, current: ScrRecord): "Promotion" | "Job Change" | "" {
  const careerChanged =
    compareField("Job Title", previous.jobTitle, current.jobTitle).changed ||
    compareField("Job Level", previous.jobLevel, current.jobLevel).changed ||
    compareField("Job Grade", previous.jobGrade, current.jobGrade).changed;
  if (!careerChanged) return "";

  const previousGrade = normalizeJobGrade(previous.jobGrade);
  const currentGrade = normalizeJobGrade(current.jobGrade);
  const previousRank = JOB_GRADE_ORDER.indexOf(previousGrade);
  const currentRank = JOB_GRADE_ORDER.indexOf(currentGrade);
  if (previousRank >= 0 && currentRank > previousRank) return "Promotion";
  if (
    previousGrade &&
    previousGrade === currentGrade &&
    /^ic/i.test(text(previous.jobLevel)) &&
    /^mr/i.test(text(current.jobLevel))
  ) {
    return "Promotion";
  }
  return "Job Change";
}

function normalizeJobGrade(value: unknown): string {
  return normalizeText(value).replace(/^0+(?=\d)/, "");
}

export function deriveAuditSubcategory(changeLabels: readonly string[]): string {
  const labels = [...new Set(changeLabels.filter(Boolean))];
  if (labels.length === 0) return "";
  if (labels.length > 1) {
    return labels.includes("Commission Amount") ? "Variable + Other Changes" : "Multiple Changes - No Variable";
  }
  return (
    {
      "Job Title": "Job Title Change Only",
      "Job Level": "Job Change - No Variable Change",
      "Job Grade": "Job Change - No Variable Change",
      "Supervisory Manager": "Manager Change Only",
      "OTE (Base+Comm)": "OTE Change Only",
      "Commission Amount": "Variable Change Only",
      "Business Unit": "Business Unit Change Only",
      Country: "Country Change Only",
      Currency: "Currency Change Only",
    }[labels[0] ?? ""] ?? "Other Change Only"
  );
}

function deriveVerificationSubcategory(results: VerificationFieldResult[]): string {
  const auditItem = results[0]?.auditItem ?? "";
  if (
    auditItem !== "Change to Existing Participant" &&
    auditItem !== "Deferred Change While on LOA" &&
    auditItem !== "LOA Return with Participant Changes"
  ) {
    return "";
  }
  const changeLabelByFieldKey: Record<string, string> = {
    hrJobTitle: "Job Title",
    jobLevel: "Job Level",
    jobGrade: "Job Grade",
    level1Manager: "Supervisory Manager",
    salary: "OTE (Base+Comm)",
    annualVariable: "Commission Amount",
    businessUnit: "Business Unit",
    country: "Country",
    salaryCurrency: "Currency",
  };
  return deriveAuditSubcategory(results.map((result) => changeLabelByFieldKey[result.fieldKey]).filter(Boolean));
}

function isNewHireInProcessingWindow(record: ScrRecord, startDate: Date, endDate: Date): boolean {
  return Boolean(record.hireDate && record.hireDate >= startDate && record.hireDate <= endDate);
}

function isRehireNewHire(record: ScrRecord): boolean {
  return isYes(record.isRehire) && Boolean(record.originalHireDate && record.hireDate && record.originalHireDate < record.hireDate);
}

function hasChanged(changes: { label: string; changed: boolean }[], label: string): boolean {
  return changes.some((item) => item.label === label && item.changed);
}

function buildMissingXactlySetupSummary(missingPeople: boolean, missingPosition: boolean): string {
  const missing = [
    missingPeople ? "People" : "",
    missingPosition ? "Position" : "",
  ].filter(Boolean);
  return `Employee is active in the current month SCR but missing from ${missing.join(" and ")}.`;
}

function buildNewHireSummary(
  hasPeopleRecord: boolean,
  hasNegativeBalance: boolean,
  missingPeople: boolean,
  missingPosition: boolean,
): string {
  const parts: string[] = [];
  if (hasPeopleRecord) parts.push("Employee exists in People.");
  if (hasNegativeBalance) parts.push("Employee has a negative payment balance.");
  if (missingPeople || missingPosition) {
    const missing = [
      missingPeople ? "People" : "",
      missingPosition ? "Position" : "",
    ].filter(Boolean);
    parts.push(`New hire pending Xactly setup: missing ${missing.join(" and ")}.`);
  }
  return combineSummaryParts(...parts);
}

function formatMaterialNegativeBalance(summary: BalanceSummary): string {
  return Object.entries(summary.negativeTotalByCurrency)
    .filter(([, amount]) => amount < 0 && Math.abs(amount) >= NEGATIVE_BALANCE_MATERIALITY_THRESHOLD)
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([currency, amount]) => `${currency} ${formatNumber(amount)}`)
    .join("; ");
}

function numericCell(value: number | null): number | "" {
  return value ?? "";
}

function combineSummaryParts(...parts: string[]): string {
  const cleaned = parts.map((part) => part.trim()).filter(Boolean);
  return [...new Set(cleaned)].join(" | ");
}

function prefixChangeSummary(prefix: string, summary: string): string {
  const trimmed = summary.trim();
  if (!trimmed) return prefix;
  if (trimmed.startsWith(prefix)) return trimmed;
  return `${prefix} ${trimmed}`;
}

function pickLatestPeopleRecord(previous: PeopleRecord | undefined, candidate: PeopleRecord): PeopleRecord {
  if (!previous) return candidate;
  return scorePeopleRecord(candidate) >= scorePeopleRecord(previous) ? candidate : previous;
}

function scorePeopleRecord(record: PeopleRecord): number {
  return Math.max(dateScore(record.planEffectiveDate), dateScore(record.effectiveStartDate), dateScore(record.uploadDate));
}

function pickLatestPositionRecord(previous: PositionRecord | undefined, candidate: PositionRecord): PositionRecord {
  if (!previous) return candidate;
  return dateScore(candidate.effectiveStartDate) >= dateScore(previous.effectiveStartDate) ? candidate : previous;
}

function pickPreferredLoaRecord(previous: LoaRecord | undefined, candidate: LoaRecord): LoaRecord {
  if (!previous) return candidate;
  const previousCompleted = dateScore(previous.dateTimeCompleted);
  const candidateCompleted = dateScore(candidate.dateTimeCompleted);
  if (candidateCompleted !== previousCompleted) return candidateCompleted > previousCompleted ? candidate : previous;
  const previousCorrection = dateScore(previous.latestCorrection);
  const candidateCorrection = dateScore(candidate.latestCorrection);
  if (candidateCorrection !== previousCorrection) return candidateCorrection > previousCorrection ? candidate : previous;
  return dateScore(candidate.estimatedLastDayOfLeave) >= dateScore(previous.estimatedLastDayOfLeave) ? candidate : previous;
}

function sortDisplayValues(values: Set<string>): string[] {
  return [...values].filter(Boolean).sort((left, right) => left.localeCompare(right));
}

function collectScrLobOptions(data: AppData): string[] {
  const lobSet = new Set<string>();
  for (const record of Object.values(data.currentScrById)) {
    if (!isYes(record.activeStatus)) continue;
    lobSet.add(deriveLob(record, data.peopleById[record.employeeId]));
  }
  for (const record of Object.values(data.previousScrById)) {
    if (!isYes(record.activeStatus)) continue;
    lobSet.add(deriveLob(record, data.peopleById[record.employeeId]));
  }
  lobSet.delete("Unmapped");
  return sortDisplayValues(lobSet);
}

function intersectKeys<T>(left: Record<string, T>, right: Record<string, T>): string[] {
  return Object.keys(left).filter((key) => key in right);
}

function dedupeStrings(values: string[]): string[] {
  return [...new Set(values)];
}

function compareMonthKeys(left: string, right: string): number {
  const leftDate = parseMonthKey(left);
  const rightDate = parseMonthKey(right);
  if (!leftDate || !rightDate) return left.localeCompare(right);
  return leftDate.getTime() - rightDate.getTime();
}

function formatMonthKey(value: Date): string {
  return `${MONTH_NAMES[value.getMonth()]}-${value.getFullYear()}`;
}

function parseMonthKey(value: string): Date | null {
  const match = /^([A-Z]{3})-(\d{4})$/i.exec(value.trim());
  if (!match) return null;
  const monthIndex = MONTH_NAMES.indexOf(match[1].toUpperCase());
  if (monthIndex < 0) return null;
  return new Date(Number(match[2]), monthIndex, 1);
}

function extractMonthColumns(headerRow: string[]): Record<string, number> {
  const columns: Record<string, number> = {};
  for (let index = 0; index < headerRow.length; index += 1) {
    const rawHeader = text(headerRow[index]);
    const match = /^\s*(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)-(\d{4})\s*\*?\s*$/i.exec(rawHeader);
    if (match) columns[`${match[1].toUpperCase()}-${match[2]}`] = index;
  }
  return columns;
}

async function readMatrixFromFile(file: File, sheetIndex: number): Promise<unknown[][]> {
  const workbook = normalizeWorkbookRanges(XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false, raw: true }));
  const sheetName = workbook.SheetNames[sheetIndex];
  if (!sheetName) throw new Error("No worksheet found in the uploaded file.");
  const sheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null }) as unknown[][];
}

function normalizeWorkbookRanges(workbook: WorkBook): WorkBook {
  for (const sheetName of workbook.SheetNames) {
    normalizeWorksheetRange(workbook.Sheets[sheetName]);
  }
  return workbook;
}

function normalizeWorksheetRange(sheet: WorkSheet | undefined): void {
  if (!sheet) return;

  const addresses = Object.keys(sheet).filter((key) => !key.startsWith("!"));
  if (addresses.length === 0) return;

  let minRow = Number.POSITIVE_INFINITY;
  let minCol = Number.POSITIVE_INFINITY;
  let maxRow = 0;
  let maxCol = 0;

  for (const address of addresses) {
    const decoded = XLSX.utils.decode_cell(address);
    if (decoded.r < minRow) minRow = decoded.r;
    if (decoded.c < minCol) minCol = decoded.c;
    if (decoded.r > maxRow) maxRow = decoded.r;
    if (decoded.c > maxCol) maxCol = decoded.c;
  }

  const actualRange = XLSX.utils.encode_range({
    s: { r: Number.isFinite(minRow) ? minRow : 0, c: Number.isFinite(minCol) ? minCol : 0 },
    e: { r: maxRow, c: maxCol },
  });
  const currentRange = sheet["!ref"];
  if (typeof currentRange !== "string") {
    sheet["!ref"] = actualRange;
    return;
  }

  try {
    const existing = XLSX.utils.decode_range(currentRange);
    const actual = XLSX.utils.decode_range(actualRange);
    const needsExpansion =
      actual.s.r < existing.s.r || actual.s.c < existing.s.c || actual.e.r > existing.e.r || actual.e.c > existing.e.c;
    if (needsExpansion) sheet["!ref"] = actualRange;
  } catch {
    sheet["!ref"] = actualRange;
  }
}

function findHeaderRow(matrix: unknown[][], requiredTokens: string[]): number {
  const maxRows = Math.min(matrix.length, 12);
  for (let rowIndex = 0; rowIndex < maxRows; rowIndex += 1) {
    const joined = (matrix[rowIndex] ?? []).map(normalizeHeader).join("|");
    if (requiredTokens.every((token) => joined.includes(token))) return rowIndex;
  }
  return -1;
}

function findColumn(header: string[], candidates: string[]): number {
  return header.findIndex((name) => candidates.some((candidate) => name.includes(candidate)));
}

function findExactColumn(header: string[], exactCandidates: string[], fallbackCandidates: string[]): number {
  const exactIndex = header.findIndex((name) => exactCandidates.includes(name));
  return exactIndex >= 0 ? exactIndex : findColumn(header, fallbackCandidates);
}

function normalizeHeader(value: unknown): string {
  return String(value ?? "")
    .toLowerCase()
    .replace(/[#*()\-\/.%_]/g, "")
    .replace(/\s+/g, "");
}

function text(value: unknown): string {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}

function normalizeText(value: unknown): string {
  return text(value).toLowerCase().replace(/\s+/g, " ");
}

function cell(row: unknown[], index: number): unknown {
  if (index < 0 || index >= row.length) return null;
  return row[index];
}

function normalizeEmployeeIdFromCell(value: unknown): string | null {
  if (typeof value === "number" && Number.isFinite(value)) {
    return String(Math.trunc(value)).padStart(6, "0");
  }
  const normalized = text(value);
  if (!/^\d+$/.test(normalized)) return null;
  return normalized.padStart(6, "0");
}

function normalizeEmployeeIdFromText(value: unknown): string | null {
  const match = /\b(\d{1,6})\b/.exec(text(value));
  return match ? match[1].padStart(6, "0") : null;
}

function toDate(value: unknown): Date | null {
  if (value === null || value === undefined || value === "") return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, parsed.S);
  }
  const parsed = new Date(String(value));
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function toNumber(value: unknown): number | null {
  if (value === null || value === undefined || value === "") return null;
  if (typeof value === "number") return Number.isFinite(value) ? value : null;
  const normalized = String(value).replace(/,/g, "").trim();
  if (!normalized) return null;
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : null;
}

function normalizeRegionValue(value: string): string {
  const upper = value.trim().toUpperCase();
  if (!upper) return "";
  if (upper === "CHINA") return "APAC";
  if (upper.includes("NAMER")) return "NAMER";
  if (upper.includes("LATAM")) return "LATAM";
  if (upper.includes("EMEA")) return "EMEA";
  if (upper.includes("APAC")) return "APAC";
  return "";
}

function dateScore(value: Date | null | undefined): number {
  return value ? value.getTime() : -1;
}

function formatDate(value: Date | null): string {
  if (!value) return "";
  const yyyy = value.getFullYear();
  const mm = String(value.getMonth() + 1).padStart(2, "0");
  const dd = String(value.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

function formatNumber(value: number | null): string {
  if (value === null || value === undefined) return "";
  return value.toLocaleString("en-US", { maximumFractionDigits: 2 });
}

function numbersEqual(left: number | null, right: number | null): boolean {
  if (left === null && right === null) return true;
  if (left === null || right === null) return false;
  return Math.abs(left - right) < 0.000001;
}

function isYes(value: string): boolean {
  return normalizeText(value) === "yes";
}

function displayOrUnmapped(value: string): string {
  return value.trim() || "Unmapped";
}
