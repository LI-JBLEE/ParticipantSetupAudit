import XLSXImport from "xlsx-js-style";
import type { ColInfo, WorkBook, WorkSheet } from "xlsx-js-style";
import type {
  AppData,
  AuditBuildResult,
  AuditRow,
  BalanceRow,
  BalanceSummary,
  FileParseResult,
  FilterOptions,
  Filters,
  LoaRecord,
  MsftTransferRecord,
  PeopleRecord,
  PositionRecord,
  ProcessingMonthOption,
  QuotaAssignmentRow,
  ScrRecord,
  UploadDefinition,
} from "./types";

const XLSX = ((XLSXImport as unknown as { default?: typeof XLSXImport }).default ?? XLSXImport) as typeof XLSXImport;

const MONTH_NAMES = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"];
const REGION_OPTIONS = ["APAC", "EMEA", "LATAM", "NAMER"];

const REQUIRED_UPLOADS: UploadDefinition[] = [
  { key: "currentScr", label: "Sales Compensation Report (Current Month)", accept: ".xlsx,.xls" },
  { key: "previousScr", label: "Sales Compensation Report (Previous Month)", accept: ".xlsx,.xls" },
  { key: "people", label: "People", accept: ".xlsx,.xls" },
  { key: "position", label: "Position", accept: ".xlsx,.xls" },
  { key: "quota", label: "Quota Assignment", accept: ".xlsx,.xls" },
  { key: "balance", label: "Payment Balance", accept: ".xls,.xlsx" },
  { key: "loa", label: "LOA Report", accept: ".xlsx,.xls" },
  { key: "msftTransfer", label: "Transfer to MSFT", accept: ".xlsx,.xls" },
];

export function getRequiredUploads(): UploadDefinition[] {
  return REQUIRED_UPLOADS;
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
  const buffer = await fetchDefaultFile("Country Region Mapping.xlsx");
  return parseCountryRegionWorkbook(normalizeWorkbookRanges(XLSX.read(buffer, { type: "array", cellDates: false })));
}

export function buildFilterOptions(data: AppData, countryToRegion: Record<string, string>): FilterOptions {
  const regionSet = new Set<string>();
  const lobSet = new Set<string>();
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
    lobSet.add(context.lob);
  }

  for (const record of Object.values(data.currentScrById)) {
    if (record.country && record.country !== "Unmapped") {
      countrySet.add(record.country);
    }
  }

  return {
    regions: REGION_OPTIONS,
    lobs: sortDisplayValues(lobSet),
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
    hireDate: findExactColumn(header, ["hiredate"], ["hiredate"]),
    isRehire: findExactColumn(header, ["isrehire"], ["isrehire"]),
    terminationDate: findColumn(header, ["terminationdate"]),
    jobTitle: findColumn(header, ["jobtitle"]),
    supervisoryManager: findColumn(header, ["supervisorymanager"]),
    oteBaseComm: findColumn(header, ["otebasecomm"]),
    commissionAmount: findColumn(header, ["commissionamount"]),
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
      hireDate: toDate(cell(row, cols.hireDate)),
      isRehire: text(cell(row, cols.isRehire)),
      terminationDate: toDate(cell(row, cols.terminationDate)),
      jobTitle: text(cell(row, cols.jobTitle)),
      supervisoryManager: text(cell(row, cols.supervisoryManager)),
      oteBaseComm: toNumber(cell(row, cols.oteBaseComm)),
      commissionAmount: toNumber(cell(row, cols.commissionAmount)),
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

export function buildAuditReport(
  processingMonth: string,
  filters: Filters,
  data: AppData,
  countryToRegion: Record<string, string>,
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

  for (const current of Object.values(data.currentScrById)) {
    if (!isYes(current.activeStatus)) continue;
    if (current.terminationDate && !isRehireNewHire(current)) continue;
    if (!current.hireDate || current.hireDate < newHireStart || current.hireDate > newHireEnd) continue;
    const context = resolveEmployeeContext(current.employeeId, data, countryToRegion);
    if (!matchesFilters(context, filters)) continue;
    const people = data.peopleById[current.employeeId];
    const balance = data.balanceById[current.employeeId];
    rows.push(
      createAuditRow("New Hire", processingMonth, current.employeeId, current.fullName || context.name, context, {
        hireDate: formatDate(current.hireDate),
        currentJobTitle: current.jobTitle,
        currentBusinessUnit: current.businessUnit,
        currentCountry: current.country,
        currentCurrency: current.currency,
        currentCommissionAmount: numericCell(current.commissionAmount),
        rehireInPeople: people ? "Yes" : "No",
        negativeBalance: people && balance ? formatNegativeBalance(balance) : "",
        changeSummary: people
          ? balance
            ? "Employee exists in People and has a negative payment balance."
            : "Employee exists in People."
          : "",
      }),
    );
  }

  const sharedActiveIds = intersectKeys(data.currentScrById, data.previousScrById).filter((employeeId) => {
    const current = data.currentScrById[employeeId];
    const previous = data.previousScrById[employeeId];
    return Boolean(current) && Boolean(previous) && isYes(current.activeStatus) && isYes(previous.activeStatus);
  });

  for (const employeeId of sharedActiveIds) {
    const current = data.currentScrById[employeeId];
    const previous = data.previousScrById[employeeId];
    if (!current || !previous) continue;
    const context = resolveEmployeeContext(employeeId, data, countryToRegion);
    if (!matchesFilters(context, filters)) continue;

    const changes = [
      compareField("Job Title", previous.jobTitle, current.jobTitle),
      compareField("Supervisory Manager", previous.supervisoryManager, current.supervisoryManager),
      compareField("OTE (Base+Comm)", previous.oteBaseComm, current.oteBaseComm),
      compareField("Commission Amount", previous.commissionAmount, current.commissionAmount),
      compareField("Business Unit", previous.businessUnit, current.businessUnit),
      compareField("Country", previous.country, current.country),
      compareField("Currency", previous.currency, current.currency),
    ].filter((item) => item.changed);

    if (changes.length > 0) {
      rows.push(
        createAuditRow("Change to Existing Participant", processingMonth, employeeId, current.fullName || context.name, context, {
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
          changeSummary: changes.map((item) => item.label).join(", "),
        }),
      );
    }

    const previousOnLeave = isYes(previous.onLeave);
    const currentOnLeave = isYes(current.onLeave);
    if (previousOnLeave !== currentOnLeave) {
      const loa = data.loaById[employeeId];
      if (!loa) warnings.push(`LOA detail not found for employee ${employeeId}.`);
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
      rows.push(
        createAuditRow("Transfer to Non-Sales", processingMonth, employeeId, previous.fullName || context.name, context, {
          transferDirection: "Sales to Non-Sales",
          previousJobTitle: previous.jobTitle,
          previousBusinessUnit: previous.businessUnit,
          previousCountry: previous.country,
          changeSummary: "Active in the previous month SCR but missing from the current month SCR.",
        }),
      );
    }
  }

  for (const [employeeId, current] of Object.entries(data.currentScrById)) {
    if (!isYes(current.activeStatus)) continue;
    if (data.previousScrById[employeeId]) continue;
    if (!current.hireDate || current.hireDate >= transferInCutoff) continue;
    const context = resolveEmployeeContext(employeeId, data, countryToRegion);
    if (!matchesFilters(context, filters)) continue;
    rows.push(
      createAuditRow("Transfer to Sales", processingMonth, employeeId, current.fullName || context.name, context, {
        transferDirection: "Non-Sales to Sales",
        hireDate: formatDate(current.hireDate),
        currentJobTitle: current.jobTitle,
        currentBusinessUnit: current.businessUnit,
        currentCountry: current.country,
        changeSummary:
          "Active in the current month SCR, not present in the previous month SCR, and hire date is earlier than the previous month 15th.",
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
    rows.push(
      createAuditRow("Termination", processingMonth, employeeId, current.fullName || previous.fullName || context.name, context, {
        terminationDate: formatDate(current.terminationDate ?? previous.terminationDate),
        microsoftTransfer: transfer ? "Yes" : "No",
        changeSummary: transfer ? "Employee also appears in the Transfer to MSFT file." : "",
      }),
    );
  }

  rows.sort((left, right) => {
    const itemCompare = left.auditItem.localeCompare(right.auditItem);
    if (itemCompare !== 0) return itemCompare;
    return left.employeeId.localeCompare(right.employeeId);
  });

  return { rows, warnings: dedupeStrings(warnings) };
}

export function buildAuditWorkbook(rows: AuditRow[], fileNames: Record<string, string>): ArrayBuffer {
  const wb = XLSX.utils.book_new();
  const reportRows = rows.map((row) => ({ ...row }));
  const reportSheet = XLSX.utils.json_to_sheet(reportRows);
  applyHeaderStyle(reportSheet);
  reportSheet["!autofilter"] = {
    ref: XLSX.utils.encode_range(
      reportSheet["!ref"] ? XLSX.utils.decode_range(reportSheet["!ref"]) : { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } },
    ),
  };
  reportSheet["!cols"] = buildColumnWidths(reportRows);
  XLSX.utils.book_append_sheet(wb, reportSheet, "Audit Report");

  const summaryRows = [
    ...Object.entries(fileNames).map(([key, value]) => ({ Section: "Uploaded File", Name: key, Value: value })),
    ...Object.entries(countByAuditItem(rows)).map(([key, value]) => ({ Section: "Audit Count", Name: key, Value: value })),
    { Section: "Audit Count", Name: "Total Rows", Value: rows.length },
  ];
  const summarySheet = XLSX.utils.json_to_sheet(summaryRows);
  applyHeaderStyle(summarySheet);
  summarySheet["!cols"] = buildColumnWidths(summaryRows);
  XLSX.utils.book_append_sheet(wb, summarySheet, "Summary");

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

function buildColumnWidths(rows: Record<string, unknown>[]): ColInfo[] {
  if (rows.length === 0) return [];
  const keys = Object.keys(rows[0]);
  return keys.map((key) => {
    const maxValueLength = Math.max(key.length, ...rows.map((row) => String(row[key] ?? "").length));
    return { wch: Math.min(Math.max(maxValueLength + 2, 14), 36) };
  });
}

function countByAuditItem(rows: AuditRow[]): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const row of rows) counts[row.auditItem] = (counts[row.auditItem] ?? 0) + 1;
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
  return {
    auditItem,
    processingMonth,
    employeeId,
    employeeName,
    region: context.region,
    lob: context.lob,
    country: context.country,
    changeSummary: mergedChangeSummary,
    peoplePlanEffectiveDate: context.peoplePlanEffectiveDate,
    peopleBusinessUnit: context.peopleBusinessUnit,
    analystName: context.analystName,
    planType: context.planType,
    hireDate: "",
    terminationDate: "",
    rehireInPeople: "",
    negativeBalance: "",
    previousJobTitle: "",
    currentJobTitle: "",
    previousSupervisoryManager: "",
    currentSupervisoryManager: "",
    previousCommissionAmount: "",
    currentCommissionAmount: "",
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
    uploadDate: context.uploadDate,
    ...restValues,
  };
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
  uploadDate: string;
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
  const lob = displayOrUnmapped(current?.businessUnit || previous?.businessUnit || people?.businessUnit || "");

  return {
    name: current?.fullName || previous?.fullName || people?.fullName || employeeId,
    region,
    lob,
    country,
    peoplePlanEffectiveDate: formatDate(people?.planEffectiveDate ?? null),
    peopleBusinessUnit: people?.businessUnit ?? "",
    analystName: people?.analystName ?? "",
    planType: people?.planType ?? "",
    uploadDate: formatDate(people?.uploadDate ?? null),
  };
}

function matchesFilters(context: ResolvedContext, filters: Filters): boolean {
  return filters.regions.includes(context.region) && filters.lobs.includes(context.lob) && filters.countries.includes(context.country);
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

function isRehireNewHire(record: ScrRecord): boolean {
  return isYes(record.isRehire) && Boolean(record.originalHireDate && record.hireDate && record.originalHireDate < record.hireDate);
}

function hasChanged(changes: { label: string; changed: boolean }[], label: string): boolean {
  return changes.some((item) => item.label === label && item.changed);
}

function formatNegativeBalance(summary: BalanceSummary): string {
  return Object.entries(summary.negativeTotalByCurrency)
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
  const workbook = normalizeWorkbookRanges(XLSX.read(await file.arrayBuffer(), { type: "array", cellDates: false }));
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

async function fetchDefaultFile(fileName: string): Promise<ArrayBuffer> {
  const encoded = encodeURI(fileName);
  const baseUrl = normalizeBaseUrl(import.meta.env.BASE_URL);
  const candidatePaths = [
    `defaults/${encoded}`,
    `./defaults/${encoded}`,
    `/defaults/${encoded}`,
    `${baseUrl}defaults/${encoded}`,
  ];

  for (const filePath of candidatePaths) {
    try {
      const response = await fetch(filePath);
      if (response.ok) return await response.arrayBuffer();
    } catch {
      // try the next path
    }
  }

  throw new Error(`Could not load local reference file: ${fileName}`);
}

function parseCountryRegionWorkbook(workbook: WorkBook): Record<string, string> {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils
    .sheet_to_json<Record<string, unknown>>(sheet, { defval: null })
    .reduce<Record<string, string>>((acc, row) => {
      const country = text(row.Country).toLowerCase();
      const region = normalizeRegionValue(text(row.Region));
      if (country && region) acc[country] = region;
      return acc;
    }, {});
}

function normalizeBaseUrl(value: unknown): string {
  const raw = typeof value === "string" ? value : "/";
  const normalized = raw.trim() || "/";
  return normalized.endsWith("/") ? normalized : `${normalized}/`;
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
