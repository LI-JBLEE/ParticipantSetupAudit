export interface ProcessingMonthOption {
  label: string;
  date: Date;
}

export interface FileParseResult<T> {
  fileName: string;
  rows: number;
  data: T;
}

export interface PeopleRecord {
  employeeId: string;
  fullName: string;
  firstName: string;
  lastName: string;
  region: string;
  businessUnit: string;
  country: string;
  analystName: string;
  planEffectiveDate: Date | null;
  planType: string;
  effectiveStartDate: Date | null;
  uploadDate: Date | null;
  employeeStatus: string;
  terminationDate: Date | null;
  salary: number | null;
  salaryCurrency: string;
  annualVariable: number | null;
  hrJobTitle: string;
  level1Manager: string;
}

export interface PositionRecord {
  employeeId: string;
  positionName: string;
  personName: string;
  title: string;
  businessGroup: string;
  effectiveStartDate: Date | null;
}

export interface BalanceRow {
  employeeId: string;
  personName: string;
  positionName: string;
  remainingBalance: number;
  currency: string;
  createdDate: Date | null;
}

export interface BalanceSummary {
  employeeId: string;
  negativeTotalByCurrency: Record<string, number>;
  rows: BalanceRow[];
}

export interface QuotaAssignmentRow {
  employeeId: string;
  quotaName: string;
  type: string;
  name: string;
  personName: string;
  effectiveStartDate: Date | null;
  monthValues: Record<string, number>;
}

export interface LoaRecord {
  employeeId: string;
  region: string;
  firstDayOfLeave: Date | null;
  estimatedLastDayOfLeave: Date | null;
  totalDaysOnLeave: string;
  dateTimeCompleted: Date | null;
  latestCorrection: Date | null;
}

export interface ScrRecord {
  employeeId: string;
  firstName: string;
  lastName: string;
  fullName: string;
  originalHireDate: Date | null;
  activeStatus: string;
  onLeave: string;
  firstDayOfLeave: Date | null;
  hireDate: Date | null;
  isRehire: string;
  terminationDate: Date | null;
  jobTitle: string;
  supervisoryManager: string;
  oteBaseComm: number | null;
  commissionAmount: number | null;
  costCenter: string;
  jobFamily: string;
  businessUnit: string;
  country: string;
  currency: string;
}

export interface MsftTransferRecord {
  employeeId: string;
  subject: string;
  effectiveDate: Date | null;
  businessProcessReason: string;
  businessUnitOrganization: string;
  region: string;
}

export interface AppData {
  peopleById: Record<string, PeopleRecord>;
  peopleHistoryById: Record<string, PeopleRecord[]>;
  positionById: Record<string, PositionRecord>;
  balanceById: Record<string, BalanceSummary>;
  quotaRows: QuotaAssignmentRow[];
  loaById: Record<string, LoaRecord>;
  currentScrById: Record<string, ScrRecord>;
  previousScrById: Record<string, ScrRecord>;
  msftTransferById: Record<string, MsftTransferRecord>;
}

export interface Filters {
  regions: string[];
  lobs: string[];
  countries: string[];
}

export interface FilterOptions {
  regions: string[];
  lobs: string[];
  countries: string[];
}

export interface AuditRow {
  auditItem: string;
  processingMonth: string;
  employeeId: string;
  employeeName: string;
  region: string;
  lob: string;
  country: string;
  currentActiveStatus: string;
  currentOnLeave: string;
  currentFirstDayOfLeave: string;
  changeSummary: string;
  peoplePlanEffectiveDate: string;
  peopleBusinessUnit: string;
  analystName: string;
  planType: string;
  hireDate: string;
  terminationDate: string;
  rehireInPeople: string;
  negativeBalance: string;
  missingPeopleSetup: string;
  missingPositionSetup: string;
  previousJobTitle: string;
  currentJobTitle: string;
  previousSupervisoryManager: string;
  currentSupervisoryManager: string;
  previousCommissionAmount: number | "";
  currentCommissionAmount: number | "";
  previousBusinessUnit: string;
  currentBusinessUnit: string;
  previousCountry: string;
  currentCountry: string;
  previousCurrency: string;
  currentCurrency: string;
  loaFirstDayOfLeave: string;
  loaEstimatedLastDay: string;
  loaTotalDays: string;
  okrStartMonth: string;
  okrEndMonth: string;
  transferDirection: string;
  microsoftTransfer: string;
  peopleUploadDate: string;
}

export interface AuditBuildResult {
  rows: AuditRow[];
  warnings: string[];
  expectations: VerificationExpectation[];
}

export type VerificationRule = "exists" | "text" | "number" | "date" | "oneOf" | "unverifiable";

export interface VerificationExpectation {
  verificationId: string;
  processingMonth: string;
  generatedAt: string;
  dueDate: string;
  employeeId: string;
  employeeName: string;
  region: string;
  lob: string;
  country: string;
  analystName: string;
  analystSource: string;
  analystConfidence: string;
  analystSampleSize: number;
  auditItem: string;
  fieldKey: string;
  fieldLabel: string;
  baselineValue: string;
  expectedValue: string;
  rule: VerificationRule;
  deferred: string;
  note: string;
}

export interface VerificationFieldResult extends VerificationExpectation {
  actualValue: string;
  matched: string;
}

export type VerificationProgressStatus =
  | "Completed"
  | "Partially Completed"
  | "Pending"
  | "Manager Mismatch Only"
  | "Deferred"
  | "Not Verifiable";

export type VerificationSlaStatus = "On Time" | "Overdue" | "Not Due" | "Not Applicable";

export interface VerificationResultRow {
  verificationId: string;
  processingMonth: string;
  employeeId: string;
  employeeName: string;
  region: string;
  lob: string;
  country: string;
  analystName: string;
  analystSource: string;
  analystConfidence: string;
  analystSampleSize: number;
  auditItem: string;
  progressStatus: VerificationProgressStatus;
  slaStatus: VerificationSlaStatus;
  baselineGeneratedAt: string;
  dueDate: string;
  followUpPeopleDate: string;
  completedDate: string;
  timely: string;
  completedFields: string;
  pendingFields: string;
  notVerifiableFields: string;
  verificationNotes: string;
}

export interface FollowUpBuildResult {
  rows: VerificationResultRow[];
  fieldResults: VerificationFieldResult[];
  warnings: string[];
}

export interface DashboardBreakdownRow {
  label: string;
  employees: number;
  required: number;
  completed: number;
  pending: number;
  overdue: number;
  inferred: number;
}

export interface DashboardModel {
  regionOptions: string[];
  commissionedEmployees: number;
  setupRequired: number;
  setupRequiredRate: number;
  completed: number;
  partiallyCompleted: number;
  pending: number;
  managerMismatchOnly: number;
  completionRate: number;
  latestPeopleDate: string;
  byRegion: DashboardBreakdownRow[];
  byLob: DashboardBreakdownRow[];
  byAnalyst: DashboardBreakdownRow[];
}

export interface UploadDefinition {
  key:
    | "people"
    | "position"
    | "balance"
    | "quota"
    | "loa"
    | "currentScr"
    | "previousScr"
    | "msftTransfer";
  label: string;
  accept: string;
}

export interface UploadStatus {
  fileName: string;
  rowCount: number;
  summary: string;
}
