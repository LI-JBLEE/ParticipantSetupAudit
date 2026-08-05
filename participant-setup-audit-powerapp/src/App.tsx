import { useEffect, useMemo, useState } from "react";
import { buildDashboardHtml, buildDashboardHtmlFileName, METRIC_DESCRIPTIONS } from "./lib/dashboardHtml";
import {
  buildAuditReport,
  buildAuditWorkbook,
  buildDashboardModel,
  buildDownloadFileName,
  buildFilterOptions,
  buildFollowUpDownloadFileName,
  buildFollowUpVerification,
  buildFollowUpWorkbook,
  buildProcessingMonthOptions,
  createEmptyAppData,
  getRequiredUploads,
  loadCountryRegionReferenceMap,
  parseBalanceFile,
  parseLoaFile,
  parseMsftTransferFile,
  parsePeopleFile,
  parsePositionFile,
  parseQuotaAssignmentFile,
  parseScrFile,
  parseVerificationBaselineFile,
  summarizeSetupExecution,
} from "./lib/engine";
import type {
  AppData,
  AuditRow,
  DashboardBreakdownRow,
  DashboardModel,
  FilterOptions,
  Filters,
  UploadDefinition,
  UploadStatus,
  VerificationExpectation,
  VerificationResultRow,
} from "./lib/types";

const TABLE_COLUMNS: Array<{ key: keyof AuditRow; label: string }> = [
  { key: "auditItem", label: "Audit Item" },
  { key: "auditSubcategory", label: "Audit Subcategory" },
  { key: "employeeId", label: "Employee ID" },
  { key: "employeeName", label: "Employee Name" },
  { key: "region", label: "Region" },
  { key: "lob", label: "LOB" },
  { key: "country", label: "Country" },
  { key: "currentActiveStatus", label: "Active Status" },
  { key: "currentOnLeave", label: "On Leave" },
  { key: "currentFirstDayOfLeave", label: "First Day of Leave" },
  { key: "changeSummary", label: "Change Summary" },
  { key: "peoplePlanEffectiveDate", label: "People Plan Effective Date" },
  { key: "peopleBusinessUnit", label: "People Business Unit" },
  { key: "analystName", label: "Analyst Name" },
  { key: "planType", label: "Plan Type" },
  { key: "hireDate", label: "Hire Date" },
  { key: "terminationDate", label: "Termination Date" },
  { key: "rehireInPeople", label: "Rehire in People" },
  { key: "negativeBalance", label: "Negative Balance" },
  { key: "missingPeopleSetup", label: "Missing People Setup" },
  { key: "missingPositionSetup", label: "Missing Position Setup" },
  { key: "previousJobTitle", label: "Previous Job Title" },
  { key: "currentJobTitle", label: "Current Job Title" },
  { key: "previousSupervisoryManager", label: "Previous Supervisory Manager" },
  { key: "currentSupervisoryManager", label: "Current Supervisory Manager" },
  { key: "previousCommissionAmount", label: "Previous Commission Amount" },
  { key: "currentCommissionAmount", label: "Current Commission Amount" },
  { key: "previousBusinessUnit", label: "Previous Business Unit" },
  { key: "currentBusinessUnit", label: "Current Business Unit" },
  { key: "previousCountry", label: "Previous Country" },
  { key: "currentCountry", label: "Current Country" },
  { key: "previousCurrency", label: "Previous Currency" },
  { key: "currentCurrency", label: "Current Currency" },
  { key: "loaFirstDayOfLeave", label: "LOA First Day" },
  { key: "loaEstimatedLastDay", label: "LOA Estimated Last Day" },
  { key: "loaTotalDays", label: "LOA Total Days" },
  { key: "okrStartMonth", label: "OKR Start Month" },
  { key: "okrEndMonth", label: "OKR End Month" },
  { key: "transferDirection", label: "Transfer Direction" },
  { key: "microsoftTransfer", label: "Microsoft Transfer" },
  { key: "peopleUploadDate", label: "peopleUploadDate" },
];

const VERIFICATION_COLUMNS: Array<{ key: keyof VerificationResultRow; label: string }> = [
  { key: "progressStatus", label: "Progress" },
  { key: "slaStatus", label: "SLA" },
  { key: "employeeId", label: "Employee ID" },
  { key: "employeeName", label: "Employee Name" },
  { key: "auditItem", label: "Audit Item" },
  { key: "auditSubcategory", label: "Audit Subcategory" },
  { key: "region", label: "Region" },
  { key: "lob", label: "LOB" },
  { key: "analystName", label: "Analyst" },
  { key: "analystSource", label: "Analyst Source" },
  { key: "analystConfidence", label: "Confidence" },
  { key: "dueDate", label: "Due Date" },
  { key: "followUpPeopleDate", label: "Follow-up People Date" },
  { key: "completedFields", label: "Completed Fields" },
  { key: "pendingFields", label: "Pending Fields" },
  { key: "notVerifiableFields", label: "Not Verifiable" },
  { key: "verificationNotes", label: "Notes" },
];

type UploadKey = UploadDefinition["key"];
type TabKey = "audit" | "dashboard" | "instruction";
type AuditMode = "initial" | "followUp";

function App() {
  const [activeTab, setActiveTab] = useState<TabKey>("audit");
  const [auditMode, setAuditMode] = useState<AuditMode>("initial");
  const [data, setData] = useState<AppData>(createEmptyAppData);
  const [uploadStatuses, setUploadStatuses] = useState<Partial<Record<UploadKey, UploadStatus>>>({});
  const [isBusy, setIsBusy] = useState(false);
  const [errors, setErrors] = useState<string[]>([]);
  const [warnings, setWarnings] = useState<string[]>([]);
  const [rows, setRows] = useState<AuditRow[]>([]);
  const [downloadBlob, setDownloadBlob] = useState<Blob | null>(null);
  const [downloadName, setDownloadName] = useState("");
  const [countryRegionMap, setCountryRegionMap] = useState<Record<string, string>>({});
  const [verificationExpectations, setVerificationExpectations] = useState<VerificationExpectation[]>([]);
  const [verificationRows, setVerificationRows] = useState<VerificationResultRow[]>([]);
  const [baselineStatus, setBaselineStatus] = useState<UploadStatus>();
  const [baselineCurrentScrById, setBaselineCurrentScrById] = useState<AppData["currentScrById"]>({});
  const [followUpPeopleStatus, setFollowUpPeopleStatus] = useState<UploadStatus>();
  const [followUpPeopleById, setFollowUpPeopleById] = useState<AppData["peopleById"]>({});
  const [verificationDownloadBlob, setVerificationDownloadBlob] = useState<Blob | null>(null);
  const [verificationDownloadName, setVerificationDownloadName] = useState("");
  const [dashboardRegion, setDashboardRegion] = useState("All Regions");

  const processingMonthOptions = useMemo(() => buildProcessingMonthOptions(), []);
  const [processingMonth, setProcessingMonth] = useState(processingMonthOptions[0]?.label ?? "");

  const filterOptions = useMemo<FilterOptions>(() => buildFilterOptions(data, countryRegionMap), [data, countryRegionMap]);
  const [filters, setFilters] = useState<Filters>({ regions: [], lobs: [], countries: [] });
  const [filtersTouched, setFiltersTouched] = useState(false);
  const visibleCountryOptions = useMemo(
    () => deriveVisibleCountries(filters.regions, filterOptions.countries, countryRegionMap),
    [filters.regions, filterOptions.countries, countryRegionMap],
  );

  useEffect(() => {
    let active = true;
    void loadCountryRegionReferenceMap()
      .then((map) => {
        if (active) setCountryRegionMap(map);
      })
      .catch((error) => {
        if (active) setErrors((prev) => [...prev, `Country Region Mapping load failed: ${toError(error)}`]);
      });
    return () => {
      active = false;
    };
  }, []);

  useEffect(() => {
    setFilters((prev) => {
      if (!filtersTouched) {
        return {
          regions: [...filterOptions.regions],
          lobs: [...filterOptions.lobs],
          countries: [...filterOptions.countries],
        };
      }
      return {
        regions: keepIntersection(prev.regions, filterOptions.regions),
        lobs: keepIntersection(prev.lobs, filterOptions.lobs),
        countries: keepIntersection(prev.countries, filterOptions.countries),
      };
    });
  }, [filterOptions, filtersTouched]);

  const uploadDefinitions = useMemo(() => getRequiredUploads(), []);
  const isReadyToGenerate =
    uploadDefinitions.every((item) => Boolean(uploadStatuses[item.key])) &&
    Boolean(processingMonth) &&
    Object.keys(countryRegionMap).length > 0;
  const summaryCounts = useMemo(() => countRowsByAuditItem(rows), [rows]);
  const verificationCounts = useMemo(() => countVerificationRows(verificationRows), [verificationRows]);
  const dashboardPeopleById = Object.keys(followUpPeopleById).length > 0 ? followUpPeopleById : data.peopleById;
  const dashboardCurrentScrById =
    Object.keys(data.currentScrById).length > 0 ? data.currentScrById : baselineCurrentScrById;
  const dashboardModels = useMemo(() => {
    const allRegions = buildDashboardModel(
      dashboardCurrentScrById,
      dashboardPeopleById,
      verificationRows,
      countryRegionMap,
    );
    return Object.fromEntries([
      ["All Regions", allRegions],
      ...allRegions.regionOptions.map((region) => [
        region,
        buildDashboardModel(
          dashboardCurrentScrById,
          dashboardPeopleById,
          verificationRows,
          countryRegionMap,
          region,
        ),
      ]),
    ]) as Record<string, DashboardModel>;
  }, [dashboardCurrentScrById, dashboardPeopleById, verificationRows, countryRegionMap]);
  const dashboardModel = dashboardModels[dashboardRegion] ?? dashboardModels["All Regions"];

  useEffect(() => {
    if (dashboardRegion !== "All Regions" && !dashboardModel.regionOptions.includes(dashboardRegion)) {
      setDashboardRegion("All Regions");
    }
  }, [dashboardModel.regionOptions, dashboardRegion]);

  async function handleUpload(uploadKey: UploadKey, file: File): Promise<void> {
    setIsBusy(true);
    setErrors([]);
    setWarnings([]);
    setRows([]);
    setDownloadBlob(null);
    setDownloadName("");

    try {
      if (uploadKey === "people") {
        const result = await parsePeopleFile(file);
        setData((prev) => ({ ...prev, peopleById: result.data.byId, peopleHistoryById: result.data.historyById }));
        setUploadStatus(uploadKey, result.fileName, result.rows, "Valid numeric employee IDs only");
      } else if (uploadKey === "position") {
        const result = await parsePositionFile(file);
        setData((prev) => ({ ...prev, positionById: result.data }));
        setUploadStatus(uploadKey, result.fileName, result.rows, "Parsed for validation and future reference");
      } else if (uploadKey === "balance") {
        const result = await parseBalanceFile(file);
        setData((prev) => ({ ...prev, balanceById: result.data }));
        setUploadStatus(uploadKey, result.fileName, result.rows, "Negative balance rows only");
      } else if (uploadKey === "quota") {
        const result = await parseQuotaAssignmentFile(file);
        setData((prev) => ({ ...prev, quotaRows: result.data }));
        setUploadStatus(uploadKey, result.fileName, result.rows, "OKR Quota rows with employee IDs");
      } else if (uploadKey === "loa") {
        const result = await parseLoaFile(file);
        setData((prev) => ({ ...prev, loaById: result.data }));
        setUploadStatus(uploadKey, result.fileName, result.rows, "Latest LOA record per employee");
      } else if (uploadKey === "currentScr") {
        const result = await parseScrFile(file);
        setData((prev) => ({ ...prev, currentScrById: result.data }));
        setUploadStatus(uploadKey, result.fileName, result.rows, "Current month SCR rows");
      } else if (uploadKey === "previousScr") {
        const result = await parseScrFile(file);
        setData((prev) => ({ ...prev, previousScrById: result.data }));
        setUploadStatus(uploadKey, result.fileName, result.rows, "Previous month SCR rows");
      } else if (uploadKey === "msftTransfer") {
        const result = await parseMsftTransferFile(file);
        setData((prev) => ({ ...prev, msftTransferById: result.data }));
        setUploadStatus(uploadKey, result.fileName, result.rows, "Transfer rows with employee IDs");
      }
    } catch (error) {
      setErrors([toError(error)]);
    } finally {
      setIsBusy(false);
    }
  }

  function setUploadStatus(uploadKey: UploadKey, fileName: string, rowCount: number, summary: string): void {
    setUploadStatuses((prev) => ({
      ...prev,
      [uploadKey]: {
        fileName,
        rowCount,
        summary,
      },
    }));
  }

  function toggleFilter(kind: keyof Filters, value: string): void {
    setFiltersTouched(true);
    setFilters((prev) => {
      if (kind === "regions") {
        const exists = prev.regions.includes(value);
        const nextRegions = exists
          ? prev.regions.filter((item) => item !== value)
          : [...prev.regions, value].sort((a, b) => a.localeCompare(b));
        const nextCountries = deriveVisibleCountries(nextRegions, filterOptions.countries, countryRegionMap);
        return {
          ...prev,
          regions: nextRegions,
          countries: nextCountries,
        };
      }
      const exists = prev[kind].includes(value);
      return {
        ...prev,
        [kind]: exists ? prev[kind].filter((item) => item !== value) : [...prev[kind], value].sort((a, b) => a.localeCompare(b)),
      };
    });
  }

  function selectAll(kind: keyof Filters, values: string[]): void {
    setFiltersTouched(true);
    setFilters((prev) => {
      if (kind === "regions") {
        return {
          ...prev,
          regions: [...values],
          countries: deriveVisibleCountries(values, filterOptions.countries, countryRegionMap),
        };
      }
      return { ...prev, [kind]: [...values] };
    });
  }

  function clearAll(kind: keyof Filters): void {
    setFiltersTouched(true);
    setFilters((prev) => {
      if (kind === "regions") {
        return {
          ...prev,
          regions: [],
          countries: deriveVisibleCountries([], filterOptions.countries, countryRegionMap),
        };
      }
      return { ...prev, [kind]: [] };
    });
  }

  function generateReport(): void {
    try {
      setErrors([]);
      const generatedAt = new Date();
      const result = buildAuditReport(processingMonth, filters, data, countryRegionMap, generatedAt);
      setRows(result.rows);
      setWarnings(result.warnings);
      setVerificationExpectations(result.expectations);
      setVerificationRows(buildFollowUpVerification(result.expectations, data.peopleById).rows);
      const fileNames = Object.fromEntries(
        uploadDefinitions.map((definition) => [definition.label, uploadStatuses[definition.key]?.fileName ?? ""]),
      );
      const workbook = buildAuditWorkbook(result.rows, fileNames, result.expectations, data.currentScrById);
      setDownloadBlob(new Blob([workbook], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      setDownloadName(buildDownloadFileName());
    } catch (error) {
      setErrors([toError(error)]);
    }
  }

  function downloadWorkbook(): void {
    if (!downloadBlob) return;
    const url = URL.createObjectURL(downloadBlob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = downloadName;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  async function handleBaselineUpload(file: File): Promise<void> {
    setIsBusy(true);
    setErrors([]);
    setWarnings([]);
    setVerificationRows([]);
    setVerificationDownloadBlob(null);
    try {
      const result = await parseVerificationBaselineFile(file);
      setVerificationExpectations(result.data.expectations);
      setBaselineCurrentScrById(result.data.currentScrById);
      setBaselineStatus({
        fileName: result.fileName,
        rowCount: result.rows,
        summary: "Verification Baseline loaded",
      });
    } catch (error) {
      setErrors([toError(error)]);
    } finally {
      setIsBusy(false);
    }
  }

  async function handleFollowUpPeopleUpload(file: File): Promise<void> {
    setIsBusy(true);
    setErrors([]);
    setWarnings([]);
    setVerificationRows([]);
    setVerificationDownloadBlob(null);
    try {
      const result = await parsePeopleFile(file);
      setFollowUpPeopleById(result.data.byId);
      setFollowUpPeopleStatus({
        fileName: result.fileName,
        rowCount: result.rows,
        summary: "Latest People records loaded",
      });
    } catch (error) {
      setErrors([toError(error)]);
    } finally {
      setIsBusy(false);
    }
  }

  function generateFollowUpVerification(): void {
    try {
      setErrors([]);
      const result = buildFollowUpVerification(verificationExpectations, followUpPeopleById);
      setVerificationRows(result.rows);
      setWarnings(result.warnings);
      const workbook = buildFollowUpWorkbook(result, {
        "Original Audit Output": baselineStatus?.fileName ?? "",
        "Latest People": followUpPeopleStatus?.fileName ?? "",
      });
      setVerificationDownloadBlob(
        new Blob([workbook], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
      );
      setVerificationDownloadName(buildFollowUpDownloadFileName());
    } catch (error) {
      setErrors([toError(error)]);
    }
  }

  function downloadFollowUpWorkbook(): void {
    if (!verificationDownloadBlob) return;
    const url = URL.createObjectURL(verificationDownloadBlob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = verificationDownloadName;
    anchor.click();
    URL.revokeObjectURL(url);
  }

  function downloadDashboardHtml(): void {
    const html = buildDashboardHtml(dashboardModels, dashboardRegion, processingMonth);
    const url = URL.createObjectURL(new Blob([html], { type: "text/html;charset=utf-8" }));
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = buildDashboardHtmlFileName(processingMonth);
    anchor.click();
    URL.revokeObjectURL(url);
  }

  function resetApp(): void {
    setActiveTab("audit");
    setAuditMode("initial");
    setData(createEmptyAppData());
    setUploadStatuses({});
    setIsBusy(false);
    setErrors([]);
    setWarnings([]);
    setRows([]);
    setDownloadBlob(null);
    setDownloadName("");
    setProcessingMonth(processingMonthOptions[0]?.label ?? "");
    setFilters({ regions: [], lobs: [], countries: [] });
    setFiltersTouched(false);
    setVerificationExpectations([]);
    setVerificationRows([]);
    setBaselineStatus(undefined);
    setBaselineCurrentScrById({});
    setFollowUpPeopleStatus(undefined);
    setFollowUpPeopleById({});
    setVerificationDownloadBlob(null);
    setVerificationDownloadName("");
    setDashboardRegion("All Regions");
  }

  return (
    <div className="app-shell">
      <header className="hero">
        <div className="hero-title">
          <h1>Participant Setup Audit</h1>
          <p className="app-version">Version 1.2</p>
        </div>
        <div className="hero-side">
          <label className="field">
            <span>Processing Month</span>
            <select value={processingMonth} onChange={(event) => setProcessingMonth(event.target.value)}>
              {processingMonthOptions.map((option) => (
                <option key={option.label} value={option.label}>
                  {option.label}
                </option>
              ))}
            </select>
          </label>
          <button className="secondary-button hero-reset-button" disabled={isBusy} onClick={resetApp}>
            Reset
          </button>
        </div>
      </header>

      <nav className="tab-row" aria-label="Primary">
        <button className={activeTab === "audit" ? "tab is-active" : "tab"} onClick={() => setActiveTab("audit")}>
          Participant Setup Audit
        </button>
        <button className={activeTab === "dashboard" ? "tab is-active" : "tab"} onClick={() => setActiveTab("dashboard")}>
          Dashboard
        </button>
        <button className={activeTab === "instruction" ? "tab is-active" : "tab"} onClick={() => setActiveTab("instruction")}>
          User Instruction
        </button>
      </nav>

      {activeTab === "audit" ? (
        <main className="content-grid">
          <section className="workflow-switch" aria-label="Audit workflow">
            <button
              className={auditMode === "initial" ? "workflow-option is-active" : "workflow-option"}
              onClick={() => setAuditMode("initial")}
            >
              <span>01</span>
              <strong>Initial Audit</strong>
              <small>Generate required setup actions</small>
            </button>
            <button
              className={auditMode === "followUp" ? "workflow-option is-active" : "workflow-option"}
              onClick={() => setAuditMode("followUp")}
            >
              <span>02</span>
              <strong>Follow-up Verification</strong>
              <small>Confirm updates after seven days</small>
            </button>
          </section>

          {(errors.length > 0 || warnings.length > 0) && (
            <section className="message-stack">
              {errors.length > 0 && (
                <div className="message error">
                  <strong>Errors</strong>
                  <ul>
                    {errors.map((error) => (
                      <li key={error}>{error}</li>
                    ))}
                  </ul>
                </div>
              )}
              {warnings.length > 0 && (
                <div className="message warning">
                  <strong>Warnings</strong>
                  <ul>
                    {warnings.map((warning) => (
                      <li key={warning}>{warning}</li>
                    ))}
                  </ul>
                </div>
              )}
            </section>
          )}

          {auditMode === "initial" ? (
            <>
          <section className="panel">
            <div className="section-head">
              <div>
                <h2>1. Upload Files</h2>
              </div>
              <div className={isBusy ? "busy-pill is-busy" : "busy-pill"}>{isBusy ? "Parsing..." : "Ready"}</div>
            </div>

            <div className="upload-grid">
              {uploadDefinitions.map((definition) => (
                <UploadCard
                  key={definition.key}
                  definition={definition}
                  status={uploadStatuses[definition.key]}
                  onFileSelect={(file) => void handleUpload(definition.key, file)}
                />
              ))}
            </div>
          </section>

          <section className="panel">
            <div className="section-head">
              <div>
                <h2>2. Global Filters</h2>
              </div>
              <button className="primary-button" disabled={!isReadyToGenerate} onClick={generateReport}>
                Generate Report
              </button>
            </div>

            <div className="filter-grid">
              <FilterPanel
                title="Region"
                values={filterOptions.regions}
                selected={filters.regions}
                onToggle={(value) => toggleFilter("regions", value)}
                onSelectAll={() => selectAll("regions", filterOptions.regions)}
                onClear={() => clearAll("regions")}
              />
              <FilterPanel
                title="LOB"
                values={filterOptions.lobs}
                selected={filters.lobs}
                onToggle={(value) => toggleFilter("lobs", value)}
                onSelectAll={() => selectAll("lobs", filterOptions.lobs)}
                onClear={() => clearAll("lobs")}
              />
              <FilterPanel
                title="Country"
                values={visibleCountryOptions}
                selected={filters.countries}
                onToggle={(value) => toggleFilter("countries", value)}
                onSelectAll={() => selectAll("countries", visibleCountryOptions)}
                onClear={() => clearAll("countries")}
              />
            </div>
          </section>

          <section className="panel results-panel">
            <div className="section-head">
              <div>
                <h2>3. Audit Results</h2>
              </div>
              <button className="primary-button" disabled={!downloadBlob} onClick={downloadWorkbook}>
                Download Excel
              </button>
            </div>

            <div className="metric-row">
              {Object.entries(summaryCounts).length > 0 ? (
                Object.entries(summaryCounts).map(([label, value]) => (
                  <div className="metric-card" key={label}>
                    <span>{label}</span>
                    <strong>{value.toLocaleString()}</strong>
                  </div>
                ))
              ) : (
                <div className="empty-state">Audit counts will appear here after report generation.</div>
              )}
            </div>

            <div className={rows.length > 0 ? "table-shell is-scrollable" : "table-shell"}>
              {rows.length > 0 ? (
                <table>
                  <thead>
                    <tr>
                      {TABLE_COLUMNS.map((column) => (
                        <th key={column.key}>{column.label}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {rows.map((row, index) => (
                      <tr key={`${row.auditItem}-${row.employeeId}-${index}`}>
                        {TABLE_COLUMNS.map((column) => (
                          <td key={column.key}>{row[column.key]}</td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              ) : (
                <div className="empty-state large">No report has been generated yet.</div>
              )}
            </div>
          </section>
            </>
          ) : (
            <FollowUpWorkspace
              baselineStatus={baselineStatus}
              peopleStatus={followUpPeopleStatus}
              isBusy={isBusy}
              ready={Boolean(baselineStatus && followUpPeopleStatus)}
              verificationRows={verificationRows}
              verificationCounts={verificationCounts}
              canDownload={Boolean(verificationDownloadBlob)}
              onBaselineUpload={(file) => void handleBaselineUpload(file)}
              onPeopleUpload={(file) => void handleFollowUpPeopleUpload(file)}
              onGenerate={generateFollowUpVerification}
              onDownload={downloadFollowUpWorkbook}
            />
          )}
        </main>
      ) : activeTab === "dashboard" ? (
        <Dashboard
          model={dashboardModel}
          selectedRegion={dashboardRegion}
          onRegionChange={setDashboardRegion}
          onHtmlExport={downloadDashboardHtml}
        />
      ) : (
        <main className="panel instruction-panel">
          <div className="instruction-heading">
            <p className="eyebrow">Operating guide</p>
            <h2>User Instruction</h2>
            <p>
              Use Initial Audit to identify participant setup actions, Follow-up Verification to confirm execution after
              approximately seven days, and Dashboard to monitor results.
            </p>
          </div>

          <div className="instruction-grid">
            <article className="instruction-section">
              <span className="instruction-step">01</span>
              <h3>Initial Audit</h3>
              <ol>
                <li>Select the Processing Month.</li>
                <li>Upload all eight required source files shown in the workspace.</li>
                <li>Adjust Region, LOB, or Country filters if Select All is not required.</li>
                <li>Click Generate Report, review warnings, and validate the Audit Results.</li>
                <li>Use Audit Subcategory to separate single-field changes, Variable + Other Changes, and Multiple Changes without Variable.</li>
                <li>Download the Excel file for analyst action and later verification.</li>
              </ol>
              <p className="instruction-note">
                The workbook includes Audit Report, Column Guide, Summary, Verification Baseline, and SCR Population.
              </p>
            </article>

            <article className="instruction-section">
              <span className="instruction-step">02</span>
              <h3>Follow-up Verification</h3>
              <ol>
                <li>After approximately seven days, select Follow-up Verification.</li>
                <li>Upload the original audit output and the latest People file.</li>
                <li>Click Generate Verification to compare the latest People values with the initial baseline.</li>
                <li>Review the result tiles and Audit Subcategory, then download the Verification Excel file.</li>
              </ol>
              <p className="instruction-note">
                The workbook includes Verification Report, Column Guide, Field Details, and Summary.
              </p>
            </article>

            <article className="instruction-section instruction-section-wide">
              <span className="instruction-step">03</span>
              <h3>Verification Status Guide</h3>
              <div className="instruction-status-grid">
                <p><strong>Completed</strong> All verifiable fields match the expected values.</p>
                <p><strong>Partially Completed</strong> Some verifiable fields match and others remain pending.</p>
                <p><strong>Pending</strong> No verifiable fields match the expected values yet.</p>
                <p><strong>Manager Mismatch Only</strong> The only outstanding field is Level 1 Manager.</p>
                <p><strong>Deferred</strong> The change is intentionally postponed while the employee is on LOA.</p>
                <p><strong>Not Verifiable</strong> The action cannot be confirmed from a People-only follow-up.</p>
              </div>
              <p className="instruction-note">
                Deferred and Not Verifiable remain visible in the Excel output but are excluded from Dashboard Setup Required
                and Completion Rate. Manager Mismatch Only is shown separately.
              </p>
            </article>

            <article className="instruction-section">
              <span className="instruction-step">04</span>
              <h3>Dashboard and Export</h3>
              <ul>
                <li>Use Region to update KPI, Analyst, Region, and LOB views.</li>
                <li>Hover over any KPI tile for a short metric description.</li>
                <li>PDF opens the browser print dialog with a compact portrait layout.</li>
                <li>HTML downloads a self-contained interactive dashboard with an offline Region filter.</li>
                <li>The HTML export contains aggregated Dashboard data, not employee-level source records.</li>
              </ul>
            </article>

            <article className="instruction-section">
              <span className="instruction-step">05</span>
              <h3>Key Calculation and Mapping Rules</h3>
              <ul>
                <li>Commissioned Employees are distinct employees with Active Status = Yes in the current SCR.</li>
                <li>A blank current SCR Active Status is treated as Termination; absence from current SCR is Transfer to Non-Sales.</li>
                <li>Setup Required = Completed + Partially Completed + Pending.</li>
                <li>Completion Rate = Completed / Setup Required.</li>
                <li>Missing Analyst assignments use Country + LOB, then Region + LOB; ties remain Unassigned.</li>
                <li>New Hire and Transfer to Sales use inferred ownership instead of the employee's historical People Analyst.</li>
                <li>Variable Change Only and Variable + Other Changes both require the Annual Variable update to be verified.</li>
              </ul>
            </article>

            <article className="instruction-section instruction-section-wide instruction-reference">
              <h3>LOB Derivation Priority</h3>
              <p>
                For active SCR employees: Cost Center containing GCP → GCP; Job Family beginning with Sales Development → SD;
                Advertising Sales/Operations → LMS; LCS Sales/Operations → LTS; Sales Solutions/Operations → LSS; Global Sales
                Operations → SD, except SalesQ VP → Global. Otherwise use People Business_Unit, displaying TS as LTS and MS as LMS.
              </p>
              <p className="instruction-note">
                Employee IDs are normalized to six digits. Non-numeric IDs are ignored as placeholder or test records. Related
                audit actions are consolidated into one actionable row where applicable.
              </p>
            </article>
          </div>
        </main>
      )}
    </div>
  );
}

function UploadCard({
  definition,
  status,
  onFileSelect,
}: {
  definition: UploadDefinition;
  status?: UploadStatus;
  onFileSelect: (file: File) => void;
}) {
  return (
    <label className={status ? "upload-card is-uploaded" : "upload-card"}>
      <input
        type="file"
        accept={definition.accept}
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) onFileSelect(file);
          event.currentTarget.value = "";
        }}
      />
      <span className="upload-label">{definition.label}</span>
      {status ? (
        <>
          <strong className="upload-file">{status.fileName}</strong>
          <span className="upload-meta">{status.rowCount.toLocaleString()} rows</span>
          <span className="upload-summary">{status.summary}</span>
        </>
      ) : (
        <span className="upload-summary">Click to choose a file.</span>
      )}
    </label>
  );
}

function FilterPanel({
  title,
  values,
  selected,
  onToggle,
  onSelectAll,
  onClear,
}: {
  title: string;
  values: string[];
  selected: string[];
  onToggle: (value: string) => void;
  onSelectAll: () => void;
  onClear: () => void;
}) {
  return (
    <div className="filter-panel">
      <div className="filter-head">
        <div>
          <h3>{title}</h3>
          <p>{selected.length}/{values.length} selected</p>
        </div>
        <div className="filter-actions">
          <button type="button" onClick={onSelectAll}>
            Select All
          </button>
          <button type="button" onClick={onClear}>
            Clear
          </button>
        </div>
      </div>
      <div className="filter-list">
        {values.length > 0 ? (
          values.map((value) => (
            <label key={value} className="checkbox-row">
              <input type="checkbox" checked={selected.includes(value)} onChange={() => onToggle(value)} />
              <span>{value}</span>
            </label>
          ))
        ) : (
          <div className="empty-state">Upload files to populate this filter.</div>
        )}
      </div>
    </div>
  );
}

function keepIntersection(previous: string[], nextValues: string[]): string[] {
  if (nextValues.length === 0) return [];
  return previous.filter((value) => nextValues.includes(value));
}

function countRowsByAuditItem(rows: AuditRow[]): Record<string, number> {
  const counts: Record<string, number> = {};
  for (const row of rows) counts[row.auditItem] = (counts[row.auditItem] ?? 0) + 1;
  return counts;
}

function countVerificationRows(rows: VerificationResultRow[]): Record<string, number> {
  if (rows.length === 0) return {};
  const summary = summarizeSetupExecution(rows);
  return {
    Completed: summary.completed,
    "Partially Completed": summary.partiallyCompleted,
    Pending: summary.pending,
    "Manager Mismatch Only": summary.managerMismatchOnly,
    "Completion Rate": summary.completionRate,
  };
}

function deriveVisibleCountries(regions: string[], countries: string[], countryRegionMap: Record<string, string>): string[] {
  if (regions.length === 0) return [...countries];
  return countries.filter((country) => {
    const mappedRegion = normalizeRegion(countryRegionMap[country.toLowerCase()] ?? "");
    return mappedRegion ? regions.includes(mappedRegion) : false;
  });
}

function normalizeRegion(value: string): string {
  const upper = value.trim().toUpperCase();
  if (upper === "CHINA") return "APAC";
  if (upper.includes("NAMER")) return "NAMER";
  if (upper.includes("LATAM")) return "LATAM";
  if (upper.includes("EMEA")) return "EMEA";
  if (upper.includes("APAC")) return "APAC";
  return "";
}

function toError(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export default App;

function FollowUpWorkspace({
  baselineStatus,
  peopleStatus,
  isBusy,
  ready,
  verificationRows,
  verificationCounts,
  canDownload,
  onBaselineUpload,
  onPeopleUpload,
  onGenerate,
  onDownload,
}: {
  baselineStatus?: UploadStatus;
  peopleStatus?: UploadStatus;
  isBusy: boolean;
  ready: boolean;
  verificationRows: VerificationResultRow[];
  verificationCounts: Record<string, number>;
  canDownload: boolean;
  onBaselineUpload: (file: File) => void;
  onPeopleUpload: (file: File) => void;
  onGenerate: () => void;
  onDownload: () => void;
}) {
  return (
    <>
      <section className="panel follow-up-intro">
        <div>
          <p className="eyebrow">Seven-day control</p>
          <h2>Verify that required People updates were completed</h2>
          <p>Upload the original audit output and the latest People file. The baseline travels inside the output workbook.</p>
        </div>
        <div className={isBusy ? "busy-pill is-busy" : "busy-pill"}>{isBusy ? "Parsing..." : "Ready"}</div>
      </section>

      <section className="panel">
        <div className="section-head">
          <div>
            <h2>1. Upload Follow-up Files</h2>
          </div>
          <button className="primary-button" disabled={!ready || isBusy} onClick={onGenerate}>
            Verify Updates
          </button>
        </div>
        <div className="upload-grid follow-up-grid">
          <FollowUpUploadCard
            step="A"
            label="Original Audit Output"
            accept=".xlsx"
            status={baselineStatus}
            onFileSelect={onBaselineUpload}
          />
          <FollowUpUploadCard
            step="B"
            label="Latest People"
            accept=".xlsx,.xls"
            status={peopleStatus}
            onFileSelect={onPeopleUpload}
          />
        </div>
      </section>

      <section className="panel results-panel">
        <div className="section-head">
          <div>
            <h2>2. Verification Results</h2>
          </div>
          <button className="primary-button" disabled={!canDownload} onClick={onDownload}>
            Download Verification Excel
          </button>
        </div>
        <div className="metric-row verification-metrics">
          {Object.entries(verificationCounts).length > 0 ? (
            Object.entries(verificationCounts).map(([label, value]) => (
              <div className="metric-card" key={label} title={METRIC_DESCRIPTIONS[label]}>
                <span>{label}</span>
                <strong>{label === "Completion Rate" ? `${Math.round(value * 100)}%` : value.toLocaleString()}</strong>
              </div>
            ))
          ) : (
            <div className="empty-state">Verification status will appear after both files are loaded.</div>
          )}
        </div>
        <div className={verificationRows.length > 0 ? "table-shell is-scrollable" : "table-shell"}>
          {verificationRows.length > 0 ? (
            <table>
              <thead>
                <tr>
                  {VERIFICATION_COLUMNS.map((column) => (
                    <th key={column.key}>{column.label}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {verificationRows.map((row) => (
                  <tr key={row.verificationId}>
                    {VERIFICATION_COLUMNS.map((column) => (
                      <td
                        key={column.key}
                        className={column.key === "progressStatus" || column.key === "slaStatus" ? "status-cell" : undefined}
                      >
                        {row[column.key]}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          ) : (
            <div className="empty-state large">No follow-up verification has been generated yet.</div>
          )}
        </div>
      </section>
    </>
  );
}

function FollowUpUploadCard({
  step,
  label,
  accept,
  status,
  onFileSelect,
}: {
  step: string;
  label: string;
  accept: string;
  status?: UploadStatus;
  onFileSelect: (file: File) => void;
}) {
  return (
    <label className={status ? "upload-card is-uploaded follow-up-upload" : "upload-card follow-up-upload"}>
      <input
        type="file"
        accept={accept}
        onChange={(event) => {
          const file = event.target.files?.[0];
          if (file) onFileSelect(file);
          event.currentTarget.value = "";
        }}
      />
      <span className="upload-step">{step}</span>
      <span className="upload-label">{label}</span>
      {status ? (
        <>
          <strong className="upload-file">{status.fileName}</strong>
          <span className="upload-meta">{status.rowCount.toLocaleString()} rows</span>
          <span className="upload-summary">{status.summary}</span>
        </>
      ) : (
        <span className="upload-summary">Click to choose a file.</span>
      )}
    </label>
  );
}

function Dashboard({
  model,
  selectedRegion,
  onRegionChange,
  onHtmlExport,
}: {
  model: DashboardModel;
  selectedRegion: string;
  onRegionChange: (region: string) => void;
  onHtmlExport: () => void;
}) {
  return (
    <main className="dashboard-shell">
      <section className="dashboard-heading">
        <div>
          <p className="eyebrow">Manager control tower</p>
          <h2>Commissioned population and setup execution</h2>
          <p>
            Employees with Active Status = Yes in the current SCR, paired with the latest verification baseline.
            {model.latestPeopleDate ? ` People snapshot: ${model.latestPeopleDate}.` : ""}
          </p>
        </div>
        <div className="dashboard-actions">
          <label className="field dashboard-filter">
            <span>Region</span>
            <select value={selectedRegion} onChange={(event) => onRegionChange(event.target.value)}>
              <option value="All Regions">All Regions</option>
              {model.regionOptions.map((region) => (
                <option value={region} key={region}>
                  {region}
                </option>
              ))}
            </select>
          </label>
          <div className="dashboard-export-actions" aria-label="Export Dashboard">
            <button className="secondary-button dashboard-export-button" type="button" onClick={() => window.print()}>
              PDF
            </button>
            <button className="secondary-button dashboard-export-button" type="button" onClick={onHtmlExport}>
              HTML
            </button>
          </div>
          <span className="dashboard-print-region">Region: {selectedRegion}</span>
        </div>
      </section>

      <section className="dashboard-kpis" aria-label="Setup metrics">
        <div className="dashboard-kpi-row dashboard-kpi-row-population">
          <KpiCard label="Commissioned Employees" value={model.commissionedEmployees} tone="navy" />
          <KpiCard label="Setup Required" value={model.setupRequired} tone="ink" />
          <KpiCard label="Setup Required (%)" value={model.setupRequiredRate} tone="blue" percent />
        </div>
        <div className="dashboard-kpi-row dashboard-kpi-row-execution">
          <KpiCard label="Completed" value={model.completed} tone="green" />
          <KpiCard label="Partially Completed" value={model.partiallyCompleted} tone="amber" />
          <KpiCard label="Pending" value={model.pending} tone="coral" />
          <KpiCard label="Manager Mismatch Only" value={model.managerMismatchOnly} tone="amber" />
          <KpiCard label="Completion Rate" value={model.completionRate} tone="blue" percent />
        </div>
      </section>

      {model.commissionedEmployees === 0 && model.setupRequired === 0 ? (
        <section className="panel dashboard-empty">
          <h3>Dashboard data is not loaded yet.</h3>
          <p>Run an initial audit to load the current SCR population and setup status.</p>
        </section>
      ) : (
        <>
          <section className="panel dashboard-table-panel analyst-ownership-panel">
            <div className="section-head">
              <div>
                <p className="eyebrow">Analyst view</p>
                <h2>Analyst setup ownership</h2>
                <p className="panel-copy">
                  Missing assignments use the unique top Analyst for Country + SCR LOB, then Region + SCR LOB.
                  Ties remain Unassigned.
                </p>
              </div>
              <span className="data-caption">Inferred assignments are estimates, not Xactly master data.</span>
            </div>
            <ProgressTable rows={model.byAnalyst} firstColumn="Analyst" showInferred />
          </section>
          <section className="dashboard-grid">
            <BreakdownBars title="Commissioned employees by Region" rows={model.byRegion} />
            <BreakdownBars title="Commissioned employees by LOB" rows={model.byLob} />
          </section>
          <section className="panel dashboard-table-panel">
            <div className="section-head">
              <div>
                <p className="eyebrow">Execution</p>
                <h2>Setup progress by LOB</h2>
              </div>
            </div>
            <ProgressTable rows={model.byLob} firstColumn="LOB" />
          </section>
        </>
      )}
    </main>
  );
}

function KpiCard({
  label,
  value,
  tone,
  percent = false,
}: {
  label: string;
  value: number;
  tone: string;
  percent?: boolean;
}) {
  return (
    <article className={`dashboard-kpi tone-${tone}`} title={METRIC_DESCRIPTIONS[label]}>
      <span>{label}</span>
      <strong>{percent ? `${Math.round(value * 100)}%` : value.toLocaleString()}</strong>
    </article>
  );
}

function BreakdownBars({ title, rows }: { title: string; rows: DashboardBreakdownRow[] }) {
  const maxEmployees = Math.max(1, ...rows.map((row) => row.employees));
  return (
    <section className="panel dashboard-chart">
      <h2>{title}</h2>
      <div className="bar-list">
        {rows.length > 0 ? (
          rows.map((row) => (
            <div className="bar-row" key={row.label}>
              <div className="bar-label">
                <span>{row.label}</span>
                <strong>{row.employees.toLocaleString()}</strong>
              </div>
              <div className="bar-track" aria-hidden="true">
                <span style={{ width: `${Math.max(2, (row.employees / maxEmployees) * 100)}%` }} />
              </div>
            </div>
          ))
        ) : (
          <div className="empty-state">No commissioned employees in this selection.</div>
        )}
      </div>
    </section>
  );
}

function ProgressTable({
  rows,
  firstColumn,
  showInferred = false,
}: {
  rows: DashboardBreakdownRow[];
  firstColumn: string;
  showInferred?: boolean;
}) {
  return (
    <div className="compact-table-shell">
      <table className="compact-table">
        <thead>
          <tr>
            <th>{firstColumn}</th>
            <th>Employees</th>
            <th>Required</th>
            <th>Completed</th>
            <th>Pending</th>
            <th>Overdue</th>
            {showInferred ? <th>Inferred</th> : null}
            <th>Completion</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr key={row.label}>
              <td>{row.label}</td>
              <td>{row.employees.toLocaleString()}</td>
              <td>{row.required.toLocaleString()}</td>
              <td>{row.completed.toLocaleString()}</td>
              <td>{row.pending.toLocaleString()}</td>
              <td>{row.overdue.toLocaleString()}</td>
              {showInferred ? <td>{row.inferred.toLocaleString()}</td> : null}
              <td>{row.required > 0 ? `${Math.round((row.completed / row.required) * 100)}%` : "—"}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
