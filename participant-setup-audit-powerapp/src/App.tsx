import { useEffect, useMemo, useState } from "react";
import {
  buildAuditReport,
  buildAuditWorkbook,
  buildDownloadFileName,
  buildFilterOptions,
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
} from "./lib/engine";
import type { AppData, AuditRow, FilterOptions, Filters, UploadDefinition, UploadStatus } from "./lib/types";

const TABLE_COLUMNS: Array<{ key: keyof AuditRow; label: string }> = [
  { key: "auditItem", label: "Audit Item" },
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

type UploadKey = UploadDefinition["key"];
type TabKey = "audit" | "instruction";

function App() {
  const [activeTab, setActiveTab] = useState<TabKey>("audit");
  const [data, setData] = useState<AppData>(createEmptyAppData);
  const [uploadStatuses, setUploadStatuses] = useState<Partial<Record<UploadKey, UploadStatus>>>({});
  const [isBusy, setIsBusy] = useState(false);
  const [errors, setErrors] = useState<string[]>([]);
  const [warnings, setWarnings] = useState<string[]>([]);
  const [rows, setRows] = useState<AuditRow[]>([]);
  const [downloadBlob, setDownloadBlob] = useState<Blob | null>(null);
  const [downloadName, setDownloadName] = useState("");
  const [countryRegionMap, setCountryRegionMap] = useState<Record<string, string>>({});

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
      const result = buildAuditReport(processingMonth, filters, data, countryRegionMap);
      setRows(result.rows);
      setWarnings(result.warnings);
      const fileNames = Object.fromEntries(
        uploadDefinitions.map((definition) => [definition.label, uploadStatuses[definition.key]?.fileName ?? ""]),
      );
      const workbook = buildAuditWorkbook(result.rows, fileNames);
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

  function resetApp(): void {
    setActiveTab("audit");
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
  }

  return (
    <div className="app-shell">
      <header className="hero">
        <div className="hero-title">
          <h1>Participant Setup Audit</h1>
          <p className="app-version">Version 1.0</p>
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
          Audit Workspace
        </button>
        <button className={activeTab === "instruction" ? "tab is-active" : "tab"} onClick={() => setActiveTab("instruction")}>
          User Instruction
        </button>
      </nav>

      {activeTab === "audit" ? (
        <main className="content-grid">
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
        </main>
      ) : (
        <main className="panel instruction-panel">
          <h2>User Instruction</h2>
          <ol>
            <li>Upload all eight files in the order shown on the Audit Workspace tab.</li>
            <li>Employee IDs are normalized to six digits. Numeric IDs shorter than six digits remain valid and are left-padded automatically.</li>
            <li>Rows with non-numeric employee IDs are ignored as placeholders or test records.</li>
            <li>Region values are limited to APAC, EMEA, LATAM, and NAMER by the bundled Country Region Mapping reference.</li>
            <li>Choose the processing month, then adjust Region, LOB, and Country if you do not want the default Select All behavior.</li>
            <li>Click Generate Report to create the audit table.</li>
            <li>Click Download Excel to export the audit output and summary sheet.</li>
          </ol>
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
