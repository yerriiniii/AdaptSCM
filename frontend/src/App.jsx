import { useEffect, useMemo, useRef, useState } from "react";
import axios from "axios";
import * as XLSX from "xlsx";

const API_BASE = import.meta.env.VITE_API_BASE_URL || "http://localhost:8000";
const OVERSEAS_UPLOAD_COUNTRIES = ["US", "JP", "TW", "SEA", "HK", "CN"];
const SETTINGS_COUNTRY_ORDER = ["KR", "US", "JP", "TW", "SEA", "HK", "CN"];

function toFixed(value, digits = 2) {
  if (value === null || value === undefined) return "-";
  const parsed = Number(value);
  if (Number.isNaN(parsed)) return String(value);
  return parsed.toFixed(digits);
}

function detectCountry(name = "") {
  const n = name.toLowerCase();
  if (n.includes("tw") || n.includes("taiwan") || n.includes("대만")) return "TW";
  if (n.includes("jp") || n.includes("japan") || n.includes("일본")) return "JP";
  if (n.includes("us") || n.includes("usa") || n.includes("미국")) return "US";
  return "KR";
}

function detectDate(name = "") {
  const n = name.toLowerCase();
  const ymd = n.match(/(20\d{2})[-_\.]?([01]\d)[-_\.]?([0-3]\d)/);
  if (ymd) return `${ymd[1]}-${ymd[2]}-${ymd[3]}`;
  return "";
}

function formatFileSize(bytes = 0) {
  const kb = Math.max(1, Math.round(bytes / 1024));
  return `${kb}KB`;
}

function extractBaseProductCode(itemno = "") {
  const raw = String(itemno ?? "").trim().toUpperCase();
  if (!raw) return "";
  // Prefix(국가/판매처), suffix(비고)와 무관하게 5자리 상품코드만 추출
  const m = raw.match(/(\d{5})/);
  return m ? m[1] : raw;
}

function countryLabel(code = "KR") {
  if (code === "KR") return "한국";
  if (code === "TW") return "대만";
  if (code === "JP") return "일본";
  if (code === "US") return "미국";
  if (code === "HK") return "홍콩";
  if (code === "CN") return "중국";
  if (code === "SEA") return "동남아";
  return code;
}

export default function App() {
  const [fileEntries, setFileEntries] = useState([]);
  const [rawInventoryRows, setRawInventoryRows] = useState([]);
  const [rawInventoryDates, setRawInventoryDates] = useState([]);
  const [inventorySummary, setInventorySummary] = useState({ item_count: 0, date_count: 0 });
  const [availableCountries, setAvailableCountries] = useState([]);
  const [inventoryLoading, setInventoryLoading] = useState(false);
  const [inventoryError, setInventoryError] = useState("");
  const [inventoryKeyword, setInventoryKeyword] = useState("");
  const [inventoryDateRange, setInventoryDateRange] = useState("all");
  const [inventoryStartDate, setInventoryStartDate] = useState("");
  const [inventoryEndDate, setInventoryEndDate] = useState("");
  const [inventoryLevelFilter, setInventoryLevelFilter] = useState("1");
  const [inventoryRequested, setInventoryRequested] = useState(false);
  const [scopeResultCache, setScopeResultCache] = useState({});
  const [scopeErrorCache, setScopeErrorCache] = useState({});
  const [showKrCompare, setShowKrCompare] = useState(false);
  const [uploadInputKey, setUploadInputKey] = useState(0);
  const [countryTabMode, setCountryTabMode] = useState("KR");
  const [selectedOverseasCountry, setSelectedOverseasCountry] = useState(OVERSEAS_UPLOAD_COUNTRIES[0]);
  const [topScrollWidth, setTopScrollWidth] = useState(0);
  const topScrollRef = useRef(null);
  const tableScrollRef = useRef(null);
  const syncingScrollRef = useRef(false);

  function matchesCountryScope(code) {
    const country = String(code || "KR");
    if (countryTabMode === "KR") return country === "KR";
    if (countryTabMode === "OVERSEAS") {
      if (country === "KR") return false;
      if (selectedOverseasCountry) return country === selectedOverseasCountry;
      return true;
    }
    return true;
  }

  function getScopeKey(mode = countryTabMode, overseasCode = selectedOverseasCountry) {
    if (mode === "KR") return "KR";
    if (mode === "OVERSEAS") return `OVERSEAS:${overseasCode || OVERSEAS_UPLOAD_COUNTRIES[0]}`;
    return "__NONE__";
  }

  const currentScopeKey = useMemo(
    () => getScopeKey(countryTabMode, selectedOverseasCountry),
    [countryTabMode, selectedOverseasCountry]
  );
  const isKRScope = countryTabMode === "KR";
  const isOverseasScope = countryTabMode === "OVERSEAS";

  const scopedFileEntries = useMemo(
    () => fileEntries.filter((entry) => matchesCountryScope(entry.country)),
    [fileEntries, countryTabMode, selectedOverseasCountry]
  );
  const inventoryFiles = useMemo(() => scopedFileEntries.map((entry) => entry.file), [scopedFileEntries]);

  useEffect(() => {
    if (currentScopeKey === "__NONE__") return;
    const cached = scopeResultCache[currentScopeKey];
    if (cached) {
      setRawInventoryRows(cached.rows || []);
      setRawInventoryDates(cached.dates || []);
      setInventorySummary(cached.summary || { item_count: 0, date_count: 0 });
      setAvailableCountries(cached.countries || []);
      setInventoryRequested(Boolean(cached.requested));
      setInventoryError(scopeErrorCache[currentScopeKey] || "");
      return;
    }
    setRawInventoryRows([]);
    setRawInventoryDates([]);
    setInventorySummary({ item_count: 0, date_count: 0 });
    setInventoryRequested(false);
    setInventoryError(scopeErrorCache[currentScopeKey] || "");
  }, [currentScopeKey, scopeResultCache, scopeErrorCache]);

  function setDatePreset(nextRange) {
    setInventoryDateRange(nextRange);
    if (nextRange !== "custom") {
      setInventoryStartDate("");
      setInventoryEndDate("");
    }
  }

  const filteredDateColumns = useMemo(() => {
    if (!rawInventoryDates.length) return [];

    const scopedDateKeys = rawInventoryDates.filter((dateKey) =>
      rawInventoryRows.some(
        (row) => matchesCountryScope(row.country) && Number(row[dateKey] || 0) !== 0
      )
    );
    const dateSource = scopedDateKeys.length ? scopedDateKeys : rawInventoryDates;

    const parsed = dateSource
      .map((dt) => new Date(dt))
      .filter((dt) => !Number.isNaN(dt.getTime()));
    if (!parsed.length) return dateSource;

    let from = null;
    let to = null;
    const maxDate = new Date(Math.max(...parsed.map((d) => d.getTime())));

    if (inventoryDateRange === "7d") {
      from = new Date(maxDate);
      from.setDate(maxDate.getDate() - 6);
    } else if (inventoryDateRange === "30d") {
      from = new Date(maxDate);
      from.setDate(maxDate.getDate() - 29);
    } else if (inventoryDateRange === "3m") {
      from = new Date(maxDate);
      from.setDate(maxDate.getDate() - 89);
    }

    if (inventoryDateRange === "custom" && inventoryStartDate) {
      const start = new Date(inventoryStartDate);
      if (!Number.isNaN(start.getTime())) from = from ? (start > from ? start : from) : start;
    }
    if (inventoryDateRange === "custom" && inventoryEndDate) {
      const end = new Date(inventoryEndDate);
      if (!Number.isNaN(end.getTime())) to = end;
    }

    return dateSource.filter((dateKey) => {
      const d = new Date(dateKey);
      if (Number.isNaN(d.getTime())) return false;
      if (from && d < from) return false;
      if (to && d > to) return false;
      return true;
    });
  }, [
    rawInventoryDates,
    rawInventoryRows,
    countryTabMode,
    selectedOverseasCountry,
    inventoryDateRange,
    inventoryStartDate,
    inventoryEndDate,
  ]);

  const filteredRows = useMemo(() => {
    if (!rawInventoryRows.length) return [];
    const needle = inventoryKeyword.trim().toLowerCase();

    return rawInventoryRows.filter((row) => {
      if (!matchesCountryScope(row.country)) return false;

      if (!isKRScope && inventoryLevelFilter !== "all") {
        const lv = (String(row.level ?? "").replace(/^0+/, "") || "0").trim();
        if (lv !== inventoryLevelFilter) return false;
      }
      if (needle) {
        const item = String(row.itemno ?? "").toLowerCase();
        const desc = String(row.description ?? "").toLowerCase();
        if (!item.includes(needle) && !desc.includes(needle)) return false;
      }
      return true;
    });
  }, [
    rawInventoryRows,
    inventoryKeyword,
    inventoryLevelFilter,
    countryTabMode,
    selectedOverseasCountry,
    isKRScope,
  ]);

  const overseasCountries = useMemo(() => OVERSEAS_UPLOAD_COUNTRIES, []);

  const groupedFileEntries = useMemo(() => {
    const groups = {};
    for (const code of SETTINGS_COUNTRY_ORDER) groups[code] = [];
    for (const entry of fileEntries) {
      const code = String(entry.country || "").toUpperCase();
      if (!groups[code]) groups[code] = [];
      groups[code].push(entry);
    }
    for (const key of Object.keys(groups)) {
      groups[key].sort((a, b) => {
        const ad = String(a.date || "").trim();
        const bd = String(b.date || "").trim();
        if (ad && bd) return ad.localeCompare(bd);
        if (ad && !bd) return -1;
        if (!ad && bd) return 1;
        return String(a.name || "").localeCompare(String(b.name || ""));
      });
    }
    return groups;
  }, [fileEntries]);

  const krCompareMap = useMemo(() => {
    if (!rawInventoryRows.length || !rawInventoryDates.length) return new Map();
    const map = new Map();
    for (const row of rawInventoryRows) {
      if (String(row.country || "KR") !== "KR") continue;
      const baseCode = extractBaseProductCode(row.itemno);
      const key = `${baseCode}`;
      const qty = rawInventoryDates.reduce((sum, dt) => sum + Number(row[dt] || 0), 0);
      map.set(key, Number(map.get(key) || 0) + qty);
    }
    return map;
  }, [rawInventoryRows, rawInventoryDates]);

  const krNameMap = useMemo(() => {
    const map = new Map();
    for (const row of rawInventoryRows) {
      if (String(row.country || "KR") !== "KR") continue;
      const code = extractBaseProductCode(row.itemno);
      if (!code) continue;
      const name = String(row.description || "").trim();
      if (!map.has(code) && name) map.set(code, name);
    }
    return map;
  }, [rawInventoryRows]);

  const totalInventory = useMemo(() => {
    return filteredRows.reduce((acc, row) => {
      const rowSum = filteredDateColumns.reduce((sum, dt) => sum + Number(row[dt] || 0), 0);
      return acc + rowSum;
    }, 0);
  }, [filteredRows, filteredDateColumns]);

  const krDisplayRows = useMemo(() => {
    if (!isKRScope) return [];
    return filteredRows.map((row) => {
      const qty = filteredDateColumns.reduce((sum, dt) => sum + Number(row[dt] || 0), 0);
      return {
        ...row,
        supplier: String(row.supplier || "").trim() || "-",
        warehouseStock: String(row.category || "").trim() || "-",
        currentQty: qty,
      };
    });
  }, [isKRScope, filteredRows, filteredDateColumns]);

  const hasTableData = useMemo(() => {
    if (isKRScope) return krDisplayRows.length > 0;
    return filteredRows.length > 0 && filteredDateColumns.length > 0;
  }, [isKRScope, krDisplayRows, filteredRows, filteredDateColumns]);

  const periodLabel = useMemo(() => {
    if (isKRScope) return "현 재고";
    const dates = scopedFileEntries
      .map((entry) => String(entry.date || ""))
      .filter((d) => /^\d{4}-\d{2}-\d{2}$/.test(d))
      .sort();
    if (!dates.length) return "-";
    return dates[0] === dates[dates.length - 1] ? dates[0] : `${dates[0]} ~ ${dates[dates.length - 1]}`;
  }, [scopedFileEntries, isKRScope]);
  const tableMinWidth = useMemo(() => {
    const base = 460;
    const dateCols = filteredDateColumns.length * 88;
    const compareCol = isOverseasScope ? 120 + (showKrCompare ? 120 : 0) : 0;
    return Math.max(900, base + dateCols + compareCol);
  }, [filteredDateColumns, isOverseasScope, showKrCompare]);

  useEffect(() => {
    const el = tableScrollRef.current;
    if (!el) return;
    const measure = () => {
      const w = el.scrollWidth || tableMinWidth;
      setTopScrollWidth(w);
    };
    measure();
    if (typeof ResizeObserver !== "undefined") {
      const observer = new ResizeObserver(measure);
      observer.observe(el);
      const table = el.querySelector("table");
      if (table) observer.observe(table);
      return () => observer.disconnect();
    }
    window.addEventListener("resize", measure);
    return () => window.removeEventListener("resize", measure);
  }, [hasTableData, tableMinWidth, filteredRows.length, filteredDateColumns.length, isKRScope, isOverseasScope]);

  function syncScroll(source) {
    if (syncingScrollRef.current) return;
    const topEl = topScrollRef.current;
    const tableEl = tableScrollRef.current;
    if (!topEl || !tableEl) return;
    syncingScrollRef.current = true;
    if (source === "top") tableEl.scrollLeft = topEl.scrollLeft;
    else topEl.scrollLeft = tableEl.scrollLeft;
    requestAnimationFrame(() => {
      syncingScrollRef.current = false;
    });
  }

  function renderDateHeader(dateKey) {
    const [y, m, d] = String(dateKey).split("-");
    if (y && m && d) {
      return (
        <span className="dateHead">
          <span>{y} -</span>
          <span>
            {m}-{d}
          </span>
        </span>
      );
    }
    return dateKey;
  }

  async function aggregateInventory() {
    if (countryTabMode === "SETTINGS") return;
    const scopeKey = getScopeKey(countryTabMode, selectedOverseasCountry);
    if (scopeKey === "__NONE__") return;
    setInventoryRequested(true);
    setInventoryError("");
    setScopeErrorCache((prev) => ({ ...prev, [scopeKey]: "" }));
    if (!scopedFileEntries.length) {
      setInventoryError("현재 탭/국가에 업로드된 파일이 없습니다.");
      setScopeErrorCache((prev) => ({
        ...prev,
        [scopeKey]: "현재 탭/국가에 업로드된 파일이 없습니다.",
      }));
      return;
    }
    setInventoryLoading(true);
    try {
      const formData = new FormData();
      const requestEntries = (() => {
        if (!isOverseasScope) return scopedFileEntries;
        const krEntries = fileEntries.filter((entry) => String(entry.country || "KR") === "KR");
        const merged = new Map();
        [...scopedFileEntries, ...krEntries].forEach((entry) => merged.set(entry.id, entry));
        return Array.from(merged.values());
      })();
      requestEntries.forEach((entry) => {
        formData.append("files", entry.file);
        formData.append("file_dates", entry.date || "");
        formData.append("file_countries", entry.country || "");
      });
      formData.append("level_filter", "all");
      const res = await axios.post(`${API_BASE}/api/inventory/aggregate`, formData, {
        headers: { "Content-Type": "multipart/form-data" },
      });
      setRawInventoryRows(res.data.rows || []);
      setRawInventoryDates(res.data.dates || []);
      setInventorySummary(res.data.summary || { item_count: 0, date_count: 0 });
      const countries = res.data.countries || [];
      setAvailableCountries(countries);
      const overseas = countries.filter((code) => code !== "KR");
      setSelectedOverseasCountry((prev) => prev || overseas[0] || "");
      setScopeResultCache((prev) => ({
        ...prev,
        [scopeKey]: {
          rows: res.data.rows || [],
          dates: res.data.dates || [],
          summary: res.data.summary || { item_count: 0, date_count: 0 },
          countries,
          requested: true,
        },
      }));
    } catch (err) {
      const detail = err?.response?.data?.detail;
      const msg = Array.isArray(detail) ? detail.join("\n") : detail || "재고 통합 중 오류";
      setInventoryError(msg);
      setScopeErrorCache((prev) => ({ ...prev, [scopeKey]: msg }));
    } finally {
      setInventoryLoading(false);
    }
  }

  function exportCurrentView() {
    if (isKRScope) {
      if (!krDisplayRows.length) return;
      const rows = krDisplayRows.map((row) => ({
        공급처: row.supplier,
        상품코드: row.itemno,
        상품명: row.description || "",
        "창고+정상 재고": row.warehouseStock,
        정상재고: Number(row.currentQty || 0),
      }));
      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "한국재고");
      XLSX.writeFile(wb, "inventory_kr_current_view.xlsx");
      return;
    }
    if (!filteredRows.length || !filteredDateColumns.length) return;
    const rows = filteredRows.map((row) => {
      const out = {
        상품코드: extractBaseProductCode(row.itemno),
        상품명: row.description || "",
        한국상품명: krNameMap.get(extractBaseProductCode(row.itemno)) || "",
      };
      if (showKrCompare) {
        out["한국 현 재고"] = Number(krCompareMap.get(`${extractBaseProductCode(row.itemno)}`) || 0);
      }
      out["카테고리"] = row.category || "";
      for (const dt of filteredDateColumns) {
        out[dt] = Number(row[dt] || 0);
      }
      return out;
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "해외재고");
    XLSX.writeFile(wb, "inventory_aggregated_view.xlsx");
  }

  function openFileInput() {
    document.getElementById("inventory-files-input")?.click();
  }

  async function fetchFileMetadata(files, overrideCountry = "") {
    const formData = new FormData();
    files.forEach((file) => {
      formData.append("files", file);
      formData.append("file_countries", overrideCountry || "");
    });
    const res = await axios.post(`${API_BASE}/api/inventory/file-metadata`, formData, {
      headers: { "Content-Type": "multipart/form-data" },
    });
    return res?.data?.files || [];
  }

  function onClickUpload() {
    if (countryTabMode === "SETTINGS") return;
    openFileInput();
  }

  function invalidateAggregatedResult() {
    setRawInventoryRows([]);
    setRawInventoryDates([]);
    setInventorySummary({ item_count: 0, date_count: 0 });
    setInventoryRequested(false);
    setInventoryError("");
    setScopeResultCache({});
    setScopeErrorCache({});
  }

  return (
    <div className={`pageSplit ${countryTabMode === "OVERSEAS" ? "split-overseas" : "split-default"}`}>
    <div className="dashboard">
      <div className="headerArea">
        <section className="hero">
          <h1 className="heroTitle">재고 분석 대시보드</h1>
          <div className="heroActions">
            <button
              className="primary"
              onClick={onClickUpload}
              disabled={countryTabMode === "SETTINGS"}
            >
              파일 업로드
            </button>
            <button
              className="primary"
              onClick={aggregateInventory}
              disabled={countryTabMode === "SETTINGS" || inventoryLoading || inventoryFiles.length === 0}
            >
              {inventoryLoading ? "통합 중..." : "재고 통합 실행"}
            </button>
            <button className="ghost" onClick={exportCurrentView}>
              내보내기
            </button>
            <input
              key={uploadInputKey}
              id="inventory-files-input"
              type="file"
              accept=".xlsx,.csv"
              multiple
              hidden
              onChange={(e) => {
              const run = async () => {
                const files = Array.from(e.target.files || []);
                if (!files.length) return;
                const override =
                  countryTabMode === "OVERSEAS"
                    ? String(selectedOverseasCountry || OVERSEAS_UPLOAD_COUNTRIES[0]).trim().toUpperCase()
                    : "KR";
                let metadata = [];
                try {
                  metadata = await fetchFileMetadata(files, override);
                } catch {
                  metadata = [];
                }
                setFileEntries((prev) => [
                  ...prev,
                  ...files.map((file, idx) => {
                    const meta = metadata[idx] || {};
                    return {
                      id: `${Date.now()}-${idx}-${file.name}`,
                      file,
                      name: file.name,
                      size: file.size,
                      country: override || meta.country || detectCountry(file.name),
                      date: meta.date || detectDate(file.name),
                    };
                  }),
                ]);
                setUploadInputKey((k) => k + 1);
              };
              run();
            }}
            />
          </div>
        </section>

        <header className="topbar">
          <div className="tabs">
            <button
              className={`tab ${countryTabMode === "KR" ? "active" : ""}`}
              onClick={() => setCountryTabMode("KR")}
            >
              한국 재고
            </button>
            <button
              className={`tab ${countryTabMode === "OVERSEAS" ? "active" : ""}`}
              onClick={() => {
                setCountryTabMode("OVERSEAS");
                setSelectedOverseasCountry((prev) => prev || OVERSEAS_UPLOAD_COUNTRIES[0]);
              }}
            >
              해외 재고
            </button>
            <button
              className={`tab ${countryTabMode === "SETTINGS" ? "active" : ""}`}
              onClick={() => setCountryTabMode("SETTINGS")}
            >
              데이터 관리
            </button>
          </div>
        </header>
      </div>

      {countryTabMode === "OVERSEAS" && (
        <div className="countryChips">
          {overseasCountries.map((code) => (
            <button
              key={code}
              className={`chip ${selectedOverseasCountry === code ? "chipActive" : ""}`}
              onClick={() => setSelectedOverseasCountry(code)}
            >
              {countryLabel(code)}
            </button>
          ))}
        </div>
      )}
      {countryTabMode !== "SETTINGS" && (
        <>
      <section className="kpiRow">
        <div className="kpiCard">
          <div className="kpiLabel">업로드된 파일</div>
          <div className="kpiValue">{inventoryFiles.length}개</div>
        </div>
        <div className="kpiCard">
          <div className="kpiLabel">분석 상품 수</div>
          <div className="kpiValue">{filteredRows.length}개</div>
        </div>
        <div className="kpiCard">
          <div className="kpiLabel">총 재고</div>
          <div className="kpiValue">{toFixed(totalInventory, 0)}</div>
        </div>
        <div className="kpiCard">
          <div className="kpiLabel">데이터 기간</div>
          <div className="kpiValue">{periodLabel}</div>
        </div>
      </section>

      <section className="tableCard">
        <div className="filterBar">
          <div className="searchWrap">
            <span className="searchIcon" aria-hidden="true">
              🔍
            </span>
            <input
              className="searchInput"
              type="text"
              value={inventoryKeyword}
              onChange={(e) => setInventoryKeyword(e.target.value)}
              placeholder="상품코드 또는 상품명 검색..."
            />
          </div>
          {!isKRScope && (
            <>
              <select value={inventoryLevelFilter} onChange={(e) => setInventoryLevelFilter(e.target.value)}>
                <option value="1">레벨 1</option>
                <option value="5">레벨 5</option>
                <option value="all">전체 레벨</option>
              </select>
              <div className="datePresetBox">
                <button
                  className={inventoryDateRange === "all" ? "preset active" : "preset"}
                  onClick={() => setDatePreset("all")}
                >
                  전체
                </button>
                <button
                  className={inventoryDateRange === "7d" ? "preset active" : "preset"}
                  onClick={() => setDatePreset("7d")}
                >
                  최근 7일
                </button>
                <button
                  className={inventoryDateRange === "30d" ? "preset active" : "preset"}
                  onClick={() => setDatePreset("30d")}
                >
                  최근 30일
                </button>
                <button
                  className={inventoryDateRange === "3m" ? "preset active" : "preset"}
                  onClick={() => setDatePreset("3m")}
                >
                  최근 3개월
                </button>
                <button
                  className={inventoryDateRange === "custom" ? "preset active" : "preset"}
                  onClick={() => setDatePreset("custom")}
                >
                  직접 설정
                </button>
              </div>
              {inventoryDateRange === "custom" && (
                <>
                  <input
                    type="date"
                    value={inventoryStartDate}
                    onChange={(e) => setInventoryStartDate(e.target.value)}
                  />
                  <input
                    type="date"
                    value={inventoryEndDate}
                    onChange={(e) => setInventoryEndDate(e.target.value)}
                  />
                </>
              )}
              {isOverseasScope && (
                <button
                  type="button"
                  className={`ghost compareToggle ${showKrCompare ? "on" : ""}`}
                  onClick={() => setShowKrCompare((v) => !v)}
                >
                  한국 현 재고 비교 {showKrCompare ? "ON" : "OFF"}
                </button>
              )}
            </>
          )}
          <button
            className="ghost resetBtn"
            onClick={() => {
              setInventoryKeyword("");
              setDatePreset("all");
              setInventoryStartDate("");
              setInventoryEndDate("");
              setInventoryLevelFilter("1");
            }}
          >
            필터 초기화
          </button>
        </div>

        {inventoryError && <pre className="error">{inventoryError}</pre>}

        {hasTableData && (
          <>
            {!isKRScope && (
              <div
                ref={topScrollRef}
                className="tableTopScroll"
                onScroll={() => syncScroll("top")}
              >
                <div style={{ width: topScrollWidth || tableMinWidth }} />
              </div>
            )}
            <div
              ref={tableScrollRef}
              className="tableWrap"
              onScroll={() => syncScroll("table")}
            >
            <table style={{ minWidth: isKRScope ? 860 : tableMinWidth }}>
              <thead>
                <tr>
                  {isKRScope && <th>공급처</th>}
                  <th>상품코드</th>
                  <th>상품명</th>
                  {isOverseasScope && <th>한국상품명</th>}
                  {isOverseasScope && showKrCompare && <th className="krCompareCol">한국 현 재고</th>}
                  <th>{isKRScope ? "창고+정상 재고" : "카테고리"}</th>
                  {isKRScope ? (
                    <th>정상재고</th>
                  ) : (
                    filteredDateColumns.map((dt) => (
                      <th key={dt}>{renderDateHeader(dt)}</th>
                    ))
                  )}
                </tr>
              </thead>
              <tbody>
                {(isKRScope ? krDisplayRows : filteredRows).map((row, idx) => (
                  <tr key={`${row.country}-${row.itemno}-${row.level}-${row.category}-${idx}`}>
                    {isKRScope && <td>{row.supplier}</td>}
                    <td>{isOverseasScope ? extractBaseProductCode(row.itemno) : row.itemno}</td>
                    <td>{row.description || "(상품명 없음)"}</td>
                    {isOverseasScope && (
                      <td>{krNameMap.get(extractBaseProductCode(row.itemno)) || "-"}</td>
                    )}
                    {isOverseasScope && showKrCompare && (
                      <td className="krCompareCol">
                        {toFixed(
                          krCompareMap.get(
                            `${extractBaseProductCode(row.itemno)}`
                          ) || 0,
                          0
                        )}
                      </td>
                    )}
                    <td>{isKRScope ? row.warehouseStock : row.category}</td>
                    {isKRScope ? (
                      <td>{toFixed(row.currentQty, 0)}</td>
                    ) : (
                      filteredDateColumns.map((dt) => (
                        <td key={`${row.itemno}-${dt}`}>{toFixed(row[dt], 0)}</td>
                      ))
                    )}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          </>
        )}

        {inventoryRequested &&
          !inventoryLoading &&
          !inventoryError &&
          !hasTableData && (
            <pre className="error">
              조건에 맞는 데이터가 없습니다.
              {!isKRScope && filteredRows.length > 0 && filteredDateColumns.length === 0
                ? "\n- 선택한 기간에 유효한 데이터가 없습니다."
                : ""}
              {"\n"}- 날짜/레벨/검색 필터를 초기화해보세요.
              {"\n"}- 먼저 파일 업로드 후 재고 통합 실행을 1회 해주세요.
            </pre>
          )}
      </section>
      </>
      )}

      {countryTabMode === "SETTINGS" && (
        <section className="settingsPane settingsCard">
          {fileEntries.length > 0 && (
            <div className="settingsTopActions">
              <button
                className="ghost settingsResetBtn"
                onClick={() => {
                  setFileEntries([]);
                  invalidateAggregatedResult();
                }}
              >
                전체 초기화
              </button>
            </div>
          )}

          {fileEntries.length === 0 && (
            <pre className="error">업로드된 파일이 없습니다. 상단의 파일 업로드 버튼을 눌러주세요.</pre>
          )}

          {Object.entries(groupedFileEntries)
            .sort(([a], [b]) => {
              const ai = SETTINGS_COUNTRY_ORDER.indexOf(a);
              const bi = SETTINGS_COUNTRY_ORDER.indexOf(b);
              const oa = ai === -1 ? 999 : ai;
              const ob = bi === -1 ? 999 : bi;
              return oa - ob || a.localeCompare(b);
            })
            .map(([country, entries]) => (
              <details key={country} className="settingsGroup" open>
                <summary className="settingsHeader">
                  <span>
                    {countryLabel(country)}
                  </span>
                  <span className="settingsHeaderRight">
                    <span>{entries.length}개 파일</span>
                    <button
                      type="button"
                      className="ghost settingsCountryResetBtn"
                      disabled={entries.length === 0}
                      onClick={(e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        if (!entries.length) return;
                        setFileEntries((prev) =>
                          prev.filter((x) => String(x.country || "").toUpperCase() !== country)
                        );
                        invalidateAggregatedResult();
                      }}
                    >
                      데이터 초기화
                    </button>
                    <span className="accordionToggle" aria-hidden="true">
                      ▾
                    </span>
                  </span>
                </summary>
                <div className="settingsRows">
                  {entries.map((entry) => (
                    <div key={entry.id} className="fileRow compact">
                      <div className="fileMeta">
                        <div className="fileName">{entry.name}</div>
                        <small>{formatFileSize(entry.size)}</small>
                      </div>
                      <div className="fileControl">
                        <span>날짜</span>
                        <input
                          type="date"
                          value={entry.date}
                          onChange={(e) => {
                            const next = e.target.value;
                            setFileEntries((prev) =>
                              prev.map((x) => (x.id === entry.id ? { ...x, date: next } : x))
                            );
                          }}
                        />
                      </div>
                      <button
                        className="ghost"
                        onClick={() => {
                          setFileEntries((prev) => prev.filter((x) => x.id !== entry.id));
                          invalidateAggregatedResult();
                        }}
                      >
                        삭제
                      </button>
                    </div>
                  ))}
                </div>
              </details>
            ))}

        </section>
      )}
    </div>
    </div>
  );
}

