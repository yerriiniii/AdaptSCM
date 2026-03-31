import { useEffect, useMemo, useRef, useState } from "react";
import axios from "axios";
import * as XLSX from "xlsx";

const API_BASE = import.meta.env.VITE_API_BASE_URL || "http://localhost:8000";
const DEFAULT_DATE_RANGE = "10d";
const OVERSEAS_UPLOAD_COUNTRIES = ["US", "JP", "TW", "SEA", "HK", "CN"];
const SETTINGS_COUNTRY_ORDER = ["KR", "US", "JP", "TW", "SEA", "HK", "CN"];

function toFixed(value, digits = 2) {
  if (value === null || value === undefined) return "-";
  const parsed = Number(value);
  if (Number.isNaN(parsed)) return String(value);
  return parsed.toFixed(digits);
}

function toSigned(value, digits = 0) {
  if (value === null || value === undefined) return "-";
  const parsed = Number(value);
  if (Number.isNaN(parsed)) return String(value);
  const fixed = parsed.toFixed(digits);
  if (parsed > 0) return `+${fixed}`;
  return fixed;
}

function toPercent(value, digits = 1) {
  if (value === null || value === undefined) return "-";
  const parsed = Number(value);
  if (Number.isNaN(parsed)) return "-";
  const fixed = parsed.toFixed(digits);
  if (parsed > 0) return `+${fixed}%`;
  return `${fixed}%`;
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

function normalizeHeaderKey(value = "") {
  return String(value)
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[_\-()]/g, "");
}

function parseDateValue(value) {
  if (value === null || value === undefined || value === "") return "";
  if (typeof value === "number" && Number.isFinite(value)) {
    const date = XLSX.SSF.parse_date_code(value);
    if (date?.y && date?.m && date?.d) {
      return `${date.y}-${String(date.m).padStart(2, "0")}-${String(date.d).padStart(2, "0")}`;
    }
  }
  const raw = String(value).trim().replace(/\.0$/, "");
  const digits = raw.replace(/\D/g, "");
  if (/^\d{8}$/.test(digits)) {
    return `${digits.slice(0, 4)}-${digits.slice(4, 6)}-${digits.slice(6, 8)}`;
  }
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return "";
  return parsed.toISOString().slice(0, 10);
}

function mostFrequent(items = []) {
  const counts = new Map();
  for (const item of items) {
    if (!item) continue;
    counts.set(item, (counts.get(item) || 0) + 1);
  }
  let best = "";
  let bestCount = 0;
  counts.forEach((count, item) => {
    if (count > bestCount) {
      best = item;
      bestCount = count;
    }
  });
  return best;
}

async function inferFileMetadata(file, overrideCountry = "") {
  const country = overrideCountry || detectCountry(file.name);
  const filenameDate = detectDate(file.name);
  if (country === "KR" && filenameDate) {
    return { country, date: filenameDate };
  }

  try {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const firstSheetName = workbook.SheetNames?.[0];
    const sheet = firstSheetName ? workbook.Sheets[firstSheetName] : null;
    if (!sheet) return { country, date: filenameDate };
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
    const header = Array.isArray(rows[0]) ? rows[0] : [];
    const normalizedHeader = header.map((cell) => normalizeHeaderKey(cell));
    const dateIndex = normalizedHeader.findIndex((key) => ["date", "일자", "날짜"].map(normalizeHeaderKey).includes(key));
    if (dateIndex === -1) return { country, date: filenameDate };
    const parsedDates = rows
      .slice(1)
      .map((row) => parseDateValue(Array.isArray(row) ? row[dateIndex] : ""))
      .filter(Boolean);
    return { country, date: mostFrequent(parsedDates) || filenameDate };
  } catch {
    return { country, date: filenameDate };
  }
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

function getLatestDateKey(dateKeys = []) {
  if (!dateKeys.length) return "";
  return [...dateKeys].sort().at(-1) || "";
}

function buildTrendData(row, dateColumns = []) {
  if (!row || !dateColumns.length) return null;

  const series = dateColumns.map((dateKey, idx) => {
    const qty = Number(row[dateKey] || 0);
    const prevQty = idx > 0 ? Number(row[dateColumns[idx - 1]] || 0) : null;
    const delta = prevQty === null ? null : qty - prevQty;
    const deltaRate = prevQty === null || prevQty === 0 ? null : ((qty - prevQty) / prevQty) * 100;
    return { dateKey, qty, prevQty, delta, deltaRate };
  });

  const latestPoint = series.at(-1) || null;
  const maxQty = Math.max(...series.map((point) => point.qty), 0);
  const minQty = Math.min(...series.map((point) => point.qty), 0);
  const range = Math.max(maxQty - minQty, 1);
  const chartWidth = 560;
  const chartHeight = 220;
  const paddingX = 30;
  const paddingY = 20;
  const innerWidth = chartWidth - paddingX * 2;
  const innerHeight = chartHeight - paddingY * 2;
  const points = series.map((point, idx) => {
    const x =
      series.length === 1
        ? chartWidth / 2
        : paddingX + (idx / Math.max(series.length - 1, 1)) * innerWidth;
    const y = paddingY + ((maxQty - point.qty) / range) * innerHeight;
    return { ...point, x, y };
  });
  const xLabelStep = Math.max(1, Math.ceil(series.length / 8));
  const showPointValueLabels = series.length <= 12;

  return {
    series,
    latestPoint,
    latestQty: latestPoint?.qty ?? 0,
    latestDelta: latestPoint?.delta ?? null,
    latestDeltaRate: latestPoint?.deltaRate ?? null,
    highestQty: maxQty,
    lowestQty: minQty,
    points,
    chartWidth,
    chartHeight,
    xLabelStep,
    showPointValueLabels,
  };
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
  const [inventoryDateRange, setInventoryDateRange] = useState(DEFAULT_DATE_RANGE);
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
  const [selectedKrTrendRowKey, setSelectedKrTrendRowKey] = useState("");
  const [selectedOverseasTrendRowKey, setSelectedOverseasTrendRowKey] = useState("");
  const [compareSelectedDate, setCompareSelectedDate] = useState("");
  const [showCautionModal, setShowCautionModal] = useState(false);
  const [settingsMutating, setSettingsMutating] = useState(false);
  const [topScrollWidth, setTopScrollWidth] = useState(0);
  const [showTopScroll, setShowTopScroll] = useState(false);
  const topScrollRef = useRef(null);
  const headerScrollRef = useRef(null);
  const tableScrollRef = useRef(null);
  const syncingScrollRef = useRef(false);
  const topbarRef = useRef(null);
  const countryChipsRef = useRef(null);
  const inventoryFilterBarRef = useRef(null);
  const compareFilterBarRef = useRef(null);
  const [stickyHeights, setStickyHeights] = useState({
    topbar: 0,
    countryChips: 0,
    inventoryFilter: 0,
    compareFilter: 0,
    topScroll: 0,
  });

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
    if (mode === "COMPARE") return "__COMPARE__";
    return "__NONE__";
  }

  const currentScopeKey = useMemo(
    () => getScopeKey(countryTabMode, selectedOverseasCountry),
    [countryTabMode, selectedOverseasCountry]
  );
  const isKRScope = countryTabMode === "KR";
  const isOverseasScope = countryTabMode === "OVERSEAS";
  const isCompareScope = countryTabMode === "COMPARE";
  const activeCountryChipsHeight = isOverseasScope ? stickyHeights.countryChips : 0;
  const activeFilterStickyTop = stickyHeights.topbar + activeCountryChipsHeight;
  const activeFilterHeight = isCompareScope ? stickyHeights.compareFilter : stickyHeights.inventoryFilter;
  const activeTopScrollHeight = showTopScroll ? stickyHeights.topScroll : 0;
  const tableHeaderTop = activeFilterStickyTop + activeFilterHeight + activeTopScrollHeight;

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

  useEffect(() => {
    const run = async () => {
      try {
        await hydratePersistedState();
      } catch (err) {
        const detail = err?.response?.data?.detail;
        setInventoryError(Array.isArray(detail) ? detail.join("\n") : detail || "저장된 데이터를 불러오는 중 오류");
      }
    };
    run();
  }, []);

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

    if (inventoryDateRange === "10d") {
      from = new Date(maxDate);
      from.setDate(maxDate.getDate() - 9);
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
    const krRows = (scopeResultCache.KR?.rows || []).filter((row) => String(row.country || "KR") === "KR");
    const krDates = scopeResultCache.KR?.dates || [];
    if (!krRows.length || !krDates.length) return new Map();
    const utcTodayKey = new Date().toISOString().slice(0, 10);
    const hasKrCurrentSnapshot = krDates.includes(utcTodayKey);
    if (!hasKrCurrentSnapshot) return new Map();
    const map = new Map();
    for (const row of krRows) {
      const baseCode = extractBaseProductCode(row.itemno);
      const key = `${baseCode}`;
      const qty = Number(row[utcTodayKey] || 0);
      map.set(key, Number(map.get(key) || 0) + qty);
    }
    return map;
  }, [scopeResultCache]);

  const hasTodayKrSnapshot = useMemo(() => {
    const utcTodayKey = new Date().toISOString().slice(0, 10);
    return (scopeResultCache.KR?.dates || []).includes(utcTodayKey);
  }, [scopeResultCache]);

  const krNameMap = useMemo(() => {
    const map = new Map();
    for (const row of scopeResultCache.KR?.rows || []) {
      if (String(row.country || "KR") !== "KR") continue;
      const code = extractBaseProductCode(row.itemno);
      if (!code) continue;
      const name = String(row.description || "").trim();
      if (!map.has(code) && name) map.set(code, name);
    }
    return map;
  }, [scopeResultCache]);

  const totalInventory = useMemo(() => {
    const latestDate = getLatestDateKey(filteredDateColumns);
    if (!latestDate) return 0;
    return filteredRows.reduce((acc, row) => acc + Number(row[latestDate] || 0), 0);
  }, [filteredRows, filteredDateColumns]);

  const krDisplayRows = useMemo(() => {
    if (!isKRScope) return [];
    return filteredRows.map((row, idx) => ({
      ...row,
      supplier: String(row.supplier || "").trim() || "-",
      warehouseStock: String(row.category || "").trim() || "-",
      trendRowKey: `${row.country}-${row.itemno}-${row.level}-${row.category}-${row.supplier}-${idx}`,
    }));
  }, [isKRScope, filteredRows]);

  const overseasDisplayRows = useMemo(() => {
    if (!isOverseasScope) return [];
    return filteredRows.map((row, idx) => ({
      ...row,
      trendRowKey: `${row.country}-${row.itemno}-${row.level}-${row.category}-${idx}`,
    }));
  }, [isOverseasScope, filteredRows]);

  const hasTableData = useMemo(() => {
    if (isKRScope) return krDisplayRows.length > 0 && filteredDateColumns.length > 0;
    return filteredRows.length > 0 && filteredDateColumns.length > 0;
  }, [isKRScope, krDisplayRows, filteredRows, filteredDateColumns]);

  const latestScopeDateLabel = useMemo(() => getLatestDateKey(filteredDateColumns) || "-", [filteredDateColumns]);

  const compareAvailableDates = useMemo(() => {
    const allDates = new Set(scopeResultCache.KR?.dates || []);
    for (const code of OVERSEAS_UPLOAD_COUNTRIES) {
      for (const dateKey of scopeResultCache[`OVERSEAS:${code}`]?.dates || []) {
        allDates.add(dateKey);
      }
    }
    return Array.from(allDates).sort();
  }, [scopeResultCache]);

  const compareLatestDateLabel = useMemo(
    () => getLatestDateKey(compareAvailableDates) || "-",
    [compareAvailableDates]
  );

  useEffect(() => {
    if (!compareAvailableDates.length) {
      if (compareSelectedDate) setCompareSelectedDate("");
      return;
    }
    if (!compareSelectedDate) {
      setCompareSelectedDate(getLatestDateKey(compareAvailableDates));
    }
  }, [compareAvailableDates, compareSelectedDate]);

  const compareCountryData = useMemo(() => {
    const allCodes = new Set();
    const krNameLookup = new Map();
    const byCountry = {};

    const collectCountry = (countryCode) => {
      const scopeKey = countryCode === "KR" ? "KR" : `OVERSEAS:${countryCode}`;
      const cached = scopeResultCache[scopeKey];
      const rows = (cached?.rows || []).filter((row) => String(row.country || "KR") === countryCode);
      const dates = cached?.dates || [];
      const hasDate = Boolean(compareSelectedDate) && dates.includes(compareSelectedDate);
      const qtyMap = new Map();

      for (const row of rows) {
        const code = extractBaseProductCode(row.itemno);
        if (!code) continue;
        allCodes.add(code);
        if (countryCode === "KR") {
          const name = String(row.description || "").trim();
          if (!krNameLookup.has(code) && name) krNameLookup.set(code, name);
        }
        if (!hasDate) continue;
        qtyMap.set(code, Number(qtyMap.get(code) || 0) + Number(row[compareSelectedDate] || 0));
      }

      byCountry[countryCode] = { hasDate, qtyMap };
    };

    collectCountry("KR");
    OVERSEAS_UPLOAD_COUNTRIES.forEach(collectCountry);

    return {
      allCodes: Array.from(allCodes).sort(),
      krNameLookup,
      byCountry,
    };
  }, [scopeResultCache, compareSelectedDate]);

  const compareMissingCountries = useMemo(() => {
    const missing = [];
    if (!compareSelectedDate) return missing;
    if (!compareCountryData.byCountry.KR?.hasDate) missing.push("한국");
    for (const code of OVERSEAS_UPLOAD_COUNTRIES) {
      if (!compareCountryData.byCountry[code]?.hasDate) missing.push(countryLabel(code));
    }
    return missing;
  }, [compareCountryData, compareSelectedDate]);

  const compareKrLatestTotal = useMemo(() => {
    const latestKrDate = getLatestDateKey(scopeResultCache.KR?.dates || []);
    if (!latestKrDate) return 0;
    return (scopeResultCache.KR?.rows || [])
      .filter((row) => String(row.country || "KR") === "KR")
      .reduce((sum, row) => sum + Number(row[latestKrDate] || 0), 0);
  }, [scopeResultCache]);

  const compareRows = useMemo(() => {
    if (!compareSelectedDate || !compareCountryData.allCodes.length) return [];
    return compareCountryData.allCodes.map((code) => {
      const row = {
        상품코드: code,
        한국상품명: compareCountryData.krNameLookup.get(code) || "",
        "한국 현 재고": compareCountryData.byCountry.KR?.hasDate
          ? Number(compareCountryData.byCountry.KR.qtyMap.get(code) || 0)
          : null,
      };
      for (const countryCode of OVERSEAS_UPLOAD_COUNTRIES) {
        row[countryLabel(countryCode)] = compareCountryData.byCountry[countryCode]?.hasDate
          ? Number(compareCountryData.byCountry[countryCode].qtyMap.get(code) || 0)
          : null;
      }
      return row;
    });
  }, [compareCountryData, compareSelectedDate]);

  const filteredCompareRows = useMemo(() => {
    if (!isCompareScope) return [];
    const needle = inventoryKeyword.trim().toLowerCase();
    if (!needle) return compareRows;
    return compareRows.filter((row) => {
      const code = String(row["상품코드"] || "").toLowerCase();
      const krName = String(row["한국상품명"] || "").toLowerCase();
      return code.includes(needle) || krName.includes(needle);
    });
  }, [compareRows, inventoryKeyword, isCompareScope]);
  const tableMinWidth = useMemo(() => {
    const base = 460;
    const dateCols = filteredDateColumns.length * 88;
    const compareCol = isOverseasScope ? 120 + (showKrCompare ? 120 : 0) : 0;
    return Math.max(900, base + dateCols + compareCol);
  }, [filteredDateColumns, isOverseasScope, showKrCompare]);
  const inventoryTableWidth = useMemo(
    () => (isKRScope ? Math.max(980, 420 + filteredDateColumns.length * 88) : tableMinWidth),
    [isKRScope, filteredDateColumns, tableMinWidth]
  );

  useEffect(() => {
    const el = tableScrollRef.current;
    if (!el) return;
    const measure = () => {
      const w = el.scrollWidth || tableMinWidth;
      setTopScrollWidth(w);
      setShowTopScroll((el.scrollWidth || 0) > (el.clientWidth || 0) + 1);
      setStickyHeights((prev) => ({
        topbar: topbarRef.current?.offsetHeight || 0,
        countryChips: countryChipsRef.current?.offsetHeight || 0,
        inventoryFilter: inventoryFilterBarRef.current?.offsetHeight || 0,
        compareFilter: compareFilterBarRef.current?.offsetHeight || 0,
        topScroll: topScrollRef.current?.offsetHeight || 0,
      }));
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
  }, [
    hasTableData,
    tableMinWidth,
    filteredRows.length,
    filteredDateColumns.length,
    filteredCompareRows.length,
    isKRScope,
    isOverseasScope,
    isCompareScope,
    showTopScroll,
  ]);

  useEffect(() => {
    if (!hasTableData || isCompareScope) return;
    const topEl = topScrollRef.current;
    const headerEl = headerScrollRef.current;
    const tableEl = tableScrollRef.current;
    if (!tableEl) return;

    const scrollToLatest = () => {
      const nextScrollLeft = Math.max(0, tableEl.scrollWidth - tableEl.clientWidth);
      if (topEl) topEl.scrollLeft = nextScrollLeft;
      if (headerEl) headerEl.scrollLeft = nextScrollLeft;
      tableEl.scrollLeft = nextScrollLeft;
    };

    scrollToLatest();
    const rafId = requestAnimationFrame(scrollToLatest);
    return () => cancelAnimationFrame(rafId);
  }, [
    hasTableData,
    isCompareScope,
    isKRScope,
    isOverseasScope,
    filteredDateColumns,
    showKrCompare,
    showTopScroll,
  ]);

  function syncScroll(source) {
    if (syncingScrollRef.current) return;
    const topEl = topScrollRef.current;
    const headerEl = headerScrollRef.current;
    const tableEl = tableScrollRef.current;
    if (!tableEl) return;
    syncingScrollRef.current = true;
    const nextScrollLeft =
      source === "top"
        ? topEl?.scrollLeft || 0
        : source === "header"
          ? headerEl?.scrollLeft || 0
          : tableEl.scrollLeft;
    if (topEl && source !== "top") topEl.scrollLeft = nextScrollLeft;
    if (headerEl && source !== "header") headerEl.scrollLeft = nextScrollLeft;
    if (source !== "table") tableEl.scrollLeft = nextScrollLeft;
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

  function renderInventoryHeaderCells() {
    return (
      <>
        {isKRScope && <th className="stickyCol stickyColSupplier">공급처</th>}
        <th className="stickyCol stickyColCode">상품코드</th>
        <th className="stickyCol stickyColName stickyColBoundary">상품명</th>
        {isOverseasScope && <th className="stickyCol stickyColKrName stickyColBoundary">한국상품명</th>}
        {(isKRScope || isOverseasScope) && (
          <th
            className={`trendActionCol stickyCol stickyColTrend ${
              !isOverseasScope || !showKrCompare ? "stickyColBoundary" : ""
            }`}
          ></th>
        )}
        {isOverseasScope && showKrCompare && (
          <th className="krCompareCol stickyCol stickyColCompare stickyColBoundary">한국 현 재고</th>
        )}
        {filteredDateColumns.map((dt) => (
          <th key={dt} className="dateCol">{renderDateHeader(dt)}</th>
        ))}
      </>
    );
  }

  function renderCompareHeaderCells() {
    return (
      <>
        <th className="compareCodeCol stickyCol stickyColCode">상품코드</th>
        <th className="compareNameCol stickyCol stickyColName stickyColBoundary">한국상품명</th>
        <th className="compareCountryCol krCompareCol">한국</th>
        {OVERSEAS_UPLOAD_COUNTRIES.map((code) => (
          <th key={code} className="compareCountryCol">{countryLabel(code)}</th>
        ))}
      </>
    );
  }

  async function aggregateInventory() {
    if (countryTabMode === "SETTINGS" || countryTabMode === "COMPARE") return;
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
      const requestEntries = scopedFileEntries.filter((entry) => entry.file && !entry.dbFileId);
      if (!requestEntries.length) {
        await hydratePersistedState();
        return;
      }
      const uploadPlans = await requestDirectUploadPlans(requestEntries);
      const planByClientId = new Map(uploadPlans.map((plan) => [String(plan.client_id), plan]));
      for (const entry of requestEntries) {
        const plan = planByClientId.get(String(entry.id));
        if (!plan) continue;
        if (!plan.file_id && plan.upload_url) {
          await uploadFileToS3(plan, entry.file);
        }
      }
      await completeDirectUploads(uploadPlans);
      await hydratePersistedState();
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
    if (isCompareScope) {
      if (!filteredCompareRows.length) return;
      const rows = filteredCompareRows.map((row) => {
        const out = {
          상품코드: row["상품코드"],
          한국상품명: row["한국상품명"] || "",
          한국: row["한국 현 재고"] === null ? "-" : Number(row["한국 현 재고"] || 0),
        };
        for (const code of OVERSEAS_UPLOAD_COUNTRIES) {
          const key = countryLabel(code);
          out[key] = row[key] === null ? "-" : Number(row[key] || 0);
        }
        return out;
      });
      const ws = XLSX.utils.json_to_sheet(rows);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "재고비교");
      XLSX.writeFile(wb, "inventory_compare_view.xlsx");
      return;
    }
    if (isKRScope) {
      if (!krDisplayRows.length) return;
      const rows = krDisplayRows.map((row) => {
        const out = {
          공급처: row.supplier,
          상품코드: row.itemno,
          상품명: row.description || "",
        };
        for (const dt of filteredDateColumns) {
          out[dt] = Number(row[dt] || 0);
        }
        return out;
      });
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
        out["한국"] = krCompareMap.size
          ? Number(krCompareMap.get(`${extractBaseProductCode(row.itemno)}`) || 0)
          : "-";
      }
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

  async function requestDirectUploadPlans(entries) {
    const payload = {
      files: entries.map((entry) => ({
        client_id: entry.id,
        original_name: entry.name,
        country_code: entry.country,
        date: entry.date || "",
        content_type: entry.file?.type || "application/octet-stream",
        size: entry.size || 0,
      })),
    };
    const res = await axios.post(`${API_BASE}/api/inventory/upload-url`, payload);
    return Array.isArray(res?.data?.files) ? res.data.files : [];
  }

  async function uploadFileToS3(plan, file) {
    await axios.put(plan.upload_url, file, {
      headers: {
        "Content-Type": file?.type || "application/octet-stream",
      },
    });
  }

  async function completeDirectUploads(plans) {
    const payload = {
      files: plans.map((plan) => ({
        client_id: plan.client_id,
        original_name: plan.original_name,
        country_code: plan.country_code,
        date: plan.date || "",
        stored_name: plan.stored_name,
        s3_key: plan.s3_key,
        size: plan.size || 0,
        content_type: plan.content_type || "application/octet-stream",
        file_id: plan.file_id || "",
      })),
    };
    return axios.post(`${API_BASE}/api/inventory/complete-upload`, payload);
  }

  function onClickUpload() {
    if (countryTabMode === "SETTINGS" || countryTabMode === "COMPARE") return;
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

  async function fetchPersistedFiles() {
    const res = await axios.get(`${API_BASE}/api/inventory/files`);
    return Array.isArray(res?.data?.files) ? res.data.files : [];
  }

  async function fetchPersistedScopeView(countryCode) {
    const res = await axios.get(`${API_BASE}/api/inventory/view`, {
      params: { country_code: countryCode },
    });
    return {
      rows: res?.data?.rows || [],
      dates: res?.data?.dates || [],
      summary: res?.data?.summary || { item_count: 0, date_count: 0 },
      countries: res?.data?.countries || [countryCode],
      requested: true,
    };
  }

  async function hydratePersistedState(options = {}) {
    const { preserveLocalOnly = true, excludeCountry = "", excludeEntryId = "" } = options;
    const [persistedFiles, ...views] = await Promise.all([
      fetchPersistedFiles(),
      fetchPersistedScopeView("KR"),
      ...OVERSEAS_UPLOAD_COUNTRIES.map((code) => fetchPersistedScopeView(code)),
    ]);

    const nextCache = {
      KR: views[0],
    };
    OVERSEAS_UPLOAD_COUNTRIES.forEach((code, idx) => {
      nextCache[`OVERSEAS:${code}`] = views[idx + 1];
    });

    const available = new Set();
    Object.values(nextCache).forEach((scope) => {
      (scope?.countries || []).forEach((code) => available.add(code));
    });

    setAvailableCountries(Array.from(available));
    setScopeResultCache(nextCache);
    setScopeErrorCache({});
    setFileEntries((prev) => {
      const persistedKeys = new Set(
        persistedFiles.map((entry) => `${entry.name}::${entry.country || ""}::${entry.date || ""}`)
      );
      const localOnlyEntries = preserveLocalOnly
        ? prev.filter((entry) => {
            if (entry.dbFileId || !entry.file) return false;
            if (excludeEntryId && entry.id === excludeEntryId) return false;
            if (excludeCountry && String(entry.country || "").toUpperCase() === excludeCountry) return false;
            const entryKey = `${entry.name}::${entry.country || ""}::${entry.date || ""}`;
            if (persistedKeys.has(entryKey)) return false;
            return true;
          })
        : [];
      const serverEntries = persistedFiles.map((entry) => ({
        id: entry.file_id,
        dbFileId: entry.file_id,
        file: null,
        name: entry.name,
        size: entry.size,
        country: entry.country,
        date: entry.date,
      }));
      return [...serverEntries, ...localOnlyEntries];
    });
  }

  async function deleteFileEntry(entry) {
    if (!entry) return;
    try {
      setSettingsMutating(true);
      if (entry.dbFileId) {
        await axios.delete(`${API_BASE}/api/inventory/files/${entry.dbFileId}`);
      }
      await hydratePersistedState({ excludeEntryId: entry.id });
    } catch (err) {
      const detail = err?.response?.data?.detail;
      window.alert(Array.isArray(detail) ? detail.join("\n") : detail || "파일 삭제 중 오류");
    } finally {
      setSettingsMutating(false);
    }
  }

  async function clearFilesByCountry(country, entries) {
    if (!entries?.length) return;
    try {
      setSettingsMutating(true);
      if (entries.some((entry) => entry.dbFileId)) {
        await axios.delete(`${API_BASE}/api/inventory/files`, {
          params: { country_code: country },
        });
      }
      await hydratePersistedState({ excludeCountry: country });
    } catch (err) {
      const detail = err?.response?.data?.detail;
      window.alert(Array.isArray(detail) ? detail.join("\n") : detail || "국가 데이터 삭제 중 오류");
    } finally {
      setSettingsMutating(false);
    }
  }

  async function clearAllFiles() {
    try {
      setSettingsMutating(true);
      if (fileEntries.some((entry) => entry.dbFileId)) {
        await axios.delete(`${API_BASE}/api/inventory/files`);
      }
      await hydratePersistedState({ preserveLocalOnly: false });
    } catch (err) {
      const detail = err?.response?.data?.detail;
      window.alert(Array.isArray(detail) ? detail.join("\n") : detail || "전체 데이터 삭제 중 오류");
    } finally {
      setSettingsMutating(false);
    }
  }

  const selectedKrTrendRow = useMemo(() => {
    if (!selectedKrTrendRowKey) return null;
    return krDisplayRows.find((row) => row.trendRowKey === selectedKrTrendRowKey) || null;
  }, [krDisplayRows, selectedKrTrendRowKey]);

  const selectedOverseasTrendRow = useMemo(() => {
    if (!selectedOverseasTrendRowKey) return null;
    return overseasDisplayRows.find((row) => row.trendRowKey === selectedOverseasTrendRowKey) || null;
  }, [overseasDisplayRows, selectedOverseasTrendRowKey]);

  const activeTrendRow = selectedKrTrendRow || selectedOverseasTrendRow;
  const activeTrendData = useMemo(
    () => buildTrendData(activeTrendRow, filteredDateColumns),
    [activeTrendRow, filteredDateColumns]
  );

  useEffect(() => {
    if (!activeTrendRow) return undefined;
    const onKeyDown = (event) => {
      if (event.key === "Escape") {
        setSelectedKrTrendRowKey("");
        setSelectedOverseasTrendRowKey("");
      }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, [activeTrendRow]);

  return (
    <div className={`pageSplit ${countryTabMode === "OVERSEAS" ? "split-overseas" : "split-default"}`}>
    <div className="dashboard">
      <div className="headerArea">
        <section className="hero">
          <div className="heroHead">
            <h1 className="heroTitle">재고 분석 대시보드</h1>
            <button
              type="button"
              className="cautionBtn"
              onClick={() => setShowCautionModal(true)}
            >
              주의사항
            </button>
          </div>
          <div className="heroActions">
            <button
              className="primary"
              onClick={onClickUpload}
              disabled={countryTabMode === "SETTINGS" || countryTabMode === "COMPARE"}
            >
              파일 업로드
            </button>
            <button
              className="primary"
              onClick={aggregateInventory}
              disabled={
                countryTabMode === "SETTINGS" ||
                countryTabMode === "COMPARE" ||
                inventoryLoading ||
                inventoryFiles.length === 0
              }
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
                const existingNames = new Set(fileEntries.map((entry) => String(entry.name || "")));
                const incomingCounts = files.reduce((acc, file) => {
                  acc[file.name] = (acc[file.name] || 0) + 1;
                  return acc;
                }, {});
                const duplicateNames = [
                  ...new Set(
                    files
                      .map((file) => file.name)
                      .filter((name) => existingNames.has(name) || incomingCounts[name] > 1)
                  ),
                ];
                const uploadableFiles = files.filter(
                  (file, index) =>
                    !existingNames.has(file.name) &&
                    files.findIndex((candidate) => candidate.name === file.name) === index
                );
                if (duplicateNames.length) {
                  window.alert(`${duplicateNames.join(", ")}\n같은 파일이 두개입니다.`);
                }
                if (!uploadableFiles.length) {
                  setUploadInputKey((k) => k + 1);
                  return;
                }
                const override =
                  countryTabMode === "OVERSEAS"
                    ? String(selectedOverseasCountry || OVERSEAS_UPLOAD_COUNTRIES[0]).trim().toUpperCase()
                    : "KR";
                const metadata = await Promise.all(
                  uploadableFiles.map((file) => inferFileMetadata(file, override))
                );
                setFileEntries((prev) => [
                  ...prev,
                  ...uploadableFiles.map((file, idx) => {
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

      </div>

      <header ref={topbarRef} className="topbar stickyTopbar">
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
            className={`tab ${countryTabMode === "COMPARE" ? "active" : ""}`}
            onClick={() => setCountryTabMode("COMPARE")}
          >
            재고 비교
          </button>
          <button
            className={`tab ${countryTabMode === "SETTINGS" ? "active" : ""}`}
            onClick={() => setCountryTabMode("SETTINGS")}
          >
            데이터 관리
          </button>
        </div>
      </header>

      {countryTabMode === "OVERSEAS" && (
        <div
          ref={countryChipsRef}
          className="countryChips stickyCountryChips"
          style={{ top: stickyHeights.topbar }}
        >
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
      {countryTabMode !== "SETTINGS" && !isCompareScope && (
        <>
      <section className="kpiRow">
        <div className="kpiCard">
          <div className="kpiLabel">업로드된 파일</div>
          <div className="kpiValue">{inventoryFiles.length}개</div>
        </div>
        <div className="kpiCard">
          <div className="kpiLabel">최신 기준일</div>
          <div className="kpiValue">{latestScopeDateLabel}</div>
        </div>
        <div className="kpiCard">
          <div className="kpiLabel">분석 상품 수</div>
          <div className="kpiValue">{filteredRows.length}개</div>
        </div>
        <div className="kpiCard">
          <div className="kpiLabel">총 재고</div>
          <div className="kpiValue">{toFixed(totalInventory, 0)}</div>
        </div>
      </section>

      <section className="tableCard">
        {isKRScope && <div className="krOnlyNotice">정상 재고만 파악합니다.</div>}
        <div
          ref={inventoryFilterBarRef}
          className="filterBar stickyFilterBar"
          style={{ top: activeFilterStickyTop }}
        >
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
          <>
            {!isKRScope && (
              <select value={inventoryLevelFilter} onChange={(e) => setInventoryLevelFilter(e.target.value)}>
                <option value="1">레벨 1</option>
                <option value="5">레벨 5</option>
                <option value="all">전체 레벨</option>
              </select>
            )}
            <div className="datePresetBox">
              <button
                className={inventoryDateRange === "10d" ? "preset active" : "preset"}
                onClick={() => setDatePreset("10d")}
              >
                최근 10일
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
                onClick={() => {
                  if (!showKrCompare && !hasTodayKrSnapshot) {
                    window.alert("오늘 일자 데이터가 존재하지 않아 한국 현 재고 비교를 할 수 없습니다.");
                    return;
                  }
                  setShowKrCompare((v) => !v);
                }}
              >
                한국 현 재고 비교 {showKrCompare ? "ON" : "OFF"}
              </button>
            )}
          </>
          <button
            className="ghost resetBtn"
            onClick={() => {
              setInventoryKeyword("");
              setDatePreset(DEFAULT_DATE_RANGE);
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
            {showTopScroll && (
              <div
                ref={topScrollRef}
                className="tableTopScroll stickyTableTopScroll"
                style={{ top: activeFilterStickyTop + activeFilterHeight }}
                onScroll={() => syncScroll("top")}
              >
                <div style={{ width: topScrollWidth || tableMinWidth }} />
              </div>
            )}
            <div className="stickyTableHeader" style={{ top: tableHeaderTop }}>
              <div
                ref={headerScrollRef}
                className="tableHeaderScroll"
                onScroll={() => syncScroll("header")}
              >
                <table
                  className={`inventoryTable stickyHeaderTable ${isKRScope ? "krTable" : "overseasTable"}`}
                  style={{ minWidth: inventoryTableWidth }}
                >
                  <thead>
                    <tr>{renderInventoryHeaderCells()}</tr>
                  </thead>
                </table>
              </div>
            </div>
            <div
              ref={tableScrollRef}
              className="tableWrap"
              onScroll={() => syncScroll("table")}
            >
            <table
              className={`inventoryTable bodyTable ${isKRScope ? "krTable" : "overseasTable"}`}
              style={{ minWidth: inventoryTableWidth }}
            >
              <tbody>
                {(isKRScope ? krDisplayRows : isOverseasScope ? overseasDisplayRows : filteredRows).map((row, idx) => (
                  <tr key={row.trendRowKey || `${row.country}-${row.itemno}-${row.level}-${row.category}-${idx}`}>
                    {isKRScope && <td className="stickyCol stickyColSupplier">{row.supplier}</td>}
                    <td className="stickyCol stickyColCode">
                      {isOverseasScope ? extractBaseProductCode(row.itemno) : row.itemno}
                    </td>
                    <td className="stickyCol stickyColName stickyColBoundary">{row.description || "(상품명 없음)"}</td>
                    {isOverseasScope && (
                      <td className="stickyCol stickyColKrName stickyColBoundary">
                        {krNameMap.get(extractBaseProductCode(row.itemno)) || "-"}
                      </td>
                    )}
                    {(isKRScope || isOverseasScope) && (
                      <td
                        className={`trendActionCell stickyCol stickyColTrend ${
                          !isOverseasScope || !showKrCompare ? "stickyColBoundary" : ""
                        }`}
                      >
                        <button
                          type="button"
                          className="ghost compareToggle on trendActionBtn"
                          onClick={() => {
                            if (isKRScope) setSelectedKrTrendRowKey(row.trendRowKey);
                            if (isOverseasScope) setSelectedOverseasTrendRowKey(row.trendRowKey);
                          }}
                        >
                          재고 변화 추이
                        </button>
                      </td>
                    )}
                    {isOverseasScope && showKrCompare && (
                      <td className="krCompareCol stickyCol stickyColCompare stickyColBoundary">
                        {krCompareMap.size
                          ? toFixed(
                              krCompareMap.get(
                                `${extractBaseProductCode(row.itemno)}`
                              ) || 0,
                              0
                            )
                          : "-"}
                      </td>
                    )}
                    {isKRScope
                      ? filteredDateColumns.map((dt) => (
                          <td key={`${row.itemno}-${dt}`} className="dateCol">{toFixed(row[dt], 0)}</td>
                        ))
                      : filteredDateColumns.map((dt) => (
                        <td key={`${row.itemno}-${dt}`} className="dateCol">{toFixed(row[dt], 0)}</td>
                      ))
                    }
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

      {isCompareScope && (
        <>
          <section className="kpiRow">
            <div className="kpiCard">
              <div className="kpiLabel">비교 가능 국가</div>
              <div className="kpiValue">{OVERSEAS_UPLOAD_COUNTRIES.length}개</div>
            </div>
            <div className="kpiCard">
              <div className="kpiLabel">최신 기준일</div>
              <div className="kpiValue">{compareLatestDateLabel}</div>
            </div>
            <div className="kpiCard">
              <div className="kpiLabel">비교 상품 수</div>
              <div className="kpiValue">{filteredCompareRows.length}개</div>
            </div>
            <div className="kpiCard">
              <div className="kpiLabel">한국 총 재고</div>
              <div className="kpiValue">{toFixed(compareKrLatestTotal, 0)}</div>
            </div>
          </section>

          <section className="tableCard">
            <div
              ref={compareFilterBarRef}
              className="filterBar stickyFilterBar"
              style={{ top: activeFilterStickyTop }}
            >
              <div className="searchWrap">
                <span className="searchIcon" aria-hidden="true">
                  🔍
                </span>
                <input
                  className="searchInput"
                  type="text"
                  value={inventoryKeyword}
                  onChange={(e) => setInventoryKeyword(e.target.value)}
                  placeholder="상품코드 또는 한국상품명 검색..."
                />
              </div>
              <input
                type="date"
                value={compareSelectedDate}
                onChange={(e) => setCompareSelectedDate(e.target.value)}
              />
              <button
                className="ghost resetBtn"
                onClick={() => {
                  setInventoryKeyword("");
                  setCompareSelectedDate(getLatestDateKey(compareAvailableDates));
                }}
              >
                필터 초기화
              </button>
            </div>

            {compareSelectedDate && compareMissingCountries.length > 0 && (
              <pre className="error">
                선택한 기준일 `{compareSelectedDate}` 에 데이터가 없는 국가: {compareMissingCountries.join(", ")}
              </pre>
            )}

            {!filteredCompareRows.length ? (
              <pre className="error">
                비교할 데이터가 없습니다.
                {"\n"}- 먼저 한국 재고 또는 해외 재고 탭에서 파일 업로드 후 재고 통합 실행을 해주세요.
              </pre>
            ) : (
              <>
                {showTopScroll && (
                  <div
                    ref={topScrollRef}
                    className="tableTopScroll stickyTableTopScroll"
                    style={{ top: activeFilterStickyTop + activeFilterHeight }}
                    onScroll={() => syncScroll("top")}
                  >
                    <div style={{ width: topScrollWidth || 980 }} />
                  </div>
                )}
              <div className="stickyTableHeader" style={{ top: tableHeaderTop }}>
                <div
                  ref={headerScrollRef}
                  className="tableHeaderScroll"
                  onScroll={() => syncScroll("header")}
                >
                  <table className="compareTable stickyHeaderTable" style={{ minWidth: 980 }}>
                    <thead>
                      <tr>{renderCompareHeaderCells()}</tr>
                    </thead>
                  </table>
                </div>
              </div>
              <div
                ref={tableScrollRef}
                className="tableWrap"
                onScroll={() => syncScroll("table")}
              >
                <table className="compareTable bodyTable" style={{ minWidth: 980 }}>
                  <tbody>
                    {filteredCompareRows.map((row) => (
                      <tr key={`compare-${row["상품코드"]}`}>
                        <td className="compareCodeCol stickyCol stickyColCode">{row["상품코드"]}</td>
                        <td className="compareNameCol stickyCol stickyColName stickyColBoundary">
                          {row["한국상품명"] || "-"}
                        </td>
                        <td className="compareCountryCol krCompareCol">
                          {row["한국 현 재고"] === null ? "-" : toFixed(row["한국 현 재고"], 0)}
                        </td>
                        {OVERSEAS_UPLOAD_COUNTRIES.map((code) => (
                          <td key={`${row["상품코드"]}-${code}`} className="compareCountryCol">
                            {row[countryLabel(code)] === null
                              ? "-"
                              : toFixed(row[countryLabel(code)], 0)}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              </>
            )}
          </section>
        </>
      )}

      {activeTrendRow && activeTrendData && (
        <div
          className="trendModalBackdrop"
          onClick={() => {
            setSelectedKrTrendRowKey("");
            setSelectedOverseasTrendRowKey("");
          }}
        >
          <div className="trendModal" onClick={(e) => e.stopPropagation()}>
            <div className="trendModalHeader">
              <div>
                <div className="trendModalTitle">재고 변화 추이</div>
                <div className="trendModalSubtitle">
                  {activeTrendRow.itemno} · {activeTrendRow.description || "(상품명 없음)"}
                </div>
              </div>
              <button
                type="button"
                className="ghost trendCloseBtn"
                onClick={() => {
                  setSelectedKrTrendRowKey("");
                  setSelectedOverseasTrendRowKey("");
                }}
              >
                닫기
              </button>
            </div>

            <div className="trendSummaryGrid">
              <div className="trendSummaryCard">
                <div className="trendSummaryLabel">최신 재고</div>
                <div className="trendSummaryValue">{toFixed(activeTrendData.latestQty, 0)}</div>
              </div>
              <div className="trendSummaryCard">
                <div className="trendSummaryLabel">직전 데이터 대비 증감</div>
                <div
                  className={`trendSummaryValue ${
                    Number(activeTrendData.latestDelta || 0) > 0
                      ? "up"
                      : Number(activeTrendData.latestDelta || 0) < 0
                        ? "down"
                        : ""
                  }`}
                >
                  {toSigned(activeTrendData.latestDelta, 0)}
                </div>
              </div>
              <div className="trendSummaryCard">
                <div className="trendSummaryLabel">직전 데이터 대비 증감률</div>
                <div
                  className={`trendSummaryValue ${
                    Number(activeTrendData.latestDeltaRate || 0) > 0
                      ? "up"
                      : Number(activeTrendData.latestDeltaRate || 0) < 0
                        ? "down"
                        : ""
                  }`}
                >
                  {toPercent(activeTrendData.latestDeltaRate, 1)}
                </div>
              </div>
              <div className="trendSummaryCard">
                <div className="trendSummaryLabel">최대-최소 변동폭</div>
                <div className="trendSummaryValue">
                  {toSigned(activeTrendData.highestQty - activeTrendData.lowestQty, 0)}
                </div>
              </div>
            </div>

            <div className="trendChartCard">
              <div className="trendSectionTitle">날짜별 재고 변화</div>
              <svg
                viewBox={`0 0 ${activeTrendData.chartWidth} ${activeTrendData.chartHeight}`}
                className="trendChart"
                role="img"
                aria-label="재고 변화 추이 차트"
              >
                <line
                  x1="30"
                  y1={activeTrendData.chartHeight - 20}
                  x2={activeTrendData.chartWidth - 30}
                  y2={activeTrendData.chartHeight - 20}
                  className="trendAxis"
                />
                <line x1="30" y1="20" x2="30" y2={activeTrendData.chartHeight - 20} className="trendAxis" />
                {activeTrendData.points.length > 1 && (
                  <polyline
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="3"
                    points={activeTrendData.points.map((point) => `${point.x},${point.y}`).join(" ")}
                    className="trendLine"
                  />
                )}
                {activeTrendData.points.map((point, pointIdx) => (
                  <g key={point.dateKey}>
                    <title>{`${point.dateKey} | 재고 ${toFixed(point.qty, 0)}`}</title>
                    <circle cx={point.x} cy={point.y} r="5" className="trendDot" />
                    {activeTrendData.showPointValueLabels && (
                      <text x={point.x} y={point.y - 12} textAnchor="middle" className="trendDotLabel">
                        {toFixed(point.qty, 0)}
                      </text>
                    )}
                    {pointIdx % activeTrendData.xLabelStep === 0 && (
                      <text
                        x={point.x}
                        y={activeTrendData.chartHeight - 2}
                        textAnchor="middle"
                        className="trendXAxisLabel"
                      >
                        {point.dateKey.slice(5)}
                      </text>
                    )}
                  </g>
                ))}
              </svg>
            </div>

            <div className="trendTableCard">
              <div className="trendSectionTitle">날짜별 변화 상세</div>
              <div className="trendTableWrap">
                <table className="trendDetailTable">
                  <thead>
                    <tr>
                      <th>기준일</th>
                      <th>재고</th>
                      <th>직전 데이터 대비</th>
                      <th>증감률</th>
                    </tr>
                  </thead>
                  <tbody>
                    {activeTrendData.series.map((point) => (
                      <tr key={`trend-${point.dateKey}`}>
                        <td>{point.dateKey}</td>
                        <td>{toFixed(point.qty, 0)}</td>
                        <td>{toSigned(point.delta, 0)}</td>
                        <td>{toPercent(point.deltaRate, 1)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      )}

      {showCautionModal && (
        <div className="cautionModalBackdrop" onClick={() => setShowCautionModal(false)}>
          <div className="cautionModal" onClick={(e) => e.stopPropagation()}>
            <div className="cautionModalHeader">
              <div>
                <div className="cautionModalTitle">주의사항</div>
                <div className="cautionModalSubtitle">데이터 형식이나 날짜 기준이 다르면 재고 파악이 정확하지 않을 수 있습니다.</div>
              </div>
              <button
                type="button"
                className="ghost cautionCloseBtn"
                onClick={() => setShowCautionModal(false)}
              >
                닫기
              </button>
            </div>

            <div className="cautionSection">
              <div className="cautionSectionTitle">공통</div>
              <ul className="cautionList">
                <li>같은 이름의 파일은 중복 업로드되지 않습니다.</li>
                <li>필수 컬럼명이나 날짜 형식이 맞지 않으면 재고가 집계되지 않을 수 있습니다.</li>
                <li>재고 비교 탭은 선택한 동일 기준일의 데이터만 비교하며, 없는 값은 대체하지 않습니다.</li>
              </ul>
            </div>

            <div className="cautionSection">
              <div className="cautionSectionTitle">한국 재고</div>
              <ul className="cautionList">
                <li>날짜는 파일명에서 읽습니다. 파일명에 `YYYYMMDD` 8자리 날짜가 포함되어야 합니다.</li>
                <li>재고는 `정상재고` 컬럼 기준으로만 파악합니다.</li>
                <li>한국 탭은 현재고 스냅샷 데이터를 날짜별로 비교하는 방식입니다.</li>
              </ul>
            </div>

            <div className="cautionSection">
              <div className="cautionSectionTitle">해외 재고</div>
              <ul className="cautionList">
                <li>날짜는 파일명 기준이 아니라 파일 내부의 `Date` 컬럼에서 읽습니다.</li>
                <li>재고는 `Quantity` 컬럼 기준으로 집계합니다.</li>
                <li>`한국 현 재고 비교`는 UTC 오늘 날짜의 한국 데이터가 있을 때만 표시됩니다.</li>
              </ul>
            </div>

            <div className="cautionSection">
              <div className="cautionSectionTitle">재고 비교</div>
              <ul className="cautionList">
                <li>선택한 날짜에 특정 국가 데이터가 없으면 `-`로 표시됩니다.</li>
                <li>전날 데이터나 최신 데이터를 임의로 끌어와 대체하지 않습니다.</li>
                <li>의미 있는 비교를 위해 가능한 한 같은 기준일의 국가별 파일을 맞춰 업로드해주세요.</li>
              </ul>
            </div>
          </div>
        </div>
      )}

      {countryTabMode === "SETTINGS" && (
        <section className="settingsPane settingsCard">
          {fileEntries.length > 0 && (
            <div className="settingsTopActions">
              <button
                className="ghost settingsResetBtn"
                disabled={settingsMutating}
                onClick={clearAllFiles}
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
                      disabled={entries.length === 0 || settingsMutating}
                      onClick={async (e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        if (!entries.length) return;
                        await clearFilesByCountry(country, entries);
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
                        disabled={settingsMutating}
                        onClick={() => deleteFileEntry(entry)}
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

