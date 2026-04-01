import { useEffect, useMemo, useRef, useState } from "react";
import axios from "axios";
import * as XLSX from "xlsx";

const API_BASE = import.meta.env.VITE_API_BASE_URL || "http://localhost:8000";
const DEFAULT_DATE_RANGE = "10d";
const OVERSEAS_UPLOAD_COUNTRIES = ["US", "TW", "VN", "SG", "AU", "UK", "AE"];
const SETTINGS_COUNTRY_ORDER = ["KR", "US", "TW", "VN", "SG", "AU", "UK", "AE"];
const SKU_MAPPING_TEMPLATE_COLUMNS = [
  "description",
  "kr",
  "us",
  "tw",
  "vn",
  "sg",
  "au",
  "uk",
  "ae",
];
const EMPTY_SKU_MAPPING_FORM = {
  description: "",
  kr: "",
  us: "",
  tw: "",
  vn: "",
  sg: "",
  au: "",
  uk: "",
  ae: "",
};
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
  if (n.includes("us") || n.includes("usa") || n.includes("미국")) return "US";
  if (n.includes("vn") || n.includes("vietnam") || n.includes("베트남")) return "VN";
  if (n.includes("sg") || n.includes("singapore") || n.includes("싱가포르")) return "SG";
  if (n.includes("au") || n.includes("australia") || n.includes("호주")) return "AU";
  if (n.includes("uk") || n.includes("england") || n.includes("britain") || n.includes("영국")) return "UK";
  if (n.includes("ae") || n.includes("uae") || n.includes("dubai") || n.includes("아랍에미리트")) return "AE";
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

function getRowSku(row) {
  return String(row?.sku ?? row?.itemno ?? "").trim().toUpperCase();
}

function getCanonicalMatchCode(row) {
  const mappedSku = String(row?.mapped_kr_sku || "").trim();
  return mappedSku || getRowSku(row) || extractBaseProductCode(row?.itemno);
}

function getPreferredKrName(row) {
  const mappedName = String(row?.mapped_kr_name || "").trim();
  if (mappedName) return mappedName;
  if (String(row?.country || "KR") === "KR") {
    return String(row?.description || "").trim();
  }
  return "";
}

function getCompareDisplayName(row) {
  const preferred = getPreferredKrName(row);
  if (preferred) return preferred;
  return String(row?.description || "").trim();
}

function getCompareMeta(row) {
  return {
    supplier: String(row?.supplier || "").trim(),
    level: String(row?.level ?? "").trim(),
    warehouse: String(row?.warehouse || row?.category || "").trim(),
  };
}

function getCompareMetaLabel(row) {
  const meta = getCompareMeta(row);
  const parts = [];
  if (meta.supplier) parts.push(meta.supplier);
  if (meta.level) parts.push(`L${meta.level}`);
  if (meta.warehouse) parts.push(meta.warehouse);
  return parts.join(" / ");
}

function getCompareIdentity(row) {
  const code = getCanonicalMatchCode(row);
  const name = getCompareDisplayName(row);
  const meta = getCompareMeta(row);
  if (!code && !name && !meta.supplier && !meta.level && !meta.warehouse) return "";
  return `${code}::${name}::${meta.supplier}::${meta.level}::${meta.warehouse}`;
}

function countryLabel(code = "KR") {
  if (code === "KR") return "한국";
  if (code === "TW") return "대만";
  if (code === "US") return "미국";
  if (code === "VN") return "베트남";
  if (code === "SG") return "싱가포르";
  if (code === "AU") return "호주";
  if (code === "UK") return "영국";
  if (code === "AE") return "아랍에미리트";
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
  const [mappingSearchKeyword, setMappingSearchKeyword] = useState("");
  const [mappingRows, setMappingRows] = useState([]);
  const [mappingRowsLoading, setMappingRowsLoading] = useState(false);
  const [mappingRowsError, setMappingRowsError] = useState("");
  const [mappingSummary, setMappingSummary] = useState({
    total_count: 0,
    updated_at: "",
    upload_updated_at: "",
    manual_updated_at: "",
    required_columns: SKU_MAPPING_TEMPLATE_COLUMNS,
  });
  const [mappingError, setMappingError] = useState("");
  const [manualMappingForm, setManualMappingForm] = useState({ ...EMPTY_SKU_MAPPING_FORM });
  const [mappingInputKey, setMappingInputKey] = useState(0);
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
  const isSettingsScope = countryTabMode === "SETTINGS";
  const isSkuMappingScope = countryTabMode === "SKU_MAPPING";
  const isProductSearchScope = countryTabMode === "PRODUCT_SEARCH";
  const isInventoryAdminScope = isSettingsScope || isSkuMappingScope || isProductSearchScope;
  const inventoryStickyWidth = useMemo(() => {
    if (isKRScope) return 700;
    if (isOverseasScope) return showKrCompare ? 998 : 892;
    return 0;
  }, [isKRScope, isOverseasScope, showKrCompare]);
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
        const item = String(getRowSku(row) ?? "").toLowerCase();
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
      const key = getCompareIdentity(row);
      if (!key) continue;
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
      const key = getCompareIdentity(row);
      if (!key) continue;
      const name = getCompareDisplayName(row);
      if (!map.has(key) && name) map.set(key, name);
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
      warehouseStock: String(row.warehouse || "").trim() || "-",
      trendRowKey: `${row.country}-${getRowSku(row)}-${row.description}-${row.level}-${row.warehouse}-${row.supplier}-${idx}`,
    }));
  }, [isKRScope, filteredRows]);

  const overseasDisplayRows = useMemo(() => {
    if (!isOverseasScope) return [];
    return filteredRows.map((row, idx) => ({
      ...row,
      trendRowKey: `${row.country}-${getRowSku(row)}-${row.description}-${row.level}-${row.warehouse}-${idx}`,
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
    const allKeys = new Set();
    const rowMetaLookup = new Map();
    const byCountry = {};

    const collectCountry = (countryCode) => {
      const scopeKey = countryCode === "KR" ? "KR" : `OVERSEAS:${countryCode}`;
      const cached = scopeResultCache[scopeKey];
      const rows = (cached?.rows || []).filter((row) => String(row.country || "KR") === countryCode);
      const dates = cached?.dates || [];
      const hasDate = Boolean(compareSelectedDate) && dates.includes(compareSelectedDate);
      const qtyMap = new Map();

      for (const row of rows) {
        const code = getCanonicalMatchCode(row);
        const displayName = getCompareDisplayName(row);
        const metaLabel = getCompareMetaLabel(row);
        const compareKey = getCompareIdentity(row);
        if (!compareKey) continue;
        allKeys.add(compareKey);
        if (!rowMetaLookup.has(compareKey)) {
          rowMetaLookup.set(compareKey, { code, name: displayName, metaLabel });
        }
        if (!hasDate) continue;
        qtyMap.set(compareKey, Number(qtyMap.get(compareKey) || 0) + Number(row[compareSelectedDate] || 0));
      }

      byCountry[countryCode] = { hasDate, qtyMap };
    };

    collectCountry("KR");
    OVERSEAS_UPLOAD_COUNTRIES.forEach(collectCountry);

    return {
      rowKeys: Array.from(allKeys).sort((a, b) => {
        const aMeta = rowMetaLookup.get(a) || {};
        const bMeta = rowMetaLookup.get(b) || {};
        return (
          String(aMeta.code || "").localeCompare(String(bMeta.code || "")) ||
          String(aMeta.name || "").localeCompare(String(bMeta.name || ""))
        );
      }),
      rowMetaLookup,
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
    if (!compareSelectedDate || !compareCountryData.rowKeys.length) return [];
    return compareCountryData.rowKeys.map((rowKey) => {
      const meta = compareCountryData.rowMetaLookup.get(rowKey) || {};
      const row = {
        _compareKey: rowKey,
        상품코드: meta.code || "",
        한국상품명: meta.name || "",
        구분: meta.metaLabel || "",
        "한국 현 재고": compareCountryData.byCountry.KR?.hasDate
          ? Number(compareCountryData.byCountry.KR.qtyMap.get(rowKey) || 0)
          : null,
      };
      for (const countryCode of OVERSEAS_UPLOAD_COUNTRIES) {
        row[countryLabel(countryCode)] = compareCountryData.byCountry[countryCode]?.hasDate
          ? Number(compareCountryData.byCountry[countryCode].qtyMap.get(rowKey) || 0)
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
      const meta = String(row["구분"] || "").toLowerCase();
      return code.includes(needle) || krName.includes(needle) || meta.includes(needle);
    });
  }, [compareRows, inventoryKeyword, isCompareScope]);
  const tableMinWidth = useMemo(() => {
    const dateCols = filteredDateColumns.length * 88;
    const overseasFixedCols = showKrCompare ? 1002 : 896;
    const fixedCols = isOverseasScope ? overseasFixedCols : 844;
    return Math.max(980, fixedCols + dateCols);
  }, [filteredDateColumns, isOverseasScope, showKrCompare]);
  const inventoryTableWidth = useMemo(
    () => (isKRScope ? Math.max(980, 844 + filteredDateColumns.length * 88) : tableMinWidth),
    [isKRScope, filteredDateColumns, tableMinWidth]
  );
  const compareTableWidth = useMemo(
    () => Math.max(1280, 590 + (OVERSEAS_UPLOAD_COUNTRIES.length + 1) * 96),
    []
  );
  const datePeekFadeStyle =
    !isCompareScope && inventoryStickyWidth > 0 && filteredDateColumns.length > 0
      ? { "--date-peek-fade-left": `${inventoryStickyWidth + 2}px` }
      : undefined;

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
    inventoryStickyWidth,
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
        <th className="compareMetaCol">구분</th>
        <th className="compareCountryCol krCompareCol">한국</th>
        {OVERSEAS_UPLOAD_COUNTRIES.map((code) => (
          <th key={code} className="compareCountryCol">{countryLabel(code)}</th>
        ))}
      </>
    );
  }

  async function aggregateInventory() {
    if (isInventoryAdminScope || countryTabMode === "COMPARE") return;
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
          구분: row["구분"] || "",
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
          상품코드: getRowSku(row),
          상품명: row.description || "",
          창고: row.warehouse || "",
          레벨: row.level || "",
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
      const matchCode = getCanonicalMatchCode(row);
      const compareKey = getCompareIdentity(row);
      const out = {
        상품코드: matchCode,
        상품명: row.description || "",
        한국상품명: getCompareDisplayName(row) || krNameMap.get(compareKey) || "",
        공급처: row.supplier || "",
        창고: row.warehouse || "",
        레벨: row.level || "",
      };
      if (showKrCompare) {
        out["한국"] = krCompareMap.size
          ? Number(krCompareMap.get(compareKey) || 0)
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
    if (isInventoryAdminScope || countryTabMode === "COMPARE") return;
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

  async function fetchSkuMappingSummary() {
    const res = await axios.get(`${API_BASE}/api/inventory/mappings`);
    return {
      total_count: Number(res?.data?.total_count || 0),
      updated_at: res?.data?.updated_at || "",
      upload_updated_at: res?.data?.upload_updated_at || "",
      manual_updated_at: res?.data?.manual_updated_at || "",
      required_columns: Array.isArray(res?.data?.required_columns)
        ? res.data.required_columns
        : SKU_MAPPING_TEMPLATE_COLUMNS,
    };
  }

  async function fetchSkuMappingItems(query = "") {
    const res = await axios.get(`${API_BASE}/api/inventory/mappings/items`, {
      params: { query: query || "", limit: 200 },
    });
    return Array.isArray(res?.data?.items) ? res.data.items : [];
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
    const [persistedFiles, latestMappingSummary, ...views] = await Promise.all([
      fetchPersistedFiles(),
      fetchSkuMappingSummary(),
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
    setMappingSummary(latestMappingSummary);
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

  useEffect(() => {
    if (!isProductSearchScope) return;
    if (!mappingSearchKeyword.trim()) {
      setMappingRows([]);
      setMappingRowsError("");
      setMappingRowsLoading(false);
      return;
    }
    const run = async () => {
      try {
        setMappingRowsLoading(true);
        setMappingRowsError("");
        const items = await fetchSkuMappingItems(mappingSearchKeyword);
        setMappingRows(items);
      } catch (err) {
        const detail = err?.response?.data?.detail;
        setMappingRowsError(Array.isArray(detail) ? detail.join("\n") : detail || "매핑 검색 중 오류");
      } finally {
        setMappingRowsLoading(false);
      }
    };
    run();
  }, [isProductSearchScope, mappingSearchKeyword]);

  function openSkuMappingInput() {
    document.getElementById("sku-mapping-input")?.click();
  }

  async function uploadSkuMappingFile(file) {
    if (!file) return;
    try {
      setSettingsMutating(true);
      setMappingError("");
      const formData = new FormData();
      formData.append("file", file);
      const res = await axios.post(`${API_BASE}/api/inventory/mappings/upload`, formData, {
        headers: { "Content-Type": "multipart/form-data" },
      });
      setMappingSummary({
        total_count: Number(res?.data?.total_count || 0),
        updated_at: res?.data?.updated_at || "",
        upload_updated_at: res?.data?.upload_updated_at || "",
        manual_updated_at: res?.data?.manual_updated_at || "",
        required_columns: Array.isArray(res?.data?.required_columns)
          ? res.data.required_columns
          : SKU_MAPPING_TEMPLATE_COLUMNS,
      });
      setMappingInputKey((prev) => prev + 1);
      await hydratePersistedState();
      window.alert(`${Number(res?.data?.replaced_count || 0)}건의 SKU 매핑을 반영했습니다.`);
    } catch (err) {
      const detail = err?.response?.data?.detail;
      const msg = Array.isArray(detail) ? detail.join("\n") : detail || "SKU 매핑 업로드 중 오류";
      setMappingError(msg);
      window.alert(msg);
    } finally {
      setSettingsMutating(false);
    }
  }

  async function clearSkuMappings() {
    if (!window.confirm("등록된 SKU 매핑 마스터를 모두 삭제할까요?")) return;
    try {
      setSettingsMutating(true);
      setMappingError("");
      await axios.delete(`${API_BASE}/api/inventory/mappings`);
      await hydratePersistedState();
    } catch (err) {
      const detail = err?.response?.data?.detail;
      const msg = Array.isArray(detail) ? detail.join("\n") : detail || "SKU 매핑 초기화 중 오류";
      setMappingError(msg);
      window.alert(msg);
    } finally {
      setSettingsMutating(false);
    }
  }

  async function saveManualSkuMapping() {
    try {
      setSettingsMutating(true);
      setMappingError("");
      const payload = Object.fromEntries(
        Object.entries(manualMappingForm).map(([key, value]) => [key, String(value || "").trim()])
      );
      await axios.post(`${API_BASE}/api/inventory/mappings/item`, payload);
      setManualMappingForm({ ...EMPTY_SKU_MAPPING_FORM });
      await hydratePersistedState();
      window.alert("SKU 정보를 저장했습니다.");
    } catch (err) {
      const detail = err?.response?.data?.detail;
      const msg = Array.isArray(detail) ? detail.join("\n") : detail || "SKU 수기 저장 중 오류";
      setMappingError(msg);
      window.alert(msg);
    } finally {
      setSettingsMutating(false);
    }
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
              disabled={isInventoryAdminScope || countryTabMode === "COMPARE"}
            >
              파일 업로드
            </button>
            <button
              className="primary"
              onClick={aggregateInventory}
              disabled={
                isInventoryAdminScope ||
                countryTabMode === "COMPARE" ||
                inventoryLoading ||
                inventoryFiles.length === 0
              }
            >
              {inventoryLoading ? "통합 중..." : "재고 통합 실행"}
            </button>
            <button className="ghost" onClick={exportCurrentView} disabled={isInventoryAdminScope}>
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
            className={`tab ${countryTabMode === "PRODUCT_SEARCH" ? "active" : ""}`}
            onClick={() => setCountryTabMode("PRODUCT_SEARCH")}
          >
            상품 매핑
          </button>
          <button
            className={`tab ${countryTabMode === "SETTINGS" ? "active" : ""}`}
            onClick={() => setCountryTabMode("SETTINGS")}
          >
            데이터 관리
          </button>
          <button
            className={`tab ${countryTabMode === "SKU_MAPPING" ? "active" : ""}`}
            onClick={() => setCountryTabMode("SKU_MAPPING")}
          >
            SKU 관리
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
      {!isInventoryAdminScope && !isCompareScope && (
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
                className={`tableHeaderScroll ${datePeekFadeStyle ? "withDatePeekFade" : ""}`}
                style={datePeekFadeStyle}
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
              className={`tableWrap ${datePeekFadeStyle ? "withDatePeekFade" : ""}`}
              style={datePeekFadeStyle}
              onScroll={() => syncScroll("table")}
            >
            <table
              className={`inventoryTable bodyTable ${isKRScope ? "krTable" : "overseasTable"}`}
              style={{ minWidth: inventoryTableWidth }}
            >
              <tbody>
                {(isKRScope ? krDisplayRows : isOverseasScope ? overseasDisplayRows : filteredRows).map((row, idx) => (
                  <tr key={row.trendRowKey || `${row.country}-${getRowSku(row)}-${row.description}-${row.level}-${row.warehouse}-${idx}`}>
                    {isKRScope && <td className="stickyCol stickyColSupplier">{row.supplier}</td>}
                    <td className="stickyCol stickyColCode">
                      {isOverseasScope ? getCanonicalMatchCode(row) : getRowSku(row)}
                    </td>
                    <td className="stickyCol stickyColName stickyColBoundary">
                      <span className="nameCellText">{row.description || "(상품명 없음)"}</span>
                    </td>
                    {isOverseasScope && (
                      <td className="stickyCol stickyColKrName stickyColBoundary">
                        <span className="nameCellText">
                          {getCompareDisplayName(row) || krNameMap.get(getCompareIdentity(row)) || "-"}
                        </span>
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
                          ? toFixed(krCompareMap.get(getCompareIdentity(row)) || 0, 0)
                          : "-"}
                      </td>
                    )}
                    {isKRScope
                      ? filteredDateColumns.map((dt) => (
                          <td key={`${getRowSku(row)}-${row.description}-${row.level}-${row.warehouse}-${dt}`} className="dateCol">{toFixed(row[dt], 0)}</td>
                        ))
                      : filteredDateColumns.map((dt) => (
                        <td key={`${getRowSku(row)}-${row.description}-${row.level}-${row.warehouse}-${dt}`} className="dateCol">{toFixed(row[dt], 0)}</td>
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
                    <div style={{ width: topScrollWidth || compareTableWidth }} />
                  </div>
                )}
              <div className="stickyTableHeader" style={{ top: tableHeaderTop }}>
                <div
                  ref={headerScrollRef}
                  className="tableHeaderScroll"
                  onScroll={() => syncScroll("header")}
                >
                  <table className="compareTable stickyHeaderTable" style={{ minWidth: compareTableWidth }}>
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
                <table className="compareTable bodyTable" style={{ minWidth: compareTableWidth }}>
                  <tbody>
                    {filteredCompareRows.map((row) => (
                      <tr key={`compare-${row._compareKey}`}>
                        <td className="compareCodeCol stickyCol stickyColCode">{row["상품코드"]}</td>
                        <td className="compareNameCol stickyCol stickyColName stickyColBoundary">
                          <span className="nameCellText">{row["한국상품명"] || "-"}</span>
                        </td>
                        <td className="compareMetaCol">{row["구분"] || "-"}</td>
                        <td className="compareCountryCol krCompareCol">
                          {row["한국 현 재고"] === null ? "-" : toFixed(row["한국 현 재고"], 0)}
                        </td>
                        {OVERSEAS_UPLOAD_COUNTRIES.map((code) => (
                          <td key={`${row._compareKey}-${code}`} className="compareCountryCol">
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
                  {getRowSku(activeTrendRow)} · {activeTrendRow.description || "(상품명 없음)"}
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

            <div className="cautionSection">
              <div className="cautionSectionTitle">SKU 매핑 마스터</div>
              <ul className="cautionList">
                <li>매핑 마스터는 `.xlsx` 파일만 업로드할 수 있습니다.</li>
                <li>필수 헤더는 `{SKU_MAPPING_TEMPLATE_COLUMNS.join("`, `")}` 입니다.</li>
                <li>헤더 순서가 다르거나 `description` 또는 국가별 SKU가 중복되면 전체 업로드가 거절됩니다.</li>
                <li>원본 파일은 저장하지 않고, 읽은 매핑 데이터만 DB에 반영합니다.</li>
              </ul>
            </div>
          </div>
        </div>
      )}

      {isSettingsScope && (
        <section className="settingsPane settingsCard">
          {fileEntries.length === 0 && (
            <pre className="error settingsNotice">
              업로드된 파일이 없습니다. 상단의 파일 업로드 버튼을 눌러주세요.
            </pre>
          )}

          <div className={`settingsToolbar ${fileEntries.length === 0 ? "settingsToolbarWithNotice" : ""}`}>
            <button
              className="ghost settingsDangerButton"
              disabled={settingsMutating || fileEntries.length === 0}
              onClick={clearAllFiles}
            >
              전체 초기화
            </button>
          </div>

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

      {isSkuMappingScope && (
        <section className="settingsPane settingsCard">
          <div className="skuManageCard">
            <div className="skuManageHeader">
              <div className="skuManageSubtitle">
                SKU 마스터 파일을 업로드하면 SKU 정보가 업데이트 됩니다. 파일은 저장되지 않습니다.
              </div>
              <button
                type="button"
                className="ghost skuManageResetButton"
                disabled={settingsMutating || mappingSummary.total_count === 0}
                onClick={clearSkuMappings}
              >
                매핑 초기화
              </button>
            </div>
            <input
              key={mappingInputKey}
              id="sku-mapping-input"
              type="file"
              accept=".xlsx"
              style={{ display: "none" }}
              onChange={(e) => uploadSkuMappingFile(e.target.files?.[0] || null)}
            />
            <div className="skuManageContent">
              <button
                type="button"
                className="skuUploadPanel"
                disabled={settingsMutating}
                onClick={openSkuMappingInput}
              >
                <span className="skuUploadMain">
                  <span className="skuUploadBadge">XLSX</span>
                  <span className="skuUploadButtonLabel">{settingsMutating ? "업로드 중..." : "파일 업로드"}</span>
                </span>
                <span className="skuUploadMeta">
                  파일 업로드 최근 반영일:{" "}
                  {mappingSummary.upload_updated_at
                    ? mappingSummary.upload_updated_at.replace("T", " ").slice(0, 19)
                    : "-"}
                </span>
              </button>
              <div className="skuManualCard">
                <div className="skuManualTitle">수기 SKU 입력</div>
                <div className="skuManualSubtitle">
                  파일 업로드가 어려운 경우 아래 칸에 직접 입력해서 SKU 정보를 갱신할 수 있습니다.
                </div>
                <div className="skuManualGrid">
                  <label className="skuField">
                    <span>상품명</span>
                    <input
                      type="text"
                      value={manualMappingForm.description}
                      onChange={(e) => setManualMappingForm((prev) => ({ ...prev, description: e.target.value }))}
                      placeholder="한국 상품명"
                    />
                  </label>
                  <label className="skuField">
                    <span>한국 SKU</span>
                    <input
                      type="text"
                      value={manualMappingForm.kr}
                      onChange={(e) => setManualMappingForm((prev) => ({ ...prev, kr: e.target.value }))}
                      placeholder="필수"
                    />
                  </label>
                  <label className="skuField">
                    <span>미국 SKU</span>
                    <input
                      type="text"
                      value={manualMappingForm.us}
                      onChange={(e) => setManualMappingForm((prev) => ({ ...prev, us: e.target.value }))}
                    />
                  </label>
                  <label className="skuField">
                    <span>대만 SKU</span>
                    <input
                      type="text"
                      value={manualMappingForm.tw}
                      onChange={(e) => setManualMappingForm((prev) => ({ ...prev, tw: e.target.value }))}
                    />
                  </label>
                  <label className="skuField">
                    <span>베트남 SKU</span>
                    <input
                      type="text"
                      value={manualMappingForm.vn}
                      onChange={(e) => setManualMappingForm((prev) => ({ ...prev, vn: e.target.value }))}
                    />
                  </label>
                  <label className="skuField">
                    <span>싱가포르 SKU</span>
                    <input
                      type="text"
                      value={manualMappingForm.sg}
                      onChange={(e) => setManualMappingForm((prev) => ({ ...prev, sg: e.target.value }))}
                    />
                  </label>
                  <label className="skuField">
                    <span>호주 SKU</span>
                    <input
                      type="text"
                      value={manualMappingForm.au}
                      onChange={(e) => setManualMappingForm((prev) => ({ ...prev, au: e.target.value }))}
                    />
                  </label>
                  <label className="skuField">
                    <span>영국 SKU</span>
                    <input
                      type="text"
                      value={manualMappingForm.uk}
                      onChange={(e) => setManualMappingForm((prev) => ({ ...prev, uk: e.target.value }))}
                    />
                  </label>
                  <label className="skuField">
                    <span>아랍에미리트 SKU</span>
                    <input
                      type="text"
                      value={manualMappingForm.ae}
                      onChange={(e) => setManualMappingForm((prev) => ({ ...prev, ae: e.target.value }))}
                    />
                  </label>
                </div>
                <div className="skuManualActions">
                  <div className="skuManageMeta">
                    수기 입력 최근 반영일:{" "}
                    {mappingSummary.manual_updated_at
                      ? mappingSummary.manual_updated_at.replace("T", " ").slice(0, 19)
                      : "-"}
                  </div>
                  <div className="skuManualButtons">
                    <button type="button" className="primary" disabled={settingsMutating} onClick={saveManualSkuMapping}>
                      저장
                    </button>
                  </div>
                </div>
              </div>
            </div>
            {mappingError && <pre className="error mappingError">{mappingError}</pre>}
          </div>
        </section>
      )}

      {isProductSearchScope && (
        <section className="settingsPane settingsCard">
          <section className="tableCard searchPanelCard">
            <div className="mappingSearchHead">
              <div>
                <div className="mappingCardTitle">상품 매핑</div>
                <div className="mappingCardSubtitle">
                  SKU 또는 한국상품명을 검색하면 국가별 SKU 매핑을 확인할 수 있습니다.
                </div>
              </div>
            </div>
            <div className="filterBar">
              <div className="searchWrap">
                <span className="searchIcon" aria-hidden="true">
                  🔍
                </span>
                <input
                  className="searchInput"
                  type="text"
                  value={mappingSearchKeyword}
                  onChange={(e) => setMappingSearchKeyword(e.target.value)}
                  placeholder="한국상품명 또는 SKU 검색..."
                />
              </div>
            </div>
            {mappingRowsError && <pre className="error">{mappingRowsError}</pre>}
            {mappingRowsLoading ? (
              <div className="searchEmptyState">검색 중...</div>
            ) : !mappingSearchKeyword.trim() ? (
              <div className="searchEmptyState">SKU 또는 한국상품명을 입력하면 매핑 결과가 표시됩니다.</div>
            ) : !mappingRows.length ? (
              <div className="searchEmptyState">일치하는 매핑 결과가 없습니다.</div>
            ) : (
              <div className="tableWrap">
                <table className="searchTable">
                  <thead>
                    <tr>
                      <th>상품명</th>
                      <th>한국</th>
                      <th>미국</th>
                      <th>대만</th>
                      <th>베트남</th>
                      <th>싱가포르</th>
                      <th>호주</th>
                      <th>영국</th>
                      <th>아랍에미리트</th>
                    </tr>
                  </thead>
                  <tbody>
                    {mappingRows.map((row, idx) => (
                      <tr key={`${row.description}-${row.kr}-${idx}`}>
                        <td>{row.description || "-"}</td>
                        <td>{row.kr || "-"}</td>
                        <td>{row.us || "-"}</td>
                        <td>{row.tw || "-"}</td>
                        <td>{row.vn || "-"}</td>
                        <td>{row.sg || "-"}</td>
                        <td>{row.au || "-"}</td>
                        <td>{row.uk || "-"}</td>
                        <td>{row.ae || "-"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </section>
        </section>
      )}
    </div>
    </div>
  );
}

