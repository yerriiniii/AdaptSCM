import { Fragment, useEffect, useMemo, useRef, useState } from "react";
import axios from "axios";
import * as XLSX from "xlsx";

const API_BASE = import.meta.env.VITE_API_BASE_URL || "http://localhost:8000";
const DEFAULT_DATE_RANGE = "10d";
const OVERSEAS_UPLOAD_COUNTRIES = ["US", "TW", "HK", "VN", "SG", "AU", "UK", "AE"];
const SETTINGS_COUNTRY_ORDER = ["KR", "US", "TW", "HK", "VN", "SG", "AU", "UK", "AE"];
const SKU_MAPPING_FIELDS = [
  { code: "KR", label: "한국", nameKey: "kr_name", skuKey: "kr_sku" },
  { code: "US", label: "미국", nameKey: "us_name", skuKey: "us_sku" },
  { code: "TW", label: "대만", nameKey: "tw_name", skuKey: "tw_sku" },
  { code: "HK", label: "홍콩", nameKey: "hk_name", skuKey: "hk_sku" },
  { code: "VN", label: "베트남", nameKey: "vn_name", skuKey: "vn_sku" },
  { code: "SG", label: "싱가포르", nameKey: "sg_name", skuKey: "sg_sku" },
  { code: "AU", label: "호주", nameKey: "au_name", skuKey: "au_sku" },
  { code: "UK", label: "영국", nameKey: "uk_name", skuKey: "uk_sku" },
  { code: "AE", label: "아랍에미리트", nameKey: "ae_name", skuKey: "ae_sku" },
];
const SKU_MAPPING_TEMPLATE_COLUMNS = SKU_MAPPING_FIELDS.flatMap(({ nameKey, skuKey }) => [nameKey, skuKey]);
const SKU_MAPPING_OPTIONAL_COLUMNS = ["option", "brand"];
const EMPTY_SKU_MAPPING_FORM = Object.fromEntries([
  ...SKU_MAPPING_FIELDS.flatMap(({ nameKey, skuKey }) => [
    [nameKey, ""],
    [skuKey, ""],
  ]),
  ["brand", ""],
  ["option", ""],
]);
const PRODUCT_MAPPING_SEARCH_CHIPS = SKU_MAPPING_FIELDS.map(({ code }) => code);
const SKU_MAPPING_OVERSEAS_FIELDS = SKU_MAPPING_FIELDS.filter(({ code }) => code !== "KR");

/** 수기 입력 브랜드 드롭다운 (순서 유지) */
const MANUAL_BRAND_PRESETS = Object.freeze([
  "8APM",
  "95PROBLEM",
  "cs",
  "SC08",
  "글라센",
  "뉴트리스토리",
  "데이알",
  "듀오렉신",
  "랍셍스",
  "레이어",
  "레이어옵티컬",
  "리이오",
  "림트",
  "마시밀레",
  "매그드레인",
  "먼로우",
  "블랙홀",
  "슈럭",
  "스칸디대디",
  "스킨빌더스",
  "알브단스",
  "에스마켓",
  "에이페",
  "엠마녹스",
  "오브제",
  "옵스테드",
  "자연미식",
  "컬러풀선데이",
  "클릭앤블락",
  "페닐롭",
  "페트리스",
  "푸드올로지",
  "풀리",
  "플릭",
  "필린",
]);

function normalizeManualBrandSearch(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "");
}

function ManualBrandCombobox({ value, onChange, disabled }) {
  const rootRef = useRef(null);
  const searchInputRef = useRef(null);
  const customInputRef = useRef(null);
  const presetSet = useMemo(() => new Set(MANUAL_BRAND_PRESETS), []);
  const [open, setOpen] = useState(false);
  const [search, setSearch] = useState("");
  /** 빈 칸에서 「직접 입력」으로 들어온 상태(값을 지워도 목록으로 보내기 전까지 유지) */
  const [customMode, setCustomMode] = useState(() => Boolean(value && !presetSet.has(value)));
  /** 직접 입력 중에도 프리셋 목록 UI로 전환 */
  const [pickFromListUi, setPickFromListUi] = useState(false);

  const isPresetValue = Boolean(value && presetSet.has(value));
  const showCustomRow =
    !pickFromListUi && (customMode || (Boolean(value) && !presetSet.has(value)));

  useEffect(() => {
    if (isPresetValue) {
      setCustomMode(false);
      setPickFromListUi(false);
    }
  }, [isPresetValue]);

  useEffect(() => {
    if (!open) return;
    const onDoc = (e) => {
      if (rootRef.current && !rootRef.current.contains(e.target)) {
        setOpen(false);
        setSearch("");
      }
    };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [open]);

  useEffect(() => {
    if (open) searchInputRef.current?.focus();
  }, [open]);

  useEffect(() => {
    if (showCustomRow && customMode) customInputRef.current?.focus();
  }, [showCustomRow, customMode]);

  const filteredPresets = useMemo(() => {
    const q = normalizeManualBrandSearch(search);
    if (!q) return [...MANUAL_BRAND_PRESETS];
    return MANUAL_BRAND_PRESETS.filter((b) => normalizeManualBrandSearch(b).includes(q));
  }, [search]);

  if (showCustomRow) {
    return (
      <div className="manualBrandCombobox manualBrandComboboxCustom" ref={rootRef}>
        <div className="manualBrandCustomRow">
          <input
            ref={customInputRef}
            type="text"
            className="manualBrandCustomInput"
            value={value}
            onChange={(e) => onChange(e.target.value)}
            disabled={disabled}
            placeholder="브랜드 직접 입력"
            autoComplete="off"
          />
          <button
            type="button"
            className="manualBrandToPresetsBtn"
            disabled={disabled}
            onClick={() => setPickFromListUi(true)}
          >
            목록에서 선택
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className={`manualBrandCombobox${open ? " manualBrandComboboxOpen" : ""}`} ref={rootRef}>
      <button
        type="button"
        className="manualBrandTrigger"
        disabled={disabled}
        onClick={() => setOpen((o) => !o)}
        aria-expanded={open}
        aria-haspopup="listbox"
      >
        <span className={`manualBrandTriggerText${value ? "" : " isPlaceholder"}`}>
          {value || "브랜드 선택"}
        </span>
        <span className="manualBrandChevron" aria-hidden="true">
          ▼
        </span>
      </button>
      {open ? (
        <div className="manualBrandPopover" role="listbox">
          <input
            ref={searchInputRef}
            type="text"
            className="manualBrandSearch"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="검색…"
            autoComplete="off"
            onMouseDown={(e) => e.stopPropagation()}
          />
          <ul className="manualBrandList">
            {filteredPresets.length ? (
              filteredPresets.map((b) => (
                <li key={b}>
                  <button
                    type="button"
                    className="manualBrandOption"
                    onClick={() => {
                      onChange(b);
                      setOpen(false);
                      setSearch("");
                      setCustomMode(false);
                      setPickFromListUi(false);
                    }}
                  >
                    {b}
                  </button>
                </li>
              ))
            ) : (
              <li className="manualBrandListEmpty">일치하는 브랜드가 없습니다.</li>
            )}
          </ul>
          <button
            type="button"
            className="manualBrandOption manualBrandOptionDirect"
            onClick={() => {
              setOpen(false);
              setSearch("");
              setPickFromListUi(false);
              setCustomMode(true);
              const keepCustom = Boolean(value && !presetSet.has(value));
              if (!keepCustom) onChange("");
            }}
          >
            직접 입력…
          </button>
        </div>
      ) : null}
    </div>
  );
}

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
  if (n.includes("hongkong") || n.includes("hong kong") || n.includes("香港") || n.includes("홍콩")) return "HK";
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

/** 한국 재고 '오늘' 스냅샷 판별용 — UTC가 아니라 Asia/Seoul 달력 날짜(YYYY-MM-DD) */
function getKoreaDateKey(date = new Date()) {
  return date.toLocaleDateString("sv-SE", { timeZone: "Asia/Seoul" });
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
  const mappedSku = String(row?.mapped_kr_sku || "").trim().toUpperCase();
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
  if (code === "HK") return "홍콩";
  if (code === "US") return "미국";
  if (code === "VN") return "베트남";
  if (code === "SG") return "싱가포르";
  if (code === "AU") return "호주";
  if (code === "UK") return "영국";
  if (code === "AE") return "아랍에미리트";
  return code;
}

function getProductMappingCountries(row = {}) {
  const locales = Array.isArray(row?.locales) ? row.locales : [];
  return locales
    .map((locale) => {
      const code = String(locale?.country_code || locale?.countryCode || "").trim().toUpperCase();
      const sku = String(locale?.sku ?? "").trim();
      const name = String(locale?.name ?? "").trim();
      if (!code) return null;
      // SKU 없이 이름만 있는 로케일도 표시(백엔드·데이터에 따라 빈 SKU가 올 수 있음)
      if (!sku && !name) return null;
      return {
        code,
        label: countryLabel(code),
        sku: sku || "—",
        description: name || "—",
      };
    })
    .filter(Boolean);
}

function getLatestDateKey(dateKeys = []) {
  if (!dateKeys.length) return "";
  return [...dateKeys].sort().at(-1) || "";
}

/** hydratePersistedState 실패 시 — 네트워크/DB/500 구분에 도움 */
function formatPersistedLoadError(err, apiBase) {
  const detail = err?.response?.data?.detail;
  const fromApi = Array.isArray(detail) ? detail.join("\n") : detail != null ? String(detail) : "";
  const status = err?.response?.status;
  const code = err?.code;
  const message = String(err?.message || "");

  if (code === "ERR_NETWORK" || message === "Network Error") {
    return [
      "백엔드 API에 연결할 수 없습니다.",
      `요청 기준 URL: ${apiBase}`,
      "백엔드(uvicorn) 실행 여부, VITE_API_BASE_URL, 방화벽·VPN을 확인하세요.",
    ].join("\n");
  }

  if (fromApi) return fromApi;
  const parts = [];
  if (status) parts.push(`HTTP ${status}`);
  if (message) parts.push(message);
  return parts.filter(Boolean).join(" · ") || "저장된 데이터를 불러오는 중 오류";
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
  const [inventoryLevelFilter, setInventoryLevelFilter] = useState("all");
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
    optional_columns: SKU_MAPPING_OPTIONAL_COLUMNS,
  });
  const [mappingError, setMappingError] = useState("");
  const mappingErrorRef = useRef(null);
  const [manualMappingForm, setManualMappingForm] = useState({ ...EMPTY_SKU_MAPPING_FORM });
  const [manualSkuFormKey, setManualSkuFormKey] = useState(0);
  const [mappingInputKey, setMappingInputKey] = useState(0);
  const [skuManageMode, setSkuManageMode] = useState("UPLOAD");

  useEffect(() => {
    if (!mappingError) return;
    mappingErrorRef.current?.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }, [mappingError]);

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
  const productMappingCards = useMemo(() => {
    const cards = mappingRows
      .map((row, idx) => ({
        ...row,
        _id: row.group_id || `${row.kr_name || "mapping"}-${idx}`,
        _countries: getProductMappingCountries(row),
      }))
      .filter((row) => row._countries.length > 0);
    // 검색 전에만: 국가가 많은 그룹 우선(백엔드도 동일 정렬). 검색 중에는 API 응답 순서 유지
    if (!mappingSearchKeyword.trim()) {
      cards.sort((a, b) => {
        const byLen = b._countries.length - a._countries.length;
        if (byLen !== 0) return byLen;
        return String(a.kr_name || "").localeCompare(String(b.kr_name || ""), "ko");
      });
    }
    return cards;
  }, [mappingRows, mappingSearchKeyword]);

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
    if (isKRScope) return 680;
    if (isOverseasScope) return showKrCompare ? 960 : 852;
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
        setInventoryError(formatPersistedLoadError(err, API_BASE));
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

  const availableInventoryLevels = useMemo(() => {
    if (isKRScope) return [];
    const values = Array.from(
      new Set(
        rawInventoryRows
          .filter((row) => matchesCountryScope(row.country))
          .map((row) => String(row.level ?? "").trim())
          .filter(Boolean)
      )
    );
    return values.sort((a, b) => Number(a) - Number(b) || a.localeCompare(b));
  }, [rawInventoryRows, selectedOverseasCountry, countryTabMode, isKRScope]);
  const hasInventoryLevels = availableInventoryLevels.length > 0;
  const defaultInventoryLevelFilter = availableInventoryLevels[0] || "all";

  const filteredRows = useMemo(() => {
    if (!rawInventoryRows.length) return [];
    const needle = inventoryKeyword.trim().toLowerCase();

    return rawInventoryRows.filter((row) => {
      if (!matchesCountryScope(row.country)) return false;

      if (!isKRScope && hasInventoryLevels && inventoryLevelFilter !== "all") {
        const lv = (String(row.level ?? "").replace(/^0+/, "") || "0").trim();
        if (lv !== inventoryLevelFilter) return false;
      }
      if (needle) {
        const item = String(getRowSku(row) ?? "").toLowerCase();
        const desc = String(row.description ?? "").toLowerCase();
        const krSku = String(row.mapped_kr_sku ?? "").toLowerCase();
        if (!item.includes(needle) && !desc.includes(needle) && !krSku.includes(needle)) return false;
      }
      return true;
    });
  }, [
    rawInventoryRows,
    inventoryKeyword,
    inventoryLevelFilter,
    hasInventoryLevels,
    countryTabMode,
    selectedOverseasCountry,
    isKRScope,
  ]);

  useEffect(() => {
    if (isKRScope || !hasInventoryLevels) {
      if (inventoryLevelFilter !== "all") setInventoryLevelFilter("all");
      return;
    }
    if (!availableInventoryLevels.includes(inventoryLevelFilter)) {
      setInventoryLevelFilter(defaultInventoryLevelFilter);
    }
  }, [
    isKRScope,
    hasInventoryLevels,
    inventoryLevelFilter,
    availableInventoryLevels,
    defaultInventoryLevelFilter,
    currentScopeKey,
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

  /** 해외 행과 맞출 때는 공급처/창고/레벨이 국가마다 달라 getCompareIdentity로는 키가 안 맞음 → 한국 SKU(매핑 우선) 기준으로 합산 */
  const krCompareMap = useMemo(() => {
    const krRows = (scopeResultCache.KR?.rows || []).filter((row) => String(row.country || "KR") === "KR");
    const krDates = scopeResultCache.KR?.dates || [];
    if (!krRows.length || !krDates.length) return new Map();
    const koreaTodayKey = getKoreaDateKey();
    const hasKrCurrentSnapshot = krDates.includes(koreaTodayKey);
    if (!hasKrCurrentSnapshot) return new Map();
    const map = new Map();
    for (const row of krRows) {
      const key = getCanonicalMatchCode(row);
      if (!key) continue;
      const qty = Number(row[koreaTodayKey] || 0);
      map.set(key, Number(map.get(key) || 0) + qty);
    }
    return map;
  }, [scopeResultCache]);

  const hasTodayKrSnapshot = useMemo(() => {
    const koreaTodayKey = getKoreaDateKey();
    return (scopeResultCache.KR?.dates || []).includes(koreaTodayKey);
  }, [scopeResultCache]);

  const krNameMap = useMemo(() => {
    const map = new Map();
    for (const row of scopeResultCache.KR?.rows || []) {
      if (String(row.country || "KR") !== "KR") continue;
      const key = getCanonicalMatchCode(row);
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
    const overseasFixedCols = showKrCompare ? 964 : 856;
    const fixedCols = isOverseasScope ? overseasFixedCols : 824;
    return Math.max(980, fixedCols + dateCols);
  }, [filteredDateColumns, isOverseasScope, showKrCompare]);
  const inventoryTableWidth = useMemo(
    () => (isKRScope ? Math.max(980, 824 + filteredDateColumns.length * 88) : tableMinWidth),
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
      const krMatchKey = getCanonicalMatchCode(row);
      const out = {
        상품코드: getRowSku(row),
        상품명: row.description || "",
        한국상품명: getCompareDisplayName(row) || krNameMap.get(krMatchKey) || "",
        공급처: row.supplier || "",
        창고: row.warehouse || "",
        레벨: row.level || "",
      };
      if (showKrCompare) {
        out["한국"] = krCompareMap.size
          ? Number(krCompareMap.get(krMatchKey) || 0)
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
      optional_columns: Array.isArray(res?.data?.optional_columns)
        ? res.data.optional_columns
        : SKU_MAPPING_OPTIONAL_COLUMNS,
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
    const run = async () => {
      try {
        setMappingRowsLoading(true);
        setMappingRowsError("");
        const items = await fetchSkuMappingItems(mappingSearchKeyword.trim());
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

  async function uploadSkuMappingFiles(files) {
    const uploadFiles = Array.from(files || []).filter(Boolean);
    if (!uploadFiles.length) return;
    try {
      setSettingsMutating(true);
      setMappingError("");
      const formData = new FormData();
      uploadFiles.forEach((file) => formData.append("files", file));
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
        optional_columns: Array.isArray(res?.data?.optional_columns)
          ? res.data.optional_columns
          : SKU_MAPPING_OPTIONAL_COLUMNS,
      });
      setMappingInputKey((prev) => prev + 1);
      await hydratePersistedState();
      window.alert(
        `${Number(res?.data?.processed_file_count || 0)}개 파일에서 ${Number(
          res?.data?.merged_item_count || 0
        )}개의 SKU 매핑을 병합 반영했습니다.`
      );
    } catch (err) {
      const detail = err?.response?.data?.detail;
      const msg = Array.isArray(detail) ? detail.join("\n") : detail || "SKU 매핑 업로드 중 오류";
      setMappingError(msg);
      window.alert(
        "SKU 매핑 업로드가 중단되었습니다.\n\n충돌하는 item_id, 엑셀 행에 적힌 값, DB 상품별 SKU·이름 요약은 이 화면 아래 빨간 오류 박스에 전체가 표시됩니다. (alert 는 긴 메시지를 잘라 보여 줄 수 있어요.)"
      );
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
      if (!payload.kr_sku) {
        window.alert("한국 SKU를 입력해 주세요.");
        return;
      }
      if (!payload.kr_name) {
        window.alert("한국 상품명을 입력해 주세요.");
        return;
      }
      if (!payload.brand) {
        window.alert("브랜드를 선택하거나 직접 입력해 주세요.");
        return;
      }
      await axios.post(`${API_BASE}/api/inventory/mappings/item`, payload);
      setManualMappingForm({ ...EMPTY_SKU_MAPPING_FORM });
      setManualSkuFormKey((k) => k + 1);
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

  function resetManualSkuMappingForm() {
    setManualMappingForm({ ...EMPTY_SKU_MAPPING_FORM });
    setManualSkuFormKey((k) => k + 1);
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
      await axios.delete(`${API_BASE}/api/inventory/files`, {
        params: { country_code: country },
      });
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
      await axios.delete(`${API_BASE}/api/inventory/files`);
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
    <div className="pageSplit">
    <div className="dashboard">
      <div className="dashboardHeroBand">
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
      </div>

      <header
        ref={topbarRef}
        className={`topbar stickyTopbar dashboardStripWhite${countryTabMode !== "OVERSEAS" ? " isBottomCapsule" : ""}`}
      >
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
          className="countryChips stickyCountryChips dashboardStripWhite isBottomCapsule"
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
            {!isKRScope && hasInventoryLevels && (
              <select value={inventoryLevelFilter} onChange={(e) => setInventoryLevelFilter(e.target.value)}>
                {availableInventoryLevels.map((level) => (
                  <option key={level} value={level}>
                    레벨 {level}
                  </option>
                ))}
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
              setInventoryLevelFilter(defaultInventoryLevelFilter);
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
                      {getRowSku(row)}
                    </td>
                    <td className="stickyCol stickyColName stickyColBoundary">
                      <span className="nameCellText">{row.description || "(상품명 없음)"}</span>
                    </td>
                    {isOverseasScope && (
                      <td
                        className="stickyCol stickyColKrName stickyColBoundary"
                        title={
                          String(row.mapped_kr_sku || "").trim()
                            ? `한국 SKU: ${String(row.mapped_kr_sku).trim()}`
                            : "등록된 한국 SKU가 없습니다"
                        }
                      >
                        <span className="nameCellText">
                          {getCompareDisplayName(row) || krNameMap.get(getCanonicalMatchCode(row)) || "-"}
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
                          ? toFixed(krCompareMap.get(getCanonicalMatchCode(row)) || 0, 0)
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
                <li>필수 컬럼명이나 날짜 형식이 다르면 업로드나 집계가 실패할 수 있습니다.</li>
                <li>
                  재고 파일에 선택 열 option(별칭: 옵션, variant 등)이 있으면 상품명(description) 뒤에 공백과 함께
                  붙여 집계·저장합니다. SKU 매핑의 option 규칙과 같습니다.
                </li>
                <li>
                  재고 파일의 각 SKU는 DB의 item_mapping에 해당 국가 코드로 미리 등록되어 있어야 업로드됩니다. 미등록 SKU가
                  있으면 전체가 거절되며, 등록된 행이어도 같은 item에 한국(KR) SKU가 없으면 화면의 한국상품명 열은 하이픈(-)으로
                  표시됩니다.
                </li>
                <li>재고 비교는 선택한 동일 기준일 데이터만 사용하며, 없는 값은 임의로 대체하지 않습니다.</li>
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
                <li>`한국 현 재고 비교`는 한국 시간(Asia/Seoul) 기준 오늘 날짜의 한국 데이터가 있을 때만 켤 수 있습니다.</li>
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
              <div className="cautionSectionTitle">SKU 관리</div>
              <ul className="cautionList">
                <li>
                  SKU 마스터는 여러 개의 .xlsx를 올릴 수 있으며, 필수 열은{" "}
                  {SKU_MAPPING_TEMPLATE_COLUMNS.join(", ")} 입니다. 선택 열은{" "}
                  <code>option</code>(별칭: 옵션, variant 등)와 <code>brand</code>(또는 <code>브랜드</code>)입니다. option이
                  있으면 각 국가 상품명 뒤에 공백을 두고 붙여 저장합니다. 기존 item에만 매칭되는 행은 브랜드 열이 없거나 비어
                  있어도 되며 그때는 <code>item.brand</code>를 바꾸지 않습니다. DB에 없어 신규 item으로 생기는 행은 브랜드 값이
                  필수입니다.
                </li>
                <li>파일마다 일부 국가 컬럼만 있어도 되지만, 각 행에는 최소 한 국가의 SKU 값이 필요합니다.</li>
                <li>같은 국가의 같은 SKU가 다른 상품과 충돌하면 전체 업로드가 거절되며 아무 데이터도 반영되지 않습니다.</li>
                <li>원본 파일은 저장하지 않고, 읽은 매핑 데이터만 DB에 반영합니다.</li>
                <li>파일 업로드는 기존 매핑을 지우지 않고 병합 업데이트하며, 수기 입력은 한국(`KR`) 상품명과 SKU가 필수입니다.</li>
                <li>파일 업로드가 어려우면 아래 수기 입력 영역에서 국가별 상품명·SKU와 브랜드를 직접 저장할 수 있습니다.</li>
              </ul>
            </div>

            <div className="cautionSection">
              <div className="cautionSectionTitle">상품 매핑</div>
              <ul className="cautionList">
                <li>상품명 또는 SKU로 검색하면 등록된 SKU 매핑 기준으로 같은 상품의 국가별 상품명과 SKU를 함께 확인할 수 있습니다.</li>
                <li>검색 전에는 결과가 표시되지 않으며, 검색어와 일치하는 매핑이 없으면 결과가 비어 보일 수 있습니다.</li>
                <li>재고 매칭은 국가별 SKU를 우선 사용하고, 필요 시 같은 국가의 상품명 기준으로도 연결됩니다.</li>
              </ul>
            </div>
          </div>
        </div>
      )}

      {isSettingsScope && (
        <section className="settingsPane settingsCard">
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
            <input
              key={mappingInputKey}
              id="sku-mapping-input"
              type="file"
              accept=".xlsx"
              multiple
              style={{ display: "none" }}
              onChange={(e) => uploadSkuMappingFiles(e.target.files || [])}
            />
            <div className="skuManageHeader">
              <div className="skuManageTabs">
                <button
                  type="button"
                  className={`skuManageTab ${skuManageMode === "UPLOAD" ? "active" : ""}`}
                  onClick={() => setSkuManageMode("UPLOAD")}
                >
                  SKU 파일 업로드
                </button>
                <button
                  type="button"
                  className={`skuManageTab ${skuManageMode === "MANUAL" ? "active" : ""}`}
                  onClick={() => setSkuManageMode("MANUAL")}
                >
                  수기 작성
                </button>
              </div>
            </div>
            <div className={`skuManageContent ${skuManageMode === "MANUAL" ? "manual-only" : "upload-only"}`}>
              {skuManageMode === "UPLOAD" ? (
                <div className="skuUploadStage skuManageSinglePanel">
                  <button
                    type="button"
                    className="skuUploadPanel"
                    disabled={settingsMutating}
                    onClick={openSkuMappingInput}
                  >
                    <span className="skuUploadMain">
                      <span className="skuUploadBadge">XLSX</span>
                      <span className="skuUploadButtonLabel">
                        {settingsMutating ? "업로드 중..." : "SKU 파일 업로드"}
                      </span>
                    </span>
                  </button>
                  <div className="skuUploadFooter">
                    파일 업로드 최근 반영일:{" "}
                    {mappingSummary.upload_updated_at
                      ? mappingSummary.upload_updated_at.replace("T", " ").slice(0, 19)
                      : "-"}
                  </div>
                </div>
              ) : (
                <div className="skuManualCard skuManageSinglePanel">
                  <div className="skuManualHead skuManualHeadCompact">
                    <div className="skuManualHeadActions">
                      <button type="button" className="ghost" disabled={settingsMutating} onClick={resetManualSkuMappingForm}>
                        초기화
                      </button>
                      <button type="button" className="primary" disabled={settingsMutating} onClick={saveManualSkuMapping}>
                        저장
                      </button>
                    </div>
                  </div>

                  <div className="skuManualSection skuManualSectionKr">
                    <div className="skuManualBlock">
                      <p className="skuManualNoticeHint">
                        한국&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;브랜드, SKU, 상품명 필수
                      </p>
                      <div className="skuManualKrGrid">
                      <label className="skuManualKrField">
                        <span className="skuManualKrFieldHead">
                          브랜드 <abbr title="필수">*</abbr>
                        </span>
                        <div className="skuManualKrFieldBody">
                          <ManualBrandCombobox
                            key={manualSkuFormKey}
                            value={manualMappingForm.brand}
                            onChange={(next) =>
                              setManualMappingForm((prev) => ({ ...prev, brand: next }))
                            }
                            disabled={settingsMutating}
                          />
                        </div>
                      </label>
                      <label className="skuManualKrField">
                        <span className="skuManualKrFieldHead">
                          한국 SKU <abbr title="필수">*</abbr>
                        </span>
                        <div className="skuManualKrFieldBody">
                          <input
                            type="text"
                            value={manualMappingForm.kr_sku}
                            onChange={(e) =>
                              setManualMappingForm((prev) => ({ ...prev, kr_sku: e.target.value }))
                            }
                            placeholder="한국 SKU"
                            autoComplete="off"
                          />
                        </div>
                      </label>
                      <label className="skuManualKrField">
                        <span className="skuManualKrFieldHead">
                          한국 상품명 <abbr title="필수">*</abbr>
                        </span>
                        <div className="skuManualKrFieldBody">
                          <input
                            type="text"
                            value={manualMappingForm.kr_name}
                            onChange={(e) =>
                              setManualMappingForm((prev) => ({ ...prev, kr_name: e.target.value }))
                            }
                            placeholder="한국 상품명"
                            autoComplete="off"
                          />
                        </div>
                      </label>
                      <label className="skuManualKrField">
                        <span className="skuManualKrFieldHead">옵션</span>
                        <div className="skuManualKrFieldBody">
                          <input
                            type="text"
                            value={manualMappingForm.option}
                            onChange={(e) =>
                              setManualMappingForm((prev) => ({ ...prev, option: e.target.value }))
                            }
                            placeholder="예: S / M / 13호 아이보리"
                            autoComplete="off"
                          />
                        </div>
                      </label>
                    </div>
                    </div>
                  </div>

                  <div className="skuManualSection skuManualSectionOverseas">
                    <div className="skuManualBlock">
                      <p className="skuManualNoticeHint">
                        해외&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;선택. 비우면 해당 국가 제외
                      </p>
                      <div className="skuManualTable skuManualTableOverseas">
                      <div className="skuManualTableHead">
                        <div>국가</div>
                        <div>SKU</div>
                        <div>상품명</div>
                      </div>
                      {SKU_MAPPING_OVERSEAS_FIELDS.map((field) => (
                        <div key={field.code} className="skuManualRow">
                          <div className="skuManualCountryCell">
                            <span className="skuManualCountryLabel">{field.label}</span>
                          </div>
                          <div className="skuManualInputCell">
                            <input
                              type="text"
                              value={manualMappingForm[field.skuKey]}
                              onChange={(e) =>
                                setManualMappingForm((prev) => ({ ...prev, [field.skuKey]: e.target.value }))
                              }
                              placeholder={`${field.label} SKU`}
                              autoComplete="off"
                            />
                          </div>
                          <div className="skuManualInputCell">
                            <input
                              type="text"
                              value={manualMappingForm[field.nameKey]}
                              onChange={(e) =>
                                setManualMappingForm((prev) => ({ ...prev, [field.nameKey]: e.target.value }))
                              }
                              placeholder={`${field.label} 상품명`}
                              autoComplete="off"
                            />
                          </div>
                        </div>
                      ))}
                    </div>
                    </div>
                  </div>
                  <div className="skuManualActions">
                    <div className="skuManageMeta">
                      수기 입력 최근 반영일:{" "}
                      {mappingSummary.manual_updated_at
                        ? mappingSummary.manual_updated_at.replace("T", " ").slice(0, 19)
                        : "-"}
                    </div>
                    <div className="skuManualButtons" />
                  </div>
                </div>
              )}
            </div>
            {mappingError && (
              <pre ref={mappingErrorRef} className="error mappingError">
                {mappingError}
              </pre>
            )}
          </div>
        </section>
      )}

      {isProductSearchScope && (
        <section className="settingsPane settingsCard">
          <section className="tableCard productMappingCard">
            <div className="productMappingHero">
              <div className="productMappingChipRow">
                {PRODUCT_MAPPING_SEARCH_CHIPS.map((code) => (
                  <span key={code} className="productMappingChip">
                    {code}
                  </span>
                ))}
              </div>
              <div className="productMappingHeroTitle">어느 국가든 상품명이나 SKU를 검색하세요</div>
              <div className="productMappingHeroSubtitle">
                카드 상단은 한국 상품명 기준 제목입니다. 아래 국가 줄은 DB에 SKU·상품명이 함께 등록된 로케일만 보입니다.
              </div>
              <div className="productMappingSearchRow">
                <div className="searchWrap productMappingSearchWrap">
                  <span className="searchIcon" aria-hidden="true">
                    🔍
                  </span>
                  <input
                    className="searchInput productMappingSearchInput"
                    type="text"
                    value={mappingSearchKeyword}
                    onChange={(e) => setMappingSearchKeyword(e.target.value)}
                    placeholder="상품명 또는 SKU 입력..."
                  />
                </div>
              </div>
            </div>
            {mappingRowsError && <pre className="error">{mappingRowsError}</pre>}
            <div className="productMappingResultsColumn">
            {mappingRowsLoading ? (
              <div className="searchEmptyState">검색 중...</div>
            ) : !productMappingCards.length && mappingSearchKeyword.trim() ? (
              <div className="productMappingNoResult">일치하는 매핑 결과가 없습니다.</div>
            ) : !productMappingCards.length ? null : (
              <div className="productMappingList">
                {productMappingCards.map((row) => (
                  <article key={row._id} className="productMappingItemCard">
                    <div
                      className="productMappingAlignGrid"
                      title={`${row.brand || "—"} | ${row.kr_name || "상품명 없음"}`}
                    >
                      <div className="productMappingGridHeadBand">
                        <div className="productMappingItemBrandPart productMappingGridHeadBrand">
                          {row.brand || "—"}
                        </div>
                        <span className="productMappingPipe productMappingGridHeadPipe" aria-hidden="true">
                          |
                        </span>
                        <div className="productMappingItemNamePart productMappingGridHeadName">
                          {row.kr_name || "상품명 없음"}
                        </div>
                        <div className="productMappingItemMeta productMappingGridHeadMeta">
                          {row._countries.length}개 국가
                        </div>
                      </div>
                      {row._countries.map((country, countryIdx) => (
                        <Fragment key={`${row._id}-${country.code}`}>
                          {countryIdx > 0 ? (
                            <div className="productMappingGridRowRule" aria-hidden="true" />
                          ) : null}
                          <div className="productMappingCountryPrefix productMappingGridCountryPrefix">
                            <span className="productMappingCountryCode">{country.code}</span>
                            <span className="productMappingCountryLocaleName" title={country.label}>
                              {country.label}
                            </span>
                          </div>
                          <span className="productMappingPipe productMappingGridCountryPipe" aria-hidden="true">
                            |
                          </span>
                          <span className="productMappingCountryName productMappingGridCountryName" title={country.description}>
                            {country.description}
                          </span>
                          <span className="productMappingCountrySkuInline productMappingGridCountrySku" title={country.sku}>
                            {country.sku}
                          </span>
                        </Fragment>
                      ))}
                    </div>
                  </article>
                ))}
              </div>
            )}
            </div>
          </section>
        </section>
      )}
    </div>
    </div>
  );
}

