import { useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import axios from "axios";
import { API_BASE } from "./apiClient";
import { Download, Filter, Minus, Plus, Upload } from "lucide-react";

/** 백엔드 `domains/item/routers/item.py` master-rows 엔드포인트와 동기화 */
const ITEM_MASTER_API = {
  rows: `${API_BASE}/api/adaptscm/mappings/master-rows`,
  upload: `${API_BASE}/api/adaptscm/mappings/master-rows/upload`,
  extraColumns: `${API_BASE}/api/adaptscm/mappings/master-rows/extra-columns`,
  extraColumn: (fieldKey) =>
    `${API_BASE}/api/adaptscm/mappings/master-rows/extra-columns/${encodeURIComponent(fieldKey)}`,
};

const ITEM_MASTER_FILTER_POPOVER_WIDTH_PX = 320;

const ITEM_MASTER_FIELDS = [
  { key: "brand", label: "브랜드" },
  { key: "representative_code", label: "대표코드" },
  { key: "kr_sku", label: "상품코드" },
  { key: "version", label: "Ver." },
  { key: "kr_name", label: "상품명" },
  { key: "segment", label: "구분" },
  { key: "stock_category", label: "재고구분" },
  { key: "fcst_grade", label: "FCST등급" },
  { key: "stock_grade", label: "재고등급" },
  { key: "release_month", label: "출시월" },
  { key: "code_registered_at", label: "코드 등록 일자" },
  { key: "kr_grade", label: "한국 등급" },
  { key: "us_grade", label: "미국 등급" },
  { key: "tw_grade", label: "대만 등급" },
  { key: "hk_grade", label: "홍콩 등급" },
  { key: "jp_grade", label: "일본 등급" },
];

const ITEM_MASTER_FILTER_SPECS = [
  { key: "brand", label: "브랜드" },
  { key: "segment", label: "구분" },
  { key: "version", label: "Ver." },
  { key: "stock_category", label: "재고구분" },
  { key: "fcst_grade", label: "FCST등급" },
  { key: "stock_grade", label: "재고등급" },
];

const EMPTY_ITEM_MASTER_FILTERS = Object.fromEntries(
  ITEM_MASTER_FILTER_SPECS.map(({ key }) => [key, ""])
);

const EMPTY_ITEM_MASTER_DRAFT = Object.fromEntries(ITEM_MASTER_FIELDS.map(({ key }) => [key, ""]));
const ITEM_MASTER_TABLE_MIN_WIDTH_PX = 1320;
/** CSS `--item-master-extra-col-width` 와 동기화 (사용자 정의 열만 JS에서 추가) */
const ITEM_MASTER_EXTRA_COL_GRID = "var(--item-master-extra-col-width)";
const ITEM_MASTER_EXTRA_COL_MIN_WIDTH_PX = 72;
/** 스크롤 전 툴바 아래 여백 — `.itemMasterStickyToolbar { padding-bottom }` */
const ITEM_MASTER_FLOW_TOOLBAR_BOTTOM_PAD_PX = 12;
/** sticky 고정 시 — 스크롤 전보다 살짝 좁게 */
const ITEM_MASTER_STICKY_TOOLBAR_TOP_GAP_PX = 14;
const ITEM_MASTER_STICKY_TOOLBAR_BOTTOM_GAP_PX = 14;
const PRODUCT_EDIT_DISCONTINUED_SEGMENT_DISPLAY = "(X) 단종";

function SearchFieldIcon({ className, size = 16, strokeWidth = 2, ...rest }) {
  return (
    <svg
      className={className}
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={strokeWidth}
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
      {...rest}
    >
      <circle cx="11" cy="11" r="8" />
      <path d="m21 21-4.3-4.3" />
    </svg>
  );
}

function productEditSegmentImpliesDiscontinued(raw) {
  let s = String(raw ?? "").trim();
  if (!s) return false;
  s = s.replace(/^[(\uFF08]\s*[XxＸ]\s*[)\uFF09]\s*/, "").trim();
  return s === "단종";
}

function masterDraftFromRow(row, extraColumns = []) {
  const draft = { ...EMPTY_ITEM_MASTER_DRAFT, discontinued: false, extra_fields: {} };
  ITEM_MASTER_FIELDS.forEach(({ key }) => {
    let val = row?.[key] == null ? "" : String(row[key]);
    if (key === "code_registered_at") {
      val = formatCodeRegisteredAtDisplay(val);
    }
    draft[key] = val;
  });
  const ef = row?.extra_fields && typeof row.extra_fields === "object" ? row.extra_fields : {};
  extraColumns.forEach(({ field_key }) => {
    draft.extra_fields[field_key] = ef[field_key] == null ? "" : String(ef[field_key]);
  });
  const seg = String(row?.segment ?? draft.segment ?? "").trim();
  draft.discontinued = productEditSegmentImpliesDiscontinued(seg);
  if (draft.discontinued) {
    draft.segment = PRODUCT_EDIT_DISCONTINUED_SEGMENT_DISPLAY;
  }
  return draft;
}

function formatCodeRegisteredAtDisplay(raw) {
  const s = String(raw ?? "").trim();
  if (!s) return "";
  if (s.includes("T")) return s.split("T", 1)[0];
  if (s.includes(" ") && /^\d{4}-\d{2}-\d{2}/.test(s)) return s.split(" ", 1)[0];
  return s;
}

function itemMasterTooltipText(value) {
  const text = String(value ?? "").trim();
  return text || undefined;
}

function masterCellTooltip(key, draft, row = null) {
  let raw = masterCellRawValue(key, draft);
  if (row && row[key] != null && String(row[key]).trim()) {
    raw = String(row[key]);
  }
  return itemMasterTooltipText(raw);
}

function masterCellRawValue(key, draft) {
  if (key === "segment" && draft.discontinued) {
    return PRODUCT_EDIT_DISCONTINUED_SEGMENT_DISPLAY;
  }
  return draft[key] ?? "";
}

function masterCellDisplayValue(key, draft) {
  const raw = masterCellRawValue(key, draft);
  if (key === "code_registered_at") {
    return formatCodeRegisteredAtDisplay(raw);
  }
  if (key === "segment" && draft.discontinued) {
    return PRODUCT_EDIT_DISCONTINUED_SEGMENT_DISPLAY;
  }
  return raw;
}

function buildItemMasterSavePayload(groupId, draft, extraColumns = []) {
  const segRaw = String(draft.segment || "").trim();
  const isDiscontinued = Boolean(draft.discontinued) || productEditSegmentImpliesDiscontinued(segRaw);
  const segmentOut = isDiscontinued
    ? "단종"
    : (() => {
        if (!segRaw || segRaw === PRODUCT_EDIT_DISCONTINUED_SEGMENT_DISPLAY) return null;
        return segRaw;
      })();
  const payload = { group_id: groupId };
  ITEM_MASTER_FIELDS.forEach(({ key }) => {
    if (key === "segment") {
      payload.segment = segmentOut;
      return;
    }
    const v = String(draft[key] ?? "").trim();
    payload[key] = v || null;
  });
  payload.brand = String(draft.brand || "").trim();
  payload.kr_sku = String(draft.kr_sku || "").trim();
  payload.kr_name = String(draft.kr_name || "").trim();
  if (extraColumns.length) {
    payload.extra_fields = {};
    extraColumns.forEach(({ field_key }) => {
      const v = String(draft.extra_fields?.[field_key] ?? "").trim();
      payload.extra_fields[field_key] = v || null;
    });
  }
  return payload;
}

function excelColumnWidthFromPxApprox(px) {
  const p = Number(px);
  if (!Number.isFinite(p) || p <= 0) return 15;
  return Math.round((((p - 5) / 7) + Number.EPSILON) * 100) / 100;
}

async function buildItemMasterTemplateWorkbookBuffer() {
  const ExcelJS = (await import("exceljs")).default;
  const FONT_9 = { name: "맑은 고딕", size: 9 };
  const headers = ITEM_MASTER_FIELDS.map(({ label }) => label);
  const columnWidthsPx = {
    브랜드: 100,
    대표코드: 88,
    상품코드: 88,
    "Ver.": 56,
    상품명: 240,
    구분: 120,
    재고구분: 80,
    FCST등급: 72,
    재고등급: 72,
    출시월: 72,
    "코드 등록 일자": 100,
    "한국 등급": 72,
    "미국 등급": 72,
    "대만 등급": 72,
    "홍콩 등급": 72,
    "일본 등급": 72,
  };
  const exampleHintRow = [
    "예: 95PROBLEM",
    "예: 06231",
    "예: 06231",
    "예: V0",
    "예: 95PROBLEM 알패치(R) (4매입 - 파우치)",
    "예: (X) 단종",
    "예: 일반",
    "예: 정기발주",
    "예: C",
    "예: 2026년 1월",
    "",
    "",
    "",
    "",
    "",
    "",
  ];
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1", { views: [{ showGridLines: true }] });
  const hr = ws.addRow(headers);
  hr.height = 18;
  ws.addRow(exampleHintRow).height = 16;
  headers.forEach((h, i) => {
    const px = columnWidthsPx[h];
    ws.getColumn(i + 1).width =
      px != null ? excelColumnWidthFromPxApprox(px) : String(h).length > 12 ? 22 : 15;
  });
  ws.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      cell.font = { ...FONT_9, bold: rowNumber === 1 };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      if (rowNumber === 1) {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFFF00" } };
      }
    });
  });
  return wb.xlsx.writeBuffer();
}

async function downloadItemMasterTemplateWorkbook() {
  const buf = await buildItemMasterTemplateWorkbookBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "상품마스터_템플릿.xlsx";
  a.click();
  URL.revokeObjectURL(a.href);
}

function rowNeedsSave(row, draft, extraColumns = []) {
  const gid = row?.group_id;
  if (!gid || !draft) return false;
  const nextPayload = buildItemMasterSavePayload(gid, draft, extraColumns);
  const prevPayload = buildItemMasterSavePayload(gid, masterDraftFromRow(row, extraColumns), extraColumns);
  return JSON.stringify(nextPayload) !== JSON.stringify(prevPayload);
}

/** 사용자 정의 열이 있을 때만 CSS 기본 14열 뒤에 열 정의를 붙임 */
function buildItemMasterGridStyle(extraColumns) {
  const extras = (extraColumns || [])
    .map(() => ITEM_MASTER_EXTRA_COL_GRID)
    .join(" ");
  if (!extras) return undefined;
  return { gridTemplateColumns: `var(--item-master-grid-columns) ${extras}` };
}

function formatInt(n) {
  return Number(n || 0).toLocaleString("ko-KR");
}

export default function ItemMasterTab({
  active,
  embedded = false,
  settingsMutating,
  setSettingsMutating,
  stickyBaseTop = 0,
  onFilesUploaded,
}) {
  const [masterRows, setMasterRows] = useState([]);
  const [masterExtraColumns, setMasterExtraColumns] = useState([]);
  const [masterAddingColumn, setMasterAddingColumn] = useState(false);
  const [masterDeletingColumnKey, setMasterDeletingColumnKey] = useState("");
  const [masterLoading, setMasterLoading] = useState(false);
  const [masterError, setMasterError] = useState("");
  const [masterQuery, setMasterQuery] = useState("");
  const [masterFilters, setMasterFilters] = useState({ ...EMPTY_ITEM_MASTER_FILTERS });
  const [masterFilterOpen, setMasterFilterOpen] = useState(false);
  const [filterPopoverStyle, setFilterPopoverStyle] = useState({ top: 0, left: 0 });
  const masterFilterMenuRef = useRef(null);
  const masterFilterBtnRef = useRef(null);
  const masterFilterPopoverRef = useRef(null);
  const [masterDraftById, setMasterDraftById] = useState({});
  const [masterSavingAll, setMasterSavingAll] = useState(false);
  const [masterUploadBusy, setMasterUploadBusy] = useState(false);
  const [masterEditMode, setMasterEditMode] = useState(false);
  const [masterInputKey, setMasterInputKey] = useState(0);
  const masterToolbarRef = useRef(null);
  const masterBodyScrollRef = useRef(null);
  const masterHeaderScrollRef = useRef(null);
  const masterTopScrollRef = useRef(null);
  const masterScrollSyncingRef = useRef(false);
  /** 행 수·검색 필터와 무관하게 유지할 표 레이아웃 너비(한 번 넓게 잡히면 줄이지 않음) */
  const masterTableLayoutWidthRef = useRef(ITEM_MASTER_TABLE_MIN_WIDTH_PX);
  const [masterShowTopScroll, setMasterShowTopScroll] = useState(false);
  const [masterTopScrollWidth, setMasterTopScrollWidth] = useState(ITEM_MASTER_TABLE_MIN_WIDTH_PX);
  const [masterToolbarHeight, setMasterToolbarHeight] = useState(0);
  const [masterTopStripHeight, setMasterTopStripHeight] = useState(0);

  useEffect(() => {
    if (!active) return;
    const run = async () => {
      try {
        setMasterLoading(true);
        setMasterError("");
        setMasterEditMode(false);
        const res = await axios.get(ITEM_MASTER_API.rows, {
          params: { query: masterQuery || "" },
        });
        const items = Array.isArray(res?.data?.items) ? res.data.items : [];
        const extraCols = Array.isArray(res?.data?.extra_columns) ? res.data.extra_columns : [];
        setMasterExtraColumns(extraCols);
        setMasterRows(items);
        setMasterDraftById(Object.fromEntries(items.map((r) => [r.group_id, masterDraftFromRow(r, extraCols)])));
      } catch (err) {
        const detail = err?.response?.data?.detail;
        setMasterError(Array.isArray(detail) ? detail.join("\n") : detail || "로우데이터를 불러오지 못했습니다.");
      } finally {
        setMasterLoading(false);
      }
    };
    void run();
  }, [active, masterQuery]);

  useEffect(() => {
    if (!masterFilterOpen) return;
    const onDocPointer = (e) => {
      const target = e.target;
      if (masterFilterBtnRef.current?.contains(target)) return;
      if (masterFilterPopoverRef.current?.contains(target)) return;
      setMasterFilterOpen(false);
    };
    document.addEventListener("mousedown", onDocPointer);
    return () => document.removeEventListener("mousedown", onDocPointer);
  }, [masterFilterOpen]);

  useLayoutEffect(() => {
    if (!masterFilterOpen || !masterFilterBtnRef.current) return;
    const updatePos = () => {
      const btn = masterFilterBtnRef.current;
      if (!btn) return;
      const rect = btn.getBoundingClientRect();
      const width = ITEM_MASTER_FILTER_POPOVER_WIDTH_PX;
      const left = Math.min(
        Math.max(8, rect.right - width),
        window.innerWidth - width - 8
      );
      setFilterPopoverStyle({
        top: rect.bottom + 8,
        left,
      });
    };
    updatePos();
    window.addEventListener("resize", updatePos);
    window.addEventListener("scroll", updatePos, true);
    return () => {
      window.removeEventListener("resize", updatePos);
      window.removeEventListener("scroll", updatePos, true);
    };
  }, [masterFilterOpen, masterToolbarHeight]);

  useLayoutEffect(() => {
    if (!active) {
      setMasterToolbarHeight(0);
      return;
    }
    const measure = () => {
      setMasterToolbarHeight(masterToolbarRef.current?.offsetHeight ?? 0);
    };
    measure();
    let ro;
    if (typeof ResizeObserver !== "undefined" && masterToolbarRef.current) {
      ro = new ResizeObserver(measure);
      ro.observe(masterToolbarRef.current);
    }
    window.addEventListener("resize", measure);
    return () => {
      window.removeEventListener("resize", measure);
      ro?.disconnect();
    };
  }, [active, masterEditMode, masterError, masterFilters]);

  useLayoutEffect(() => {
    if (!active || !masterShowTopScroll) {
      setMasterTopStripHeight(0);
      return;
    }
    const measure = () => {
      setMasterTopStripHeight(masterTopScrollRef.current?.offsetHeight ?? 0);
    };
    measure();
    let ro;
    if (typeof ResizeObserver !== "undefined" && masterTopScrollRef.current) {
      ro = new ResizeObserver(measure);
      ro.observe(masterTopScrollRef.current);
    }
    window.addEventListener("resize", measure);
    return () => {
      window.removeEventListener("resize", measure);
      ro?.disconnect();
    };
  }, [active, masterShowTopScroll]);

  const masterToolbarStickyTop = stickyBaseTop + ITEM_MASTER_STICKY_TOOLBAR_TOP_GAP_PX;
  const masterTopScrollStickyTop =
    masterToolbarStickyTop +
    masterToolbarHeight -
    ITEM_MASTER_FLOW_TOOLBAR_BOTTOM_PAD_PX +
    ITEM_MASTER_STICKY_TOOLBAR_BOTTOM_GAP_PX;
  const masterHeaderStickyTop =
    masterTopScrollStickyTop + (masterShowTopScroll ? masterTopStripHeight : 0);
  const masterToolbarBottomGapStickyTop =
    masterToolbarStickyTop + masterToolbarHeight - ITEM_MASTER_FLOW_TOOLBAR_BOTTOM_PAD_PX;

  const masterGridStyle = buildItemMasterGridStyle(masterExtraColumns);
  const masterTableMinWidthPx =
    ITEM_MASTER_TABLE_MIN_WIDTH_PX + masterExtraColumns.length * ITEM_MASTER_EXTRA_COL_MIN_WIDTH_PX;

  const masterFilterOptions = useMemo(() => {
    const out = {};
    ITEM_MASTER_FILTER_SPECS.forEach(({ key }) => {
      const vals = new Set();
      masterRows.forEach((row) => {
        const v = String(row[key] ?? "").trim();
        if (v) vals.add(v);
      });
      out[key] = [...vals].sort((a, b) => a.localeCompare(b, "ko"));
    });
    return out;
  }, [masterRows]);

  const displayedMasterRows = useMemo(() => {
    return masterRows.filter((row) => {
      for (const { key } of ITEM_MASTER_FILTER_SPECS) {
        const selected = String(masterFilters[key] || "").trim();
        if (!selected) continue;
        if (String(row[key] ?? "").trim() !== selected) return false;
      }
      return true;
    });
  }, [masterRows, masterFilters]);

  const masterFiltersActive = Object.values(masterFilters).some((v) => String(v || "").trim());
  const masterActiveFilterCount = Object.values(masterFilters).filter((v) =>
    String(v || "").trim()
  ).length;

  useLayoutEffect(() => {
    if (!active || masterLoading || !masterRows.length) {
      setMasterShowTopScroll(false);
      return;
    }
    let cancelled = false;
    const measure = () => {
      if (cancelled) return;
      const body = masterBodyScrollRef.current;
      const head = masterHeaderScrollRef.current;
      const bodyInner = body?.querySelector(".itemMasterTableInnerBody");
      const headInner = head?.querySelector(".itemMasterTableInnerHead");
      if (!body || !bodyInner) return;
      const wrap = body.closest(".itemMasterTableWrap");
      const cw = wrap?.clientWidth ?? body.clientWidth;
      bodyInner.style.width = "";
      bodyInner.style.minWidth = "";
      if (headInner) {
        headInner.style.width = "";
        headInner.style.minWidth = "";
      }
      const layoutW = Math.max(
        masterTableMinWidthPx,
        headInner?.scrollWidth ?? 0,
        bodyInner.scrollWidth,
        headInner?.offsetWidth ?? 0,
        bodyInner.offsetWidth
      );
      if (layoutW > masterTableLayoutWidthRef.current) {
        masterTableLayoutWidthRef.current = layoutW;
      }
      const w = Math.max(masterTableMinWidthPx, cw, masterTableLayoutWidthRef.current);
      bodyInner.style.width = `${w}px`;
      bodyInner.style.minWidth = `${w}px`;
      if (headInner) {
        headInner.style.width = `${w}px`;
        headInner.style.minWidth = `${w}px`;
      }
      setMasterTopScrollWidth(w);
      setMasterShowTopScroll(w > cw + 1);
      if (head) {
        const scrollbarPad = Math.max(0, body.offsetWidth - body.clientWidth);
        head.style.paddingRight = scrollbarPad ? `${scrollbarPad}px` : "";
      }
    };
    const onResize = () => {
      masterTableLayoutWidthRef.current = masterTableMinWidthPx;
      measure();
    };
    measure();
    const t1 = requestAnimationFrame(measure);
    const t2 = requestAnimationFrame(() => requestAnimationFrame(measure));
    let ro;
    if (typeof ResizeObserver !== "undefined") {
      ro = new ResizeObserver(measure);
      const body = masterBodyScrollRef.current;
      const head = masterHeaderScrollRef.current;
      if (body) {
        ro.observe(body);
        const inner = body.querySelector(".itemMasterTableInnerBody");
        if (inner) ro.observe(inner);
      }
      if (head) ro.observe(head);
    }
    window.addEventListener("resize", onResize);
    return () => {
      cancelled = true;
      cancelAnimationFrame(t1);
      cancelAnimationFrame(t2);
      ro?.disconnect();
      window.removeEventListener("resize", onResize);
      if (masterHeaderScrollRef.current) {
        masterHeaderScrollRef.current.style.paddingRight = "";
      }
      const bodyInner = masterBodyScrollRef.current?.querySelector(".itemMasterTableInnerBody");
      const headInner = masterHeaderScrollRef.current?.querySelector(".itemMasterTableInnerHead");
      if (bodyInner) {
        bodyInner.style.width = "";
        bodyInner.style.minWidth = "";
      }
      if (headInner) {
        headInner.style.width = "";
        headInner.style.minWidth = "";
      }
    };
  }, [active, masterLoading, masterRows, masterEditMode, masterExtraColumns.length, masterTableMinWidthPx]);

  function syncMasterTableScroll(source) {
    if (masterScrollSyncingRef.current) return;
    const topEl = masterTopScrollRef.current;
    const headerEl = masterHeaderScrollRef.current;
    const bodyEl = masterBodyScrollRef.current;
    if (!bodyEl) return;
    masterScrollSyncingRef.current = true;
    const raw =
      source === "top"
        ? topEl?.scrollLeft ?? 0
        : source === "header"
          ? headerEl?.scrollLeft ?? 0
          : bodyEl.scrollLeft ?? 0;
    const sl = Math.max(0, Math.round(Number(raw)));
    if (topEl) topEl.scrollLeft = sl;
    if (headerEl) headerEl.scrollLeft = sl;
    bodyEl.scrollLeft = sl;
    requestAnimationFrame(() => {
      if (topEl) topEl.scrollLeft = bodyEl.scrollLeft;
      if (headerEl) headerEl.scrollLeft = bodyEl.scrollLeft;
      masterScrollSyncingRef.current = false;
    });
  }

  function enterMasterEditMode() {
    setMasterDraftById(
      Object.fromEntries(masterRows.map((r) => [r.group_id, masterDraftFromRow(r, masterExtraColumns)]))
    );
    setMasterEditMode(true);
  }

  async function addMasterExtraColumn() {
    const label = window.prompt("추가할 열 이름을 입력하세요.", "");
    if (label == null) return;
    const trimmed = String(label).trim();
    if (!trimmed) return;
    try {
      setMasterAddingColumn(true);
      setMasterError("");
      const res = await axios.post(ITEM_MASTER_API.extraColumns, { label: trimmed });
      const col = res?.data;
      if (!col?.field_key) throw new Error("invalid column");
      setMasterExtraColumns((prev) => [...prev, col]);
      setMasterDraftById((prev) => {
        const next = { ...prev };
        Object.keys(next).forEach((gid) => {
          next[gid] = {
            ...next[gid],
            extra_fields: { ...(next[gid]?.extra_fields || {}), [col.field_key]: "" },
          };
        });
        return next;
      });
    } catch (err) {
      const detail = err?.response?.data?.detail;
      const msg = Array.isArray(detail) ? detail.join("\n") : detail || "열 추가 중 오류";
      setMasterError(msg);
      window.alert(msg);
    } finally {
      setMasterAddingColumn(false);
    }
  }

  async function deleteMasterExtraColumn(fieldKey, label) {
    const key = String(fieldKey || "").trim();
    if (!key) return;
    if (
      !window.confirm(
        `「${label}」 열을 삭제할까요?\n해당 열에 저장된 모든 상품 값도 함께 삭제됩니다.`
      )
    ) {
      return;
    }
    try {
      setMasterDeletingColumnKey(key);
      setMasterError("");
      await axios.delete(ITEM_MASTER_API.extraColumn(key));
      setMasterExtraColumns((prev) => prev.filter((c) => c.field_key !== key));
      setMasterRows((prev) =>
        prev.map((row) => {
          const ef = { ...(row.extra_fields || {}) };
          delete ef[key];
          return { ...row, extra_fields: ef };
        })
      );
      setMasterDraftById((prev) => {
        const next = { ...prev };
        Object.keys(next).forEach((gid) => {
          const ef = { ...(next[gid]?.extra_fields || {}) };
          delete ef[key];
          next[gid] = { ...next[gid], extra_fields: ef };
        });
        return next;
      });
      masterTableLayoutWidthRef.current = Math.max(
        ITEM_MASTER_TABLE_MIN_WIDTH_PX,
        ITEM_MASTER_TABLE_MIN_WIDTH_PX +
          Math.max(0, masterExtraColumns.length - 1) * ITEM_MASTER_EXTRA_COL_MIN_WIDTH_PX
      );
    } catch (err) {
      const detail = err?.response?.data?.detail;
      const msg = Array.isArray(detail) ? detail.join("\n") : detail || "열 삭제 중 오류";
      setMasterError(msg);
      window.alert(msg);
    } finally {
      setMasterDeletingColumnKey("");
    }
  }

  async function saveAllMasterEdits() {
    const dirtyRows = masterRows.filter((row) =>
      rowNeedsSave(row, masterDraftById[row.group_id], masterExtraColumns)
    );
    if (!dirtyRows.length) {
      setMasterEditMode(false);
      return;
    }
    try {
      setMasterSavingAll(true);
      setMasterError("");
      let nextRows = masterRows;
      let nextDrafts = { ...masterDraftById };
      for (const row of dirtyRows) {
        const gid = row.group_id;
        const draft = masterDraftById[gid];
        if (!gid || !draft) continue;
        const res = await axios.patch(
          ITEM_MASTER_API.rows,
          buildItemMasterSavePayload(gid, draft, masterExtraColumns)
        );
        const saved = res?.data;
        nextRows = nextRows.map((r) => (r.group_id === gid ? saved : r));
        nextDrafts = { ...nextDrafts, [gid]: masterDraftFromRow(saved, masterExtraColumns) };
      }
      setMasterRows(nextRows);
      setMasterDraftById(nextDrafts);
      setMasterEditMode(false);
    } catch (err) {
      const detail = err?.response?.data?.detail;
      const msg = Array.isArray(detail) ? detail.join("\n") : detail || "저장 중 오류";
      setMasterError(msg);
      window.alert(msg);
    } finally {
      setMasterSavingAll(false);
    }
  }

  function handleToolbarEditSaveClick() {
    if (masterEditMode) {
      void saveAllMasterEdits();
      return;
    }
    enterMasterEditMode();
  }

  async function uploadItemMasterFilesHandler(fileList) {
    const files = Array.from(fileList || []).filter(Boolean);
    if (!files.length) return;
    try {
      setMasterUploadBusy(true);
      setSettingsMutating(true);
      setMasterError("");
      const form = new FormData();
      files.forEach((file) => form.append("files", file));
      const res = await axios.post(ITEM_MASTER_API.upload, form, { timeout: 600_000 });
      const result = res?.data || {};
      onFilesUploaded?.(files);
      setMasterInputKey((k) => k + 1);
      setMasterEditMode(false);
      try {
        const listRes = await axios.get(ITEM_MASTER_API.rows, {
          params: { query: masterQuery || "" },
          timeout: 120_000,
        });
        const items = Array.isArray(listRes?.data?.items) ? listRes.data.items : [];
        const extraCols = Array.isArray(listRes?.data?.extra_columns) ? listRes.data.extra_columns : [];
        setMasterExtraColumns(extraCols);
        setMasterRows(items);
        setMasterDraftById(Object.fromEntries(items.map((r) => [r.group_id, masterDraftFromRow(r, extraCols)])));
      } catch (listErr) {
        console.warn("item master list refresh after upload", listErr);
        setMasterError("업로드는 완료되었으나 목록 새로고침에 실패했습니다. 페이지를 새로고침해 주세요.");
      }
      window.alert(
        `${formatInt(result.processed_file_count)}개 파일, ${formatInt(result.processed_row_count)}행 처리`
      );
    } catch (err) {
      const detail = err?.response?.data?.detail;
      const msg = detail
        ? Array.isArray(detail)
          ? detail.join("\n")
          : String(detail)
        : err?.code === "ECONNABORTED"
          ? "업로드 시간이 초과되었습니다. 파일 행 수가 많으면 잠시 후 다시 시도해 주세요."
          : err?.message || "로우데이터 업로드 중 오류";
      setMasterError(msg);
      window.alert(msg);
    } finally {
      setMasterUploadBusy(false);
      setSettingsMutating(false);
    }
  }

  if (!active) return null;

  const panel = (
    <>
        <input
          key={masterInputKey}
          id="item-master-input"
          type="file"
          accept=".xlsx"
          multiple
          style={{ display: "none" }}
          onChange={(e) => uploadItemMasterFilesHandler(e.target.files || [])}
        />
        <div
          className="itemMasterStickySectionGap itemMasterStickySectionGapBelowHeader"
          style={{
            top: stickyBaseTop,
            height: ITEM_MASTER_STICKY_TOOLBAR_TOP_GAP_PX,
            marginBottom: -ITEM_MASTER_STICKY_TOOLBAR_TOP_GAP_PX,
          }}
          aria-hidden="true"
        />
        <div
          ref={masterToolbarRef}
          className="itemMasterToolbarRow itemMasterStickyToolbar"
          style={{ top: masterToolbarStickyTop }}
        >
          <button
            type="button"
            className="poOrderFileTemplateBtn itemMasterTemplateBtn"
            disabled={settingsMutating || masterEditMode}
            onClick={() => {
              downloadItemMasterTemplateWorkbook().catch(() =>
                window.alert("엑셀 양식을 만드는 중 오류가 났습니다.")
              );
            }}
          >
            <Download size={16} strokeWidth={2} aria-hidden="true" />
            엑셀 양식 다운로드
          </button>
          <button
            type="button"
            className="poOrderFileTemplateBtn itemMasterTemplateBtn itemMasterUploadBtn"
            disabled={settingsMutating || masterEditMode || masterUploadBusy}
            onClick={() => document.getElementById("item-master-input")?.click()}
          >
            <Upload size={16} strokeWidth={2} aria-hidden="true" />
            {masterUploadBusy ? "업로드 중…" : "엑셀 파일 업로드"}
          </button>
          <div className="searchWrap skuProductEditSearchWrap itemMasterToolbarSearch">
            <SearchFieldIcon className="searchIcon" size={16} strokeWidth={2} />
            <input
              className="searchInput skuProductEditSearchInput"
              type="text"
              value={masterQuery}
              onChange={(e) => setMasterQuery(e.target.value)}
              placeholder="브랜드·대표코드·상품코드·상품명 검색"
              autoComplete="off"
              aria-label="상품마스터 검색"
              disabled={masterEditMode}
            />
          </div>
          <div className="itemMasterFilterMenuWrap" ref={masterFilterMenuRef}>
            <button
              ref={masterFilterBtnRef}
              type="button"
              className={`itemMasterFilterToggleBtn${masterFiltersActive ? " isActive" : ""}`}
              disabled={masterEditMode || masterLoading}
              aria-expanded={masterFilterOpen}
              aria-haspopup="dialog"
              onClick={() => setMasterFilterOpen((open) => !open)}
            >
              <Filter size={16} strokeWidth={2} aria-hidden="true" />
              필터
              {masterFiltersActive ? (
                <span className="itemMasterFilterActiveDot" aria-hidden="true" />
              ) : null}
            </button>
          </div>
          {masterFilterOpen && typeof document !== "undefined"
            ? createPortal(
                <div
                  ref={masterFilterPopoverRef}
                  className="itemMasterFilterPopover itemMasterFilterPopoverFloating"
                  role="dialog"
                  aria-label="상품마스터 필터"
                  style={{
                    top: filterPopoverStyle.top,
                    left: filterPopoverStyle.left,
                    width: ITEM_MASTER_FILTER_POPOVER_WIDTH_PX,
                  }}
                >
                  <div className="itemMasterFilterPopoverHead">
                    <div className="itemMasterFilterPopoverHeadMain">
                      <Filter size={16} strokeWidth={2} aria-hidden="true" />
                      <span>상품 필터</span>
                    </div>
                    {masterFiltersActive ? (
                      <span className="itemMasterFilterPopoverApplied">
                        {masterActiveFilterCount}개 적용
                      </span>
                    ) : null}
                  </div>
                  <div className="itemMasterFilterPopoverBody">
                    {ITEM_MASTER_FILTER_SPECS.map(({ key, label }) => (
                      <label key={key} className="itemMasterFilterPopoverField">
                        <span className="itemMasterFilterLabel">{label}</span>
                        <select
                          className="itemMasterFilterSelect"
                          value={masterFilters[key]}
                          onChange={(e) =>
                            setMasterFilters((prev) => ({ ...prev, [key]: e.target.value }))
                          }
                          disabled={masterEditMode || masterLoading}
                        >
                          <option value="">전체</option>
                          {(masterFilterOptions[key] || []).map((opt) => (
                            <option key={opt} value={opt}>
                              {opt}
                            </option>
                          ))}
                        </select>
                      </label>
                    ))}
                  </div>
                  {masterFiltersActive ? (
                    <div className="itemMasterFilterPopoverFoot">
                      <button
                        type="button"
                        className="itemMasterFilterReset"
                        disabled={masterEditMode}
                        onClick={() => setMasterFilters({ ...EMPTY_ITEM_MASTER_FILTERS })}
                      >
                        필터 초기화
                      </button>
                    </div>
                  ) : null}
                </div>,
                document.body
              )
            : null}
          <button
            type="button"
            className={`itemMasterToolbarEditBtn${masterEditMode ? " itemMasterToolbarEditBtnActive" : ""}`}
            disabled={settingsMutating || masterLoading || !masterRows.length || masterSavingAll}
            onClick={handleToolbarEditSaveClick}
            title={masterEditMode ? "변경 사항을 DB에 저장합니다" : "스프레드에서 수정 후 저장하면 DB에 반영됩니다"}
          >
            {masterSavingAll ? "DB 저장 중…" : masterEditMode ? "저장" : "수정"}
          </button>
          {masterEditMode ? (
            <button
              type="button"
              className="itemMasterAddColumnBtn"
              disabled={settingsMutating || masterSavingAll || masterAddingColumn}
              onClick={() => void addMasterExtraColumn()}
            >
              <Plus size={16} strokeWidth={2} aria-hidden="true" />
              {masterAddingColumn ? "열 추가 중…" : "열 추가"}
            </button>
          ) : null}
        </div>
        <div
          className="itemMasterStickySectionGap itemMasterStickySectionGapBelowToolbar"
          style={{
            top: masterToolbarBottomGapStickyTop,
            height: ITEM_MASTER_STICKY_TOOLBAR_BOTTOM_GAP_PX,
            marginBottom: -ITEM_MASTER_STICKY_TOOLBAR_BOTTOM_GAP_PX,
          }}
          aria-hidden="true"
        />
        {masterError ? <pre className="error skuProductEditError">{masterError}</pre> : null}
        {masterUploadBusy ? (
          <div className="searchEmptyState">엑셀 업로드 처리 중…</div>
        ) : masterLoading ? (
          <div className="searchEmptyState">불러오는 중...</div>
        ) : !masterRows.length ? (
          <div className="skuProductEditEmpty">등록된 상품이 없습니다. 엑셀을 업로드해 주세요.</div>
        ) : !displayedMasterRows.length ? (
          <div className="skuProductEditEmpty">필터 조건에 맞는 상품이 없습니다.</div>
        ) : (
          <div
            className={`skuProductEditTableWrap itemMasterTableWrap${
              masterShowTopScroll ? " itemMasterShowTopScrollPair" : ""
            }`}
          >
            <div
              className={`itemMasterScrollPair${
                masterShowTopScroll ? " itemMasterShowTopScrollPair" : ""
              }`}
            >
              <div
                ref={masterTopScrollRef}
                className={`tableTopScroll itemMasterTopScroll stickyTableTopScroll${
                  masterShowTopScroll ? "" : " itemMasterTopScrollHidden"
                }`}
                style={masterShowTopScroll ? { top: masterTopScrollStickyTop } : undefined}
                onScroll={() => syncMasterTableScroll("top")}
                aria-hidden={!masterShowTopScroll}
              >
                <div style={{ width: masterTopScrollWidth }} />
              </div>
              <div
                className="itemMasterStickyHeaderShell"
                style={{ top: masterHeaderStickyTop }}
              >
                <div
                  ref={masterHeaderScrollRef}
                  className="itemMasterHeaderScroll"
                  onScroll={() => syncMasterTableScroll("header")}
                >
                  <div className="itemMasterTableInner itemMasterTableInnerHead">
                    <div
                      className="skuProductEditTableHead itemMasterTableHead"
                      style={masterGridStyle}
                    >
                      {ITEM_MASTER_FIELDS.map(({ key, label }) => (
                        <div key={key} className="skuProductEditHeadCell itemMasterHeadCell">
                          {label}
                        </div>
                      ))}
                      {masterExtraColumns.map(({ field_key, label }) => (
                        <div
                          key={field_key}
                          className="skuProductEditHeadCell itemMasterHeadCell itemMasterExtraHeadCell"
                        >
                          {masterEditMode ? (
                            <div className="itemMasterExtraHeadInner">
                              <button
                                type="button"
                                className="itemMasterExtraColDeleteBtn"
                                disabled={
                                  settingsMutating ||
                                  masterSavingAll ||
                                  masterDeletingColumnKey === field_key
                                }
                                onClick={() => void deleteMasterExtraColumn(field_key, label)}
                                aria-label={`${label} 열 삭제`}
                                title={`${label} 열 삭제`}
                              >
                                <Minus size={14} strokeWidth={2.5} aria-hidden="true" />
                              </button>
                              <span className="itemMasterExtraHeadLabel">{label}</span>
                            </div>
                          ) : (
                            label
                          )}
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
              <div
                ref={masterBodyScrollRef}
                className="itemMasterBodyPane"
                onScroll={() => syncMasterTableScroll("body")}
              >
                <div className="itemMasterTableInner itemMasterTableInnerBody">
                  {displayedMasterRows.map((row, idx) => {
                    const gid = row.group_id;
                    const d = masterDraftById[gid] || masterDraftFromRow(row);
                    const isEditing = masterEditMode;
                    return (
                      <div
                        key={gid || `master-${idx}`}
                        className="skuProductEditRow itemMasterTableRow"
                        style={masterGridStyle}
                      >
                        {ITEM_MASTER_FIELDS.map(({ key, label }) => {
                          const cellRaw = masterCellRawValue(key, d);
                          const cellDisplay = masterCellDisplayValue(key, d);
                          const cellTitle = masterCellTooltip(key, d, row);
                          return (
                          <div key={key} className="skuProductEditCell itemMasterCell">
                            {isEditing ? (
                              key === "segment" ? (
                                <input
                                  type="text"
                                  className={`skuProductEditInput${
                                    productEditSegmentImpliesDiscontinued(d.segment)
                                      ? " skuProductEditInputDiscontinuedSegment"
                                      : ""
                                  }`}
                                  value={d.segment ?? ""}
                                  disabled={settingsMutating || masterSavingAll || !gid}
                                  title={cellTitle}
                                  onChange={(e) => {
                                    const v = e.target.value;
                                    setMasterDraftById((prev) => ({
                                      ...prev,
                                      [gid]: {
                                        ...d,
                                        discontinued: productEditSegmentImpliesDiscontinued(v),
                                        segment: v,
                                      },
                                    }));
                                  }}
                                  autoComplete="off"
                                  aria-label={label}
                                />
                              ) : (
                                <input
                                  type="text"
                                  className="skuProductEditInput"
                                  value={key === "code_registered_at" ? cellDisplay : (d[key] ?? "")}
                                  disabled={settingsMutating || masterSavingAll || !gid}
                                  title={cellTitle}
                                  onChange={(e) =>
                                    setMasterDraftById((prev) => ({
                                      ...prev,
                                      [gid]: { ...d, [key]: e.target.value },
                                    }))
                                  }
                                  autoComplete="off"
                                  aria-label={label}
                                />
                              )
                            ) : (
                              <span
                                className={`itemMasterCellReadonly${
                                  key === "segment" && d.discontinued
                                    ? " itemMasterCellReadonlyDiscontinued"
                                    : ""
                                }`}
                                title={cellTitle}
                              >
                                {cellDisplay}
                              </span>
                            )}
                          </div>
                          );
                        })}
                        {masterExtraColumns.map(({ field_key, label }) => {
                          const extraRaw = String(d.extra_fields?.[field_key] ?? "");
                          const extraTitle = itemMasterTooltipText(extraRaw);
                          return (
                          <div key={field_key} className="skuProductEditCell itemMasterCell itemMasterExtraCell">
                            {isEditing ? (
                              <input
                                type="text"
                                className="skuProductEditInput"
                                value={extraRaw}
                                disabled={settingsMutating || masterSavingAll || !gid}
                                title={extraTitle}
                                onChange={(e) =>
                                  setMasterDraftById((prev) => ({
                                    ...prev,
                                    [gid]: {
                                      ...d,
                                      extra_fields: {
                                        ...(d.extra_fields || {}),
                                        [field_key]: e.target.value,
                                      },
                                    },
                                  }))
                                }
                                autoComplete="off"
                                aria-label={label}
                              />
                            ) : (
                              <span className="itemMasterCellReadonly" title={extraTitle}>
                                {extraRaw}
                              </span>
                            )}
                          </div>
                          );
                        })}
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          </div>
        )}
    </>
  );

  if (embedded) return panel;
  return (
    <section className="settingsPane settingsCard">
      <div className="skuManageCard">{panel}</div>
    </section>
  );
}
