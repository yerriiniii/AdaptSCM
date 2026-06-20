import { useEffect, useLayoutEffect, useRef, useState } from "react";
import axios from "axios";
import { API_BASE } from "./apiClient";
import { Download, Upload } from "lucide-react";

/** 백엔드 `domains/item/routers/item.py` master-rows 엔드포인트와 동기화 */
const ITEM_MASTER_API = {
  rows: `${API_BASE}/api/inventory/mappings/master-rows`,
  upload: `${API_BASE}/api/inventory/mappings/master-rows/upload`,
};

const ITEM_MASTER_FIELDS = [
  { key: "brand", label: "브랜드" },
  { key: "segment", label: "구분" },
  { key: "version", label: "Ver." },
  { key: "kr_sku", label: "상품코드" },
  { key: "kr_name", label: "상품명" },
  { key: "stock_category", label: "재고구분" },
  { key: "fcst_grade", label: "FCST등급" },
  { key: "stock_grade", label: "재고등급" },
  { key: "release_month", label: "출시월" },
  { key: "code_registered_at", label: "코드 등록 일자" },
  { key: "us_grade", label: "미국 등급" },
  { key: "tw_grade", label: "대만 등급" },
  { key: "hk_grade", label: "홍콩 등급" },
  { key: "jp_grade", label: "일본 등급" },
];

const EMPTY_ITEM_MASTER_DRAFT = Object.fromEntries(ITEM_MASTER_FIELDS.map(({ key }) => [key, ""]));
const ITEM_MASTER_TABLE_MIN_WIDTH_PX = 1180;
/** 스크롤 전 툴바 아래 여백 — `.itemMasterStickyToolbar { padding-bottom }` */
const ITEM_MASTER_FLOW_TOOLBAR_BOTTOM_PAD_PX = 12;
/** sticky 고정 시 — 스크롤 전보다 살짝 좁게 */
const ITEM_MASTER_STICKY_TOOLBAR_TOP_GAP_PX = 14;
const ITEM_MASTER_STICKY_TOOLBAR_BOTTOM_GAP_PX = 6;
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

function masterDraftFromRow(row) {
  const draft = { ...EMPTY_ITEM_MASTER_DRAFT, discontinued: false };
  ITEM_MASTER_FIELDS.forEach(({ key }) => {
    draft[key] = row?.[key] == null ? "" : String(row[key]);
  });
  const seg = String(row?.segment ?? draft.segment ?? "").trim();
  draft.discontinued = productEditSegmentImpliesDiscontinued(seg);
  if (draft.discontinued) {
    draft.segment = PRODUCT_EDIT_DISCONTINUED_SEGMENT_DISPLAY;
  }
  return draft;
}

function masterCellDisplayValue(key, draft) {
  if (key === "segment" && draft.discontinued) {
    return PRODUCT_EDIT_DISCONTINUED_SEGMENT_DISPLAY;
  }
  return draft[key] ?? "";
}

function buildItemMasterSavePayload(groupId, draft) {
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
    구분: 120,
    "Ver.": 56,
    상품코드: 88,
    상품명: 240,
    재고구분: 80,
    FCST등급: 72,
    재고등급: 72,
    출시월: 72,
    "코드 등록 일자": 100,
    "미국 등급": 72,
    "대만 등급": 72,
    "홍콩 등급": 72,
    "일본 등급": 72,
  };
  const exampleHintRow = [
    "예: 95PROBLEM",
    "예: (X) 단종",
    "예: V0",
    "예: 06231",
    "예: 95PROBLEM 알패치(R) (4매입 - 파우치)",
    "예: 일반",
    "예: 정기발주",
    "예: C",
    "예: 202601",
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

function rowNeedsSave(row, draft) {
  const gid = row?.group_id;
  if (!gid || !draft) return false;
  const nextPayload = buildItemMasterSavePayload(gid, draft);
  const prevPayload = buildItemMasterSavePayload(gid, masterDraftFromRow(row));
  return JSON.stringify(nextPayload) !== JSON.stringify(prevPayload);
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
}) {
  const [masterRows, setMasterRows] = useState([]);
  const [masterLoading, setMasterLoading] = useState(false);
  const [masterError, setMasterError] = useState("");
  const [masterQuery, setMasterQuery] = useState("");
  const [masterDraftById, setMasterDraftById] = useState({});
  const [masterSavingAll, setMasterSavingAll] = useState(false);
  const [masterEditMode, setMasterEditMode] = useState(false);
  const [masterInputKey, setMasterInputKey] = useState(0);
  const masterToolbarRef = useRef(null);
  const masterBodyScrollRef = useRef(null);
  const masterHeaderScrollRef = useRef(null);
  const masterTopScrollRef = useRef(null);
  const masterScrollSyncingRef = useRef(false);
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
        setMasterRows(items);
        setMasterDraftById(Object.fromEntries(items.map((r) => [r.group_id, masterDraftFromRow(r)])));
      } catch (err) {
        const detail = err?.response?.data?.detail;
        setMasterError(Array.isArray(detail) ? detail.join("\n") : detail || "로우데이터를 불러오지 못했습니다.");
      } finally {
        setMasterLoading(false);
      }
    };
    void run();
  }, [active, masterQuery]);

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
  }, [active, masterEditMode, masterError]);

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
      const w = Math.max(
        ITEM_MASTER_TABLE_MIN_WIDTH_PX,
        cw,
        bodyInner.scrollWidth,
        bodyInner.offsetWidth,
        headInner?.scrollWidth ?? 0,
        headInner?.offsetWidth ?? 0
      );
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
    window.addEventListener("resize", measure);
    return () => {
      cancelled = true;
      cancelAnimationFrame(t1);
      cancelAnimationFrame(t2);
      ro?.disconnect();
      window.removeEventListener("resize", measure);
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
  }, [active, masterLoading, masterRows, masterEditMode]);

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
    setMasterDraftById(Object.fromEntries(masterRows.map((r) => [r.group_id, masterDraftFromRow(r)])));
    setMasterEditMode(true);
  }

  async function saveAllMasterEdits() {
    const dirtyRows = masterRows.filter((row) => rowNeedsSave(row, masterDraftById[row.group_id]));
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
        const res = await axios.patch(ITEM_MASTER_API.rows, buildItemMasterSavePayload(gid, draft));
        const saved = res?.data;
        nextRows = nextRows.map((r) => (r.group_id === gid ? saved : r));
        nextDrafts = { ...nextDrafts, [gid]: masterDraftFromRow(saved) };
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
      setSettingsMutating(true);
      setMasterError("");
      const form = new FormData();
      files.forEach((file) => form.append("files", file));
      const res = await axios.post(ITEM_MASTER_API.upload, form);
      const result = res?.data || {};
      setMasterInputKey((k) => k + 1);
      setMasterEditMode(false);
      const listRes = await axios.get(ITEM_MASTER_API.rows, {
        params: { query: masterQuery || "" },
      });
      const items = Array.isArray(listRes?.data?.items) ? listRes.data.items : [];
      setMasterRows(items);
      setMasterDraftById(Object.fromEntries(items.map((r) => [r.group_id, masterDraftFromRow(r)])));
      window.alert(
        `${formatInt(result.processed_file_count)}개 파일, ${formatInt(result.processed_row_count)}행 처리`
      );
    } catch (err) {
      const detail = err?.response?.data?.detail;
      const msg = Array.isArray(detail) ? detail.join("\n") : detail || "로우데이터 업로드 중 오류";
      setMasterError(msg);
      window.alert(msg);
    } finally {
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
            disabled={settingsMutating || masterEditMode}
            onClick={() => document.getElementById("item-master-input")?.click()}
          >
            <Upload size={16} strokeWidth={2} aria-hidden="true" />
            엑셀 파일 업로드
          </button>
          <div className="searchWrap skuProductEditSearchWrap itemMasterToolbarSearch">
            <SearchFieldIcon className="searchIcon" size={16} strokeWidth={2} />
            <input
              className="searchInput skuProductEditSearchInput"
              type="text"
              value={masterQuery}
              onChange={(e) => setMasterQuery(e.target.value)}
              placeholder="브랜드·상품코드·상품명 검색"
              autoComplete="off"
              aria-label="상품마스터 검색"
              disabled={masterEditMode}
            />
          </div>
          <button
            type="button"
            className={`itemMasterToolbarEditBtn${masterEditMode ? " itemMasterToolbarEditBtnActive" : ""}`}
            disabled={settingsMutating || masterLoading || !masterRows.length || masterSavingAll}
            onClick={handleToolbarEditSaveClick}
          >
            {masterSavingAll ? "저장 중…" : masterEditMode ? "저장" : "수정"}
          </button>
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
        {masterLoading ? (
          <div className="searchEmptyState">불러오는 중...</div>
        ) : !masterRows.length ? (
          <div className="skuProductEditEmpty">등록된 상품이 없습니다. 엑셀을 업로드해 주세요.</div>
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
                    <div className="skuProductEditTableHead itemMasterTableHead">
                      {ITEM_MASTER_FIELDS.map(({ key, label }) => (
                        <div key={key} className="skuProductEditHeadCell itemMasterHeadCell">
                          {label}
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
                  {masterRows.map((row, idx) => {
                    const gid = row.group_id;
                    const d = masterDraftById[gid] || masterDraftFromRow(row);
                    const isEditing = masterEditMode;
                    return (
                      <div key={gid || `master-${idx}`} className="skuProductEditRow itemMasterTableRow">
                        {ITEM_MASTER_FIELDS.map(({ key, label }) => (
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
                                  value={d[key] ?? ""}
                                  disabled={settingsMutating || masterSavingAll || !gid}
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
                              >
                                {masterCellDisplayValue(key, d)}
                              </span>
                            )}
                          </div>
                        ))}
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
