import { Fragment, useEffect, useLayoutEffect, useMemo, useRef, useState } from "react";
import { createPortal } from "react-dom";
import axios from "axios";
import { API_BASE, clearAccessToken } from "./apiClient";
import {
  Ban,
  Barcode,
  ClipboardList,
  Database,
  FileText,
  GitCompare,
  Globe2,
  Home,
  Link2,
  LogOut,
  Truck,
  UserRound,
} from "lucide-react";
import MyPage from "./MyPage.jsx";
import * as XLSX from "xlsx";

/** 발주 목록 GET이 응답 없이 멈출 때 UI가 「불러오는 중」에 고정되지 않도록 */
const PURCHASE_ORDERS_LIST_TIMEOUT_MS = 45_000;
/** 한국 재고 `.tableTopScroll`: 높이 14px + 테두리로 보통 offsetHeight ≈16, 아래 `margin-bottom` 6px */
const INVENTORY_MATCH_TOP_SCROLL_STRIP_FALLBACK_PX = 16;
const INVENTORY_MATCH_TOP_SCROLL_MARGIN_BELOW_PX = 6;
const DEFAULT_DATE_RANGE = "10d";
const OVERSEAS_UPLOAD_COUNTRIES = ["US", "TW", "HK", "JP", "SG", "DE", "UK", "AU", "AE", "VN", "TH"];
/** hydrate 시 해외 `/view` 동시 요청 수 — DB·연결 풀 부하 시 전체가 한꺼번에 막히는 것 완화 */
const HYDRATE_OVERSEAS_VIEW_CONCURRENCY = 3;
const SETTINGS_COUNTRY_ORDER = ["KR", "US", "TW", "HK", "JP", "SG", "DE", "UK", "AU", "AE", "VN", "TH", "SHIPMENT"];
/** 출고 현황 칩 = 엑셀 시트 이름(백엔드 SHIPMENT_MATRIX_SHEET_NAMES와 동일 순서) */
const SHIPMENT_MATRIX_CHIPS = [
  "B2B 통합",
  "B2C 통합",
  "해외 통합",
  "B2B 현황",
  "쿠팡 현황",
  "올리브영 현황",
  "자사몰 현황",
  "외부몰 현황",
  "대만 현황",
  "홍콩 현황",
  "일본 현황",
  "그 외 해외 현황",
];
/** 이 칩 앞에 세로 구분 막대 삽입(`styles.css` `.shipmentChipSep`) */
const SHIPMENT_MATRIX_CHIP_DIVIDER_BEFORE = new Set(["B2B 현황", "자사몰 현황", "대만 현황"]);
/** 차트 등 레거시 변수명 호환 */
const SHIPMENT_SHEET_CHANNELS = SHIPMENT_MATRIX_CHIPS;
/** 출고 현황 표 고정 열 너비(px). 열 순서: 상품코드·브랜드·상품명·… — `styles.css` `.stickyColShip*` 의 width·left·`shipmentTotalMergedCell` 과 동기화. */
const SHIPMENT_MATRIX_COL_BRAND = 103;
const SHIPMENT_MATRIX_COL_CODE = 74;
const SHIPMENT_MATRIX_COL_NAME = 318;
const SHIPMENT_MATRIX_COL_MKT = 99;
const SHIPMENT_MATRIX_COL_SEGMENT = 96;
const SHIPMENT_MATRIX_STICKY_TOTAL_PX =
  SHIPMENT_MATRIX_COL_BRAND +
  SHIPMENT_MATRIX_COL_CODE +
  SHIPMENT_MATRIX_COL_NAME +
  SHIPMENT_MATRIX_COL_MKT +
  SHIPMENT_MATRIX_COL_SEGMENT;
/** 한국·해외·출고 통합 표 날짜/월 열 너비 — `styles.css` `.inventoryTable .dateCol`, `.compareTable .dateCol` 과 동기화. */
const INVENTORY_DATE_COL_PX = 81;
/** 고정 열 합(날짜 제외) — 통합 표 글자 12px(13px 대비 12/13)에 맞춘 값. */
const INVENTORY_KR_FIXED_TOTAL_PX = 761;
const INVENTORY_OVERSEAS_FIXED_TOTAL_PX = 791;
const INVENTORY_OVERSEAS_FIXED_WITH_KR_COMPARE_PX = 891;
const INVENTORY_KR_STICKY_TOTAL_PX = 628;
const INVENTORY_OVERSEAS_STICKY_TOTAL_PX = 787;
const INVENTORY_OVERSEAS_STICKY_WITH_KR_COMPARE_PX = 887;
const COMPARE_TABLE_MIN_VIEWPORT_PX = Math.round((1280 * 12) / 13);
const COMPARE_TABLE_MIN_BASE_PX = Math.round((590 * 12) / 13);
const COMPARE_TABLE_PER_COUNTRY_PX = Math.round((96 * 12) / 13);
/** 한국·해외·출고(일자/월별 뷰) KPI 카드 아이콘 — `frontend/public/image/` (URL `/image/...`). */
const INVENTORY_KPI_CARD_ICON_SRC = {
  uploadedFiles: "/image/업로드된파일.png",
  latestBaseDate: "/image/최신기준일.png",
  inventoryBasis: "/image/재고파악기준.png",
  analyzedSkuCount: "/image/분석상품수.png",
};
/** 재고 비교 탭 KPI (전용 아이콘이 없으면 위와 동일 파일 재사용). */
const COMPARE_KPI_CARD_ICON_SRC = {
  compareCountries: "/image/비교국가수.png",
  latestBaseDate: "/image/최신기준일.png",
  compareSkuCount: "/image/분석상품수.png",
};
/** 상단 메인 탭 아이콘 크기 — Lucide SVG (`styles.css` `.tabIcon`과 맞춤). */
const TOP_TAB_ICON_SIZE_PX = 20;
/** `.kpiCardIcon` 표시 크기와 동기화 (`styles.css`). */
const KPI_CARD_ICON_DISPLAY_PX = 50;
const SKU_MAPPING_FIELDS = [
  { code: "KR", label: "한국", nameKey: "kr_name", skuKey: "kr_sku" },
  { code: "US", label: "미국", nameKey: "us_name", skuKey: "us_sku" },
  { code: "TW", label: "대만", nameKey: "tw_name", skuKey: "tw_sku" },
  { code: "HK", label: "홍콩", nameKey: "hk_name", skuKey: "hk_sku" },
  { code: "JP", label: "일본", nameKey: "jp_name", skuKey: "jp_sku" },
  { code: "SG", label: "싱가폴", nameKey: "sg_name", skuKey: "sg_sku" },
  { code: "DE", label: "독일", nameKey: "de_name", skuKey: "de_sku" },
  { code: "UK", label: "영국", nameKey: "uk_name", skuKey: "uk_sku" },
  { code: "AU", label: "호주", nameKey: "au_name", skuKey: "au_sku" },
  { code: "AE", label: "UAE", nameKey: "ae_name", skuKey: "ae_sku" },
  { code: "VN", label: "베트남", nameKey: "vn_name", skuKey: "vn_sku" },
  { code: "TH", label: "태국", nameKey: "th_name", skuKey: "th_sku" },
];
const SKU_MAPPING_TEMPLATE_COLUMNS = SKU_MAPPING_FIELDS.flatMap(({ nameKey, skuKey }) => [nameKey, skuKey]);
const SKU_MAPPING_OPTIONAL_COLUMNS = ["option", "brand", "barcode", "segment"];
const EMPTY_SKU_MAPPING_FORM = Object.fromEntries([
  ...SKU_MAPPING_FIELDS.flatMap(({ nameKey, skuKey }) => [
    [nameKey, ""],
    [skuKey, ""],
  ]),
  ["brand", ""],
  ["barcode", ""],
  ["segment", ""],
  ["option", ""],
]);

/** 검색 입력 옆 돋보기 — Lucide Search와 같은 원·두께, 대각 손잡이만 더 길게 */
function SearchFieldIcon({ className, size = 16, strokeWidth = 2, ...rest }) {
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={strokeWidth}
      strokeLinecap="round"
      strokeLinejoin="round"
      className={className}
      aria-hidden
      {...rest}
    >
      <circle cx="11" cy="11" r="8" />
      <path d="M 23.15 23.15 L 16.65 16.65" />
    </svg>
  );
}

/** 저장된 발주 수정 시 상품유형 → 프리셋/직접입력 */
function purchaseOrderProductTypeToFields(productType) {
  const t = String(productType || "").trim();
  if (t === "본품" || t === "") return { preset: "본품", custom: "" };
  return { preset: "직접 입력", custom: t };
}

/** 엑셀 일괄 등록 시 ERP 없음 → DB에만 쓰이는 접두사 (화면에서는 en dash –) */
const PO_NO_ERP_PREFIX = "__NO_ERP__";
const PO_NO_SKU_PREFIX = "__NO_SKU__";

/** 발주 엑셀 업로드 — 헤더는 템플릿과 동일한 한글 열 이름 */
const PO_INBOUND_SHEET_HEADERS = [
  "발주일자",
  "ERP PO번호",
  "상품유형",
  "상품번호",
  "제조사",
  "총 발주수량",
  "납품가능일",
  "입고예정일",
  "실제입고일",
  "입고수량",
  "입고여부",
];

/** 엑셀 날짜 직렬(1900 기준) → YYYY-MM-DD. 타임존과 무관하게 캘린더 일만 맞춤. */
function excelSerialToIsoDate(serial) {
  const whole = Math.floor(Number(serial));
  const ms = Math.round((whole - 25569) * 86400 * 1000);
  const d = new Date(ms);
  if (Number.isNaN(d.getTime())) return "";
  return d.toISOString().slice(0, 10);
}

/** treatNumberAsExcelDate: 숫자 열(수량 등)은 엑셀 날짜 직렬과 겹칠 수 있어 날짜 열에서만 직렬→ISO 변환 */
function normalizeInboundSheetCell(value, treatNumberAsExcelDate = false) {
  if (value == null || value === "") return "";
  if (
    treatNumberAsExcelDate &&
    typeof value === "number" &&
    Number.isFinite(value) &&
    value > 20000 &&
    value < 600000
  ) {
    return excelSerialToIsoDate(value);
  }
  if (treatNumberAsExcelDate && value instanceof Date && !Number.isNaN(value.getTime())) {
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }
  if (typeof value === "number" && Number.isFinite(value)) return String(value);
  return String(value).trim();
}

function parsePurchaseOrderInboundSheet(fileArrayBuffer) {
  const errors = [];
  const wb = XLSX.read(fileArrayBuffer, { type: "array", cellDates: false });
  const sheetName = wb.SheetNames[0];
  if (!sheetName) {
    return { rows: [], errors: ["시트가 비어 있습니다."] };
  }
  const ws = wb.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", blankrows: false });
  if (!data.length) {
    return { rows: [], errors: ["데이터가 없습니다."] };
  }
  const headerRow = data[0].map((c) => String(c ?? "").trim());
  const colIndex = {};
  for (let i = 0; i < PO_INBOUND_SHEET_HEADERS.length; i += 1) {
    const want = PO_INBOUND_SHEET_HEADERS[i];
    const idx = headerRow.indexOf(want);
    if (idx < 0) {
      errors.push(`헤더에「${want}」열이 없습니다. 제공한 템플릿을 사용했는지 확인해 주세요.`);
    } else {
      colIndex[want] = idx;
    }
  }
  if (errors.length) {
    return { rows: [], errors };
  }
  const rows = [];
  for (let r = 1; r < data.length; r += 1) {
    const row = data[r];
    if (!row || !row.length) continue;
    const excelRowNum = r + 1;
    const get = (key, dateCol = false) => {
      const j = colIndex[key];
      return j == null ? "" : normalizeInboundSheetCell(row[j], dateCol);
    };
    const pack = {
      row_number: excelRowNum,
      order_date: get("발주일자", true),
      erp_po_number: get("ERP PO번호"),
      product_type: get("상품유형"),
      product_number: get("상품번호"),
      manufacturer: get("제조사"),
      total_quantity: get("총 발주수량"),
      delivery_available_date: get("납품가능일", true) || null,
      expected_inbound_date: get("입고예정일", true) || null,
      actual_inbound: get("실제입고일", true) || null,
      inbound_quantity: get("입고수량"),
      inbound_status: get("입고여부"),
    };
    if (/^\s*예\s*:/i.test(String(pack.order_date || ""))) {
      continue;
    }
    const emptyRow =
      !String(pack.erp_po_number || "").trim() &&
      !String(pack.product_number || "").trim() &&
      !String(pack.order_date || "").trim() &&
      !String(pack.total_quantity || "").trim();
    if (emptyRow) continue;
    rows.push(pack);
  }
  return { rows, errors: [] };
}

async function downloadPoInboundTemplateXlsx() {
  const ExcelJS = (await import("exceljs")).default;
  const FONT_9 = { name: "맑은 고딕", size: 9 };
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("발주_템플릿", { views: [{ showGridLines: true }] });
  const hr = ws.addRow(PO_INBOUND_SHEET_HEADERS);
  hr.height = 18;
  const hint = ws.addRow([
    "예: 2025-08-07",
    "예: CMS20250807",
    "예: 본품",
    "예: 05971",
    "예: 코스모코스",
    "예: 20000",
    "예: 2025-11-07",
    "예: 2025-12-11",
    "",
    "예: 11520",
    "예: 입고 예정",
  ]);
  hint.height = 16;
  PO_INBOUND_SHEET_HEADERS.forEach((_, i) => {
    ws.getColumn(i + 1).width = 16;
  });
  ws.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      cell.font = {
        ...FONT_9,
        bold: rowNumber === 1,
      };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      if (rowNumber === 1) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFFF00" },
        };
      }
    });
  });
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "발주_템플릿.xlsx";
  a.click();
  URL.revokeObjectURL(a.href);
}

/** 재고 대시보드 내보내기: 헤더 노란 배경·볼드, 본문 9pt (ExcelJS) */
async function downloadInventoryDashboardXlsx(filename, sheetName, rows) {
  const ExcelJS = (await import("exceljs")).default;
  const FONT_9 = { name: "맑은 고딕", size: 9 };
  const thinBorder = {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  };
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheetName, { views: [{ showGridLines: true }] });
  if (!rows.length) return;
  const headers = Object.keys(rows[0]);
  ws.addRow(headers);
  for (const r of rows) {
    ws.addRow(headers.map((h) => r[h]));
  }
  ws.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      cell.font = { ...FONT_9, bold: rowNumber === 1 };
      cell.border = thinBorder;
      if (rowNumber === 1) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFFF00" },
        };
      }
    });
  });
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  const base = String(filename || "").replace(/\.xlsx$/i, "");
  a.download = `${base}.xlsx`;
  a.click();
  URL.revokeObjectURL(a.href);
}

/** Excel 기본 열 너비(문자 단위) ≈ 픽셀 대응 — ExcelJS는 문자 폭만 지원해 근사 변환 */
function excelColumnWidthFromPxApprox(px) {
  const p = Number(px);
  if (!Number.isFinite(p) || p <= 0) return 15;
  return Math.round((((p - 5) / 7) + Number.EPSILON) * 100) / 100;
}

/** SKU 탭: 한·미·대·홍 전용 .xlsx를 ZIP으로 내려받기 (파일마다 첫 시트만 업로드 시 읽힘) */
const SKU_MAPPING_TEMPLATE_ZIP_SPECS = [
  {
    code: "KR",
    fileLabel: "한국",
    headers: ["한국 SKU", "브랜드", "한국 상품명", "옵션", "바코드", "구분"],
    columnWidthsPx: {
      "한국 상품명": 200,
      구분: 213,
    },
    exampleHintRow: [
      "예: 05803",
      "예: 푸드올로지",
      "예: 푸드올로지 보틀 500ml",
      "예: 레드 (공백 허용)",
      "예: X0041I3ECT",
      "예: 단종 (단종이 아니면 공백)",
    ],
    headerNotes: {
      구분:
        "선택. 실제 단종만 DB에 반영됩니다. 단종·(X) 단종 → 단종. 단종 예정·알뜰상품·행사상품·작업용 등 그 외 값은 저장하지 않습니다(빈 칸과 동일).",
      옵션: "선택. 있으면 각 국가 상품명 끝에 공백과 함께 붙여 저장됩니다.",
      바코드: "선택. 있으면 상품(item)의 바코드로 저장되며, 업로드 시 값이 있을 때만 갱신됩니다.",
    },
  },
  {
    code: "US",
    fileLabel: "미국",
    headers: ["미국 SKU", "미국 상품명", "한국 SKU"],
    headerNotes: {
      "한국 SKU": "연결할 기존 한국 상품의 SKU입니다. 이 값으로 item을 찾아 US 로케일을 붙입니다.",
    },
  },
  { code: "TW", fileLabel: "대만", headers: ["대만 SKU", "대만 상품명"], headerNotes: {} },
  { code: "HK", fileLabel: "홍콩", headers: ["홍콩 SKU", "홍콩 상품명"], headerNotes: {} },
];

async function buildSkuMappingTemplateWorkbookBuffer(spec) {
  const ExcelJS = (await import("exceljs")).default;
  const FONT_9 = { name: "맑은 고딕", size: 9 };
  const { headers, headerNotes = {}, exampleHintRow, columnWidthsPx = {} } = spec;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1", { views: [{ showGridLines: true }] });
  const hr = ws.addRow(headers);
  hr.height = 18;
  if (Array.isArray(exampleHintRow) && exampleHintRow.length === headers.length) {
    const hint = ws.addRow(exampleHintRow);
    hint.height = 16;
  } else {
    ws.addRow(headers.map(() => ""));
  }
  headers.forEach((h, i) => {
    const px = columnWidthsPx[h];
    ws.getColumn(i + 1).width =
      px != null ? excelColumnWidthFromPxApprox(px) : String(h).length > 12 ? 22 : 15;
  });
  ws.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      cell.font = {
        ...FONT_9,
        bold: rowNumber === 1,
      };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
      if (rowNumber === 1) {
        cell.fill = {
          type: "pattern",
          pattern: "solid",
          fgColor: { argb: "FFFFFF00" },
        };
      }
    });
  });
  headers.forEach((h, i) => {
    const note = headerNotes[h];
    if (note) {
      ws.getRow(1).getCell(i + 1).note = note;
    }
  });
  return wb.xlsx.writeBuffer();
}

async function downloadSkuMappingCountryTemplatesZip() {
  const JSZip = (await import("jszip")).default;
  const zip = new JSZip();

  for (const spec of SKU_MAPPING_TEMPLATE_ZIP_SPECS) {
    const buf = await buildSkuMappingTemplateWorkbookBuffer(spec);
    zip.file(`${spec.fileLabel}_SKU_매핑_템플릿.xlsx`, buf);
  }

  const blob = await zip.generateAsync({ type: "blob" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "한미대홍_SKU_매핑_템플릿.zip";
  a.click();
  URL.revokeObjectURL(a.href);
}

/** 데이터 관리 파일 기준일: 캘린더 피커 툴팁. 기타 화면 텍스트 입력 힌트로도 사용 */
const DATE_TEXT_INPUT_HINT = "YYYY-MM-DD";

function toDateInputValue(raw) {
  const s = String(raw || "").trim();
  if (!s) return "";
  const m = s.match(/^(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : "";
}
const ACTUAL_INBOUND_TEXT_PLACEHOLDER = "YYYY-MM-DD 또는 직접 입력";
/** 저장된 발주(표·발주 수정 모달): 날짜 입력란 힌트 없음. 새 발주 등록 탭은 위 상수 유지 */
const PO_SAVED_DATE_PLACEHOLDER = "";
const PO_SAVED_ACTUAL_INBOUND_PLACEHOLDER = "";
/** 저장 발주 표: 입고여부 O인 차수에서 입고예정일 연필 시 */
const PO_SAVED_EXPECTED_INBOUND_BLOCKED_O_MSG =
  "이미 입고 완료된 건입니다. 입고 여부를 확인해주세요.";

function purchaseOrderErpForDisplay(erp) {
  const s = String(erp || "").trim();
  if (s.startsWith(PO_NO_ERP_PREFIX)) return "\u2013";
  return s;
}

function purchaseOrderSkuForDisplay(sku) {
  const s = String(sku || "").trim();
  if (s.startsWith(PO_NO_SKU_PREFIX)) return "\u2013";
  return s || "\u2013";
}

/** 저장된 발주 표 – 입고여부: 완료 O, 미완료·예정 X → 예정 */
function formatPoInboundStatus(code) {
  const u = String(code || "").trim().toUpperCase();
  if (u === "O") return "O";
  if (u === "X") return "예정";
  const s = String(code || "").trim();
  return s || "–";
}

/** 실제입고일: 비고(예외 입고·무상 입고 등)가 있으면 비고를, 없으면 날짜 */
function formatPoInboundActualDisplay(line) {
  if (!line) return "–";
  const note = String(line.actual_inbound_note || "").trim();
  if (note) return note;
  return line.actual_inbound_date || "–";
}

/** 발주일: 발주 예정 비고가 있으면 그걸, 없으면 날짜 */
function formatPoOrderDateDisplay(po) {
  if (!po) return "–";
  const n = String(po.order_date_note || "").trim();
  if (n) return n;
  return po.order_date || "–";
}

function isOrderDatePlannedNote(s) {
  const t = String(s || "")
    .trim()
    .replace(/\s+/g, " ");
  const c = t.replace(/ /g, "");
  return t === "발주 예정" || c === "발주예정";
}

/** 저장된 발주 표·수정: 차수 1개면 ERP만, 2개 이상만 접미사. 내부용 PO는 _1, _2 */
function formatSavedPoErpCol(po, line, inboundLineCount) {
  const raw = String(po.erp_po_number || "").trim();
  const isNoErp = raw.startsWith(PO_NO_ERP_PREFIX);
  if (line == null) {
    return purchaseOrderErpForDisplay(po.erp_po_number).trim() || "–";
  }
  if (inboundLineCount <= 1) {
    return purchaseOrderErpForDisplay(po.erp_po_number).trim() || "–";
  }
  const n = Number(line.line_no) || 0;
  if (isNoErp) return `_${n}`;
  return `${raw}_${n}`;
}

function formatPoInboundRefPreview(erpInput, line, inboundLineCount) {
  const raw = String(erpInput || "").trim();
  const isNoErp = raw.startsWith(PO_NO_ERP_PREFIX);
  const n = Number(line.line_no) || 0;
  if (inboundLineCount <= 1) {
    if (isNoErp) return "–";
    return raw || "…";
  }
  if (isNoErp) return `_${n}`;
  return `${raw || "…"}_${n}`;
}

function poLineIsInboundPending(line) {
  if (!line) return false;
  return String(line.inbound_status || "").trim().toUpperCase() === "X";
}

function poHasAnyCompletedInbound(po) {
  if (!po?.inbound_lines?.length) return false;
  return (po.inbound_lines || []).some(
    (l) => String(l.inbound_status || "").trim().toUpperCase() === "O"
  );
}

/** 입고 완료(O)면 입고예정일은 비움(표시). */
function formatPoLineExpectedInboundDisplay(line, po) {
  if (!line) return po.expected_inbound_date || "미정";
  if (String(line.inbound_status || "").trim().toUpperCase() === "O") return "–";
  return line.expected_inbound_date || po.expected_inbound_date || "미정";
}

/** 저장된 발주 표: 같은 PO 묶음 2행째부터 빈 칸 (en dash) */
const PO_SAVED_SAME_GROUP_CELL = "\u2013";

const EMPTY_PURCHASE_ORDER_FORM = {
  order_date: "",
  order_date_note: "",
  erp_po_number: "",
  product_type_preset: "본품",
  product_type_custom: "",
  sku: "",
  brand: "",
  product_name: "",
  manufacturer: "",
  total_quantity: "",
  delivery_available_date: "",
  delivery_available_tbd: false,
  expected_inbound_date: "",
  expected_inbound_tbd: false,
};

/** 입고 차수 추가(표 인라인·카드 하단) 공통 초기값 */
const EMPTY_INBOUND_LINE_DRAFT = {
  delivery_available_date: "",
  delivery_available_tbd: false,
  expected_inbound_date: "",
  expected_inbound_tbd: false,
  actual_inbound_input: "",
  quantity: "",
  inbound_status: "X",
};

function isIsoDateOnly(value) {
  return /^\d{4}-\d{2}-\d{2}$/.test(String(value || "").trim());
}

/** 저장된 발주 표 정렬용: 날짜 → 해당 일의 타임스탬프(정오 기준), 파싱 불가면 null */
function parsePoSortDateMs(value) {
  if (value == null || value === "") return null;
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) return null;
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, "0");
    const d = String(value.getDate()).padStart(2, "0");
    const t = new Date(`${y}-${m}-${d}T12:00:00`).getTime();
    return Number.isNaN(t) ? null : t;
  }
  const s = String(value).trim();
  if (!s) return null;
  const isoHead = s.match(/^(\d{4}-\d{2}-\d{2})/);
  if (isoHead) {
    const day = isoHead[1];
    if (!isIsoDateOnly(day)) return null;
    const t = new Date(`${day}T12:00:00`).getTime();
    return Number.isNaN(t) ? null : t;
  }
  const parsed = Date.parse(s);
  if (!Number.isNaN(parsed)) {
    const dt = new Date(parsed);
    const y = dt.getFullYear();
    const m = String(dt.getMonth() + 1).padStart(2, "0");
    const d = String(dt.getDate()).padStart(2, "0");
    const t = new Date(`${y}-${m}-${d}T12:00:00`).getTime();
    return Number.isNaN(t) ? null : t;
  }
  return null;
}

/** 납품가능일: 차수별·발주 공통 값 중 최소·최대(다차수 대응) */
function savedPoDeliverySortBounds(po) {
  if (!po) return { min: null, max: null };
  const lines = po.inbound_lines || [];
  const vals = [];
  const push = (raw) => {
    const t = parsePoSortDateMs(raw);
    if (t != null) vals.push(t);
  };
  if (!lines.length) {
    push(po.delivery_available_date);
  } else {
    for (const line of lines) {
      push(line.delivery_available_date || po.delivery_available_date);
    }
  }
  if (!vals.length) return { min: null, max: null };
  return { min: Math.min(...vals), max: Math.max(...vals) };
}

function poLineIsInboundCompleted(line) {
  return String(line?.inbound_status || "").trim().toUpperCase() === "O";
}

/**
 * 미입고 차수만, 표시와 동일: line.expected_inbound_date ?? po.expected_inbound_date
 * (입고 완료 O는 API에서 line 날짜가 비어 있어도 여기서는 건너뜀)
 */
function savedPoExpectedPendingSortBounds(po) {
  if (!po) return { min: null, max: null };
  const lines = [...(po.inbound_lines || [])].sort(
    (a, b) => (Number(a.line_no) || 0) - (Number(b.line_no) || 0)
  );
  const vals = [];
  const pushEffectivePendingLine = (line) => {
    if (poLineIsInboundCompleted(line)) return;
    const raw = line.expected_inbound_date || po.expected_inbound_date;
    const t = parsePoSortDateMs(raw);
    if (t != null) vals.push(t);
  };
  if (!lines.length) {
    const t = parsePoSortDateMs(po.expected_inbound_date);
    if (t != null) vals.push(t);
  } else {
    for (const line of lines) pushEffectivePendingLine(line);
  }
  if (!vals.length) return { min: null, max: null };
  return { min: Math.min(...vals), max: Math.max(...vals) };
}

/**
 * 임박순: line_no 순으로 볼 때 첫 미입고 차수부터, 날짜가 잡힌 첫 입고예정일.
 * (뒤 차수가 더 이른 날이어도, 앞 차수가 아직 열려 있으면 그 차수 기준 — 화면에서 위에 보이는 '다음 입고'와 맞춤)
 */
function savedPoExpectedInboundSortKeyImminent(po) {
  if (!po) return null;
  const lines = [...(po.inbound_lines || [])].sort(
    (a, b) => (Number(a.line_no) || 0) - (Number(b.line_no) || 0)
  );
  if (!lines.length) return parsePoSortDateMs(po.expected_inbound_date);
  for (const line of lines) {
    if (poLineIsInboundCompleted(line)) continue;
    const raw = line.expected_inbound_date || po.expected_inbound_date;
    const t = parsePoSortDateMs(raw);
    if (t != null) return t;
  }
  return null;
}

function cmpSavedPoNullableNumber(a, b, desc) {
  if (a == null && b == null) return 0;
  if (a == null) return 1;
  if (b == null) return -1;
  const d = desc ? b - a : a - b;
  return d < 0 ? -1 : d > 0 ? 1 : 0;
}

function compareSavedPoSpreadsheetGroups(a, b, mode) {
  let c = 0;
  switch (mode) {
    case "order_date_desc":
      c = cmpSavedPoNullableNumber(
        parsePoSortDateMs(a.po.order_date),
        parsePoSortDateMs(b.po.order_date),
        true
      );
      break;
    case "delivery_asc": {
      const ba = savedPoDeliverySortBounds(a.po);
      const bb = savedPoDeliverySortBounds(b.po);
      c = cmpSavedPoNullableNumber(ba.min, bb.min, false);
      break;
    }
    case "delivery_desc": {
      const ba = savedPoDeliverySortBounds(a.po);
      const bb = savedPoDeliverySortBounds(b.po);
      c = cmpSavedPoNullableNumber(ba.max, bb.max, true);
      break;
    }
    case "expected_asc": {
      const ka = savedPoExpectedInboundSortKeyImminent(a.po);
      const kb = savedPoExpectedInboundSortKeyImminent(b.po);
      c = cmpSavedPoNullableNumber(ka, kb, false);
      break;
    }
    case "expected_desc": {
      const ba = savedPoExpectedPendingSortBounds(a.po);
      const bb = savedPoExpectedPendingSortBounds(b.po);
      c = cmpSavedPoNullableNumber(ba.max, bb.max, true);
      break;
    }
    case "order_date_asc":
    default:
      c = cmpSavedPoNullableNumber(
        parsePoSortDateMs(a.po.order_date),
        parsePoSortDateMs(b.po.order_date),
        false
      );
      break;
  }
  if (c !== 0) return c;
  const erpa = String(a.po.erp_po_number || "");
  const erpb = String(b.po.erp_po_number || "");
  if (erpa !== erpb) return erpa.localeCompare(erpb);
  return String(a.po.sku || "").localeCompare(String(b.po.sku || ""));
}

function normalizeActualInboundInput(rawInput) {
  const raw = String(rawInput || "").trim();
  if (!raw) return { actual_inbound_date: null, actual_inbound_note: null };
  if (isIsoDateOnly(raw)) {
    return { actual_inbound_date: raw, actual_inbound_note: null };
  }
  return { actual_inbound_date: null, actual_inbound_note: raw };
}

function normalizeOrderDateInput(rawInput) {
  const raw = String(rawInput || "").trim();
  if (!raw) return { order_date: null, order_date_note: null };
  if (isIsoDateOnly(raw)) return { order_date: raw, order_date_note: null };
  if (isOrderDatePlannedNote(raw)) return { order_date: null, order_date_note: "발주 예정" };
  return { order_date: null, order_date_note: raw };
}

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

/** 발주 상품유형: SKU 수기 브랜드와 동일(목록 ▼ ↔ 직접입력 칸 + 목록에서 선택) */
function PurchaseProductTypeField({ preset, custom, onPresetChange, onCustomChange }) {
  const rootRef = useRef(null);
  const customInputRef = useRef(null);
  const [open, setOpen] = useState(false);
  const isDirectMode = preset === "직접 입력";

  useEffect(() => {
    if (!open) return;
    const onDoc = (e) => {
      if (rootRef.current && !rootRef.current.contains(e.target)) {
        setOpen(false);
      }
    };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [open]);

  useEffect(() => {
    if (isDirectMode) customInputRef.current?.focus();
  }, [isDirectMode]);

  if (isDirectMode) {
    return (
      <div className="manualBrandCombobox manualBrandComboboxCustom" ref={rootRef}>
        <div className="manualBrandCustomRow">
          <input
            ref={customInputRef}
            type="text"
            className="manualBrandCustomInput"
            value={custom}
            onChange={(e) => onCustomChange(e.target.value)}
            placeholder="유형 직접 입력"
            autoComplete="off"
          />
          <button
            type="button"
            className="manualBrandToPresetsBtn"
            onClick={() => {
              onPresetChange("본품");
              onCustomChange("");
            }}
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
        onClick={() => setOpen((o) => !o)}
        aria-expanded={open}
        aria-haspopup="listbox"
      >
        <span className="manualBrandTriggerText">{preset}</span>
        <span className="manualBrandChevron" aria-hidden="true">
          ▼
        </span>
      </button>
      {open ? (
        <div className="manualBrandPopover poProductTypePopover" role="listbox">
          <ul className="manualBrandList">
            <li>
              <button
                type="button"
                className="manualBrandOption"
                onClick={() => {
                  onPresetChange("본품");
                  onCustomChange("");
                  setOpen(false);
                }}
              >
                본품
              </button>
            </li>
          </ul>
          <button
            type="button"
            className="manualBrandOption manualBrandOptionDirect"
            onClick={() => {
              onPresetChange("직접 입력");
              setOpen(false);
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

/** 화면 표시용 정수·천 단위 콤마 (내보내기·저장용은 `toFixed` 등 유지) */
function formatInt(value) {
  if (value === null || value === undefined) return "-";
  const parsed = Number(value);
  if (Number.isNaN(parsed)) return String(value);
  return parsed.toLocaleString("ko-KR", { maximumFractionDigits: 0 });
}

/** 재고 표 날짜 열 헤더 문자열 — `renderDateHeader`와 동일한 한 셀 표기 */
function formatInventoryDateHeaderForExport(dateKey) {
  const [y, m, d] = String(dateKey).split("-");
  if (y && m && d) return `${y} - ${m}-${d}`;
  return String(dateKey);
}

function toSigned(value, digits = 0) {
  if (value === null || value === undefined) return "-";
  const parsed = Number(value);
  if (Number.isNaN(parsed)) return String(value);
  const formatted = parsed.toLocaleString("ko-KR", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
  if (parsed > 0) return `+${formatted}`;
  return formatted;
}

function toPercent(value, digits = 1) {
  if (value === null || value === undefined) return "-";
  const parsed = Number(value);
  if (Number.isNaN(parsed)) return "-";
  const formatted = parsed.toLocaleString("ko-KR", {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
  if (parsed > 0) return `+${formatted}%`;
  return `${formatted}%`;
}

function detectCountry(name = "") {
  const n = name.toLowerCase();
  if (n.includes("tw") || n.includes("taiwan") || n.includes("대만")) return "TW";
  if (n.includes("hongkong") || n.includes("hong kong") || n.includes("香港") || n.includes("홍콩")) return "HK";
  if (n.includes("us") || n.includes("usa") || n.includes("미국")) return "US";
  if (n.includes("vn") || n.includes("vietnam") || n.includes("베트남")) return "VN";
  if (n.includes("sg") || n.includes("singapore") || n.includes("싱가포르") || n.includes("싱가폴")) return "SG";
  if (n.includes("au") || n.includes("australia") || n.includes("호주")) return "AU";
  if (n.includes("uk") || n.includes("england") || n.includes("britain") || n.includes("영국")) return "UK";
  if (n.includes("ae") || n.includes("uae") || n.includes("dubai") || n.includes("아랍에미리트")) return "AE";
  if (n.includes("jp") || n.includes("japan") || n.includes("일본")) return "JP";
  if (n.includes("de") || n.includes("germany") || n.includes("독일")) return "DE";
  if (n.includes("th") || n.includes("thailand") || n.includes("태국")) return "TH";
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

/** 한국 재고 '오늘' 스냅샷 판별용 – UTC가 아니라 Asia/Seoul 달력 날짜(YYYY-MM-DD) */
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
  const lv = row?.level;
  let levelStr = "";
  if (lv != null && lv !== "") {
    levelStr = String(lv).trim().replace(/\.0+$/, "");
  }
  return {
    supplier: String(row?.supplier || "").trim(),
    level: levelStr,
    warehouse: String(row?.warehouse || row?.category || "").trim(),
  };
}

function getCompareMetaLabel(row) {
  const meta = getCompareMeta(row);
  const parts = [];
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

/** 한국 행(메타 비어 있음) 먼저 등록된 뒤 해외 행이 오면 더 자세한 구분 문자열로 갱신 */
function pickRicherCompareMetaLabel(prev, next) {
  const a = String(prev || "").trim();
  const b = String(next || "").trim();
  const ac = a ? a.split(" / ").filter(Boolean).length : 0;
  const bc = b ? b.split(" / ").filter(Boolean).length : 0;
  if (bc > ac) return b;
  if (ac > bc) return a;
  return b || a;
}

/** identity 키 끝의 level·warehouse (브랜드는 구분 표시에 넣지 않음) */
function metaLabelFromCompareIdentityKey(compareKey) {
  const parts = String(compareKey || "").split("::");
  if (parts.length < 5) return "";
  const levelRaw = (parts[parts.length - 2] || "").trim().replace(/\.0+$/, "");
  const warehouse = (parts[parts.length - 1] || "").trim();
  const out = [];
  if (levelRaw) out.push(/^L/i.test(levelRaw) ? levelRaw : `L${levelRaw}`);
  if (warehouse) out.push(warehouse);
  return out.join(" / ");
}

function countryLabel(code = "KR") {
  if (code === "SHIPMENT") return "출고";
  if (code === "KR") return "한국";
  if (code === "TW") return "대만";
  if (code === "HK") return "홍콩";
  if (code === "US") return "미국";
  if (code === "VN") return "베트남";
  if (code === "SG") return "싱가폴";
  if (code === "AU") return "호주";
  if (code === "UK") return "영국";
  if (code === "AE") return "UAE";
  if (code === "JP") return "일본";
  if (code === "DE") return "독일";
  if (code === "TH") return "태국";
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
        sku: sku || "–",
        description: name || "–",
      };
    })
    .filter(Boolean);
}

function getLatestDateKey(dateKeys = []) {
  if (!dateKeys.length) return "";
  return [...dateKeys].sort().at(-1) || "";
}

/** 재고 표 정렬: 화면에 보이는 날짜 열 중 가장 최근 일자의 수량(내림차순 기준) */
function inventoryRowLatestQty(row, dateCols) {
  const latest = getLatestDateKey(dateCols);
  if (!latest) return 0;
  return Number(row[latest] || 0);
}

function compareInventoryRowsByLatestQty(a, b, dateCols) {
  const dq = inventoryRowLatestQty(b, dateCols) - inventoryRowLatestQty(a, dateCols);
  if (dq !== 0) return dq;
  const c = String(getRowSku(a) || "").localeCompare(String(getRowSku(b) || ""), "ko");
  if (c !== 0) return c;
  return String(a.description || "").localeCompare(String(b.description || ""), "ko");
}

/** 출고 셀: 숫자 또는 "20 · 창고이동 5" 등에서 합계용 숫자 추출 */
function shipmentCellNumericTotal(val) {
  if (val == null || val === "") return 0;
  if (typeof val === "number" && !Number.isNaN(val)) return val;
  const s = String(val).trim();
  if (!s || s === "-") return 0;
  let sum = 0;
  const re = /(\d+)/g;
  let m;
  while ((m = re.exec(s)) !== null) sum += Number(m[1]) || 0;
  return sum;
}

/** 출고 와이드 표: 날짜 셀 표시(0이면 –, 창고이동 병기 문자열은 유지) */
function formatShipmentWideDateCell(val) {
  const n = shipmentCellNumericTotal(val);
  if (n === 0) return "–";
  if (typeof val === "string" && val.includes("창고이동")) return val;
  if (typeof val === "number" && !Number.isNaN(val)) return formatInt(val);
  return String(val ?? "–");
}

/** 출고 일자별 셀 호버: DB에 저장된 판매처 원문 기준 분해 */
function shipmentDateCellVendorTitle(row, dateKey) {
  const br = row?.date_vendor_breakdown?.[dateKey];
  if (!Array.isArray(br) || !br.length) return undefined;
  const lines = [];
  for (const part of br) {
    const vendor = String(part?.vendor ?? "");
    const sale = Number(part?.sale) || 0;
    const wh = Number(part?.wh) || 0;
    if (sale) lines.push(`${vendor}: ${formatInt(sale)}`);
    if (wh) lines.push(`${vendor} (창고이동): ${formatInt(wh)}`);
  }
  return lines.length ? lines.join("\n") : undefined;
}

/** 출고 현황(매트릭스) 수량 셀: 판매처 분해가 있으면 우선, 없으면 값 출처 안내 */
function shipmentMatrixQtyCellTitle(row, kind, ctx) {
  if (kind === "day") {
    const vendorTitle = shipmentDateCellVendorTitle(row, ctx.dateKey);
    if (vendorTitle) return vendorTitle;
    const ch = String(row?.channel || "").trim() || "선택 시트";
    const dayL = ctx.dayHeadLabel || ctx.dateKey;
    const raw = row[ctx.dateKey];
    const n = shipmentCellNumericTotal(raw);
    const showVal =
      raw != null &&
      String(raw).trim() !== "" &&
      String(raw).trim() !== "–" &&
      n !== 0;
    if (!row?.is_total) {
      const lines = [`출고 매트릭스 엑셀 시트 「${ch}」`, `${dayL} 열에 저장된 수량`];
      if (showVal) {
        lines.push(
          `표시: ${typeof raw === "string" && raw.includes("창고이동") ? raw : formatInt(n)}`,
        );
      } else {
        lines.push("저장된 수량 없음(0 또는 빈 셀)");
      }
      return lines.join("\n");
    }
    const lines = [`${dayL} — 이 날짜 열 전체 상품 합계`, `합계: ${formatInt(n)}`];
    return lines.join("\n");
  }
  if (kind === "month") {
    const ch = String(row?.channel || "").trim() || "선택 시트";
    const ml = ctx.monthColLabel || `${ctx.monthKey}월 합계`;
    const v = row.month_totals?.[ctx.monthKey];
    const nv = Number(v) || 0;
    if (!row?.is_total) {
      const lines = [`출고 매트릭스 엑셀 시트 「${ch}」`, `「${ml}」 열에 저장된 값`];
      lines.push(nv !== 0 ? `표시: ${formatInt(nv)}` : "저장된 수량 없음(0 또는 빈 셀)");
      return lines.join("\n");
    }
    return [`${ml} — 전체 상품 합계`, `합계: ${formatInt(nv)}`].join("\n");
  }
  return undefined;
}

/** 출고 차트: 선택 SKU의 모든 행·일자에 대해 판매처(창고이동) 합산 */
function aggregateShipmentChartVendorTotals(rows, dateKeys) {
  const by = new Map();
  for (const row of rows || []) {
    const br = row.date_vendor_breakdown;
    if (!br || typeof br !== "object") continue;
    for (const dk of dateKeys) {
      const parts = br[dk];
      if (!Array.isArray(parts)) continue;
      for (const part of parts) {
        const v = String(part?.vendor ?? "").trim() || "(미상)";
        const sale = Number(part?.sale) || 0;
        const wh = Number(part?.wh) || 0;
        const cur = by.get(v) || { sale: 0, wh: 0 };
        by.set(v, { sale: cur.sale + sale, wh: cur.wh + wh });
      }
    }
  }
  return by;
}

/** 동일 SKU 다행·일자별 차트 점: 판매처 분해 합산(분해 없으면 행 channel로 합계 표시) */
function shipmentDateCellVendorTitleAggregated(rows, dateKey) {
  if (!rows?.length) return undefined;
  const byVendor = new Map();
  for (const row of rows) {
    const br = row?.date_vendor_breakdown?.[dateKey];
    if (!Array.isArray(br)) continue;
    for (const part of br) {
      const vendor = String(part?.vendor ?? "");
      const sale = Number(part?.sale) || 0;
      const wh = Number(part?.wh) || 0;
      if (!vendor && !sale && !wh) continue;
      const cur = byVendor.get(vendor) || { sale: 0, wh: 0 };
      byVendor.set(vendor, { sale: cur.sale + sale, wh: cur.wh + wh });
    }
  }
  const lines = [];
  for (const [vendor, { sale, wh }] of byVendor) {
    if (sale) lines.push(`${vendor}: ${formatInt(sale)}`);
    if (wh) lines.push(`${vendor} (창고이동): ${formatInt(wh)}`);
  }
  if (lines.length) return lines.join("\n");
  const byCh = new Map();
  for (const row of rows) {
    const n = shipmentCellNumericTotal(row[dateKey]);
    if (!n) continue;
    const ch = String(row.channel || "").trim() || "(미지정)";
    byCh.set(ch, (byCh.get(ch) || 0) + n);
  }
  const fb = [];
  for (const [ch, q] of byCh) {
    if (q) fb.push(`${ch}: ${formatInt(q)}`);
  }
  return fb.length ? fb.join("\n") : undefined;
}

/** 출고: 선택 기간 내 날짜 열 합계(정렬·합계용) */
function shipmentRowSumInDateCols(row, dateCols) {
  return dateCols.reduce((s, d) => s + shipmentCellNumericTotal(row[d]), 0);
}

/** 재고 비교 표: 선택 기준일 기준 각국 수량 합(없는 국가·null은 0) — 내림차순 정렬용 */
function comparePivotRowTotalQty(row) {
  let s = 0;
  const kr = row["한국 현 재고"];
  if (kr != null) s += Number(kr) || 0;
  for (const code of OVERSEAS_UPLOAD_COUNTRIES) {
    const v = row[countryLabel(code)];
    if (v != null) s += Number(v) || 0;
  }
  return s;
}

function comparePivotRowsByTotalQty(a, b) {
  const d = comparePivotRowTotalQty(b) - comparePivotRowTotalQty(a);
  if (d !== 0) return d;
  return String(a["상품코드"] || "").localeCompare(String(b["상품코드"] || ""), "ko");
}

/** hydratePersistedState 실패 시 – 네트워크/DB/500 구분에 도움 */
function formatPersistedLoadError(err, apiBase) {
  const detail = err?.response?.data?.detail;
  const fromApi = Array.isArray(detail) ? detail.join("\n") : detail != null ? String(detail) : "";
  const status = err?.response?.status;
  const code = err?.code;
  const message = String(err?.message || "");

  if (code === "ERR_NETWORK" || message === "Network Error") {
    const baseHint =
      apiBase != null && String(apiBase).trim() !== ""
        ? apiBase
        : typeof window !== "undefined"
          ? `${window.location.origin} (동일 출처 /api/...)`
          : "동일 출처 (/api/...)";
    return [
      "백엔드 API에 연결할 수 없습니다.",
      `요청 기준 URL: ${baseHint}`,
      "Nginx에 /api/ 프록시·uvicorn(8000) 실행, 또는 오래된 프론트 번들이면 재빌드·재배포를 확인하세요.",
    ].join("\n");
  }

  if (fromApi) return fromApi;
  const parts = [];
  if (status) parts.push(`HTTP ${status}`);
  if (message) parts.push(message);
  return parts.filter(Boolean).join(" · ") || "저장된 데이터를 불러오는 중 오류";
}

/** 백엔드 UNKNOWN_SKUS(재고·출고 통합) — 사용자 알림·에러 배너용 한 덩어리 문장 */
function formatUnknownSkusUserMessage(detail) {
  if (!detail || typeof detail !== "object" || detail.code !== "UNKNOWN_SKUS") return null;
  const intro = String(detail.message || "").trim();
  const blocks = [];
  if (intro) blocks.push(intro);
  const items = Array.isArray(detail.items) ? detail.items : [];
  if (items.length) {
    blocks.push("", "등록이 필요한 상품코드:");
    for (const it of items) {
      const cc = String(it.country_code || it.country || "").trim();
      const sku = it.sku != null ? String(it.sku) : "";
      const where = cc ? `${countryLabel(cc)} (${cc})` : "";
      blocks.push(where ? `- ${where} · ${sku}` : `- ${sku}`);
    }
  }
  const skus = Array.isArray(detail.skus) ? detail.skus : [];
  if (skus.length && !items.length) {
    blocks.push("", "등록이 필요한 상품코드:");
    for (const s of skus) blocks.push(`- ${s}`);
  }
  const emptyRows = Array.isArray(detail.sku_empty_rows) ? detail.sku_empty_rows : [];
  if (emptyRows.length) {
    blocks.push("", "함께 확인할 행(비어 있거나 인식되지 않은 SKU):");
    for (const line of emptyRows) blocks.push(String(line));
  }
  return blocks.join("\n");
}

/** 재고 통합: S3 직접 PUT·네트워크 오류 등으로 `response.data.detail`이 없을 때 메시지 보강 */
function formatInventoryAggregateFailureMessage(err, fallback = "재고 통합 중 오류") {
  const res = err?.response;
  const detail = res?.data?.detail;

  if (Array.isArray(detail)) {
    const lines = detail.map((item) =>
      typeof item === "string" ? item : typeof item?.msg === "string" ? item.msg : JSON.stringify(item),
    );
    const joined = lines.join("\n").trim();
    if (joined) return joined;
  }
  if (typeof detail === "string" && detail.trim()) return detail.trim();
  if (detail != null && typeof detail === "object") {
    const m = String(detail.message || "").trim();
    if (m) return m;
    try {
      return JSON.stringify(detail);
    } catch {
      /* ignore */
    }
  }

  const status = res?.status;
  const raw = res?.data;
  if (typeof raw === "string" && raw.trim()) {
    const snippet = raw.replace(/\s+/g, " ").trim().slice(0, 280);
    return `응답${status != null ? ` HTTP ${status}` : ""}: ${snippet}`;
  }

  const msg = String(err?.message || "").trim();
  const parts = [];
  if (status != null) parts.push(`HTTP ${status}`);
  if (msg) parts.push(msg);
  if (parts.length) {
    let out = parts.join(" — ");
    if (/network error/i.test(msg) || res == null) {
      out +=
        "\n\n· 브라우저가 S3에 직접 업로드(PUT)합니다. S3 버킷 CORS에 이 페이지 출처(예: http://localhost:5173, 배포 도메인)와 PUT 허용을 넣었는지 확인하세요.";
    }
    return out;
  }
  return fallback;
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

/** 출고 차트: 일자·월 누적 수량 라인 (points: { key, qty, labelShort? }[]) */
function buildShipmentQtyLineChart(points, options = {}) {
  if (!points?.length) return null;
  const {
    chartWidth: chartWidthOpt,
    chartHeight: chartHeightOpt,
    xLabelStep: xLabelStepOpt,
    showPointValueLabels: showPointValueLabelsOpt,
    paddingXLeft: padLeftOpt,
    paddingXRight: padRightOpt,
    paddingX: paddingXSym = 30,
  } = options;
  const series = points.map((p) => ({
    dateKey: p.key,
    qty: Number(p.qty) || 0,
    labelShort:
      p.labelShort ??
      (p.key.length >= 10 ? String(Number(p.key.slice(8))) : p.key.slice(5)),
    vendorTooltip: p.vendorTooltip,
  }));
  const maxQty = Math.max(...series.map((p) => p.qty), 1);
  const chartWidth =
    chartWidthOpt ?? (series.length > 24 ? 960 : series.length > 14 ? 800 : 560);
  const chartHeight = chartHeightOpt ?? 240;
  const padL = padLeftOpt ?? paddingXSym;
  const padR = padRightOpt ?? paddingXSym;
  const paddingY = 20;
  const innerWidth = chartWidth - padL - padR;
  const innerHeight = chartHeight - paddingY * 2;
  const pts = series.map((point, idx) => {
    const x =
      series.length === 1
        ? padL + innerWidth / 2
        : padL + (idx / Math.max(series.length - 1, 1)) * innerWidth;
    const y = paddingY + ((maxQty - point.qty) / maxQty) * innerHeight;
    return { ...point, x, y };
  });
  const xLabelStep = xLabelStepOpt ?? Math.max(1, Math.ceil(series.length / 8));
  const showPointValueLabels = showPointValueLabelsOpt ?? series.length <= 18;
  const totalQty = series.reduce((s, p) => s + p.qty, 0);
  return {
    series,
    points: pts,
    totalQty,
    maxQty,
    chartWidth,
    chartHeight,
    padL,
    padR,
    paddingY,
    xLabelStep,
    showPointValueLabels,
  };
}

/** YYYY-MM → 해당 달 일수 (윤년 반영) */
function daysInCalendarMonth(ym) {
  const parts = String(ym || "").split("-");
  if (parts.length < 2) return 31;
  const y = Number(parts[0]);
  const m = Number(parts[1]);
  if (!Number.isFinite(y) || !Number.isFinite(m) || m < 1 || m > 12) return 31;
  return new Date(y, m, 0).getDate();
}

/** YYYY-MM → "2026년 3월" (select 표시용, value는 그대로 YYYY-MM) */
function formatShipmentMonthLabelKorean(ym) {
  const s = String(ym || "").trim();
  const m = /^(\d{4})-(\d{2})$/.exec(s);
  if (!m) return s;
  const year = Number(m[1]);
  const month = Number(m[2]);
  if (!Number.isFinite(year) || !Number.isFinite(month) || month < 1 || month > 12) return s;
  return `${year}년 ${month}월`;
}

const SHIPMENT_DAY_HEADER_WEEKDAY_KO = ["일", "월", "화", "수", "목", "금", "토"];

/** YYYY-MM-DD → `M/D(요일)` (차트 분석 일자 열 헤더) */
function formatShipmentDayHeaderLabel(dateKey) {
  const s = String(dateKey || "").trim();
  const p = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s);
  if (!p) return s;
  const y = Number(p[1]);
  const mo = Number(p[2]);
  const d = Number(p[3]);
  if (!Number.isFinite(y) || !Number.isFinite(mo) || !Number.isFinite(d)) return s;
  const dt = new Date(y, mo - 1, d);
  if (Number.isNaN(dt.getTime())) return s;
  const w = SHIPMENT_DAY_HEADER_WEEKDAY_KO[dt.getDay()];
  return `${mo}/${d}(${w})`;
}

/** 차트 월 칩: 1월~12월 (연도는 shipmentChartChipYear) */
const SHIPMENT_MONTH_CHIP_NUMS = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12];

/** 출고 차트: 일자·월 라인 SVG (차트 분석 대시보드에서 공통 사용) */
function ShipmentQtyLineChartSvg({ lineData, dailyVendorTooltips, ariaLabel }) {
  if (!lineData) return null;
  const pl = lineData.padL ?? 30;
  const pr = lineData.padR ?? 30;
  const py = lineData.paddingY ?? 20;
  const yAxisX = pl;
  const yLabelX = Math.max(4, pl - 10);
  const chartBottomY = lineData.chartHeight - py;
  return (
    <svg
      viewBox={`0 0 ${lineData.chartWidth} ${lineData.chartHeight}`}
      preserveAspectRatio="xMidYMax meet"
      className={`trendChart shipmentLineChartSvg${dailyVendorTooltips ? " shipmentDailyLineSvg" : ""}`}
      role="img"
      aria-label={ariaLabel}
    >
      {[0.25, 0.5, 0.75].map((t) => {
        const y = py + (1 - t) * (lineData.chartHeight - py * 2);
        const gridQty = Math.round(lineData.maxQty * t);
        const showQtyLabel = gridQty > 0;
        return (
          <g key={`grid-${t}`}>
            <line
              x1={yAxisX}
              y1={y}
              x2={lineData.chartWidth - pr}
              y2={y}
              className="shipmentChartSvgGridLine"
            />
            {showQtyLabel ? (
              <text x={yLabelX} y={y + 4} textAnchor="end" className="shipmentChartYGridLabel">
                {formatInt(gridQty)}
              </text>
            ) : null}
          </g>
        );
      })}
      <text
        x={yLabelX}
        y={lineData.chartHeight - 28}
        textAnchor="end"
        className="shipmentChartYGridLabel shipmentChartYAxisOrigin"
      >
        0
      </text>
      <line
        x1={yAxisX}
        y1={chartBottomY}
        x2={lineData.chartWidth - pr}
        y2={chartBottomY}
        className="trendAxis"
      />
      <line x1={yAxisX} y1={py} x2={yAxisX} y2={chartBottomY} className="trendAxis" />
      {lineData.points.length > 1 && (
        <polyline
          fill="none"
          stroke="currentColor"
          strokeWidth="1.25"
          points={lineData.points.map((point) => `${point.x},${point.y}`).join(" ")}
          className="trendLine shipmentShipmentTrendLine"
        />
      )}
      {lineData.points.map((point, pointIdx) => (
        <g key={point.dateKey}>
          <title>
            {dailyVendorTooltips && point.vendorTooltip
              ? `${point.dateKey} | 합계 ${formatInt(point.qty)}\n${point.vendorTooltip}`
              : `${point.dateKey} | ${formatInt(point.qty)}`}
          </title>
          <circle cx={point.x} cy={point.y} r="2.5" strokeWidth="1" className="trendDot shipmentShipmentDot" />
          {lineData.showPointValueLabels && (
            <text x={point.x} y={point.y - 11} textAnchor="middle" className="trendDotLabel">
              {formatInt(point.qty)}
            </text>
          )}
          {pointIdx % lineData.xLabelStep === 0 && (
            <text
              x={point.x}
              y={lineData.chartHeight - 2}
              textAnchor="middle"
              className="trendXAxisLabel"
            >
              {point.labelShort ?? point.dateKey}
            </text>
          )}
        </g>
      ))}
    </svg>
  );
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
  const [shipmentFileEntries, setShipmentFileEntries] = useState([]);
  const [shipmentUploadInputKey, setShipmentUploadInputKey] = useState(0);
  const [shipmentLoading, setShipmentLoading] = useState(false);
  const [countryTabMode, setCountryTabMode] = useState("KR");
  /** 출고 현황: 선택한 칩(엑셀 시트명). 빈 값이면 서버 기본 채널 */
  const [shipmentMatrixChannel, setShipmentMatrixChannel] = useState("");
  const shipmentMatrixChannelRef = useRef("");
  useEffect(() => {
    shipmentMatrixChannelRef.current = shipmentMatrixChannel;
  }, [shipmentMatrixChannel]);
  /** 출고: 일자 열 월 필터 — "" 이면 아래 useEffect로 데이터에 맞는 월로 보정, "__ALL__" 이면 전체 기간 */
  const [shipmentDisplayMonth, setShipmentDisplayMonth] = useState("");
  /** 출고: 1~12월 버튼에 쓰는 연도(데이터에 여러 연도가 있을 때만 드롭다운으로 변경) */
  const [shipmentMonthStripYear, setShipmentMonthStripYear] = useState(() => new Date().getFullYear());
  /** 출고: 출고 현황(매트릭스 표) | 차트 분석 */
  const [shipmentViewMode, setShipmentViewMode] = useState("status");
  /** 차트 일자별: 볼 달(YYYY-MM) — 표 상단 월 드롭다운과 별개 */
  const [shipmentChartMonth, setShipmentChartMonth] = useState("");
  /** 차트 탭 전용 검색(히어로) — 검색 전에는 차트를 띄우지 않음 */
  const [shipmentChartSearchKeyword, setShipmentChartSearchKeyword] = useState("");
  const [shipmentChartSelectedSku, setShipmentChartSelectedSku] = useState("");
  /** 차트 분석: `all_channels=1` 조회 결과(현황 탭 scope는 단일 시트만 유지) */
  const [shipmentChartAllChannels, setShipmentChartAllChannels] = useState(null);
  const [selectedOverseasCountry, setSelectedOverseasCountry] = useState(OVERSEAS_UPLOAD_COUNTRIES[0]);
  const [selectedKrTrendRowKey, setSelectedKrTrendRowKey] = useState("");
  const [selectedOverseasTrendRowKey, setSelectedOverseasTrendRowKey] = useState("");
  const [compareSelectedDate, setCompareSelectedDate] = useState("");
  const [showCautionModal, setShowCautionModal] = useState(false);
  const [mypageView, setMypageView] = useState(false);
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
  const [purchaseOrders, setPurchaseOrders] = useState([]);
  const [purchaseOrdersLoading, setPurchaseOrdersLoading] = useState(false);
  const [purchaseOrderError, setPurchaseOrderError] = useState("");
  const [purchaseOrderSuccess, setPurchaseOrderSuccess] = useState("");
  const [poForm, setPoForm] = useState({ ...EMPTY_PURCHASE_ORDER_FORM });
  const [skuResolveHint, setSkuResolveHint] = useState("");
  const [inboundDrafts, setInboundDrafts] = useState({});
  /** 저장된 발주 표: 발주별 하단 인라인 신규 입고 차수 초안 (orderId 문자열 키) */
  const [savedPoNewLineDraftByOrderId, setSavedPoNewLineDraftByOrderId] = useState({});
  /** 발주 기록 탭 내부: 발주 파일 업로드 | 새 등록 | 저장된 목록 (기본: 저장된 발주) */
  const [purchaseOrderSubTab, setPurchaseOrderSubTab] = useState("saved");
  const [poInboundUploadBusy, setPoInboundUploadBusy] = useState(false);
  const [poInboundFileKey, setPoInboundFileKey] = useState(0);
  const poSavedTableScrollRef = useRef(null);
  const poSavedHeaderScrollRef = useRef(null);
  const poSavedHeadTableRef = useRef(null);
  const poSavedBodyTableRef = useRef(null);
  const savedPoNewLineRowRefs = useRef({});
  const poMemoAutosaveTimerRef = useRef(null);
  const poMemoModalIdsRef = useRef(null);
  const poMemoDraftRef = useRef("");
  /** 입고 줄 0개일 때 자동 생성 POST 중복 방지 */
  const poInboundSeedLockRef = useRef(new Set());
  const poSavedTopScrollRef = useRef(null);
  const poSavedScrollSyncingRef = useRef(false);
  const poSubTabsBarRef = useRef(null);
  const poSavedFilterBarRef = useRef(null);
  const [poOrderStickyHeights, setPoOrderStickyHeights] = useState({ subTabs: 0, savedFilter: 0 });
  const [poSavedTopScrollWidth, setPoSavedTopScrollWidth] = useState(0);
  const [poSavedShowTopScroll, setPoSavedShowTopScroll] = useState(false);
  const [poSavedTopStripHeight, setPoSavedTopStripHeight] = useState(0);
  const [savedPoSearch, setSavedPoSearch] = useState("");
  /** 저장된 발주: 발주일 기준 기간 (전체 | 1·3개월 | 1년) */
  const [savedPoDateRange, setSavedPoDateRange] = useState("all");
  /** 저장된 발주 목록 정렬 (기본: 발주일 오래된순 — 기존 동작과 동일) */
  const [savedPoSortMode, setSavedPoSortMode] = useState("order_date_asc");
  const [editingPoId, setEditingPoId] = useState(null);
  const [poEditDraft, setPoEditDraft] = useState(null);
  /** 저장된 발주: 비고 메모 편집 모달 { orderId, lineId, draft } */
  const [savedPoMemoModal, setSavedPoMemoModal] = useState(null);
  /** { orderId, lineId, field, draft, tbd?, orderPlanned? } — orderPlanned 은 발주일자 인라인 전용 */
  const [savedPoInline, setSavedPoInline] = useState(null);
  const savedPoInlineRef = useRef(null);
  savedPoInlineRef.current = savedPoInline;
  /** 저장 발주 표: 일괄 삭제 선택 — `P|orderId`(입고 0건인 발주) 또는 `L|orderId|lineId` */
  const [savedPoBulkSelected, setSavedPoBulkSelected] = useState({});
  const savedPoBulkHeaderCbRef = useRef(null);

  const filteredPurchaseOrders = useMemo(() => {
    let rows = purchaseOrders;
    const q = savedPoSearch.trim().toLowerCase();
    if (q) {
      rows = rows.filter((po) => {
        const skuRaw = String(po.sku || "").toLowerCase();
        const name = String(po.product_name || "").toLowerCase();
        return skuRaw.includes(q) || name.includes(q);
      });
    }
    if (savedPoDateRange !== "all") {
      const now = new Date();
      now.setHours(0, 0, 0, 0);
      const cutoff = new Date(now);
      if (savedPoDateRange === "1m") cutoff.setMonth(cutoff.getMonth() - 1);
      else if (savedPoDateRange === "3m") cutoff.setMonth(cutoff.getMonth() - 3);
      else if (savedPoDateRange === "1y") cutoff.setFullYear(cutoff.getFullYear() - 1);
      rows = rows.filter((po) => {
        const od = po.order_date;
        if (!od) return true;
        const d = new Date(`${String(od).slice(0, 10)}T12:00:00`);
        return !Number.isNaN(d.getTime()) && d >= cutoff;
      });
    }
    return rows;
  }, [purchaseOrders, savedPoSearch, savedPoDateRange]);

  /** 저장된 발주: 발주 단위 그룹(같은 발주는 rowSpan으로 묶고, 줄별은 ERP_PO_차수 ref_code만 구분) */
  const savedPoSpreadsheetGroups = useMemo(() => {
    const groups = filteredPurchaseOrders.map((po) => ({
      po,
      lines: [...(po.inbound_lines || [])].sort(
        (a, b) => (Number(a.line_no) || 0) - (Number(b.line_no) || 0)
      ),
    }));
    groups.sort((a, b) => compareSavedPoSpreadsheetGroups(a, b, savedPoSortMode));
    return groups;
  }, [filteredPurchaseOrders, savedPoSortMode]);

  const savedPoBulkSelectableKeysFlat = useMemo(() => {
    const keys = [];
    for (const { po, lines } of savedPoSpreadsheetGroups) {
      if (!lines.length) keys.push(`P|${po.id}`);
      else for (const line of lines) keys.push(`L|${po.id}|${line.id}`);
    }
    return keys;
  }, [savedPoSpreadsheetGroups]);

  const savedPoBulkKeySet = useMemo(
    () => new Set(savedPoBulkSelectableKeysFlat),
    [savedPoBulkSelectableKeysFlat]
  );

  const savedPoBulkSelectedVisibleCount = useMemo(
    () => savedPoBulkSelectableKeysFlat.filter((k) => savedPoBulkSelected[k]).length,
    [savedPoBulkSelectableKeysFlat, savedPoBulkSelected]
  );

  const savedPoBulkAllVisibleSelected = useMemo(() => {
    const flat = savedPoBulkSelectableKeysFlat;
    return flat.length > 0 && flat.every((k) => savedPoBulkSelected[k]);
  }, [savedPoBulkSelectableKeysFlat, savedPoBulkSelected]);

  const savedPoBulkSomeVisibleSelected = useMemo(
    () => savedPoBulkSelectableKeysFlat.some((k) => savedPoBulkSelected[k]),
    [savedPoBulkSelectableKeysFlat, savedPoBulkSelected]
  );

  useLayoutEffect(() => {
    const el = savedPoBulkHeaderCbRef.current;
    if (!el) return;
    el.indeterminate =
      savedPoBulkSomeVisibleSelected && !savedPoBulkAllVisibleSelected;
  }, [savedPoBulkSomeVisibleSelected, savedPoBulkAllVisibleSelected]);

  useEffect(() => {
    setSavedPoBulkSelected((prev) => {
      const next = {};
      for (const k of Object.keys(prev)) {
        if (savedPoBulkKeySet.has(k)) next[k] = true;
      }
      return Object.keys(next).length === Object.keys(prev).length ? prev : next;
    });
  }, [savedPoBulkKeySet]);

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
  const shipmentChartDayMatrixScrollRef = useRef(null);
  const shipmentChartDayMatrixTopScrollRef = useRef(null);
  const syncingShipmentChartDayMatrixScrollRef = useRef(false);
  const [shipmentChartDayMatrixTopScrollWidth, setShipmentChartDayMatrixTopScrollWidth] = useState(0);
  const [showShipmentChartDayMatrixTopScroll, setShowShipmentChartDayMatrixTopScroll] = useState(false);
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
    if (mode === "SHIPMENT") return "SHIPMENT";
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
  const isPurchaseOrderScope = countryTabMode === "PURCHASE_ORDERS";
  const isShipmentScope = countryTabMode === "SHIPMENT";
  const isInventoryAdminScope =
    isSettingsScope || isSkuMappingScope || isProductSearchScope || isPurchaseOrderScope;

  /** 출고 탭: 칩(시트) 또는 초기 로드 시 뷰 조회 */
  useEffect(() => {
    if (!isShipmentScope) {
      setShipmentMatrixChannel("");
      return;
    }
    const ac = new AbortController();
    void fetchAndApplyShipmentView({
      signal: ac.signal,
      channel: shipmentMatrixChannel.trim() || undefined,
    });
    return () => ac.abort();
  }, [isShipmentScope, shipmentMatrixChannel]);

  /** 저장 발주 표: blur가 누락될 때(스크롤 트랙·레이아웃 클릭 등)에도 입력값 커밋·편집 종료 */
  useEffect(() => {
    if (!isPurchaseOrderScope || purchaseOrderSubTab !== "saved" || !savedPoInline) return;
    const onDocMouseDownCapture = (e) => {
      const target = e.target;
      if (!(target instanceof Element)) return;
      const active = document.activeElement;
      if (!active || !(active instanceof HTMLElement)) return;
      if (active === target || active.contains(target)) return;
      if (active.tagName === "INPUT" && active.classList.contains("poSavedSsInlineInput")) {
        active.blur();
      }
    };
    document.addEventListener("mousedown", onDocMouseDownCapture, true);
    return () => document.removeEventListener("mousedown", onDocMouseDownCapture, true);
  }, [isPurchaseOrderScope, purchaseOrderSubTab, savedPoInline]);

  useEffect(() => {
    if (!isPurchaseOrderScope || purchaseOrderSubTab !== "saved") return;
    const bodyEl = poSavedTableScrollRef.current;
    const headEl = poSavedHeaderScrollRef.current;
    const measure = () => {
      const body = poSavedTableScrollRef.current;
      const head = poSavedHeaderScrollRef.current;
      const scrollbarPad = body && head ? Math.max(0, body.offsetWidth - body.clientWidth) : 0;
      if (head) {
        head.style.paddingRight = scrollbarPad ? `${scrollbarPad}px` : "";
      }
      const w = Math.max(body?.scrollWidth || 0, head?.scrollWidth || 0);
      const ref = body || head;
      const cw = ref?.clientWidth || 0;
      setPoSavedTopScrollWidth(w);
      setPoSavedShowTopScroll(w > cw + 1);
    };
    measure();
    if (typeof ResizeObserver !== "undefined") {
      const ro = new ResizeObserver(measure);
      if (bodyEl) {
        ro.observe(bodyEl);
        const t = bodyEl.querySelector("table");
        if (t) ro.observe(t);
      }
      if (headEl) {
        ro.observe(headEl);
        const ht = headEl.querySelector("table");
        if (ht) ro.observe(ht);
      }
      return () => {
        ro.disconnect();
        if (poSavedHeaderScrollRef.current) poSavedHeaderScrollRef.current.style.paddingRight = "";
      };
    }
    window.addEventListener("resize", measure);
    return () => {
      window.removeEventListener("resize", measure);
      if (poSavedHeaderScrollRef.current) poSavedHeaderScrollRef.current.style.paddingRight = "";
    };
  }, [
    isPurchaseOrderScope,
    purchaseOrderSubTab,
    filteredPurchaseOrders.length,
    purchaseOrdersLoading,
  ]);

  useEffect(() => {
    const measureTopbar = () => {
      setStickyHeights((prev) => ({
        ...prev,
        topbar: topbarRef.current?.offsetHeight || 0,
      }));
    };
    measureTopbar();
    let ro;
    if (typeof ResizeObserver !== "undefined" && topbarRef.current) {
      ro = new ResizeObserver(measureTopbar);
      ro.observe(topbarRef.current);
    }
    window.addEventListener("resize", measureTopbar);
    return () => {
      window.removeEventListener("resize", measureTopbar);
      if (ro) ro.disconnect();
    };
  }, []);

  useEffect(() => {
    if (!isPurchaseOrderScope) return;
    const measure = () => {
      setPoOrderStickyHeights({
        subTabs: poSubTabsBarRef.current?.offsetHeight ?? 0,
        savedFilter: poSavedFilterBarRef.current?.offsetHeight ?? 0,
      });
    };
    measure();
    let ro;
    if (typeof ResizeObserver !== "undefined") {
      ro = new ResizeObserver(measure);
      if (poSubTabsBarRef.current) ro.observe(poSubTabsBarRef.current);
      if (poSavedFilterBarRef.current) ro.observe(poSavedFilterBarRef.current);
    }
    window.addEventListener("resize", measure);
    return () => {
      window.removeEventListener("resize", measure);
      if (ro) ro.disconnect();
    };
  }, [
    isPurchaseOrderScope,
    purchaseOrderSubTab,
    filteredPurchaseOrders.length,
    purchaseOrdersLoading,
    savedPoSearch,
    savedPoDateRange,
    savedPoSortMode,
  ]);

  useEffect(() => {
    if (!isPurchaseOrderScope || purchaseOrderSubTab !== "saved" || !poSavedShowTopScroll) {
      setPoSavedTopStripHeight(0);
      return;
    }
    const el = poSavedTopScrollRef.current;
    if (!el) return;
    const measure = () =>
      setPoSavedTopStripHeight(el.offsetHeight || INVENTORY_MATCH_TOP_SCROLL_STRIP_FALLBACK_PX);
    measure();
    if (typeof ResizeObserver !== "undefined") {
      const ro = new ResizeObserver(measure);
      ro.observe(el);
      return () => ro.disconnect();
    }
    window.addEventListener("resize", measure);
    return () => window.removeEventListener("resize", measure);
  }, [
    isPurchaseOrderScope,
    purchaseOrderSubTab,
    poSavedShowTopScroll,
    filteredPurchaseOrders.length,
    purchaseOrdersLoading,
  ]);

  const inventoryStickyWidth = useMemo(() => {
    if (isShipmentScope) return SHIPMENT_MATRIX_STICKY_TOTAL_PX;
    if (isKRScope) return INVENTORY_KR_STICKY_TOTAL_PX;
    if (isOverseasScope)
      return showKrCompare
        ? INVENTORY_OVERSEAS_STICKY_WITH_KR_COMPARE_PX
        : INVENTORY_OVERSEAS_STICKY_TOTAL_PX;
    return 0;
  }, [isShipmentScope, isKRScope, isOverseasScope, showKrCompare]);
  /** 해외 국가 칩 · 출고 일자/월 합계 서브 탭 등 2단 띠 */
  const hasSecondaryInventoryStrip = countryTabMode === "OVERSEAS" || countryTabMode === "SHIPMENT";
  const activeCountryChipsHeight = hasSecondaryInventoryStrip ? stickyHeights.countryChips : 0;
  const activeFilterStickyTop = stickyHeights.topbar + activeCountryChipsHeight;
  const activeFilterHeight = isCompareScope ? stickyHeights.compareFilter : stickyHeights.inventoryFilter;
  const activeTopScrollHeight = showTopScroll ? stickyHeights.topScroll : 0;
  const tableHeaderTop = activeFilterStickyTop + activeFilterHeight + activeTopScrollHeight;

  const poStickySubTabsTop = stickyHeights.topbar;
  const poStickySavedFilterTop = stickyHeights.topbar + poOrderStickyHeights.subTabs;
  const poSavedFilterBottomSticky = poStickySavedFilterTop + poOrderStickyHeights.savedFilter;
  /** 상단 가로 띠 있음: 스티키 헤더 top = 필터 하단 + 띠 높이(띠·헤더 사이 6px 두면 스크롤 시 틈으로 본문이 비침). 띠 없음: 재고 탭과 같은 폴백 두께 */
  const poSavedTableHeaderStickyTop = useMemo(() => {
    if (poSavedShowTopScroll) {
      const stripH =
        poSavedTopStripHeight > 0 ? poSavedTopStripHeight : INVENTORY_MATCH_TOP_SCROLL_STRIP_FALLBACK_PX;
      return poSavedFilterBottomSticky + stripH;
    }
    return (
      poSavedFilterBottomSticky +
      INVENTORY_MATCH_TOP_SCROLL_STRIP_FALLBACK_PX +
      INVENTORY_MATCH_TOP_SCROLL_MARGIN_BELOW_PX
    );
  }, [
    poSavedFilterBottomSticky,
    poSavedShowTopScroll,
    poSavedTopStripHeight,
  ]);
  const poSavedOpaqueGapHeightPx =
    INVENTORY_MATCH_TOP_SCROLL_STRIP_FALLBACK_PX + INVENTORY_MATCH_TOP_SCROLL_MARGIN_BELOW_PX;

  useLayoutEffect(() => {
    if (!isPurchaseOrderScope || purchaseOrderSubTab !== "saved" || purchaseOrdersLoading) return;
    const headTable = poSavedHeadTableRef.current;
    const bodyTable = poSavedBodyTableRef.current;
    if (!headTable || !bodyTable) return;

    const syncWidths = () => {
      const ths = headTable.querySelectorAll(":scope > thead > tr > th");
      const colCount = ths.length;
      if (!colCount) return;
      const maxW = Array.from({ length: colCount }, () => 8);
      headTable.querySelectorAll(":scope > thead > tr").forEach((tr) => {
        Array.from(tr.cells).forEach((td, i) => {
          if (i < colCount) maxW[i] = Math.max(maxW[i], td.getBoundingClientRect().width);
        });
      });
      bodyTable.querySelectorAll(":scope > tbody > tr").forEach((tr) => {
        Array.from(tr.cells).forEach((td, i) => {
          if (i < colCount) maxW[i] = Math.max(maxW[i], td.getBoundingClientRect().width);
        });
      });
      const nameColIdx = 4;
      const nameMinPx = 190;
      const nameMaxPx = 360;
      if (nameColIdx < colCount) {
        maxW[nameColIdx] = Math.min(nameMaxPx, Math.max(maxW[nameColIdx] || 0, nameMinPx));
      }
      Array.from(ths).forEach((th, i) => {
        const w = Math.ceil(maxW[i] || 80);
        th.style.width = `${w}px`;
        th.style.minWidth = `${w}px`;
        th.style.boxSizing = "border-box";
      });
      bodyTable.querySelectorAll(":scope > tbody > tr").forEach((tr) => {
        Array.from(tr.cells).forEach((td, i) => {
          if (i < colCount) {
            const w = Math.ceil(maxW[i] || 80);
            td.style.width = `${w}px`;
            td.style.minWidth = `${w}px`;
            td.style.boxSizing = "border-box";
          }
        });
      });
      headTable.style.tableLayout = "fixed";
      bodyTable.style.tableLayout = "fixed";
    };

    let raf = 0;
    const schedule = () => {
      if (raf) cancelAnimationFrame(raf);
      raf = requestAnimationFrame(() => {
        raf = 0;
        syncWidths();
      });
    };
    schedule();
    const ro = typeof ResizeObserver !== "undefined" ? new ResizeObserver(schedule) : null;
    if (ro) {
      ro.observe(bodyTable);
      const wrap = poSavedTableScrollRef.current;
      if (wrap) ro.observe(wrap);
    }
    window.addEventListener("resize", schedule);
    return () => {
      if (raf) cancelAnimationFrame(raf);
      if (ro) ro.disconnect();
      window.removeEventListener("resize", schedule);
    };
  }, [
    isPurchaseOrderScope,
    purchaseOrderSubTab,
    purchaseOrdersLoading,
    filteredPurchaseOrders,
    purchaseOrders,
  ]);

  useEffect(() => {
    if (!savedPoMemoModal) {
      poMemoModalIdsRef.current = null;
      poMemoDraftRef.current = "";
      return;
    }
    poMemoModalIdsRef.current = {
      orderId: savedPoMemoModal.orderId,
      lineId: savedPoMemoModal.lineId,
    };
    poMemoDraftRef.current = String(savedPoMemoModal.draft ?? "");
  }, [savedPoMemoModal]);

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

  const shipmentAvailableMonths = useMemo(() => {
    const s = new Set();
    for (const d of rawInventoryDates) {
      const m = String(d).slice(0, 7);
      if (/^\d{4}-\d{2}$/.test(m)) s.add(m);
    }
    return Array.from(s).sort((a, b) => a.localeCompare(b));
  }, [rawInventoryDates]);

  const shipmentDataYears = useMemo(() => {
    const ys = new Set();
    for (const m of shipmentAvailableMonths) {
      const y = Number(String(m).slice(0, 4));
      if (Number.isFinite(y)) ys.add(y);
    }
    return Array.from(ys).sort((a, b) => a - b);
  }, [shipmentAvailableMonths]);

  useEffect(() => {
    if (!isShipmentScope) return;
    if (!shipmentDataYears.length) return;
    if (!shipmentDataYears.includes(shipmentMonthStripYear)) {
      setShipmentMonthStripYear(shipmentDataYears[shipmentDataYears.length - 1]);
    }
  }, [isShipmentScope, shipmentDataYears, shipmentMonthStripYear]);

  useEffect(() => {
    if (!isShipmentScope) return;
    if (/^\d{4}-\d{2}$/.test(shipmentDisplayMonth)) {
      const y = Number(shipmentDisplayMonth.slice(0, 4));
      if (Number.isFinite(y) && shipmentDataYears.includes(y)) setShipmentMonthStripYear(y);
      return;
    }
    if (shipmentDisplayMonth !== "__ALL__") return;
    if (!shipmentAvailableMonths.length) return;
    setShipmentMonthStripYear((prev) =>
      shipmentDataYears.includes(prev) ? prev : Number(String(shipmentAvailableMonths[shipmentAvailableMonths.length - 1]).slice(0, 4)) || prev,
    );
  }, [isShipmentScope, shipmentDisplayMonth, shipmentAvailableMonths, shipmentDataYears]);

  useEffect(() => {
    if (!isShipmentScope) return;
    if (shipmentDataYears.length !== 1) return;
    setShipmentMonthStripYear(shipmentDataYears[0]);
  }, [isShipmentScope, shipmentDataYears]);

  useEffect(() => {
    if (!isShipmentScope) return;
    if (!shipmentAvailableMonths.length) return;
    const last = shipmentAvailableMonths[shipmentAvailableMonths.length - 1];
    setShipmentDisplayMonth((prev) => {
      if (prev === "__ALL__") return prev;
      if (prev && shipmentAvailableMonths.includes(prev)) return prev;
      return last;
    });
  }, [isShipmentScope, shipmentAvailableMonths]);

  useEffect(() => {
    if (!isShipmentScope || shipmentViewMode !== "chart") {
      setShipmentChartAllChannels(null);
      return;
    }
    const ac = new AbortController();
    const year = shipmentMonthStripYear;
    (async () => {
      try {
        const shipRes = await axios.get(`${API_BASE}/api/inventory/shipment/view`, {
          params: { year, all_channels: true },
          signal: ac.signal,
        });
        const data = shipRes?.data || {};
        setShipmentChartAllChannels({
          rows: data.rows || [],
          dates: data.dates || [],
        });
      } catch (err) {
        if (err?.code === "ERR_CANCELED" || err?.name === "CanceledError") return;
        setShipmentChartAllChannels(null);
      }
    })();
    return () => ac.abort();
  }, [isShipmentScope, shipmentViewMode, shipmentMonthStripYear]);

  /** 출고 일자 열: null 이면 전체 기간( __ALL__ ) */
  const shipmentMonthForDailyFilter = useMemo(() => {
    if (!isShipmentScope) return null;
    if (shipmentDisplayMonth === "__ALL__") return null;
    if (shipmentDisplayMonth && shipmentAvailableMonths.includes(shipmentDisplayMonth)) return shipmentDisplayMonth;
    return shipmentAvailableMonths.length ? shipmentAvailableMonths[shipmentAvailableMonths.length - 1] : null;
  }, [isShipmentScope, shipmentDisplayMonth, shipmentAvailableMonths]);

  /** 출고 판매처별 월 합계: 전체 기간 선택 시에도 집계 기준 월은 데이터상 최신 월 */
  const shipmentPivotMonth = useMemo(() => {
    if (!isShipmentScope) return null;
    if (!shipmentAvailableMonths.length) return null;
    if (shipmentDisplayMonth === "__ALL__") return shipmentAvailableMonths[shipmentAvailableMonths.length - 1];
    if (shipmentDisplayMonth && shipmentAvailableMonths.includes(shipmentDisplayMonth)) return shipmentDisplayMonth;
    return shipmentAvailableMonths[shipmentAvailableMonths.length - 1];
  }, [isShipmentScope, shipmentDisplayMonth, shipmentAvailableMonths]);

  const filteredDateColumns = useMemo(() => {
    if (!rawInventoryDates.length) return [];
    if (isShipmentScope) {
      const sorted = [...rawInventoryDates].sort((a, b) => String(a).localeCompare(String(b)));
      if (shipmentMonthForDailyFilter == null) return sorted;
      return sorted.filter((d) => String(d).startsWith(shipmentMonthForDailyFilter));
    }

    const rowHasQtyOnDate = (row, dateKey) => {
      if (isShipmentScope) {
        if (row.channel_date_qty && typeof row.channel_date_qty === "object") {
          const cdq = row.channel_date_qty;
          for (const ch of SHIPMENT_SHEET_CHANNELS) {
            if (Number(cdq[ch]?.[dateKey] || 0) !== 0) return true;
          }
          for (const ch of Object.keys(cdq)) {
            const byD = cdq[ch];
            if (byD && typeof byD === "object" && Number(byD[dateKey] || 0) !== 0) return true;
          }
          return false;
        }
        return shipmentCellNumericTotal(row[dateKey]) !== 0;
      }
      return matchesCountryScope(row.country) && Number(row[dateKey] || 0) !== 0;
    };

    const scopedDateKeys = rawInventoryDates.filter((dateKey) =>
      rawInventoryRows.some((row) => rowHasQtyOnDate(row, dateKey))
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
    isShipmentScope,
    shipmentMonthForDailyFilter,
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

    const rows = rawInventoryRows.filter((row) => {
      if (!matchesCountryScope(row.country)) return false;

      if (isShipmentScope && row.is_total) {
        return !needle;
      }

      if (!isKRScope && hasInventoryLevels && inventoryLevelFilter !== "all") {
        const lv = (String(row.level ?? "").replace(/^0+/, "") || "0").trim();
        if (lv !== inventoryLevelFilter) return false;
      }
      if (needle) {
        const item = String(getRowSku(row) ?? "").toLowerCase();
        const desc = String(row.description ?? "").toLowerCase();
        const krSku = String(row.mapped_kr_sku ?? "").toLowerCase();
        const mn = String(row.mapped_kr_name ?? "").toLowerCase();
        if (isShipmentScope) {
          if (!item.includes(needle) && !desc.includes(needle) && !krSku.includes(needle) && !mn.includes(needle)) {
            return false;
          }
        } else {
          const ch = String(row.channel ?? "").toLowerCase();
          const br = String(row.brand ?? "").toLowerCase();
          const mkt = String(row.mkt_priority ?? "").toLowerCase();
          const seg = String(row.segment ?? "").toLowerCase();
          if (
            !item.includes(needle) &&
            !desc.includes(needle) &&
            !krSku.includes(needle) &&
            !ch.includes(needle) &&
            !mn.includes(needle) &&
            !br.includes(needle) &&
            !mkt.includes(needle) &&
            !seg.includes(needle)
          ) {
            return false;
          }
        }
      }
      return true;
    });
    if (isShipmentScope) {
      return rows;
    }
    return [...rows].sort((a, b) => compareInventoryRowsByLatestQty(a, b, filteredDateColumns));
  }, [
    rawInventoryRows,
    inventoryKeyword,
    inventoryLevelFilter,
    hasInventoryLevels,
    countryTabMode,
    selectedOverseasCountry,
    isKRScope,
    isShipmentScope,
    filteredDateColumns,
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
      if (code === "SHIPMENT") continue;
      if (!groups[code]) groups[code] = [];
      groups[code].push(entry);
    }
    groups.SHIPMENT = [...shipmentFileEntries];
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
  }, [fileEntries, shipmentFileEntries]);

  /** 해외 행과 맞출 때 브랜드/창고/레벨이 국가마다 달라 행 단위가 다를 수 있음 → 한국 SKU(매핑 우선)로 오늘 재고 합산 */
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

  const shipmentDisplayRows = useMemo(() => {
    if (!isShipmentScope) return [];
    const sorted = [...filteredRows].sort((a, b) => {
      if (a.is_total) return -1;
      if (b.is_total) return 1;
      const ta = shipmentRowSumInDateCols(a, filteredDateColumns);
      const tb = shipmentRowSumInDateCols(b, filteredDateColumns);
      if (tb !== ta) return tb - ta;
      const c = String(getRowSku(a) || "").localeCompare(String(getRowSku(b) || ""), "ko");
      if (c !== 0) return c;
      return String(a.description || "").localeCompare(String(b.description || ""), "ko");
    });
    return sorted.map((row, idx) => ({
      ...row,
      trendRowKey: row.is_total ? "SHIP-TOTAL" : `SHIP-${getRowSku(row)}-${idx}`,
    }));
  }, [isShipmentScope, filteredRows, filteredDateColumns]);

  const shipmentChartSourceRows = useMemo(() => {
    if (!isShipmentScope || shipmentViewMode !== "chart") return rawInventoryRows;
    if (Array.isArray(shipmentChartAllChannels?.rows) && shipmentChartAllChannels.rows.length > 0) {
      return shipmentChartAllChannels.rows;
    }
    return rawInventoryRows;
  }, [isShipmentScope, shipmentViewMode, shipmentChartAllChannels, rawInventoryRows]);

  const shipmentChartSourceDates = useMemo(() => {
    if (!isShipmentScope || shipmentViewMode !== "chart") return rawInventoryDates;
    if (Array.isArray(shipmentChartAllChannels?.dates) && shipmentChartAllChannels.dates.length > 0) {
      return shipmentChartAllChannels.dates;
    }
    return rawInventoryDates;
  }, [isShipmentScope, shipmentViewMode, shipmentChartAllChannels, rawInventoryDates]);

  const shipmentChartMonthOptions = useMemo(() => {
    const s = new Set();
    for (const d of shipmentChartSourceDates) {
      const mo = String(d).slice(0, 7);
      if (/^\d{4}-\d{2}$/.test(mo)) s.add(mo);
    }
    return Array.from(s).sort((a, b) => a.localeCompare(b));
  }, [shipmentChartSourceDates]);

  /** 출고 뷰: 서버가 준 동적 채널 목록(수동 매핑 탭명 포함), 없으면 기본 시트 탭 순서 */
  const effectiveShipmentChannels = useMemo(() => {
    const ch = scopeResultCache.SHIPMENT?.channels;
    if (Array.isArray(ch) && ch.length) return ch;
    return SHIPMENT_SHEET_CHANNELS;
  }, [scopeResultCache.SHIPMENT?.channels]);

  /** 출고 차트: 상품코드 접두 또는 상품명 부분 일치(전 시트 행 기준) */
  const shipmentChartPrefixRows = useMemo(() => {
    if (!isShipmentScope || shipmentViewMode !== "chart" || !shipmentChartSourceRows.length) return [];
    const needle = shipmentChartSearchKeyword.trim();
    if (!needle) return [];
    const needleLower = needle.toLowerCase();
    return shipmentChartSourceRows.filter((row) => {
      if (!matchesCountryScope(row.country)) return false;
      if (row.is_total) return false;
      const sku = String(getRowSku(row) ?? "").trim();
      if (!sku) return false;
      const skuLower = sku.toLowerCase();
      const descLower = String(row.description ?? "").trim().toLowerCase();
      return skuLower.startsWith(needleLower) || (descLower && descLower.includes(needleLower));
    });
  }, [
    isShipmentScope,
    shipmentViewMode,
    shipmentChartSourceRows,
    shipmentChartSearchKeyword,
    countryTabMode,
    selectedOverseasCountry,
  ]);

  const shipmentChartSkuOptions = useMemo(() => {
    const s = new Set();
    for (const row of shipmentChartPrefixRows) {
      const sku = getRowSku(row);
      if (sku) s.add(sku);
    }
    return Array.from(s).sort((a, b) => a.localeCompare(b));
  }, [shipmentChartPrefixRows]);

  /** 목록에 표시할 SKU → 상품명(첫 행 기준) */
  const shipmentChartSkuOptionLabels = useMemo(() => {
    const m = new Map();
    for (const row of shipmentChartPrefixRows) {
      const sku = getRowSku(row);
      if (!sku || m.has(sku)) continue;
      m.set(sku, String(row.description || "").trim() || "(상품명 없음)");
    }
    return m;
  }, [shipmentChartPrefixRows]);

  /** 선택한 SKU의 전체 출고 행(차트 집계·시트별 행 모두 포함) */
  const shipmentChartRowsForSku = useMemo(() => {
    if (!shipmentChartSelectedSku) return [];
    return shipmentChartSourceRows.filter((row) => {
      if (!matchesCountryScope(row.country)) return false;
      if (row.is_total) return false;
      return getRowSku(row) === shipmentChartSelectedSku;
    });
  }, [
    isShipmentScope,
    shipmentChartSourceRows,
    shipmentChartSelectedSku,
    countryTabMode,
    selectedOverseasCountry,
  ]);

  /** 원시 날짜 열을 월별로 묶음 — 월별 차트 비교용 */
  const shipmentChartDatesByMonth = useMemo(() => {
    const m = new Map();
    for (const dk of shipmentChartSourceDates) {
      const mo = String(dk).slice(0, 7);
      if (!/^\d{4}-\d{2}$/.test(mo)) continue;
      if (!m.has(mo)) m.set(mo, []);
      m.get(mo).push(dk);
    }
    return m;
  }, [shipmentChartSourceDates]);

  /** 일자별 차트: 선택한 달의 1일~말일 전부(데이터 없으면 0) */
  const shipmentChartDailyMonthSeries = useMemo(() => {
    if (!shipmentChartMonth || !shipmentChartRowsForSku.length) return [];
    const dim = daysInCalendarMonth(shipmentChartMonth);
    const out = [];
    for (let d = 1; d <= dim; d += 1) {
      const dd = String(d).padStart(2, "0");
      const dk = `${shipmentChartMonth}-${dd}`;
      const qty = shipmentChartRowsForSku.reduce((sum, row) => sum + shipmentCellNumericTotal(row[dk]), 0);
      const vendorTooltip = shipmentDateCellVendorTitleAggregated(shipmentChartRowsForSku, dk);
      out.push({ key: dk, qty, labelShort: String(d), vendorTooltip });
    }
    return out;
  }, [shipmentChartMonth, shipmentChartRowsForSku]);

  /** 월별 차트: DB에 존재하는 월끼리만 비교(상단 월 필터와 무관) */
  const shipmentChartMonthlyCompareSeries = useMemo(() => {
    if (!shipmentChartRowsForSku.length || !shipmentChartSourceDates.length) return [];
    const months = [...shipmentChartDatesByMonth.keys()].sort((a, b) => a.localeCompare(b));
    return months.map((monthKey) => {
      const dks = shipmentChartDatesByMonth.get(monthKey) || [];
      let qty = 0;
      for (const row of shipmentChartRowsForSku) {
        for (const dk of dks) {
          qty += shipmentCellNumericTotal(row[dk]);
        }
      }
      return {
        key: monthKey,
        qty,
        labelShort: monthKey.length >= 7 ? monthKey.slice(2) : monthKey,
      };
    });
  }, [shipmentChartRowsForSku, shipmentChartSourceDates, shipmentChartDatesByMonth]);

  const shipmentChartTitleLabel = useMemo(() => {
    const r = shipmentChartRowsForSku[0];
    if (!r) return "";
    return String(r.description || "").trim() || "(상품명 없음)";
  }, [shipmentChartRowsForSku]);

  const shipmentChartDateKeysFiltered = useMemo(
    () => shipmentChartSourceDates.filter((dk) => /^\d{4}-\d{2}-\d{2}$/.test(String(dk))),
    [shipmentChartSourceDates],
  );

  /** 차트 탭: 선택 SKU의 출고 요약·채널별·월합계·판매처(분해 있을 때) */
  const shipmentChartSkuInsight = useMemo(() => {
    const rows = shipmentChartRowsForSku;
    if (!rows.length) return null;
    const dateKeys = shipmentChartDateKeysFiltered;
    let total = 0;
    let firstD = null;
    let lastD = null;
    for (const dk of dateKeys) {
      let daySum = 0;
      for (const row of rows) daySum += shipmentCellNumericTotal(row[dk]);
      total += daySum;
      if (daySum > 0) {
        if (firstD == null) firstD = dk;
        lastD = dk;
      }
    }
    const byChannel = rows
      .map((row) => {
        let chSum = 0;
        for (const dk of dateKeys) chSum += shipmentCellNumericTotal(row[dk]);
        const mt = row.month_totals || {};
        return {
          channel: String(row.channel || "").trim() || "(미지정)",
          brand: String(row.brand || "").trim() || "–",
          mkt_priority: String(row.mkt_priority || "").trim() || "–",
          segment: String(row.segment || "").trim() || "–",
          rowTotal: chSum,
          month_totals: { ...mt },
        };
      })
      .sort((a, b) => b.rowTotal - a.rowTotal);
    const monthAgg = {};
    for (const r of byChannel) {
      for (const [k, v] of Object.entries(r.month_totals || {})) {
        const n = Number(v) || 0;
        if (!n) continue;
        monthAgg[k] = (monthAgg[k] || 0) + n;
      }
    }
    const monthAggSorted = Object.entries(monthAgg)
      .filter(([, v]) => Number(v))
      .sort(([a], [b]) => String(a).localeCompare(String(b)));
    const vendorMap = aggregateShipmentChartVendorTotals(rows, dateKeys);
    const vendorRows = [...vendorMap.entries()]
      .map(([vendor, { sale, wh }]) => ({
        vendor,
        sale,
        wh,
        sum: sale + wh,
      }))
      .filter((x) => x.sum > 0)
      .sort((a, b) => b.sum - a.sum);
    const brands = [...new Set(rows.map((r) => String(r.brand || "").trim()).filter(Boolean))];
    const mappedKr = [...new Set(rows.map((r) => String(r.mapped_kr_sku || "").trim()).filter(Boolean))];
    const mappedNames = [...new Set(rows.map((r) => String(r.mapped_kr_name || "").trim()).filter(Boolean))];
    return {
      total,
      firstD,
      lastD,
      byChannel,
      monthAggSorted,
      vendorRows,
      brands,
      mappedKr,
      mappedNames,
      channelCount: byChannel.length,
    };
  }, [shipmentChartRowsForSku, shipmentChartDateKeysFiltered]);

  /** 일자별·판매처별 차트 월 칩에 쓰는 연도(선택 월 또는 데이터 최신 연도) */
  const shipmentChartChipYear = useMemo(() => {
    if (shipmentChartMonth && /^\d{4}-\d{2}$/.test(shipmentChartMonth)) {
      const y = Number(shipmentChartMonth.slice(0, 4));
      if (Number.isFinite(y) && shipmentChartMonthOptions.some((m) => String(m).startsWith(`${y}-`))) {
        return y;
      }
    }
    if (!shipmentChartMonthOptions.length) return new Date().getFullYear();
    const years = shipmentChartMonthOptions
      .map((m) => Number(String(m).slice(0, 4)))
      .filter((n) => Number.isFinite(n));
    return years.length ? Math.max(...years) : new Date().getFullYear();
  }, [shipmentChartMonth, shipmentChartMonthOptions]);

  const shipmentChartDailyLineData = useMemo(() => {
    if (!shipmentChartMonth || !shipmentChartDailyMonthSeries.length) return null;
    const n = shipmentChartDailyMonthSeries.length;
    return buildShipmentQtyLineChart(shipmentChartDailyMonthSeries, {
      chartWidth: Math.min(1520, Math.max(520, n * 20)),
      chartHeight: 400,
      paddingXLeft: 52,
      paddingXRight: 30,
      xLabelStep: n <= 14 ? 1 : n <= 21 ? 2 : 3,
      showPointValueLabels: false,
    });
  }, [shipmentChartMonth, shipmentChartDailyMonthSeries]);

  const shipmentChartMonthlyLineData = useMemo(() => {
    if (!shipmentChartMonthlyCompareSeries.length) return null;
    const n = shipmentChartMonthlyCompareSeries.length;
    return buildShipmentQtyLineChart(shipmentChartMonthlyCompareSeries, {
      chartWidth: Math.min(1240, Math.max(480, n * 56)),
      chartHeight: 400,
      xLabelStep: n <= 12 ? 1 : 2,
      showPointValueLabels: n <= 14,
    });
  }, [shipmentChartMonthlyCompareSeries]);

  /** 선택 월·SKU: 칩(시트)별 일자별 수량 표 */
  const shipmentChartChannelDailyMatrix = useMemo(() => {
    if (!shipmentChartMonth || !shipmentChartRowsForSku.length) return null;
    const dim = daysInCalendarMonth(shipmentChartMonth);
    const dayKeys = [];
    for (let d = 1; d <= dim; d += 1) {
      dayKeys.push(`${shipmentChartMonth}-${String(d).padStart(2, "0")}`);
    }
    const order = new Map(effectiveShipmentChannels.map((c, i) => [c, i]));
    const byChannel = new Map();
    for (const row of shipmentChartRowsForSku) {
      const ch = String(row.channel || "").trim() || "(미지정)";
      if (!byChannel.has(ch)) byChannel.set(ch, new Map());
      const m = byChannel.get(ch);
      for (const dk of dayKeys) {
        const q = shipmentCellNumericTotal(row[dk]);
        m.set(dk, (m.get(dk) || 0) + q);
      }
    }
    const channels = [...byChannel.keys()].sort((a, b) => {
      const ia = order.has(a) ? order.get(a) : 999;
      const ib = order.has(b) ? order.get(b) : 999;
      if (ia !== ib) return ia - ib;
      return a.localeCompare(b, "ko");
    });
    const rows = channels.map((ch) => {
      const m = byChannel.get(ch);
      let rowSum = 0;
      const cells = dayKeys.map((dk) => {
        const q = m.get(dk) || 0;
        rowSum += q;
        return q;
      });
      return { channel: ch, cells, rowSum };
    });
    const colTotals = dayKeys.map((_, i) => rows.reduce((s, r) => s + r.cells[i], 0));
    const grandTotal = rows.reduce((s, r) => s + r.rowSum, 0);
    return { dayKeys, rows, colTotals, grandTotal };
  }, [shipmentChartMonth, shipmentChartRowsForSku, effectiveShipmentChannels]);

  useEffect(() => {
    if (!isShipmentScope || shipmentViewMode !== "chart") return;
    if (!shipmentChartSkuOptions.length) {
      if (shipmentChartSelectedSku) setShipmentChartSelectedSku("");
      return;
    }
    if (shipmentChartSelectedSku && !shipmentChartSkuOptions.includes(shipmentChartSelectedSku)) {
      setShipmentChartSelectedSku("");
    }
  }, [isShipmentScope, shipmentViewMode, shipmentChartSkuOptions, shipmentChartSelectedSku]);

  useEffect(() => {
    if (!isShipmentScope || shipmentViewMode !== "chart") return;
    if (!shipmentChartMonthOptions.length) {
      if (shipmentChartMonth) setShipmentChartMonth("");
      return;
    }
    if (!shipmentChartMonth || !shipmentChartMonthOptions.includes(shipmentChartMonth)) {
      setShipmentChartMonth(shipmentChartMonthOptions[shipmentChartMonthOptions.length - 1]);
    }
  }, [isShipmentScope, shipmentViewMode, shipmentChartMonthOptions, shipmentChartMonth]);

  /** 검색 결과가 상품 하나뿐이면 자동 선택 */
  useEffect(() => {
    if (!isShipmentScope || shipmentViewMode !== "chart") return;
    if (!shipmentChartSearchKeyword.trim()) return;
    if (shipmentChartSkuOptions.length !== 1) return;
    const only = shipmentChartSkuOptions[0];
    if (shipmentChartSelectedSku !== only) setShipmentChartSelectedSku(only);
  }, [
    isShipmentScope,
    shipmentViewMode,
    shipmentChartSearchKeyword,
    shipmentChartSkuOptions,
    shipmentChartSelectedSku,
  ]);

  const shipmentMatrixMonthColCount = useMemo(
    () => (scopeResultCache.SHIPMENT?.shipment_month_columns || []).length,
    [scopeResultCache.SHIPMENT?.shipment_month_columns],
  );

  const hasTableData = useMemo(() => {
    if (isShipmentScope) {
      if (shipmentViewMode === "chart") {
        const chartRows =
          (Array.isArray(shipmentChartAllChannels?.rows) && shipmentChartAllChannels.rows.length > 0) ||
          rawInventoryRows.length > 0;
        const chartDates =
          (Array.isArray(shipmentChartAllChannels?.dates) && shipmentChartAllChannels.dates.length > 0) ||
          filteredDateColumns.length > 0 ||
          rawInventoryDates.length > 0;
        return chartRows && chartDates;
      }
      return (
        shipmentDisplayRows.length > 0 &&
        (shipmentMatrixMonthColCount > 0 ||
          rawInventoryDates.length > 0 ||
          filteredDateColumns.length > 0)
      );
    }
    if (isKRScope) return krDisplayRows.length > 0 && filteredDateColumns.length > 0;
    return filteredRows.length > 0 && filteredDateColumns.length > 0;
  }, [
    isShipmentScope,
    shipmentViewMode,
    shipmentDisplayRows,
    shipmentMatrixMonthColCount,
    rawInventoryRows.length,
    shipmentChartAllChannels,
    isKRScope,
    krDisplayRows,
    filteredRows,
    filteredDateColumns,
    rawInventoryDates.length,
  ]);

  const latestScopeDateLabel = useMemo(() => getLatestDateKey(filteredDateColumns) || "-", [filteredDateColumns]);

  const inventoryBasisLabel = useMemo(() => {
    if (isShipmentScope) return "출고 수량";
    if (isKRScope) return "가용재고";
    if (isOverseasScope) return "해당 국가 창고 재고";
    return "–";
  }, [isShipmentScope, isKRScope, isOverseasScope]);

  const compareAvailableDates = useMemo(() => {
    const allDates = new Set(scopeResultCache.KR?.dates || []);
    for (const code of OVERSEAS_UPLOAD_COUNTRIES) {
      for (const dateKey of scopeResultCache[`OVERSEAS:${code}`]?.dates || []) {
        allDates.add(dateKey);
      }
    }
    return Array.from(allDates).sort();
  }, [scopeResultCache]);

  const compareDatePickerExtent = useMemo(() => {
    const sorted = [...compareAvailableDates].filter(Boolean).sort();
    if (!sorted.length) return {};
    return { min: sorted[0], max: sorted[sorted.length - 1] };
  }, [compareAvailableDates]);

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
        const existing = rowMetaLookup.get(compareKey);
        if (!existing) {
          rowMetaLookup.set(compareKey, { code, name: displayName, metaLabel });
        } else {
          rowMetaLookup.set(compareKey, {
            code: existing.code || code,
            name: String(existing.name || "").trim() || displayName,
            metaLabel: pickRicherCompareMetaLabel(existing.metaLabel, metaLabel),
          });
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
          String(aMeta.code || "").localeCompare(String(bMeta.code || ""), "ko") ||
          String(aMeta.name || "").localeCompare(String(bMeta.name || ""), "ko")
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

  const compareRows = useMemo(() => {
    if (!compareSelectedDate || !compareCountryData.rowKeys.length) return [];
    return compareCountryData.rowKeys.map((rowKey) => {
      const meta = compareCountryData.rowMetaLookup.get(rowKey) || {};
      const row = {
        _compareKey: rowKey,
        상품코드: meta.code || "",
        한국상품명: meta.name || "",
        구분: String(meta.metaLabel || "").trim() || metaLabelFromCompareIdentityKey(rowKey),
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
    const base = !needle
      ? compareRows
      : compareRows.filter((row) => {
          const code = String(row["상품코드"] || "").toLowerCase();
          const krName = String(row["한국상품명"] || "").toLowerCase();
          const meta = String(row["구분"] || "").toLowerCase();
          return code.includes(needle) || krName.includes(needle) || meta.includes(needle);
        });
    return [...base].sort(comparePivotRowsByTotalQty);
  }, [compareRows, inventoryKeyword, isCompareScope]);
  const tableMinWidth = useMemo(() => {
    const dateCols = filteredDateColumns.length * INVENTORY_DATE_COL_PX;
    const overseasFixedCols = showKrCompare
      ? INVENTORY_OVERSEAS_FIXED_WITH_KR_COMPARE_PX
      : INVENTORY_OVERSEAS_FIXED_TOTAL_PX;
    const fixedCols = isOverseasScope ? overseasFixedCols : INVENTORY_KR_FIXED_TOTAL_PX;
    return Math.max(980, fixedCols + dateCols);
  }, [filteredDateColumns, isOverseasScope, showKrCompare]);
  const inventoryTableWidth = useMemo(() => {
    if (isShipmentScope) {
      if (shipmentViewMode === "chart") return Math.max(980, tableMinWidth);
      return Math.max(
        980,
        SHIPMENT_MATRIX_STICKY_TOTAL_PX +
          shipmentMatrixMonthColCount * INVENTORY_DATE_COL_PX +
          filteredDateColumns.length * INVENTORY_DATE_COL_PX,
      );
    }
    if (isKRScope)
      return Math.max(980, INVENTORY_KR_FIXED_TOTAL_PX + filteredDateColumns.length * INVENTORY_DATE_COL_PX);
    return tableMinWidth;
  }, [
    isShipmentScope,
    shipmentViewMode,
    filteredDateColumns,
    tableMinWidth,
    shipmentMatrixMonthColCount,
  ]);
  /** 헤더/본문 두 테이블의 열 너비를 동일하게 고정(스티키 left와 실제 경계 일치) */
  const shipmentMatrixColGroup = useMemo(() => {
    if (!isShipmentScope || shipmentViewMode !== "status") return null;
    const monthCols = scopeResultCache.SHIPMENT?.shipment_month_columns || [];
    return (
      <colgroup>
        <col style={{ width: SHIPMENT_MATRIX_COL_CODE }} />
        <col style={{ width: SHIPMENT_MATRIX_COL_BRAND }} />
        <col style={{ width: SHIPMENT_MATRIX_COL_NAME }} />
        <col style={{ width: SHIPMENT_MATRIX_COL_MKT }} />
        <col style={{ width: SHIPMENT_MATRIX_COL_SEGMENT }} />
        {monthCols.map((c) => (
          <col key={`ship-col-m-${c.key}`} style={{ width: INVENTORY_DATE_COL_PX }} />
        ))}
        {filteredDateColumns.map((dt) => (
          <col key={`ship-col-d-${dt}`} style={{ width: INVENTORY_DATE_COL_PX }} />
        ))}
      </colgroup>
    );
  }, [
    isShipmentScope,
    shipmentViewMode,
    scopeResultCache.SHIPMENT?.shipment_month_columns,
    filteredDateColumns,
  ]);
  /** 출고 표: 합계 행 spacer용 전체 열 수(고정 5 + 월 + 일) */
  const shipmentBodyColCount = useMemo(() => {
    if (!isShipmentScope || shipmentViewMode !== "status") return 0;
    return (
      5 +
      (scopeResultCache.SHIPMENT?.shipment_month_columns || []).length +
      filteredDateColumns.length
    );
  }, [
    isShipmentScope,
    shipmentViewMode,
    scopeResultCache.SHIPMENT?.shipment_month_columns,
    filteredDateColumns,
  ]);
  const compareTableWidth = useMemo(
    () =>
      Math.max(
        COMPARE_TABLE_MIN_VIEWPORT_PX,
        COMPARE_TABLE_MIN_BASE_PX +
          (OVERSEAS_UPLOAD_COUNTRIES.length + 1) * COMPARE_TABLE_PER_COUNTRY_PX,
      ),
    [],
  );
  const datePeekFadeStyle =
    !isCompareScope &&
    inventoryStickyWidth > 0 &&
    (filteredDateColumns.length > 0 || isShipmentScope)
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
      const chips = countryChipsRef.current;
      if (chips) observer.observe(chips);
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
    isShipmentScope,
    countryTabMode,
    shipmentViewMode,
    showTopScroll,
  ]);

  useEffect(() => {
    if (!isShipmentScope || shipmentViewMode !== "chart") {
      setShowShipmentChartDayMatrixTopScroll(false);
      setShipmentChartDayMatrixTopScrollWidth(0);
      return;
    }
    const el = shipmentChartDayMatrixScrollRef.current;
    if (!el) return;
    const measure = () => {
      const w = el.scrollWidth || 0;
      setShipmentChartDayMatrixTopScrollWidth(w);
      setShowShipmentChartDayMatrixTopScroll(w > (el.clientWidth || 0) + 1);
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
    isShipmentScope,
    shipmentViewMode,
    shipmentChartSelectedSku,
    shipmentChartMonth,
    shipmentChartChannelDailyMatrix,
  ]);

  useEffect(() => {
    if (!hasTableData || isCompareScope) return;
    const topEl = topScrollRef.current;
    const headerEl = headerScrollRef.current;
    const tableEl = tableScrollRef.current;
    if (!tableEl) return;

    const scrollToLatest = () => {
      const nextScrollLeft = Math.max(
        0,
        Math.round(tableEl.scrollWidth - tableEl.clientWidth),
      );
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
    const raw =
      source === "top"
        ? topEl?.scrollLeft ?? 0
        : source === "header"
          ? headerEl?.scrollLeft ?? 0
          : tableEl.scrollLeft;
    const sl = Math.max(0, Math.round(Number(raw)));

    if (source === "header" && headerEl) {
      headerEl.scrollLeft = sl;
      tableEl.scrollLeft = headerEl.scrollLeft;
    } else if (source === "top" && topEl) {
      topEl.scrollLeft = sl;
      tableEl.scrollLeft = topEl.scrollLeft;
    } else {
      tableEl.scrollLeft = sl;
    }

    const canon = tableEl.scrollLeft;
    if (headerEl) headerEl.scrollLeft = canon;
    if (topEl) topEl.scrollLeft = canon;

    requestAnimationFrame(() => {
      const c = tableEl.scrollLeft;
      if (headerEl) headerEl.scrollLeft = c;
      if (topEl) topEl.scrollLeft = c;
      syncingScrollRef.current = false;
    });
  }

  function syncShipmentChartDayMatrixScroll(source) {
    if (syncingShipmentChartDayMatrixScrollRef.current) return;
    const topEl = shipmentChartDayMatrixTopScrollRef.current;
    const bodyEl = shipmentChartDayMatrixScrollRef.current;
    if (!bodyEl) return;
    syncingShipmentChartDayMatrixScrollRef.current = true;
    const raw = source === "top" ? topEl?.scrollLeft ?? 0 : bodyEl.scrollLeft;
    const sl = Math.max(0, Math.round(Number(raw)));
    if (source === "top" && topEl) {
      topEl.scrollLeft = sl;
      bodyEl.scrollLeft = topEl.scrollLeft;
    } else {
      bodyEl.scrollLeft = sl;
      if (topEl) topEl.scrollLeft = sl;
    }
    const canon = bodyEl.scrollLeft;
    if (topEl) topEl.scrollLeft = canon;
    requestAnimationFrame(() => {
      const c = bodyEl.scrollLeft;
      if (topEl) topEl.scrollLeft = c;
      syncingShipmentChartDayMatrixScrollRef.current = false;
    });
  }

  function syncPoSavedScroll(source) {
    if (poSavedScrollSyncingRef.current) return;
    const topEl = poSavedTopScrollRef.current;
    const headerEl = poSavedHeaderScrollRef.current;
    const tableEl = poSavedTableScrollRef.current;
    if (!headerEl && !tableEl) return;
    poSavedScrollSyncingRef.current = true;
    const raw =
      source === "top"
        ? topEl?.scrollLeft ?? 0
        : source === "header"
          ? headerEl?.scrollLeft ?? 0
          : tableEl?.scrollLeft ?? 0;
    const sl = Math.max(0, Math.round(Number(raw)));

    if (source === "header" && headerEl && tableEl) {
      headerEl.scrollLeft = sl;
      tableEl.scrollLeft = headerEl.scrollLeft;
    } else if (source === "top" && topEl && tableEl) {
      topEl.scrollLeft = sl;
      tableEl.scrollLeft = topEl.scrollLeft;
    } else if (tableEl) {
      tableEl.scrollLeft = sl;
    }

    const canon = tableEl?.scrollLeft ?? sl;
    if (headerEl) headerEl.scrollLeft = canon;
    if (topEl) topEl.scrollLeft = canon;

    requestAnimationFrame(() => {
      if (!tableEl) {
        poSavedScrollSyncingRef.current = false;
        return;
      }
      const c = tableEl.scrollLeft;
      if (headerEl) headerEl.scrollLeft = c;
      if (topEl) topEl.scrollLeft = c;
      poSavedScrollSyncingRef.current = false;
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

  /** 출고 현황: 상단 thead(th) / 합계 아래 보조 행(td)에 동일 라벨·열 구조 */
  function renderShipmentStatusHeaderCells(asSubBodyRow) {
    const shipMeta = scopeResultCache.SHIPMENT;
    const monthCols = shipMeta?.shipment_month_columns || [];
    const dayLabels = shipMeta?.shipment_day_labels || {};
    const subCls = asSubBodyRow ? " shipmentMatrixSubHeaderCell" : "";
    const fixed = (cls, k, label) =>
      asSubBodyRow ? (
        <td key={k} className={`${cls}${subCls}`}>
          {label}
        </td>
      ) : (
        <th key={k} className={`${cls}${subCls}`}>
          {label}
        </th>
      );
    return (
      <>
        {fixed("stickyCol stickyColShipCode", "h-code", "상품코드")}
        {fixed("stickyCol stickyColShipBrand", "h-brand", "브랜드")}
        {fixed("stickyCol stickyColShipName", "h-name", "상품명")}
        {fixed("stickyCol stickyColShipMkt", "h-mkt", "마케팅 우선순위")}
        {fixed("stickyCol stickyColShipSegment stickyColBoundary", "h-seg", "구분")}
        {monthCols.map((col) =>
          asSubBodyRow ? (
            <td key={`m-${col.key}`} className={`dateCol${subCls}`}>
              {col.label}
            </td>
          ) : (
            <th key={`m-${col.key}`} className={`dateCol${subCls}`}>
              {col.label}
            </th>
          ),
        )}
        {filteredDateColumns.map((dt) =>
          asSubBodyRow ? (
            <td key={dt} className={`dateCol${subCls}`}>
              {dayLabels[dt] || renderDateHeader(dt)}
            </td>
          ) : (
            <th key={dt} className={`dateCol${subCls}`}>
              {dayLabels[dt] || renderDateHeader(dt)}
            </th>
          ),
        )}
      </>
    );
  }

  function renderInventoryHeaderCells() {
    if (isShipmentScope && shipmentViewMode === "status") {
      return renderShipmentStatusHeaderCells(false);
    }
    return (
      <>
        {isKRScope && <th className="stickyCol stickyColCode">상품코드</th>}
        {isKRScope && <th className="stickyCol stickyColBrand">브랜드</th>}
        {!isKRScope && <th className="stickyCol stickyColCode">상품코드</th>}
        <th className="stickyCol stickyColName stickyColBoundary">상품명</th>
        {isOverseasScope && (
          <th className="stickyCol stickyColKrName stickyColBoundary">한국상품명</th>
        )}
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
    if (isShipmentScope || isInventoryAdminScope || countryTabMode === "COMPARE") return;
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
      const unknownSkusMsg = formatUnknownSkusUserMessage(detail);
      if (unknownSkusMsg) {
        window.alert(unknownSkusMsg);
        setInventoryError(unknownSkusMsg);
        setScopeErrorCache((prev) => ({ ...prev, [scopeKey]: unknownSkusMsg }));
      } else {
        const msg = formatInventoryAggregateFailureMessage(err);
        setInventoryError(msg);
        setScopeErrorCache((prev) => ({ ...prev, [scopeKey]: msg }));
      }
    } finally {
      setInventoryLoading(false);
    }
  }

  async function runShipmentAggregate() {
    if (!isShipmentScope) return false;
    const pending = shipmentFileEntries.filter((e) => e.file);
    if (!pending.length) {
      setInventoryError("출고 엑셀 파일을 먼저 업로드해 주세요.");
      return false;
    }
    setInventoryRequested(true);
    setInventoryError("");
    setScopeErrorCache((prev) => ({ ...prev, SHIPMENT: "" }));
    setShipmentLoading(true);
    try {
      const formData = new FormData();
      pending.forEach((e) => formData.append("files", e.file));
      const res = await axios.post(`${API_BASE}/api/inventory/shipment/aggregate`, formData, {
        headers: { "Content-Type": "multipart/form-data" },
      });
      const data = res?.data || {};
      setScopeResultCache((prev) => ({
        ...prev,
        SHIPMENT: {
          rows: data.rows || [],
          dates: data.dates || [],
          channels: data.channels?.length ? data.channels : SHIPMENT_MATRIX_CHIPS,
          summary: data.summary || { item_count: 0, date_count: 0 },
          countries: data.countries || ["KR"],
          shipment_month_columns: data.shipment_month_columns || [],
          shipment_day_labels: data.shipment_day_labels || {},
          shipment_totals: data.shipment_totals || {},
          shipment_active_channel: data.shipment_active_channel ?? null,
          shipment_channels_with_data: data.shipment_channels_with_data || [],
          requested: true,
        },
      }));
      const syncCh = data.shipment_active_channel ? String(data.shipment_active_channel) : "";
      if (syncCh) {
        shipmentMatrixChannelRef.current = syncCh;
        setShipmentMatrixChannel(syncCh);
      }
      try {
        await hydratePersistedState();
      } catch (syncErr) {
        console.warn("출고 집계 후 동기화 실패 — 재시도 후에도 무시 가능", syncErr);
        await new Promise((r) => setTimeout(r, 900));
        await hydratePersistedState().catch(() => {});
      }
      return true;
    } catch (err) {
      const res = err?.response;
      const detail = res?.data?.detail;
      const unknownSkusMsg = formatUnknownSkusUserMessage(detail);
      if (unknownSkusMsg) {
        window.alert(unknownSkusMsg);
        setInventoryError(unknownSkusMsg);
        setScopeErrorCache((prev) => ({ ...prev, SHIPMENT: unknownSkusMsg }));
        return false;
      }
      let msg = "출고 통합 중 오류";
      if (Array.isArray(detail)) msg = detail.join("\n");
      else if (typeof detail === "string") msg = detail;
      else if (detail && typeof detail === "object" && detail.message) msg = String(detail.message);
      else if (err?.message) msg = err.message;
      setInventoryError(msg);
      setScopeErrorCache((prev) => ({ ...prev, SHIPMENT: msg }));
      return false;
    } finally {
      setShipmentLoading(false);
    }
  }

  async function exportCurrentView() {
    if (isShipmentScope) {
      if (shipmentViewMode === "chart") return;
      if (!shipmentDisplayRows.length) return;
      const monthCols = scopeResultCache.SHIPMENT?.shipment_month_columns || [];
      const dayLabels = scopeResultCache.SHIPMENT?.shipment_day_labels || {};
      const dateCols =
        filteredDateColumns.length > 0 ? filteredDateColumns : rawInventoryDates;
      const chLabel = shipmentMatrixChannel.trim() || scopeResultCache.SHIPMENT?.shipment_active_channel || "출고";
      const matrixRows = shipmentDisplayRows.map((row) => {
        const out = {
          상품코드: getRowSku(row) || "–",
          브랜드: String(row.brand || "").trim() || "–",
          상품명: row.description ? row.description : "(상품명 없음)",
          마케팅우선순위: String(row.mkt_priority || "").trim() || "–",
          구분: String(row.segment || "").trim() || "–",
        };
        const mt = row.month_totals || {};
        for (const col of monthCols) {
          const v = mt[col.key];
          out[col.label] = v != null && Number(v) !== 0 ? toFixed(v, 0) : "-";
        }
        for (const dt of dateCols) {
          const head = dayLabels[dt] || dt;
          const val = row[dt];
          const n = shipmentCellNumericTotal(val);
          if (n === 0) out[head] = "-";
          else if (typeof val === "string" && val.includes("창고이동")) out[head] = val;
          else out[head] = toFixed(n, 0);
        }
        return out;
      });
      await downloadInventoryDashboardXlsx(`출고_${chLabel}`, "출고현황", matrixRows);
      return;
    }
    if (isCompareScope) {
      if (!filteredCompareRows.length) return;
      const rows = filteredCompareRows.map((row) => {
        const out = {
          상품코드: row["상품코드"],
          한국상품명: row["한국상품명"] || "-",
          구분: row["구분"] || "-",
          한국: row["한국 현 재고"] === null ? "-" : toFixed(row["한국 현 재고"], 0),
        };
        for (const code of OVERSEAS_UPLOAD_COUNTRIES) {
          const key = countryLabel(code);
          out[key] = row[key] === null ? "-" : toFixed(row[key], 0);
        }
        return out;
      });
      await downloadInventoryDashboardXlsx("재고비교_대시보드", "재고비교", rows);
      return;
    }
    if (isKRScope) {
      if (!krDisplayRows.length || !filteredDateColumns.length) return;
      const rows = krDisplayRows.map((row) => {
        const out = {
          브랜드: row.supplier,
          상품코드: getRowSku(row),
          상품명: row.description ? row.description : "(상품명 없음)",
        };
        for (const dt of filteredDateColumns) {
          out[formatInventoryDateHeaderForExport(dt)] = toFixed(row[dt], 0);
        }
        return out;
      });
      await downloadInventoryDashboardXlsx("한국_현 재고_대시보드", "한국재고", rows);
      return;
    }
    if (!overseasDisplayRows.length || !filteredDateColumns.length) return;
    const rows = overseasDisplayRows.map((row) => {
      const krMatchKey = getCanonicalMatchCode(row);
      const out = {
        상품코드: getRowSku(row),
        상품명: row.description ? row.description : "(상품명 없음)",
        한국상품명: getCompareDisplayName(row) || krNameMap.get(krMatchKey) || "-",
      };
      if (showKrCompare) {
        out["한국 현 재고"] = krCompareMap.size
          ? toFixed(krCompareMap.get(krMatchKey) || 0, 0)
          : "-";
      }
      for (const dt of filteredDateColumns) {
        out[formatInventoryDateHeaderForExport(dt)] = toFixed(row[dt], 0);
      }
      return out;
    });
    const overseasFileBase = `${countryLabel(selectedOverseasCountry)}_재고_대시보드`;
    await downloadInventoryDashboardXlsx(overseasFileBase, "해외재고", rows);
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
    if (isShipmentScope) {
      document.getElementById("shipment-files-input")?.click();
      return;
    }
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

  async function fetchPurchaseOrdersList() {
    const res = await axios.get(`${API_BASE}/api/inventory/purchase-orders`, {
      timeout: PURCHASE_ORDERS_LIST_TIMEOUT_MS,
    });
    setPurchaseOrders(Array.isArray(res?.data?.items) ? res.data.items : []);
  }

  /** 저장된 발주 패널용: 로딩·에러를 묶어 처리 (범위 이탈·탭 전환 시에도 끊김 없이) */
  async function reloadPurchaseOrdersUi() {
    setPurchaseOrdersLoading(true);
    setPurchaseOrderError("");
    try {
      await fetchPurchaseOrdersList();
    } catch (err) {
      const det = err?.response?.data?.detail;
      setPurchaseOrderError(
        Array.isArray(det) ? det.join("\n") : det || formatPersistedLoadError(err, API_BASE)
      );
    } finally {
      setPurchaseOrdersLoading(false);
    }
  }

  async function deletePurchaseOrder(poId, options = {}) {
    const confirmMessage =
      options.confirmMessage ??
      "이 발주를 삭제할까요? 연결된 입고 차수도 함께 삭제되며 되돌릴 수 없습니다.";
    if (!window.confirm(confirmMessage)) {
      return;
    }
    setPurchaseOrderError("");
    setPurchaseOrderSuccess("");
    try {
      await axios.delete(`${API_BASE}/api/inventory/purchase-orders/${poId}`);
      if (String(editingPoId) === String(poId)) {
        cancelPoEdit();
      }
      setSavedPoNewLineDraftByOrderId((prev) => {
        if (!prev[String(poId)]) return prev;
        const next = { ...prev };
        delete next[String(poId)];
        return next;
      });
      setInboundDrafts((prev) => {
        const next = { ...prev };
        delete next[poId];
        return next;
      });
      await fetchPurchaseOrdersList();
    } catch (err) {
      const det = err?.response?.data?.detail;
      setPurchaseOrderError(Array.isArray(det) ? det.join("\n") : det || err?.message || "삭제 실패");
    }
  }

  function toggleSavedPoBulkKey(rawKey) {
    setSavedPoBulkSelected((prev) => {
      const next = { ...prev };
      if (next[rawKey]) {
        delete next[rawKey];
        return next;
      }
      next[rawKey] = true;
      if (rawKey.startsWith("P|")) {
        const oid = rawKey.slice(2);
        for (const k of Object.keys(next)) {
          if (k.startsWith(`L|${oid}|`)) delete next[k];
        }
      } else if (rawKey.startsWith("L|")) {
        const parts = rawKey.split("|");
        const oid = parts[1];
        if (oid) delete next[`P|${oid}`];
      }
      return next;
    });
  }

  function setSavedPoBulkSelectAllVisible(selectAll) {
    if (!selectAll) {
      setSavedPoBulkSelected({});
      return;
    }
    setSavedPoBulkSelected(Object.fromEntries(savedPoBulkSelectableKeysFlat.map((k) => [k, true])));
  }

  async function bulkDeleteSavedPoSelections() {
    const keys = Object.keys(savedPoBulkSelected).filter((k) => savedPoBulkKeySet.has(k));
    if (!keys.length) return;

    const wholePo = new Set();
    const linePairs = [];
    for (const k of keys) {
      if (k.startsWith("P|")) {
        const id = k.slice(2);
        if (id) wholePo.add(id);
      } else if (k.startsWith("L|")) {
        const parts = k.split("|");
        const oid = parts[1];
        const lid = parts[2];
        if (oid && lid) linePairs.push({ orderId: oid, lineId: lid });
      }
    }
    const lineFiltered = linePairs.filter(({ orderId }) => !wholePo.has(orderId));
    const n = wholePo.size + lineFiltered.length;
    if (!window.confirm(`선택한 입고 ${formatInt(n)}건을 삭제할까요? 되돌릴 수 없습니다.`)) {
      return;
    }

    setPurchaseOrderError("");
    setPurchaseOrderSuccess("");
    const touchedOrderIds = new Set([...wholePo, ...lineFiltered.map((x) => x.orderId)]);

    try {
      for (const { orderId, lineId } of lineFiltered) {
        await axios.delete(
          `${API_BASE}/api/inventory/purchase-orders/${orderId}/inbound-lines/${lineId}`
        );
      }
      for (const poId of wholePo) {
        await axios.delete(`${API_BASE}/api/inventory/purchase-orders/${poId}`);
        if (String(editingPoId) === String(poId)) {
          cancelPoEdit();
        }
      }
      setSavedPoInline(null);
      setSavedPoBulkSelected({});
      setSavedPoNewLineDraftByOrderId((prev) => {
        const next = { ...prev };
        for (const id of touchedOrderIds) delete next[String(id)];
        return next;
      });
      setInboundDrafts((prev) => {
        const next = { ...prev };
        for (const id of touchedOrderIds) delete next[id];
        return next;
      });
      await fetchPurchaseOrdersList();
    } catch (err) {
      const det = err?.response?.data?.detail;
      setPurchaseOrderError(Array.isArray(det) ? det.join("\n") : det || err?.message || "삭제 실패");
      await fetchPurchaseOrdersList();
    }
  }

  async function resolvePurchaseOrderSku() {
    const sku = String(poForm.sku || "").trim();
    if (!sku) {
      setSkuResolveHint("");
      return;
    }
    try {
      const res = await axios.get(`${API_BASE}/api/inventory/purchase-orders/sku-hint`, {
        params: { sku },
      });
      const d = res?.data || {};
      if (d.matched) {
        setPoForm((p) => ({
          ...p,
          brand: d.brand != null ? String(d.brand) : "",
          product_name: d.product_name != null ? String(d.product_name) : "",
        }));
        setSkuResolveHint("");
      } else {
        setSkuResolveHint(String(d.message || "등록된 상품코드가 없습니다."));
      }
    } catch (err) {
      const det = err?.response?.data?.detail;
      setSkuResolveHint(Array.isArray(det) ? det.join("\n") : det || "상품코드 조회 중 오류");
    }
  }

  async function submitPurchaseOrder() {
    setPurchaseOrderError("");
    setPurchaseOrderSuccess("");
    const pt =
      poForm.product_type_preset === "본품"
        ? "본품"
        : String(poForm.product_type_custom || "").trim() || "직접입력";
    const qtyRaw = String(poForm.total_quantity || "").replace(/,/g, "").trim();
    const odTrim = String(poForm.order_date || "").trim();
    const odnTrim = String(poForm.order_date_note || "").trim();
    if (!qtyRaw || Number.isNaN(Number(qtyRaw))) {
      setPurchaseOrderError("총 발주수량을 올바른 숫자로 입력하세요.");
      return;
    }
    try {
      await axios.post(`${API_BASE}/api/inventory/purchase-orders`, {
        order_date: odTrim || null,
        order_date_note: odTrim ? null : isOrderDatePlannedNote(odnTrim) ? "발주 예정" : null,
        erp_po_number: String(poForm.erp_po_number || "").trim(),
        product_type: pt,
        sku: String(poForm.sku || "").trim(),
        brand: String(poForm.brand || "").trim(),
        product_name: String(poForm.product_name || "").trim(),
        manufacturer: String(poForm.manufacturer || "").trim(),
        total_quantity: Number(qtyRaw),
        delivery_available_date: poForm.delivery_available_tbd ? null : poForm.delivery_available_date || null,
        expected_inbound_date: poForm.expected_inbound_tbd ? null : poForm.expected_inbound_date || null,
      });
      setPurchaseOrderSuccess("발주가 저장되었습니다.");
      setPoForm({ ...EMPTY_PURCHASE_ORDER_FORM });
      setSkuResolveHint("");
      setPurchaseOrderSubTab("saved");
      await fetchPurchaseOrdersList();
    } catch (err) {
      const det = err?.response?.data?.detail;
      setPurchaseOrderError(Array.isArray(det) ? det.join("\n") : det || err?.message || "저장 실패");
    }
  }

  async function downloadPoInboundTemplateClick() {
    setPurchaseOrderError("");
    setPurchaseOrderSuccess("");
    try {
      await downloadPoInboundTemplateXlsx();
    } catch (err) {
      setPurchaseOrderError(err?.message || "템플릿을 만드는 중 오류가 났습니다.");
    }
  }

  function openPoOrderFileInput() {
    const el = document.getElementById("po-order-file-input");
    if (el) el.click();
  }

  async function handlePoInboundFileSelected(ev) {
    const f = ev.target.files?.[0];
    ev.target.value = "";
    if (!f) return;
    setPurchaseOrderError("");
    setPurchaseOrderSuccess("");
    setPoInboundUploadBusy(true);
    try {
      const buf = await f.arrayBuffer();
      const { rows, errors } = parsePurchaseOrderInboundSheet(buf);
      if (errors.length) {
        setPurchaseOrderError(errors.join("\n"));
        return;
      }
      if (!rows.length) {
        setPurchaseOrderError("업로드할 데이터 행이 없습니다. 템플릿 2행 이후에 내용을 입력했는지 확인해 주세요.");
        return;
      }
      await axios.post(`${API_BASE}/api/inventory/purchase-orders/import`, { rows });
      setPurchaseOrderSuccess(`${formatInt(rows.length)}건이 저장된 발주에 반영되었습니다.`);
      setPurchaseOrderSubTab("saved");
      await fetchPurchaseOrdersList();
    } catch (err) {
      const det = err?.response?.data?.detail;
      setPurchaseOrderError(Array.isArray(det) ? det.join("\n") : det || err?.message || "업로드 실패");
    } finally {
      setPoInboundUploadBusy(false);
      setPoInboundFileKey((k) => k + 1);
    }
  }

  /** `options.fromSavedSheet` + `options.draft` 는 표 인라인 신규 차수 저장 시 사용 */
  async function submitInboundLine(orderId, options = {}) {
    const fromSavedSheet = Boolean(options.fromSavedSheet);
    const d =
      options.draft ??
      inboundDrafts[orderId] ?? {
        ...EMPTY_INBOUND_LINE_DRAFT,
      };
    setPurchaseOrderError("");
    setPurchaseOrderSuccess("");
    const qRaw = String(d.quantity || "").replace(/,/g, "").trim();
    if (!qRaw || Number.isNaN(Number(qRaw))) {
      setPurchaseOrderError("입고수량을 올바른 숫자로 입력하세요.");
      return;
    }
    const actParsed = normalizeActualInboundInput(String(d.actual_inbound_input || "").trim());
    const hasInDate = String(actParsed.actual_inbound_date || "").trim();
    const hasInNote = String(actParsed.actual_inbound_note || "").trim();
    let nextInboundStatus = String(d.inbound_status || "X").trim().toUpperCase();
    if (hasInDate || hasInNote) {
      nextInboundStatus = "O";
    }
    if (nextInboundStatus === "O" && !hasInDate && !hasInNote) {
      setPurchaseOrderError("입고 완료(O)이면 실제입고일 또는 비고(예: 예외 입고·무상 입고) 중 하나는 입력하세요.");
      return;
    }
    try {
      const res = await axios.post(`${API_BASE}/api/inventory/purchase-orders/${orderId}/inbounds`, {
        delivery_available_date: d.delivery_available_tbd ? null : d.delivery_available_date?.trim() || null,
        expected_inbound_date:
          String(nextInboundStatus || "").toUpperCase() === "O" || d.expected_inbound_tbd
            ? null
            : d.expected_inbound_date?.trim() || null,
        actual_inbound_date: actParsed.actual_inbound_date,
        actual_inbound_note: hasInNote ? actParsed.actual_inbound_note : null,
        quantity: Number(qRaw),
        inbound_status: nextInboundStatus,
      });
      if (fromSavedSheet) {
        setSavedPoNewLineDraftByOrderId((p) => {
          const next = { ...p };
          delete next[String(orderId)];
          return next;
        });
      } else {
        setInboundDrafts((p) => ({
          ...p,
          [orderId]: {
            ...EMPTY_INBOUND_LINE_DRAFT,
          },
        }));
      }
      const newLine = res?.data;
      if (newLine && String(editingPoId) === String(orderId)) {
        setPoEditDraft((prev) => {
          if (!prev) return prev;
          return {
            ...prev,
            lines: [
              ...prev.lines,
              {
                id: String(newLine.id),
                line_no: newLine.line_no,
                ref_code: newLine.ref_code,
                delivery_available_date: newLine.delivery_available_date || "",
                delivery_available_tbd: !newLine.delivery_available_date,
                expected_inbound_date: newLine.expected_inbound_date || "",
                actual_inbound_date: newLine.actual_inbound_date || "",
                actual_inbound_note: newLine.actual_inbound_note || "",
                line_memo: newLine.line_memo || "",
                quantity: newLine.quantity != null ? String(newLine.quantity) : "",
                inbound_status: newLine.inbound_status || "O",
              },
            ],
          };
        });
      }
      await fetchPurchaseOrdersList();
    } catch (err) {
      const det = err?.response?.data?.detail;
      setPurchaseOrderError(Array.isArray(det) ? det.join("\n") : det || err?.message || "입고 저장 실패");
    }
  }

  async function patchInboundLineField(orderId, lineId, patch) {
    setPurchaseOrderError("");
    setPurchaseOrderSuccess("");
    try {
      await axios.patch(`${API_BASE}/api/inventory/purchase-orders/${orderId}/inbound-lines/${lineId}`, patch);
      setSavedPoInline(null);
      await fetchPurchaseOrdersList();
      return true;
    } catch (err) {
      const det = err?.response?.data?.detail;
      setPurchaseOrderError(Array.isArray(det) ? det.join("\n") : det || err?.message || "저장 실패");
      return false;
    }
  }

  /** 입고 차수가 없으면 1차 입고를 만들고 새 줄 id 반환 (저장 표 연필용) */
  async function ensureFirstInboundLine(po) {
    const sorted = [...(po.inbound_lines || [])].sort(
      (a, b) => (Number(a.line_no) || 0) - (Number(b.line_no) || 0)
    );
    if (sorted.length > 0) {
      return String(sorted[0].id);
    }
    const pid = String(po.id);
    if (poInboundSeedLockRef.current.has(pid)) {
      return null;
    }
    const q = po.total_quantity != null ? Number(po.total_quantity) : NaN;
    if (!Number.isFinite(q)) {
      setPurchaseOrderError("입고 행이 없습니다. 총 발주수량을 확인한 뒤 다시 시도해 주세요.");
      return null;
    }
    poInboundSeedLockRef.current.add(pid);
    setPurchaseOrderError("");
    try {
      const res = await axios.post(`${API_BASE}/api/inventory/purchase-orders/${pid}/inbounds`, {
        delivery_available_date: po.delivery_available_date || null,
        expected_inbound_date: po.expected_inbound_date || null,
        actual_inbound_date: null,
        actual_inbound_note: null,
        quantity: q,
        inbound_status: "X",
      });
      const newId = res?.data?.id;
      await fetchPurchaseOrdersList();
      return newId != null ? String(newId) : null;
    } catch (err) {
      const det = err?.response?.data?.detail;
      setPurchaseOrderError(
        Array.isArray(det) ? det.join("\n") : det || err?.message || "입고 행 생성 실패"
      );
      return null;
    } finally {
      poInboundSeedLockRef.current.delete(pid);
    }
  }

  function buildPurchaseOrderUpdatePayloadFromPo(po, patch = {}) {
    return {
      order_date: po.order_date || null,
      order_date_note: po.order_date_note || null,
      erp_po_number: String(po.erp_po_number || "").trim(),
      product_type: String(po.product_type || "본품").trim() || "본품",
      sku: String(po.sku || "").trim(),
      brand: String(po.brand || "").trim(),
      product_name: String(po.product_name || "").trim(),
      manufacturer: String(po.manufacturer || "").trim(),
      total_quantity: Number(po.total_quantity),
      delivery_available_date: po.delivery_available_date || null,
      expected_inbound_date: po.expected_inbound_date || null,
      ...patch,
    };
  }

  async function patchPurchaseOrderField(orderId, patch) {
    setPurchaseOrderError("");
    setPurchaseOrderSuccess("");
    const po = purchaseOrders.find((p) => String(p.id) === String(orderId));
    if (!po) {
      setPurchaseOrderError("발주 정보를 찾을 수 없습니다. 목록을 새로고침합니다.");
      await fetchPurchaseOrdersList();
      return false;
    }
    try {
      await axios.patch(
        `${API_BASE}/api/inventory/purchase-orders/${orderId}`,
        buildPurchaseOrderUpdatePayloadFromPo(po, patch)
      );
      setSavedPoInline(null);
      await fetchPurchaseOrdersList();
      return true;
    } catch (err) {
      const det = err?.response?.data?.detail;
      setPurchaseOrderError(Array.isArray(det) ? det.join("\n") : det || err?.message || "저장 실패");
      return false;
    }
  }

  async function closeSavedPoMemoModal() {
    if (poMemoAutosaveTimerRef.current) {
      clearTimeout(poMemoAutosaveTimerRef.current);
      poMemoAutosaveTimerRef.current = null;
    }
    const ids = poMemoModalIdsRef.current;
    const draft = poMemoDraftRef.current;
    setSavedPoMemoModal(null);
    poMemoModalIdsRef.current = null;
    poMemoDraftRef.current = "";
    if (ids) {
      await patchInboundLineField(ids.orderId, ids.lineId, {
        line_memo: String(draft || "").trim() || null,
      });
    }
  }

  async function deleteSavedPoMemo() {
    if (poMemoAutosaveTimerRef.current) {
      clearTimeout(poMemoAutosaveTimerRef.current);
      poMemoAutosaveTimerRef.current = null;
    }
    const ids = poMemoModalIdsRef.current;
    if (!ids) return;
    poMemoDraftRef.current = "";
    const ok = await patchInboundLineField(ids.orderId, ids.lineId, { line_memo: null });
    if (ok) {
      setSavedPoMemoModal(null);
      poMemoModalIdsRef.current = null;
    }
  }

  function startPoEdit(po) {
    const oid = String(po.id);
    setSavedPoNewLineDraftByOrderId((prev) => {
      if (!prev[oid]) return prev;
      const next = { ...prev };
      delete next[oid];
      return next;
    });
    const { preset, custom } = purchaseOrderProductTypeToFields(po.product_type);
    setEditingPoId(oid);
    setPoEditDraft({
      order_date: po.order_date || "",
      order_date_note: po.order_date_note || "",
      erp_po_number: po.erp_po_number || "",
      product_type_preset: preset,
      product_type_custom: custom,
      sku: po.sku || "",
      brand: po.brand || "",
      product_name: po.product_name || "",
      manufacturer: po.manufacturer || "",
      total_quantity: po.total_quantity != null ? String(po.total_quantity) : "",
      delivery_available_date: po.delivery_available_date || "",
      expected_inbound_date: po.expected_inbound_date || "",
      expected_inbound_tbd: !po.expected_inbound_date,
      skuResolveHint: "",
      lines: (po.inbound_lines || []).map((l) => ({
        id: l.id,
        line_no: l.line_no,
        ref_code: l.ref_code,
        delivery_available_date: l.delivery_available_date || "",
        expected_inbound_date:
          String(l.inbound_status || "").trim().toUpperCase() === "O"
            ? ""
            : l.expected_inbound_date || "",
        actual_inbound_date: l.actual_inbound_date || "",
        actual_inbound_note: l.actual_inbound_note || "",
        line_memo: l.line_memo || "",
        quantity: l.quantity != null ? String(l.quantity) : "",
        inbound_status: l.inbound_status || "O",
      })),
    });
  }

  function cancelPoEdit() {
    setEditingPoId(null);
    setPoEditDraft(null);
  }

  function openSavedPoNewLineRow(po) {
    const oid = String(po.id);
    setEditingPoId(null);
    setPoEditDraft(null);
    setSavedPoNewLineDraftByOrderId((prev) =>
      prev[oid] ? prev : { ...prev, [oid]: { ...EMPTY_INBOUND_LINE_DRAFT } }
    );
    window.setTimeout(() => {
      savedPoNewLineRowRefs.current[oid]?.scrollIntoView({
        block: "nearest",
        inline: "nearest",
        behavior: "smooth",
      });
    }, 0);
  }

  function cancelSavedPoNewLineRow(orderId) {
    const oid = String(orderId);
    setSavedPoNewLineDraftByOrderId((prev) => {
      if (!prev[oid]) return prev;
      const next = { ...prev };
      delete next[oid];
      return next;
    });
  }

  function updateSavedPoNewLineDraft(orderId, patchOrUpdater) {
    const oid = String(orderId);
    setSavedPoNewLineDraftByOrderId((prev) => {
      const cur = prev[oid] || { ...EMPTY_INBOUND_LINE_DRAFT };
      const patch =
        typeof patchOrUpdater === "function" ? patchOrUpdater(cur) : patchOrUpdater;
      return { ...prev, [oid]: { ...cur, ...patch } };
    });
  }

  async function resolvePoEditSku() {
    if (!poEditDraft) return;
    const sku = String(poEditDraft.sku || "").trim();
    if (!sku) {
      setPoEditDraft((p) => (p ? { ...p, skuResolveHint: "" } : p));
      return;
    }
    try {
      const res = await axios.get(`${API_BASE}/api/inventory/purchase-orders/sku-hint`, {
        params: { sku },
      });
      const d = res?.data || {};
      if (d.matched) {
        setPoEditDraft((p) =>
          p
            ? {
                ...p,
                brand: d.brand != null ? String(d.brand) : p.brand,
                product_name: d.product_name != null ? String(d.product_name) : p.product_name,
                skuResolveHint: "",
              }
            : p
        );
      } else {
        setPoEditDraft((p) =>
          p ? { ...p, skuResolveHint: String(d.message || "등록된 상품코드가 없습니다.") } : p
        );
      }
    } catch (err) {
      const det = err?.response?.data?.detail;
      setPoEditDraft((p) =>
        p ? { ...p, skuResolveHint: Array.isArray(det) ? det.join("\n") : det || "상품코드 조회 중 오류" } : p
      );
    }
  }

  async function submitPoEditSave() {
    if (!editingPoId || !poEditDraft) return;
    setPurchaseOrderError("");
    setPurchaseOrderSuccess("");
    const d = poEditDraft;
    const pt =
      d.product_type_preset === "본품"
        ? "본품"
        : String(d.product_type_custom || "").trim() || "직접입력";
    const qtyRaw = String(d.total_quantity || "").replace(/,/g, "").trim();
    const odTrim = String(d.order_date || "").trim();
    const odnTrim = String(d.order_date_note || "").trim();
    if (!String(d.erp_po_number || "").trim()) {
      setPurchaseOrderError("ERP PO 번호는 필수입니다.");
      return;
    }
    if (!String(d.sku || "").trim()) {
      setPurchaseOrderError("상품코드(SKU)는 필수입니다.");
      return;
    }
    if (!qtyRaw || Number.isNaN(Number(qtyRaw))) {
      setPurchaseOrderError("총 발주수량을 올바른 숫자로 입력하세요.");
      return;
    }
    for (const ln of d.lines) {
      const lq = String(ln.quantity || "").replace(/,/g, "").trim();
      if (!lq || Number.isNaN(Number(lq))) {
        setPurchaseOrderError("입고 차수의 수량을 모두 올바른 숫자로 입력하세요.");
        return;
      }
      const lDate = String(ln.actual_inbound_date || "").trim();
      const lNote = String(ln.actual_inbound_note || "").trim();
      if ((ln.inbound_status || "O") === "O" && !lDate && !lNote) {
        setPurchaseOrderError("입고 완료(O) 차수는 실제입고일 또는 비고(예: 예외 입고·무상 입고)가 필요합니다.");
        return;
      }
    }
    try {
      await axios.patch(`${API_BASE}/api/inventory/purchase-orders/${editingPoId}`, {
        order_date: odTrim || null,
        order_date_note: odTrim ? null : isOrderDatePlannedNote(odnTrim) ? "발주 예정" : null,
        erp_po_number: String(d.erp_po_number || "").trim(),
        product_type: pt,
        sku: String(d.sku || "").trim(),
        brand: String(d.brand || "").trim(),
        product_name: String(d.product_name || "").trim(),
        manufacturer: String(d.manufacturer || "").trim(),
        total_quantity: Number(qtyRaw),
        delivery_available_date: d.delivery_available_date || null,
        expected_inbound_date: d.expected_inbound_tbd ? null : d.expected_inbound_date || null,
        inbound_lines: d.lines.map((ln) => ({
          id: ln.id,
          delivery_available_date: ln.delivery_available_date?.trim() || null,
          expected_inbound_date: ln.expected_inbound_date?.trim() || null,
          actual_inbound_date: ln.actual_inbound_date || null,
          actual_inbound_note: String(ln.actual_inbound_note || "").trim() || null,
          line_memo: String(ln.line_memo || "").trim() || null,
          quantity: Number(String(ln.quantity || "").replace(/,/g, "")),
          inbound_status: ln.inbound_status || "O",
        })),
      });
      setEditingPoId(null);
      setPoEditDraft(null);
      await fetchPurchaseOrdersList();
    } catch (err) {
      const det = err?.response?.data?.detail;
      setPurchaseOrderError(Array.isArray(det) ? det.join("\n") : det || err?.message || "수정 실패");
    }
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

  async function fetchAndApplyShipmentView(options = {}) {
    const { signal, channel } = options;
    try {
      const params = { year: 2026 };
      if (channel) params.channel = channel;
      const shipRes = await axios.get(`${API_BASE}/api/inventory/shipment/view`, { params, signal });
      const data = shipRes?.data || {};
      const active = data.shipment_active_channel ? String(data.shipment_active_channel) : "";
      if (active && !channel) {
        setShipmentMatrixChannel(active);
      }
      setScopeResultCache((prev) => ({
        ...prev,
        SHIPMENT: {
          rows: data.rows || [],
          dates: data.dates || [],
          channels: data.channels?.length ? data.channels : SHIPMENT_MATRIX_CHIPS,
          summary: data.summary || { item_count: 0, date_count: 0 },
          countries: data.countries || ["KR"],
          shipment_month_columns: data.shipment_month_columns || [],
          shipment_day_labels: data.shipment_day_labels || {},
          shipment_totals: data.shipment_totals || {},
          shipment_active_channel: data.shipment_active_channel ?? null,
          shipment_channels_with_data: data.shipment_channels_with_data || [],
          requested: true,
        },
      }));
      setScopeErrorCache((prev) => ({ ...prev, SHIPMENT: "" }));
    } catch (err) {
      if (err?.code === "ERR_CANCELED" || err?.name === "CanceledError") return;
      const st = err?.response?.status;
      if (st === 503) return;
      const detail = err?.response?.data?.detail;
      const msg = Array.isArray(detail) ? detail.join("\n") : detail || err?.message || "출고 뷰를 불러오지 못했습니다.";
      setScopeErrorCache((prev) => ({ ...prev, SHIPMENT: msg }));
    }
  }

  async function hydratePersistedState(options = {}) {
    const { preserveLocalOnly = true, excludeCountry = "", excludeEntryId = "" } = options;
    const [persistedFiles, latestMappingSummary] = await Promise.all([
      fetchPersistedFiles(),
      fetchSkuMappingSummary(),
    ]);

    setMappingSummary(latestMappingSummary);
    setScopeErrorCache({});

    const krView = await fetchPersistedScopeView("KR");
    setScopeResultCache((prev) => ({ ...prev, KR: krView }));
    setAvailableCountries((prev) => {
      const s = new Set(prev);
      (krView.countries || ["KR"]).forEach((c) => s.add(String(c).toUpperCase()));
      return Array.from(s);
    });

    for (let i = 0; i < OVERSEAS_UPLOAD_COUNTRIES.length; i += HYDRATE_OVERSEAS_VIEW_CONCURRENCY) {
      const chunk = OVERSEAS_UPLOAD_COUNTRIES.slice(i, i + HYDRATE_OVERSEAS_VIEW_CONCURRENCY);
      const chunkViews = await Promise.all(chunk.map((code) => fetchPersistedScopeView(code)));
      setScopeResultCache((prev) => {
        const next = { ...prev };
        chunk.forEach((code, j) => {
          next[`OVERSEAS:${code}`] = chunkViews[j];
        });
        return next;
      });
      setAvailableCountries((prev) => {
        const s = new Set(prev);
        chunkViews.forEach((v) => {
          (v?.countries || []).forEach((c) => s.add(String(c).toUpperCase()));
        });
        return Array.from(s);
      });
    }

    await fetchAndApplyShipmentView({
      channel: shipmentMatrixChannelRef.current.trim() || undefined,
    });

    setFileEntries((prev) => {
      const invPersisted = persistedFiles.filter((e) => (e.file_domain || "inventory") !== "shipment");
      const persistedKeys = new Set(
        invPersisted.map((entry) => `${entry.name}::${entry.country || ""}::${entry.date || ""}`)
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
      const serverEntries = invPersisted.map((entry) => ({
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

    setShipmentFileEntries((prev) => {
      const shipPersisted = persistedFiles.filter((e) => (e.file_domain || "inventory") === "shipment");
      const persistedKeys = new Set(shipPersisted.map((entry) => `${entry.name}::SHIPMENT`));
      const localOnlyShip = preserveLocalOnly
        ? prev.filter((entry) => {
            if (entry.dbFileId || !entry.file) return false;
            if (excludeEntryId && entry.id === excludeEntryId) return false;
            if (excludeCountry && String(entry.country || "").toUpperCase() === excludeCountry) return false;
            const entryKey = `${entry.name}::SHIPMENT`;
            if (persistedKeys.has(entryKey)) return false;
            return true;
          })
        : [];
      const serverShip = shipPersisted.map((entry) => ({
        id: entry.file_id,
        dbFileId: entry.file_id,
        file: null,
        name: entry.name,
        size: entry.size,
        country: "SHIPMENT",
        date: "",
      }));
      return [...serverShip, ...localOnlyShip];
    });
  }

  useEffect(() => {
    if (!isPurchaseOrderScope) {
      setPurchaseOrdersLoading(false);
      return;
    }
    void reloadPurchaseOrdersUi();
  }, [isPurchaseOrderScope]);

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
        `${formatInt(Number(res?.data?.processed_file_count || 0))}개 파일에서 ${formatInt(
          Number(res?.data?.merged_item_count || 0)
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
    const label = String(entry.name || "").trim() || "이 파일";
    const isShipmentPersistedFile = String(entry.country || "").toUpperCase() === "SHIPMENT";
    if (
      !window.confirm(
        `「${label}」을(를) 삭제할까요?\n저장된 재고·출고 데이터에서 이 파일에 해당하는 내용이 제거됩니다.\n이 작업은 되돌릴 수 없습니다.`
      )
    ) {
      return;
    }
    try {
      setSettingsMutating(true);
      if (entry.dbFileId) {
        await axios.delete(`${API_BASE}/api/inventory/files/${entry.dbFileId}`);
      }
      // hydrate는 KR/해외 뷰 조회 등 중간에 실패할 수 있어, 삭제 직후 목록에서 먼저 제거해야
      // (실패 시 아래 catch에서 전체 hydrate로 재동기화)
      setFileEntries((prev) => prev.filter((x) => x.id !== entry.id));
      setShipmentFileEntries((prev) => prev.filter((x) => x.id !== entry.id));
      /** 출고 매트릭스 뷰는 `scopeResultCache`에 남으면 삭제 후에도 표가 그대로 보일 수 있음 → 즉시 비움 */
      if (isShipmentPersistedFile) {
        setScopeResultCache((prev) => ({
          ...prev,
          SHIPMENT: {
            rows: [],
            dates: [],
            channels: SHIPMENT_MATRIX_CHIPS,
            summary: { item_count: 0, date_count: 0 },
            countries: ["KR"],
            shipment_month_columns: [],
            shipment_day_labels: {},
            shipment_totals: {},
            shipment_active_channel: null,
            shipment_channels_with_data: [],
            requested: true,
          },
        }));
        setScopeErrorCache((prev) => ({ ...prev, SHIPMENT: "" }));
      }
      await hydratePersistedState({ excludeEntryId: entry.id });
    } catch (err) {
      const detail = err?.response?.data?.detail;
      window.alert(Array.isArray(detail) ? detail.join("\n") : detail || "파일 삭제 중 오류");
      await hydratePersistedState().catch(() => {});
    } finally {
      setSettingsMutating(false);
    }
  }

  async function patchInventoryFileBaseDate(entry, nextIso) {
    if (!entry?.dbFileId) return;
    const prevIso = entry.date || "";
    setSettingsMutating(true);
    setFileEntries((prev) => prev.map((x) => (x.id === entry.id ? { ...x, date: nextIso } : x)));
    try {
      await axios.patch(`${API_BASE}/api/inventory/files/${entry.dbFileId}`, {
        base_date: nextIso && String(nextIso).trim() ? String(nextIso).trim() : null,
      });
      await hydratePersistedState({ preserveLocalOnly: true });
    } catch (err) {
      setFileEntries((prev) => prev.map((x) => (x.id === entry.id ? { ...x, date: prevIso } : x)));
      const detail = err?.response?.data?.detail;
      window.alert(Array.isArray(detail) ? detail.join("\n") : detail || err?.message || "기준일 저장 실패");
    } finally {
      setSettingsMutating(false);
    }
  }

  async function clearFilesByCountry(country, entries) {
    if (!entries?.length) return;
    const label = countryLabel(country);
    if (
      !window.confirm(
        `「${label}」에 저장된 파일과 재고 데이터를 모두 삭제할까요?\n이 작업은 되돌릴 수 없습니다.`
      )
    ) {
      return;
    }
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
    if (
      !window.confirm(
        "모든 국가·출고에 저장된 업로드 파일과 재고·출고 데이터를 모두 삭제할까요?\n이 작업은 되돌릴 수 없습니다."
      )
    ) {
      return;
    }
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
    {mypageView && (
      <div className="mypageOverlayShell">
        <MyPage onBack={() => { setMypageView(false); }} />
      </div>
    )}
    <div
      className="dashboard"
      style={mypageView ? { display: "none" } : undefined}
      aria-hidden={mypageView ? "true" : undefined}
    >
      <div className="dashboardHeroBand">
        <div className="headerArea">
        <section className="hero">
          <div className="heroHead heroHeadDash">
            <h1 className="heroTitle">재고 분석 대시보드</h1>
            <div className="heroHeadAccount">
              <button
                type="button"
                className="ghost heroHeadAccountBtn"
                onClick={() => { setMypageView(true); }}
                title="마이페이지"
              >
                <UserRound className="tabIcon" size={18} strokeWidth={2} aria-hidden />
                마이페이지
              </button>
              <button
                type="button"
                className="ghost heroHeadAccountBtn"
                title="로그아웃 후 로그인 화면으로 이동합니다."
                onClick={() => {
                  clearAccessToken();
                  window.location.reload();
                }}
              >
                <LogOut className="tabIcon" size={18} strokeWidth={2} aria-hidden />
                로그아웃
              </button>
            </div>
          </div>
          <div className="heroActions">
            <button
              type="button"
              className="cautionBtn heroActionCautionShaped"
              onClick={() => setShowCautionModal(true)}
            >
              안내문
            </button>
            <button
              className="primary"
              onClick={onClickUpload}
              disabled={isInventoryAdminScope || countryTabMode === "COMPARE"}
            >
              파일 업로드
            </button>
            <button
              className="primary"
              onClick={() => void (isShipmentScope ? runShipmentAggregate() : aggregateInventory())}
              disabled={
                isInventoryAdminScope ||
                countryTabMode === "COMPARE" ||
                (isShipmentScope
                  ? shipmentLoading || shipmentFileEntries.filter((e) => e.file).length === 0
                  : inventoryLoading || inventoryFiles.length === 0)
              }
            >
              {isShipmentScope
                ? shipmentLoading
                  ? "출고 통합 중..."
                  : "출고 통합 실행"
                : inventoryLoading
                  ? "통합 중..."
                  : "재고 통합 실행"}
            </button>
            <button
              className="ghost"
              onClick={exportCurrentView}
              disabled={isInventoryAdminScope || (isShipmentScope && shipmentViewMode === "chart")}
              title={
                isShipmentScope && shipmentViewMode === "chart"
                  ? "상품별 출고 현황에서는 내보내기를 사용할 수 없습니다."
                  : undefined
              }
            >
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
            <input
              key={shipmentUploadInputKey}
              id="shipment-files-input"
              type="file"
              accept=".xlsx,.xls"
              multiple
              hidden
              onChange={(e) => {
                const run = async () => {
                  const files = Array.from(e.target.files || []);
                  if (!files.length) return;
                  const existingNames = new Set(shipmentFileEntries.map((entry) => String(entry.name || "")));
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
                    setShipmentUploadInputKey((k) => k + 1);
                    return;
                  }
                  setShipmentFileEntries((prev) => [
                    ...prev,
                    ...uploadableFiles.map((file, idx) => ({
                      id: `ship-${Date.now()}-${idx}-${file.name}`,
                      file,
                      name: file.name,
                      size: file.size,
                      country: "SHIPMENT",
                      date: "",
                    })),
                  ]);
                  setShipmentUploadInputKey((k) => k + 1);
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
        className={`topbar stickyTopbar dashboardStripWhite${!hasSecondaryInventoryStrip ? " isBottomCapsule" : ""}`}
      >
        <div className="tabs">
          <button
            type="button"
            className={`tab tabWithIcon ${countryTabMode === "KR" ? "active" : ""}`}
            onClick={() => setCountryTabMode("KR")}
          >
            <Home className="tabIcon" size={TOP_TAB_ICON_SIZE_PX} strokeWidth={2} aria-hidden />
            <span>한국 재고</span>
          </button>
          <button
            type="button"
            className={`tab tabWithIcon ${countryTabMode === "OVERSEAS" ? "active" : ""}`}
            onClick={() => {
              setCountryTabMode("OVERSEAS");
              setSelectedOverseasCountry((prev) => prev || OVERSEAS_UPLOAD_COUNTRIES[0]);
            }}
          >
            <Globe2 className="tabIcon" size={TOP_TAB_ICON_SIZE_PX} strokeWidth={2} aria-hidden />
            <span>해외 재고</span>
          </button>
          <button
            type="button"
            className={`tab tabWithIcon ${countryTabMode === "COMPARE" ? "active" : ""}`}
            onClick={() => setCountryTabMode("COMPARE")}
          >
            <GitCompare className="tabIcon" size={TOP_TAB_ICON_SIZE_PX} strokeWidth={2} aria-hidden />
            <span>재고 비교</span>
          </button>
          <button
            type="button"
            className={`tab tabWithIcon ${countryTabMode === "SHIPMENT" ? "active" : ""}`}
            onClick={() => setCountryTabMode("SHIPMENT")}
          >
            <Truck className="tabIcon" size={TOP_TAB_ICON_SIZE_PX} strokeWidth={2} aria-hidden />
            <span>출고 기록</span>
          </button>
          <button
            type="button"
            className={`tab tabWithIcon ${countryTabMode === "PURCHASE_ORDERS" ? "active" : ""}`}
            onClick={() => setCountryTabMode("PURCHASE_ORDERS")}
          >
            <ClipboardList className="tabIcon" size={TOP_TAB_ICON_SIZE_PX} strokeWidth={2} aria-hidden />
            <span>발주 기록</span>
          </button>
          <button
            type="button"
            className={`tab tabWithIcon ${countryTabMode === "PRODUCT_SEARCH" ? "active" : ""}`}
            onClick={() => setCountryTabMode("PRODUCT_SEARCH")}
          >
            <Link2 className="tabIcon tabIconProductMapping" size={TOP_TAB_ICON_SIZE_PX + 2} strokeWidth={2} aria-hidden />
            <span>상품 매핑</span>
          </button>
          <button
            type="button"
            className={`tab tabWithIcon ${countryTabMode === "SETTINGS" ? "active" : ""}`}
            onClick={() => setCountryTabMode("SETTINGS")}
          >
            <Database className="tabIcon tabIconDataManagement" size={TOP_TAB_ICON_SIZE_PX - 1} strokeWidth={2} aria-hidden />
            <span>데이터 관리</span>
          </button>
          <button
            type="button"
            className={`tab tabWithIcon ${countryTabMode === "SKU_MAPPING" ? "active" : ""}`}
            onClick={() => setCountryTabMode("SKU_MAPPING")}
          >
            <Barcode className="tabIcon" size={TOP_TAB_ICON_SIZE_PX} strokeWidth={2} aria-hidden />
            <span>SKU 관리</span>
          </button>
        </div>
      </header>

      {countryTabMode === "OVERSEAS" && (
        <div
          ref={countryChipsRef}
          className="countryChips stickyCountryChips dashboardStripWhite isBottomCapsule"
          style={{ top: stickyHeights.topbar }}
        >
          <div className="tabs overseasCountryTabs">
            {overseasCountries.map((code) => (
              <button
                key={code}
                type="button"
                className={`tab ${selectedOverseasCountry === code ? "active" : ""}`}
                onClick={() => setSelectedOverseasCountry(code)}
              >
                {countryLabel(code)}
              </button>
            ))}
          </div>
        </div>
      )}
      {countryTabMode === "SHIPMENT" && (
        <div
          ref={countryChipsRef}
          className="countryChips stickyCountryChips dashboardStripWhite isBottomCapsule"
          style={{ top: stickyHeights.topbar }}
        >
          <div className="tabs overseasCountryTabs" role="tablist" aria-label="출고 보기 방식">
            <button
              type="button"
              role="tab"
              aria-selected={shipmentViewMode === "status"}
              className={`tab ${shipmentViewMode === "status" ? "active" : ""}`}
              onClick={() => setShipmentViewMode("status")}
            >
              전체 출고 현황
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={shipmentViewMode === "chart"}
              className={`tab ${shipmentViewMode === "chart" ? "active" : ""}`}
              onClick={() => setShipmentViewMode("chart")}
            >
              상품별 출고 현황
            </button>
          </div>
        </div>
      )}
      {!isInventoryAdminScope && !isCompareScope && (
        <>
      {!(isShipmentScope && shipmentViewMode === "chart") && (
      <section className="kpiRow inventoryKpiRow">
        <div className="kpiCard">
          <div className="kpiCardLayout">
            <div className="kpiCardIcon" aria-hidden="true">
              {INVENTORY_KPI_CARD_ICON_SRC.uploadedFiles ? (
                <img
                  src={INVENTORY_KPI_CARD_ICON_SRC.uploadedFiles}
                  alt=""
                  width={KPI_CARD_ICON_DISPLAY_PX}
                  height={KPI_CARD_ICON_DISPLAY_PX}
                />
              ) : null}
            </div>
            <div className="kpiCardMain">
              <div className="kpiLabel">업로드된 파일</div>
              <div className="kpiValue kpiValueWithSuffix">
                <span className="kpiValueNum">
                  {formatInt(isShipmentScope ? shipmentFileEntries.length : inventoryFiles.length)}
                </span>
                <span className="kpiValueSuffix">개</span>
              </div>
            </div>
          </div>
        </div>
        <div className="kpiCard">
          <div className="kpiCardLayout">
            <div className="kpiCardIcon" aria-hidden="true">
              {INVENTORY_KPI_CARD_ICON_SRC.latestBaseDate ? (
                <img
                  src={INVENTORY_KPI_CARD_ICON_SRC.latestBaseDate}
                  alt=""
                  width={KPI_CARD_ICON_DISPLAY_PX}
                  height={KPI_CARD_ICON_DISPLAY_PX}
                />
              ) : null}
            </div>
            <div className="kpiCardMain">
              <div className="kpiLabel">최신 기준일</div>
              <div className="kpiValue kpiValuePlain">{latestScopeDateLabel}</div>
            </div>
          </div>
        </div>
        <div className="kpiCard">
          <div className="kpiCardLayout">
            <div className="kpiCardIcon" aria-hidden="true">
              {INVENTORY_KPI_CARD_ICON_SRC.inventoryBasis ? (
                <img
                  src={INVENTORY_KPI_CARD_ICON_SRC.inventoryBasis}
                  alt=""
                  width={KPI_CARD_ICON_DISPLAY_PX}
                  height={KPI_CARD_ICON_DISPLAY_PX}
                />
              ) : null}
            </div>
            <div className="kpiCardMain">
              <div className="kpiLabel">재고 파악 기준</div>
              <div className="kpiValue kpiValuePlain">{inventoryBasisLabel}</div>
            </div>
          </div>
        </div>
        <div className="kpiCard">
          <div className="kpiCardLayout">
            <div className="kpiCardIcon" aria-hidden="true">
              {INVENTORY_KPI_CARD_ICON_SRC.analyzedSkuCount ? (
                <img
                  src={INVENTORY_KPI_CARD_ICON_SRC.analyzedSkuCount}
                  alt=""
                  width={KPI_CARD_ICON_DISPLAY_PX}
                  height={KPI_CARD_ICON_DISPLAY_PX}
                />
              ) : null}
            </div>
            <div className="kpiCardMain">
              <div className="kpiLabel">분석 상품 수</div>
              <div className="kpiValue kpiValueWithSuffix">
                <span className="kpiValueNum">
                  {formatInt(
                    isShipmentScope
                      ? shipmentDisplayRows.filter((r) => !r.is_total).length
                      : filteredRows.length
                  )}
                </span>
                <span className="kpiValueSuffix">개</span>
              </div>
            </div>
          </div>
        </div>
      </section>
      )}

      <section
        className={`tableCard${isShipmentScope && shipmentViewMode === "chart" ? " tableCardShipmentChart" : ""}`}
      >
        {isShipmentScope && shipmentViewMode === "status" && (
          <div
            className="countryChips countryChipsShipment shipmentChannelChipsInTableCard shipmentChannelChipsCompact"
            role="toolbar"
            aria-label="출고 시트(채널)"
          >
            {SHIPMENT_MATRIX_CHIPS.map((ch) => {
              const hasData = (scopeResultCache.SHIPMENT?.shipment_channels_with_data || []).includes(ch);
              const active = shipmentMatrixChannel === ch;
              const showDivider = SHIPMENT_MATRIX_CHIP_DIVIDER_BEFORE.has(ch);
              return (
                <Fragment key={ch}>
                  {showDivider && <span className="shipmentChipSep" aria-hidden />}
                  <button
                    type="button"
                    className={`chip ${active ? "chipActive" : ""}`}
                    onClick={() => setShipmentMatrixChannel(ch)}
                    title={hasData ? `${ch} 데이터 있음` : `${ch} 시트 없음`}
                  >
                    {ch}
                  </button>
                </Fragment>
              );
            })}
          </div>
        )}
        {isShipmentScope && shipmentViewMode === "chart" ? (
          <div ref={inventoryFilterBarRef} className="shipmentChartFilterSpacer" aria-hidden />
        ) : (
        <div
          ref={inventoryFilterBarRef}
          className="filterBar stickyFilterBar"
          style={{ top: activeFilterStickyTop }}
        >
          <div className="searchWrap">
            <SearchFieldIcon className="searchIcon" size={16} strokeWidth={2} />
            <input
              className="searchInput"
              type="text"
              value={inventoryKeyword}
              onChange={(e) => setInventoryKeyword(e.target.value)}
              placeholder="상품코드 또는 상품명 검색..."
            />
          </div>
          {isShipmentScope && (
            <div className="shipmentStatusMonthToolbar" role="group" aria-label="출고 일자 표시 월">
              {shipmentDataYears.length > 1 && (
                <>
                  <label className="sr-only" htmlFor="shipment-month-strip-year">
                    연도
                  </label>
                  <select
                    id="shipment-month-strip-year"
                    className="inventoryFilterDate shipmentMonthStripYearSelect"
                    value={shipmentMonthStripYear}
                    onChange={(e) => {
                      const y = Number(e.target.value);
                      if (!Number.isFinite(y)) return;
                      setShipmentMonthStripYear(y);
                      if (shipmentDisplayMonth !== "__ALL__" && /^\d{4}-\d{2}$/.test(shipmentDisplayMonth)) {
                        const mm = shipmentDisplayMonth.slice(5, 7);
                        const next = `${y}-${mm}`;
                        if (shipmentAvailableMonths.includes(next)) setShipmentDisplayMonth(next);
                      }
                    }}
                    title="월 버튼에 적용할 연도"
                  >
                    {shipmentDataYears.map((y) => (
                      <option key={y} value={y}>
                        {y}년
                      </option>
                    ))}
                  </select>
                </>
              )}
              <div className="datePresetBox shipmentMonthStrip" role="group" aria-label="월 선택">
                {SHIPMENT_MONTH_CHIP_NUMS.map((monthNum) => {
                  const ym = `${shipmentMonthStripYear}-${String(monthNum).padStart(2, "0")}`;
                  const hasMonth = shipmentAvailableMonths.includes(ym);
                  const active = shipmentDisplayMonth === ym;
                  return (
                    <button
                      key={ym}
                      type="button"
                      className={`preset shipmentMonthStripBtn ${active ? "active" : ""}`}
                      disabled={!hasMonth}
                      onClick={() => setShipmentDisplayMonth(ym)}
                      title={
                        hasMonth
                          ? formatShipmentMonthLabelKorean(ym)
                          : `${shipmentMonthStripYear}년 ${monthNum}월 데이터 없음`
                      }
                    >
                      {monthNum}월
                    </button>
                  );
                })}
              </div>
              <div className="datePresetBox shipmentMonthAllRangeBox">
                <button
                  type="button"
                  className={`preset ${shipmentDisplayMonth === "__ALL__" ? "active" : ""}`}
                  aria-pressed={shipmentDisplayMonth === "__ALL__"}
                  onClick={() => setShipmentDisplayMonth("__ALL__")}
                  title="모든 일자 열 표시"
                >
                  전체 기간
                </button>
              </div>
            </div>
          )}
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
            {!isShipmentScope && (
              <>
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
                      className="inventoryFilterDate"
                      aria-label="기간 시작일"
                      value={inventoryStartDate}
                      onChange={(e) => setInventoryStartDate(e.target.value)}
                    />
                    <input
                      type="date"
                      className="inventoryFilterDate"
                      aria-label="기간 종료일"
                      value={inventoryEndDate}
                      onChange={(e) => setInventoryEndDate(e.target.value)}
                    />
                  </>
                )}
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
              if (isShipmentScope) {
                setShipmentDisplayMonth("");
                setShipmentViewMode("status");
                setShipmentChartMonth("");
                setShipmentChartSearchKeyword("");
                setShipmentChartSelectedSku("");
              }
            }}
          >
            필터 초기화
          </button>
        </div>
        )}

        {inventoryError && <pre className="error">{inventoryError}</pre>}

        {hasTableData && (
          <>
            {isShipmentScope && shipmentViewMode === "chart" ? (
              <div className="shipmentChartMappingCard shipmentChartPage">
                <div className="shipmentChartStickyBand" style={{ top: activeFilterStickyTop }}>
                <header className="shipmentChartToolbar" aria-label="출고 상품 검색">
                  <div className="shipmentChartToolbarLead">
                    <div className="shipmentChartToolbarSearchAndProduct">
                      <div className="shipmentChartToolbarSearchStack">
                        <div className="shipmentChartToolbarSearchRow">
                          <div className="shipmentChartToolbarSearch">
                            <div className="searchWrap productMappingSearchWrap shipmentChartToolbarSearchWrap">
                              <SearchFieldIcon className="searchIcon shipmentChartSearchIcon" size={16} strokeWidth={2} />
                              <input
                                className="searchInput productMappingSearchInput"
                                type="search"
                                enterKeyHint="search"
                                value={shipmentChartSearchKeyword}
                                onChange={(e) => setShipmentChartSearchKeyword(e.target.value)}
                                placeholder="상품코드 또는 상품명 검색"
                              />
                            </div>
                          </div>
                        </div>
                        {shipmentChartSearchKeyword.trim() &&
                        shipmentChartSkuOptions.length > 1 &&
                        !shipmentChartSelectedSku ? (
                          <div className="shipmentChartToolbarResults shipmentChartToolbarResultsScroll">
                            <ul className="shipmentChartToolbarPickList" role="listbox" aria-label="상품 선택">
                              {shipmentChartSkuOptions.map((sku) => (
                                <li key={sku} className="shipmentChartToolbarPickLi">
                                  <button
                                    type="button"
                                    className="shipmentChartToolbarPickItem"
                                    role="option"
                                    onClick={() => setShipmentChartSelectedSku(sku)}
                                  >
                                    <span className="shipmentChartToolbarPickCode">{sku}</span>
                                    <span className="shipmentChartToolbarPickName">
                                      {shipmentChartSkuOptionLabels.get(sku) || "—"}
                                    </span>
                                  </button>
                                </li>
                              ))}
                            </ul>
                          </div>
                        ) : null}
                      </div>
                      {shipmentChartSelectedSku ? (
                        <div className="shipmentChartToolbarIdentity" aria-label="선택된 상품">
                          <span className="shipmentChartSkuPlain">{shipmentChartSelectedSku}</span>
                          {shipmentChartTitleLabel ? (
                            <>
                              <span className="shipmentChartSkuSep" aria-hidden>
                                |
                              </span>
                              <span className="shipmentChartProductName">{shipmentChartTitleLabel}</span>
                            </>
                          ) : null}
                        </div>
                      ) : null}
                    </div>
                  </div>
                </header>
                {shipmentChartSelectedSku && shipmentChartMonthOptions.length > 0 ? (
                  <div className="shipmentChartMainHead">
                    <div
                      className="countryChips countryChipsShipment shipmentChannelChipsCompact shipmentChartMonthChips"
                      role="group"
                      aria-label="일자·시트 표·차트에 적용할 달"
                    >
                      {SHIPMENT_MONTH_CHIP_NUMS.map((monthNum) => {
                        const ym = `${shipmentChartChipYear}-${String(monthNum).padStart(2, "0")}`;
                        const enabled = shipmentChartMonthOptions.includes(ym);
                        const active = shipmentChartMonth === ym;
                        return (
                          <button
                            key={ym}
                            type="button"
                            disabled={!enabled}
                            className={`chip ${active ? "chipActive" : ""}`}
                            onClick={() => {
                              if (enabled) setShipmentChartMonth(ym);
                            }}
                            title={
                              enabled
                                ? `${formatShipmentMonthLabelKorean(ym)} 일자·시트별 상세`
                                : "해당 월 출고 데이터 없음"
                            }
                          >
                            {monthNum}월
                          </button>
                        );
                      })}
                    </div>
                  </div>
                ) : null}
                </div>
                <div className="shipmentChartMain">
                  {!shipmentChartSearchKeyword.trim() ? (
                    <div className="shipmentChartMainPlaceholder">
                      상단에서 상품코드 또는 상품명을 검색하면 해당 상품의 출고 추이를 볼 수 있습니다.
                    </div>
                  ) : !shipmentChartSkuOptions.length ? (
                    <div className="productMappingNoResult shipmentChartMainMessage">
                      일치하는 출고 상품이 없습니다. 검색어를 바꿔 보세요.
                    </div>
                  ) : !shipmentChartSelectedSku ? (
                    <div className="shipmentChartMainPlaceholder">
                      위에서 상품을 선택하면 데이터가 표시됩니다.
                    </div>
                  ) : (
                    <>
                      {shipmentChartSkuInsight ? (
                        <div
                          className="shipmentChartInsightPanel shipmentChartInsightPanelProductFirst"
                          aria-label="선택 상품 출고 상세"
                        >
                          <div className="shipmentChartInsightSection shipmentChartDataLead">
                            {shipmentChartChannelDailyMatrix &&
                            shipmentChartChannelDailyMatrix.rows.length > 0 ? (
                              <div className="shipmentChartDayMatrixScrollHost">
                                {showShipmentChartDayMatrixTopScroll ? (
                                  <div
                                    ref={shipmentChartDayMatrixTopScrollRef}
                                    className="tableTopScroll shipmentMatrixTopScroll shipmentChartDayMatrixTopScroll"
                                    onScroll={() => syncShipmentChartDayMatrixScroll("top")}
                                  >
                                    <div
                                      style={{
                                        width: shipmentChartDayMatrixTopScrollWidth || undefined,
                                      }}
                                    />
                                  </div>
                                ) : null}
                                <div
                                  ref={shipmentChartDayMatrixScrollRef}
                                  className="shipmentChannelDayMatrixWrap shipmentMatrixBodyWrap shipmentChartDayMatrixScrollSync"
                                  onScroll={() => syncShipmentChartDayMatrixScroll("table")}
                                >
                                <table className="shipmentChannelDayMatrix">
                                  <colgroup>
                                    <col className="shipmentChannelDayMatrixColRowLabel" />
                                    <col className="shipmentChannelDayMatrixColTotal" />
                                    {shipmentChartChannelDailyMatrix.dayKeys.map((dk) => (
                                      <col key={dk} className="shipmentChannelDayMatrixColDay" />
                                    ))}
                                  </colgroup>
                                  <thead>
                                    <tr>
                                      <th
                                        scope="col"
                                        className="shipmentChannelDayMatrixRowLabel"
                                        aria-label="시트"
                                      >
                                        <span className="shipmentChannelDayMatrixCornerPh" aria-hidden="true" />
                                      </th>
                                      <th
                                        scope="col"
                                        className="shipmentChannelDayMatrixTotalCol"
                                        aria-label="판매처별 합계"
                                      >
                                        판매처별
                                        <br />
                                        합계
                                      </th>
                                      {shipmentChartChannelDailyMatrix.dayKeys.map((dk) => (
                                        <th
                                          key={dk}
                                          scope="col"
                                          className="shipmentChannelDayMatrixDayCol"
                                        >
                                          {formatShipmentDayHeaderLabel(dk)}
                                        </th>
                                      ))}
                                    </tr>
                                  </thead>
                                  <tbody>
                                    <tr className="shipmentChannelDayMatrixDailyTotalRow">
                                      <th scope="row" className="shipmentChannelDayMatrixRowLabel">
                                        일별 합계
                                      </th>
                                      <td className="shipmentChannelDayMatrixTotalCol">–</td>
                                      {shipmentChartChannelDailyMatrix.colTotals.map((q, i) => (
                                        <td
                                          key={`coltot-${shipmentChartChannelDailyMatrix.dayKeys[i]}`}
                                          className="shipmentChannelDayMatrixDayCol"
                                        >
                                          {q ? formatInt(q) : "–"}
                                        </td>
                                      ))}
                                    </tr>
                                    {shipmentChartChannelDailyMatrix.rows.map((r) => (
                                      <tr key={r.channel}>
                                        <th scope="row" className="shipmentChannelDayMatrixRowLabel">
                                          {r.channel}
                                        </th>
                                        <td className="shipmentChannelDayMatrixTotalCol">
                                          {r.rowSum ? formatInt(r.rowSum) : "–"}
                                        </td>
                                        {r.cells.map((q, i) => (
                                          <td
                                            key={shipmentChartChannelDailyMatrix.dayKeys[i]}
                                            className="shipmentChannelDayMatrixDayCol"
                                          >
                                            {q ? formatInt(q) : "–"}
                                          </td>
                                        ))}
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                                </div>
                              </div>
                            ) : (
                              <p className="shipmentChartEmpty">시트·일자 표를 만들 데이터가 없습니다.</p>
                            )}
                          </div>
                          {shipmentChartSkuInsight.vendorRows.length > 0 ? (
                            <div className="shipmentChartInsightSection">
                              <h3 className="shipmentChartInsightSectionTitle">판매처·창고 분해</h3>
                              <div className="shipmentChartInsightTableWrap">
                                <table className="shipmentChartInsightTable">
                                  <thead>
                                    <tr>
                                      <th>판매처</th>
                                      <th className="shipmentChartInsightNumCol">판매</th>
                                      <th className="shipmentChartInsightNumCol">창고이동</th>
                                      <th className="shipmentChartInsightNumCol">합계</th>
                                    </tr>
                                  </thead>
                                  <tbody>
                                    {shipmentChartSkuInsight.vendorRows.map((v) => (
                                      <tr key={v.vendor}>
                                        <td>{v.vendor}</td>
                                        <td className="shipmentChartInsightNumCol">{formatInt(v.sale)}</td>
                                        <td className="shipmentChartInsightNumCol">{formatInt(v.wh)}</td>
                                        <td className="shipmentChartInsightNumCol">{formatInt(v.sum)}</td>
                                      </tr>
                                    ))}
                                  </tbody>
                                </table>
                              </div>
                            </div>
                          ) : null}
                          <div className="shipmentChartChartsBlock">
                            <div className="shipmentChartChartsRow">
                              <section className="shipmentChartBlock shipmentChartBlockCompact" aria-label="일자별 출고 추이">
                                <span className="shipmentChartChartTypeChip">일자별</span>
                                <div className="trendChartCard shipmentLineChartCard shipmentChartTrendSurface shipmentChartSvgCompact">
                                  {shipmentChartDailyLineData ? (
                                    <ShipmentQtyLineChartSvg
                                      lineData={shipmentChartDailyLineData}
                                      dailyVendorTooltips
                                      ariaLabel="선택 월 일자별 출고 라인 차트"
                                    />
                                  ) : (
                                    <p className="shipmentChartEmpty shipmentChartEmptyInCard">
                                      이 달·상품에 표시할 일자 데이터가 없습니다.
                                    </p>
                                  )}
                                </div>
                              </section>
                              <section className="shipmentChartBlock shipmentChartBlockCompact" aria-label="월별 출고 추이">
                                <span className="shipmentChartChartTypeChip">월별</span>
                                <div className="trendChartCard shipmentLineChartCard shipmentChartTrendSurface shipmentChartSvgCompact">
                                  {shipmentChartMonthlyLineData ? (
                                    <ShipmentQtyLineChartSvg
                                      lineData={shipmentChartMonthlyLineData}
                                      dailyVendorTooltips={false}
                                      ariaLabel="월별 출고 라인 차트"
                                    />
                                  ) : (
                                    <p className="shipmentChartEmpty shipmentChartEmptyInCard">
                                      월별 집계를 표시할 데이터가 없습니다.
                                    </p>
                                  )}
                                </div>
                              </section>
                            </div>
                          </div>
                        </div>
                      ) : null}
                    </>
                  )}
                </div>
              </div>
            ) : (
              <>
            {showTopScroll && (
              <div
                ref={topScrollRef}
                className={`tableTopScroll stickyTableTopScroll${
                  isShipmentScope && shipmentViewMode === "status" ? " shipmentMatrixTopScroll" : ""
                }`}
                style={{ top: activeFilterStickyTop + activeFilterHeight }}
                onScroll={() => syncScroll("top")}
              >
                <div style={{ width: topScrollWidth || tableMinWidth }} />
              </div>
            )}
            <div
              className={`stickyTableHeader${
                isShipmentScope && shipmentViewMode === "status" ? " shipmentMatrixStickyHead" : ""
              }`}
              style={{ top: tableHeaderTop }}
            >
              <div
                ref={headerScrollRef}
                className={`tableHeaderScroll ${datePeekFadeStyle ? "withDatePeekFade" : ""}`}
                style={datePeekFadeStyle}
                onScroll={() => syncScroll("header")}
              >
                <table
                  className={`inventoryTable stickyHeaderTable ${
                    isKRScope || isShipmentScope ? "krTable" : "overseasTable"
                  } ${isShipmentScope ? "shipmentMatrixTable" : ""}`}
                  style={{ minWidth: inventoryTableWidth }}
                >
                  {shipmentMatrixColGroup}
                  <thead>
                    <tr>{renderInventoryHeaderCells()}</tr>
                  </thead>
                </table>
              </div>
            </div>
            <div
              ref={tableScrollRef}
              className={`tableWrap ${datePeekFadeStyle ? "withDatePeekFade" : ""}${
                isShipmentScope && shipmentViewMode === "status" ? " shipmentMatrixBodyWrap" : ""
              }`}
              style={datePeekFadeStyle}
              onScroll={() => syncScroll("table")}
            >
            <table
              className={`inventoryTable bodyTable ${
                isKRScope || isShipmentScope ? "krTable" : "overseasTable"
              } ${isShipmentScope ? "shipmentMatrixTable" : ""}`}
              style={{ minWidth: inventoryTableWidth }}
            >
              {shipmentMatrixColGroup}
              <tbody>
                {(isShipmentScope ? shipmentDisplayRows : isKRScope
                    ? krDisplayRows
                    : isOverseasScope
                      ? overseasDisplayRows
                      : filteredRows
                ).map((row, idx, tableBodyRows) => {
                  const rowKey =
                    row.trendRowKey ||
                    `${row.country}-${getRowSku(row)}-${row.description}-${row.level}-${row.warehouse}-${idx}`;
                  const shipmentTotalGapBefore =
                    isShipmentScope && row.is_total && shipmentBodyColCount > 0 && idx > 0 ? (
                      <>
                        <tr className="shipmentTotalGapRow" aria-hidden="true">
                          <td colSpan={shipmentBodyColCount} />
                        </tr>
                        <tr className="shipmentTotalGapRow" aria-hidden="true">
                          <td colSpan={shipmentBodyColCount} />
                        </tr>
                      </>
                    ) : null;
                  const shipmentSubHeaderAfterTotal =
                    isShipmentScope && row.is_total && idx < tableBodyRows.length - 1 ? (
                      <tr key={`${rowKey}-subcolumns`} className="shipmentMatrixSubHeaderRow">
                        {renderShipmentStatusHeaderCells(true)}
                      </tr>
                    ) : null;
                  return (
                  <Fragment key={rowKey}>
                    {shipmentTotalGapBefore}
                  <tr className={isShipmentScope && row.is_total ? "shipmentTotalRow" : undefined}>
                    {isShipmentScope ? (
                      <>
                        {row.is_total ? (
                          <td colSpan={5} className="stickyCol shipmentTotalMergedCell">
                            합 계
                          </td>
                        ) : (
                          <>
                            <td className="stickyCol stickyColShipCode">{getRowSku(row) || "–"}</td>
                            <td
                              className="stickyCol stickyColShipBrand"
                              title={
                                String(row.mapped_barcode || "").trim()
                                  ? `바코드: ${String(row.mapped_barcode).trim()}`
                                  : "등록된 바코드가 없습니다"
                              }
                            >
                              {String(row.brand || "").trim() || "–"}
                            </td>
                            <td className="stickyCol stickyColShipName">
                              <span className="nameCellText">{row.description || "(상품명 없음)"}</span>
                            </td>
                            <td className="stickyCol stickyColShipMkt">
                              {String(row.mkt_priority || "").trim() || "–"}
                            </td>
                            <td className="stickyCol stickyColShipSegment stickyColBoundary">
                              {String(row.segment || "").trim() || "–"}
                            </td>
                          </>
                        )}
                        {(scopeResultCache.SHIPMENT?.shipment_month_columns || []).map((col) => {
                          const mt = row.month_totals || {};
                          const v = mt[col.key];
                          return (
                            <td
                              key={`m-${col.key}`}
                              className="dateCol"
                              title={shipmentMatrixQtyCellTitle(row, "month", {
                                monthKey: col.key,
                                monthColLabel: col.label,
                              })}
                            >
                              {v != null && Number(v) !== 0 ? formatInt(v) : "–"}
                            </td>
                          );
                        })}
                        {filteredDateColumns.map((dt) => (
                          <td
                            key={dt}
                            className="dateCol"
                            title={shipmentMatrixQtyCellTitle(row, "day", {
                              dateKey: dt,
                              dayHeadLabel:
                                scopeResultCache.SHIPMENT?.shipment_day_labels?.[dt] || dt,
                            })}
                          >
                            {formatShipmentWideDateCell(row[dt])}
                          </td>
                        ))}
                      </>
                    ) : (
                      <>
                    {isKRScope && (
                      <td
                        className="stickyCol stickyColCode"
                        title={
                          String(row.mapped_barcode || "").trim()
                            ? `바코드: ${String(row.mapped_barcode).trim()}`
                            : "등록된 바코드가 없습니다"
                        }
                      >
                        {getRowSku(row)}
                      </td>
                    )}
                    {isKRScope && <td className="stickyCol stickyColBrand">{row.supplier}</td>}
                    {!isKRScope && (
                      <td
                        className="stickyCol stickyColCode"
                        title={
                          String(row.mapped_barcode || "").trim()
                            ? `바코드: ${String(row.mapped_barcode).trim()}`
                            : "등록된 바코드가 없습니다"
                        }
                      >
                        {getRowSku(row)}
                      </td>
                    )}
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
                          ? formatInt(krCompareMap.get(getCanonicalMatchCode(row)) || 0)
                          : "-"}
                      </td>
                    )}
                    {filteredDateColumns.map((dt) => (
                      <td
                        key={`${getRowSku(row)}-${row.description}-${row.level}-${row.warehouse}-${dt}`}
                        className="dateCol"
                        title={
                          row.classification_by_date?.[dt]
                            ? `유통/분류: ${row.classification_by_date[dt]}`
                            : undefined
                        }
                      >
                        {formatInt(row[dt])}
                      </td>
                    ))}
                      </>
                    )}
                  </tr>
                    {shipmentSubHeaderAfterTotal}
                  </Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
              </>
            )}
          </>
        )}

        {inventoryRequested &&
          !inventoryLoading &&
          !shipmentLoading &&
          !inventoryError &&
          !hasTableData && (
            <pre className="error">
              조건에 맞는 데이터가 없습니다.
              {!isKRScope && !isShipmentScope && filteredRows.length > 0 && filteredDateColumns.length === 0
                ? "\n- 선택한 기간에 유효한 데이터가 없습니다."
                : ""}
              {"\n"}- 날짜/레벨/검색 필터를 초기화해보세요.
              {isShipmentScope
                ? "\n- 출고 탭에서는 엑셀 업로드 후 「출고 통합 실행」을 해 주세요."
                : "\n- 먼저 파일 업로드 후 재고 통합 실행을 1회 해주세요."}
            </pre>
          )}
      </section>
      </>
      )}

      {isCompareScope && (
        <>
          <section className="kpiRow compareKpiRow">
            <div className="kpiCard">
              <div className="kpiCardLayout">
                <div className="kpiCardIcon" aria-hidden="true">
                  {COMPARE_KPI_CARD_ICON_SRC.compareCountries ? (
                    <img
                      src={COMPARE_KPI_CARD_ICON_SRC.compareCountries}
                      alt=""
                      width={KPI_CARD_ICON_DISPLAY_PX}
                      height={KPI_CARD_ICON_DISPLAY_PX}
                    />
                  ) : null}
                </div>
                <div className="kpiCardMain">
                  <div className="kpiLabel">비교 가능 국가</div>
                  <div className="kpiValue kpiValueWithSuffix">
                    <span className="kpiValueNum">{formatInt(OVERSEAS_UPLOAD_COUNTRIES.length)}</span>
                    <span className="kpiValueSuffix">개</span>
                  </div>
                </div>
              </div>
            </div>
            <div className="kpiCard">
              <div className="kpiCardLayout">
                <div className="kpiCardIcon" aria-hidden="true">
                  {COMPARE_KPI_CARD_ICON_SRC.latestBaseDate ? (
                    <img
                      src={COMPARE_KPI_CARD_ICON_SRC.latestBaseDate}
                      alt=""
                      width={KPI_CARD_ICON_DISPLAY_PX}
                      height={KPI_CARD_ICON_DISPLAY_PX}
                    />
                  ) : null}
                </div>
                <div className="kpiCardMain">
                  <div className="kpiLabel">최신 기준일</div>
                  <div className="kpiValue kpiValuePlain">{compareLatestDateLabel}</div>
                </div>
              </div>
            </div>
            <div className="kpiCard">
              <div className="kpiCardLayout">
                <div className="kpiCardIcon" aria-hidden="true">
                  {COMPARE_KPI_CARD_ICON_SRC.compareSkuCount ? (
                    <img
                      src={COMPARE_KPI_CARD_ICON_SRC.compareSkuCount}
                      alt=""
                      width={KPI_CARD_ICON_DISPLAY_PX}
                      height={KPI_CARD_ICON_DISPLAY_PX}
                    />
                  ) : null}
                </div>
                <div className="kpiCardMain">
                  <div className="kpiLabel">비교 상품 수</div>
                  <div className="kpiValue kpiValueWithSuffix">
                    <span className="kpiValueNum">{formatInt(filteredCompareRows.length)}</span>
                    <span className="kpiValueSuffix">개</span>
                  </div>
                </div>
              </div>
            </div>
          </section>

          <section className="tableCard">
            <div
              ref={compareFilterBarRef}
              className="filterBar stickyFilterBar"
              style={{ top: activeFilterStickyTop }}
            >
              <div className="searchWrap">
                <SearchFieldIcon className="searchIcon" size={16} strokeWidth={2} />
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
                className="inventoryFilterDate"
                aria-label="재고 비교 기준일"
                min={compareDatePickerExtent.min}
                max={compareDatePickerExtent.max}
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
                          {row["한국 현 재고"] === null ? "-" : formatInt(row["한국 현 재고"])}
                        </td>
                        {OVERSEAS_UPLOAD_COUNTRIES.map((code) => (
                          <td key={`${row._compareKey}-${code}`} className="compareCountryCol">
                            {row[countryLabel(code)] === null
                              ? "-"
                              : formatInt(row[countryLabel(code)])}
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
                <div className="trendSummaryValue">{formatInt(activeTrendData.latestQty)}</div>
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
                    <title>{`${point.dateKey} | 재고 ${formatInt(point.qty)}`}</title>
                    <circle cx={point.x} cy={point.y} r="5" className="trendDot" />
                    {activeTrendData.showPointValueLabels && (
                      <text x={point.x} y={point.y - 12} textAnchor="middle" className="trendDotLabel">
                        {formatInt(point.qty)}
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
                        <td>{formatInt(point.qty)}</td>
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

      {isPurchaseOrderScope && (
        <section className="settingsPane settingsCard purchaseOrderSection">
          {purchaseOrderError ? <pre className="error purchaseOrderError">{purchaseOrderError}</pre> : null}
          {purchaseOrderSuccess ? <div className="purchaseOrderSuccess">{purchaseOrderSuccess}</div> : null}

          <div
            ref={poSubTabsBarRef}
            className={`poSubTabsBar${isPurchaseOrderScope ? " poOrderStickySubTabs" : ""}`}
            role="tablist"
            aria-label="발주 하위 메뉴"
            style={isPurchaseOrderScope ? { top: poStickySubTabsTop } : undefined}
          >
            <button
              type="button"
              role="tab"
              aria-selected={purchaseOrderSubTab === "orderFileUpload"}
              className={`tab poSubTab ${purchaseOrderSubTab === "orderFileUpload" ? "active" : ""}`}
              onClick={() => setPurchaseOrderSubTab("orderFileUpload")}
            >
              발주 파일 업로드
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={purchaseOrderSubTab === "register"}
              className={`tab poSubTab ${purchaseOrderSubTab === "register" ? "active" : ""}`}
              onClick={() => setPurchaseOrderSubTab("register")}
            >
              새 발주 등록
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={purchaseOrderSubTab === "saved"}
              className={`tab poSubTab ${purchaseOrderSubTab === "saved" ? "active" : ""}`}
              onClick={() => {
                setPurchaseOrderSubTab("saved");
                void reloadPurchaseOrdersUi();
              }}
            >
              저장된 발주
            </button>
          </div>

          {purchaseOrderSubTab === "register" ? (
          <div className="poFormCard poRegisterPanel">
            <div className="skuManualSection skuManualSectionKr poFormManualSection">
              <div className="skuManualBlock">
                <div className="skuManualKrGrid">
                  <label className="skuManualKrField">
                    <span className="skuManualKrFieldHead">발주일자</span>
                    <div className="skuManualKrFieldBody poExpectedInboundBody poRegisterDateTbdRow">
                      <input
                        type="text"
                        placeholder={DATE_TEXT_INPUT_HINT}
                        disabled={isOrderDatePlannedNote(poForm.order_date_note)}
                        value={isOrderDatePlannedNote(poForm.order_date_note) ? "" : poForm.order_date}
                        onChange={(e) =>
                          setPoForm((p) => ({
                            ...p,
                            order_date: e.target.value,
                            order_date_note: e.target.value ? "" : p.order_date_note,
                          }))
                        }
                      />
                      <label className="poExpectedInboundTbd">
                        <input
                          type="checkbox"
                          checked={isOrderDatePlannedNote(poForm.order_date_note)}
                          onChange={(e) => {
                            const planned = e.target.checked;
                            setPoForm((p) => ({
                              ...p,
                              order_date: planned ? "" : p.order_date,
                              order_date_note: planned ? "발주 예정" : "",
                            }));
                          }}
                        />
                        발주 예정
                      </label>
                    </div>
                  </label>
                  <label className="skuManualKrField">
                    <span className="skuManualKrFieldHead">ERP PO 번호</span>
                    <div className="skuManualKrFieldBody">
                      <input
                        type="text"
                        value={poForm.erp_po_number}
                        onChange={(e) => setPoForm((p) => ({ ...p, erp_po_number: e.target.value }))}
                        placeholder="예: PO2511000012"
                        autoComplete="off"
                      />
                    </div>
                  </label>
                  <label className="skuManualKrField">
                    <span className="skuManualKrFieldHead">상품유형</span>
                    <div className="skuManualKrFieldBody">
                      <PurchaseProductTypeField
                        preset={poForm.product_type_preset}
                        custom={poForm.product_type_custom}
                        onPresetChange={(v) => setPoForm((p) => ({ ...p, product_type_preset: v }))}
                        onCustomChange={(v) => setPoForm((p) => ({ ...p, product_type_custom: v }))}
                      />
                    </div>
                  </label>
                  <label className="skuManualKrField">
                    <span className="skuManualKrFieldHead">상품코드(SKU)</span>
                    <div className="skuManualKrFieldBody poKrFieldBodyStack">
                      <input
                        type="text"
                        value={poForm.sku}
                        onChange={(e) => {
                          setPoForm((p) => ({ ...p, sku: e.target.value }));
                          setSkuResolveHint("");
                        }}
                        onBlur={() => resolvePurchaseOrderSku()}
                        placeholder="입력하면 자동으로 브랜드와 상품명이 입력됩니다"
                        autoComplete="off"
                      />
                      {skuResolveHint ? <small className="poSkuHint">{skuResolveHint}</small> : null}
                    </div>
                  </label>
                  <label className="skuManualKrField">
                    <span className="skuManualKrFieldHead">브랜드</span>
                    <div className="skuManualKrFieldBody">
                      <input
                        type="text"
                        value={poForm.brand}
                        onChange={(e) => setPoForm((p) => ({ ...p, brand: e.target.value }))}
                        autoComplete="off"
                      />
                    </div>
                  </label>
                  <label className="skuManualKrField">
                    <span className="skuManualKrFieldHead">상품명</span>
                    <div className="skuManualKrFieldBody">
                      <input
                        type="text"
                        value={poForm.product_name}
                        onChange={(e) => setPoForm((p) => ({ ...p, product_name: e.target.value }))}
                        autoComplete="off"
                      />
                    </div>
                  </label>
                  <label className="skuManualKrField">
                    <span className="skuManualKrFieldHead">제조사</span>
                    <div className="skuManualKrFieldBody">
                      <input
                        type="text"
                        value={poForm.manufacturer}
                        onChange={(e) => setPoForm((p) => ({ ...p, manufacturer: e.target.value }))}
                        autoComplete="off"
                      />
                    </div>
                  </label>
                  <label className="skuManualKrField">
                    <span className="skuManualKrFieldHead">총 발주수량</span>
                    <div className="skuManualKrFieldBody">
                      <input
                        type="number"
                        min={0}
                        step="any"
                        value={poForm.total_quantity}
                        onChange={(e) => setPoForm((p) => ({ ...p, total_quantity: e.target.value }))}
                      />
                    </div>
                  </label>
                  <label className="skuManualKrField">
                    <span className="skuManualKrFieldHead">납품가능일</span>
                    <div className="skuManualKrFieldBody poExpectedInboundBody poRegisterDateTbdRow">
                      <input
                        type="text"
                        placeholder={DATE_TEXT_INPUT_HINT}
                        disabled={poForm.delivery_available_tbd}
                        value={poForm.delivery_available_tbd ? "" : poForm.delivery_available_date}
                        onChange={(e) =>
                          setPoForm((p) => ({
                            ...p,
                            delivery_available_date: e.target.value,
                            delivery_available_tbd: false,
                          }))
                        }
                      />
                      <label className="poExpectedInboundTbd">
                        <input
                          type="checkbox"
                          checked={poForm.delivery_available_tbd}
                          onChange={(e) => {
                            const tbd = e.target.checked;
                            setPoForm((p) => ({
                              ...p,
                              delivery_available_tbd: tbd,
                              delivery_available_date: tbd ? "" : p.delivery_available_date,
                            }));
                          }}
                        />
                        미정
                      </label>
                    </div>
                  </label>
                  <div className="skuManualKrField">
                    <span className="skuManualKrFieldHead">입고예정일</span>
                    <div className="skuManualKrFieldBody poExpectedInboundBody poRegisterDateTbdRow">
                      <input
                        type="text"
                        placeholder={DATE_TEXT_INPUT_HINT}
                        disabled={poForm.expected_inbound_tbd}
                        value={poForm.expected_inbound_tbd ? "" : poForm.expected_inbound_date}
                        onChange={(e) =>
                          setPoForm((p) => ({
                            ...p,
                            expected_inbound_date: e.target.value,
                            expected_inbound_tbd: false,
                          }))
                        }
                      />
                      <label className="poExpectedInboundTbd">
                        <input
                          type="checkbox"
                          checked={poForm.expected_inbound_tbd}
                          onChange={(e) => {
                            const tbd = e.target.checked;
                            setPoForm((p) => ({
                              ...p,
                              expected_inbound_tbd: tbd,
                              expected_inbound_date: tbd ? "" : p.expected_inbound_date,
                            }));
                          }}
                        />
                        미정
                      </label>
                    </div>
                  </div>
                </div>
              </div>
            </div>
            <div className="poRegisterFooterActions">
              <button type="button" className="primary" onClick={() => submitPurchaseOrder()}>
                발주 저장
              </button>
            </div>
          </div>
          ) : null}

          {purchaseOrderSubTab === "orderFileUpload" ? (
            <div className="poOrderFileUploadCard">
              <input
                key={poInboundFileKey}
                id="po-order-file-input"
                type="file"
                accept=".xlsx,.xls"
                style={{ display: "none" }}
                disabled={poInboundUploadBusy}
                onChange={(e) => void handlePoInboundFileSelected(e)}
              />
              <div className="skuUploadStage skuManageSinglePanel poOrderFileUploadStage">
                <div className="settingsNotice poOrderFileUploadNotice">
                  <p className="poOrderFileUploadLead">
                    <strong>
                      엑셀 템플릿을 받아 작성한 뒤 업로드하면 「저장된 발주」에 그대로 저장됩니다.
                    </strong>
                  </p>
                  <p className="poOrderFileUploadSub">
                    상품명과 브랜드는 상품번호를 통해 자동 입력되므로 상품코드만 입력하면 됩니다.
                  </p>
                </div>
                <div className="poOrderFileTemplateRow">
                  <button
                    type="button"
                    className="poOrderFileTemplateBtn"
                    disabled={poInboundUploadBusy}
                    onClick={() => void downloadPoInboundTemplateClick()}
                  >
                    발주 양식 액셀 템플릿 다운로드
                  </button>
                </div>
                <button
                  type="button"
                  className="skuUploadPanel"
                  disabled={poInboundUploadBusy}
                  onClick={() => openPoOrderFileInput()}
                >
                  <span className="skuUploadMain">
                    <span className="skuUploadBadge">XLSX</span>
                    <span className="skuUploadButtonLabel">
                      {poInboundUploadBusy ? "업로드 중..." : "발주 파일 업로드"}
                    </span>
                  </span>
                </button>
              </div>
            </div>
          ) : null}

          {purchaseOrderSubTab === "saved" ? (
          <div className="poFormCard poSavedOrdersPanel">
            <div
              ref={poSavedFilterBarRef}
              className="filterBar poSavedFilterBar poOrderStickySavedFilter"
              style={{ top: poStickySavedFilterTop }}
            >
              <div className="searchWrap">
                <SearchFieldIcon className="searchIcon" size={16} strokeWidth={2} />
                <input
                  className="searchInput poSavedFilterSearch"
                  type="search"
                  placeholder="상품코드 또는 상품명 검색..."
                  value={savedPoSearch}
                  onChange={(e) => setSavedPoSearch(e.target.value)}
                  aria-label="저장된 발주 검색 (상품코드 또는 상품명)"
                  autoComplete="off"
                />
              </div>
              <select
                className="poSavedSortSelect"
                value={savedPoSortMode}
                onChange={(e) => setSavedPoSortMode(e.target.value)}
                aria-label="저장된 발주 정렬"
              >
                <option value="order_date_asc">발주일 · 오래된순</option>
                <option value="order_date_desc">발주일 · 최신순</option>
                <option value="delivery_asc">납품가능일 · 오래된순</option>
                <option value="delivery_desc">납품가능일 · 최신순</option>
                <option value="expected_asc">입고예정일 · 임박순 (미입고)</option>
                <option value="expected_desc">입고예정일 · 나중순 (미입고)</option>
              </select>
              <div className="datePresetBox" role="group" aria-label="발주일 기준 기간">
                <button
                  type="button"
                  className={savedPoDateRange === "all" ? "preset active" : "preset"}
                  onClick={() => setSavedPoDateRange("all")}
                >
                  전체
                </button>
                <button
                  type="button"
                  className={savedPoDateRange === "1m" ? "preset active" : "preset"}
                  onClick={() => setSavedPoDateRange("1m")}
                >
                  1개월
                </button>
                <button
                  type="button"
                  className={savedPoDateRange === "3m" ? "preset active" : "preset"}
                  onClick={() => setSavedPoDateRange("3m")}
                >
                  3개월
                </button>
                <button
                  type="button"
                  className={savedPoDateRange === "1y" ? "preset active" : "preset"}
                  onClick={() => setSavedPoDateRange("1y")}
                >
                  1년
                </button>
              </div>
              <button
                type="button"
                className="ghost resetBtn"
                onClick={() => {
                  setSavedPoSearch("");
                  setSavedPoDateRange("all");
                  setSavedPoSortMode("order_date_asc");
                }}
              >
                필터 초기화
              </button>
              <button
                type="button"
                className="ghost poSavedBulkDeleteBtn"
                disabled={savedPoBulkSelectedVisibleCount === 0}
                onClick={() => void bulkDeleteSavedPoSelections()}
              >
                선택 삭제 ({savedPoBulkSelectedVisibleCount})
              </button>
            </div>
            {purchaseOrdersLoading ? (
              <div className="searchEmptyState">불러오는 중…</div>
            ) : !purchaseOrders.length ? (
              <div className="searchEmptyState">등록된 발주가 없습니다.</div>
            ) : !filteredPurchaseOrders.length ? (
              <div className="searchEmptyState">검색·기간 조건에 맞는 발주가 없습니다.</div>
            ) : (
              <>
              {!poSavedShowTopScroll ? (
                <div
                  className="poSavedStickyFilterTableGap"
                  style={{
                    top: poSavedFilterBottomSticky,
                    height: poSavedOpaqueGapHeightPx,
                  }}
                  aria-hidden="true"
                />
              ) : null}
              <div
                className={poSavedShowTopScroll ? "poSavedShowTopScrollPair" : undefined}
                style={{ width: "100%" }}
              >
                {poSavedShowTopScroll ? (
                  <div
                    ref={poSavedTopScrollRef}
                    className="tableTopScroll poSavedTableTopScroll stickyTableTopScroll"
                    style={{ top: poSavedFilterBottomSticky }}
                    onScroll={() => syncPoSavedScroll("top")}
                  >
                    <div style={{ width: poSavedTopScrollWidth || "100%" }} />
                  </div>
                ) : null}
                <div className="poSavedSpreadsheetWrap">
                  <div
                    className="poSavedStickyHeaderShell"
                    style={{ top: poSavedTableHeaderStickyTop }}
                  >
                    <div
                      ref={poSavedHeaderScrollRef}
                      className="poSavedHeaderScroll"
                      onScroll={() => syncPoSavedScroll("header")}
                    >
                      <table ref={poSavedHeadTableRef} className="poSavedSpreadsheet poSavedSpreadsheetHeadTable">
                        <thead>
                          <tr>
                            <th className="poSavedSsThAction">입고 추가</th>
                            <th>발주일자</th>
                            <th>상품번호</th>
                            <th>브랜드</th>
                            <th className="poSavedSsThName">상품명</th>
                            <th>상품유형</th>
                            <th>제조사</th>
                            <th>총 발주수량</th>
                            <th>납품가능일</th>
                            <th>입고예정일</th>
                            <th>ERP PO</th>
                            <th>실제입고일</th>
                            <th>입고수량</th>
                            <th>입고여부</th>
                            <th>비고</th>
                            <th className="poSavedSsThSelect">
                              <input
                                ref={savedPoBulkHeaderCbRef}
                                type="checkbox"
                                className="poSavedSsRowCheckbox"
                                disabled={!savedPoBulkSelectableKeysFlat.length}
                                checked={savedPoBulkAllVisibleSelected}
                                onChange={(e) => setSavedPoBulkSelectAllVisible(e.target.checked)}
                                aria-label="보이는 입고 전체 선택"
                              />
                            </th>
                          </tr>
                        </thead>
                      </table>
                    </div>
                  </div>
                  <div
                    ref={poSavedTableScrollRef}
                    className="poSavedBodyScroll"
                    onScroll={() => syncPoSavedScroll("table")}
                  >
                    <table ref={poSavedBodyTableRef} className="poSavedSpreadsheet poSavedSpreadsheetBodyTable">
                    {savedPoSpreadsheetGroups.map(({ po, lines }) => {
                      const lineRows = lines.length ? lines : [null];
                      const inboundCnt = lines.length;
                      const samePoDash = PO_SAVED_SAME_GROUP_CELL;
                      const samePoBlank = "";
                      const sheetNewDraft = savedPoNewLineDraftByOrderId[String(po.id)];
                      return (
                        <tbody key={po.id} className="poSavedSsGroup">
                          {lineRows.map((line, idx) => {
                            const head = idx === 0;
                            const linePending = poLineIsInboundPending(line);
                            const inboundBg = linePending ? "poSavedSsPendingInboundCols" : "";
                            const memoTrim = line ? String(line.line_memo || "").trim() : "";
                            const hasMemo = Boolean(memoTrim);
                            return (
                              <tr key={line ? `${po.id}-${line.id}` : `${po.id}-none`} className="poSavedSsDataRow">
                                <td className={`poSavedSsTd ${!head ? "poSavedSsTdSamePo" : ""}`}>
                                  {head ? (
                                    sheetNewDraft ? (
                                      <span className="poSavedSsAddInboundPlaceholder">{samePoBlank}</span>
                                    ) : (
                                      <button
                                        type="button"
                                        className="poSavedSsAddInboundBtn poSavedPreventPoRowDbl"
                                        onClick={() => openSavedPoNewLineRow(po)}
                                        aria-label="입고 차수 추가"
                                      >
                                        +
                                      </button>
                                    )
                                  ) : (
                                    samePoBlank
                                  )}
                                </td>
                                <td
                                  className={`poSavedSsColDate poSavedSsTd ${!head ? "poSavedSsTdSamePo" : ""}`}
                                >
                                  {head ? (
                                    <span className="poSavedSsHeadDateCell">
                                      {String(savedPoInline?.orderId) === String(po.id) &&
                                      savedPoInline?.field === "order_date" ? (
                                        <div className="poExpectedInboundBody poSavedSsDateColEditBody">
                                          <input
                                            type="text"
                                            className="poSavedSsInlineInput"
                                            disabled={Boolean(savedPoInline.orderPlanned)}
                                            value={
                                              savedPoInline.orderPlanned
                                                ? ""
                                                : String(savedPoInline.draft ?? "")
                                            }
                                            placeholder={PO_SAVED_DATE_PLACEHOLDER}
                                            onChange={(e) =>
                                              setSavedPoInline((s) =>
                                                s
                                                  ? {
                                                      ...s,
                                                      draft: e.target.value,
                                                      orderPlanned: false,
                                                    }
                                                  : s
                                              )
                                            }
                                            onBlur={(e) => {
                                              const s = savedPoInlineRef.current;
                                              if (!s || s.field !== "order_date") return;
                                              if (s.orderPlanned) {
                                                void patchPurchaseOrderField(po.id, {
                                                  order_date: null,
                                                  order_date_note: "발주 예정",
                                                });
                                                return;
                                              }
                                              const patch = normalizeOrderDateInput(e.currentTarget.value);
                                              void patchPurchaseOrderField(po.id, patch);
                                            }}
                                            autoFocus
                                            aria-label="발주일자 수정"
                                          />
                                          <label className="poExpectedInboundTbd">
                                            <input
                                              type="checkbox"
                                              checked={Boolean(savedPoInline.orderPlanned)}
                                              onMouseDown={(e) => e.preventDefault()}
                                              onChange={(e) => {
                                                const planned = e.target.checked;
                                                if (planned && poHasAnyCompletedInbound(po)) {
                                                  window.alert(PO_SAVED_EXPECTED_INBOUND_BLOCKED_O_MSG);
                                                  return;
                                                }
                                                setSavedPoInline((s) =>
                                                  s && s.field === "order_date"
                                                    ? {
                                                        ...s,
                                                        orderPlanned: planned,
                                                        draft: planned ? "" : s.draft,
                                                      }
                                                    : s
                                                );
                                              }}
                                            />
                                            발주 예정
                                          </label>
                                        </div>
                                      ) : (
                                        <span className="poSavedSsCellWithPencil">
                                          <span className="poSavedSsHeadDateText">{formatPoOrderDateDisplay(po)}</span>
                                          <button
                                            type="button"
                                            className="poSavedSsPencilBtn poSavedPreventPoRowDbl"
                                            aria-label="발주일자 수정"
                                            onClick={() =>
                                              setSavedPoInline({
                                                orderId: String(po.id),
                                                lineId: null,
                                                field: "order_date",
                                                draft: isOrderDatePlannedNote(po.order_date_note)
                                                  ? ""
                                                  : String(po.order_date || ""),
                                                orderPlanned: isOrderDatePlannedNote(po.order_date_note),
                                              })
                                            }
                                          >
                                            ✎
                                          </button>
                                        </span>
                                      )}
                                    </span>
                                  ) : (
                                    samePoBlank
                                  )}
                                </td>
                                <td
                                  className={`poSavedSsColSku poSavedSsTd ${!head ? "poSavedSsTdSamePo" : ""}`}
                                >
                                  {head ? purchaseOrderSkuForDisplay(po.sku) : samePoBlank}
                                </td>
                                <td className={`poSavedSsTd ${!head ? "poSavedSsTdSamePo" : ""}`}>
                                  {head ? po.brand || "–" : samePoBlank}
                                </td>
                                <td
                                  className={`poSavedSsTd poSavedSsTdName ${!head ? "poSavedSsTdSamePo" : ""}`}
                                >
                                  {head ? po.product_name || "–" : samePoBlank}
                                </td>
                                <td className={`poSavedSsTd ${!head ? "poSavedSsTdSamePo" : ""}`}>
                                  {head ? po.product_type || "–" : samePoBlank}
                                </td>
                                <td className={`poSavedSsTd ${!head ? "poSavedSsTdSamePo" : ""}`}>
                                  {head ? po.manufacturer || "–" : samePoBlank}
                                </td>
                                <td className={`poSavedSsTd ${!head ? "poSavedSsTdSamePo" : ""}`}>
                                  {head ? (
                                    String(savedPoInline?.orderId) === String(po.id) &&
                                    savedPoInline?.field === "total_quantity" ? (
                                      <input
                                        type="number"
                                        className="poSavedSsInlineInput poSavedSsInlineNumber"
                                        min={0}
                                        step="any"
                                        value={savedPoInline.draft}
                                        onChange={(e) =>
                                          setSavedPoInline((s) => (s ? { ...s, draft: e.target.value } : s))
                                        }
                                        onBlur={(e) => {
                                          const raw = e.currentTarget.value.replace(/,/g, "").trim();
                                          if (!raw || Number.isNaN(Number(raw))) {
                                            setSavedPoInline(null);
                                            return;
                                          }
                                          void patchPurchaseOrderField(po.id, {
                                            total_quantity: Number(raw),
                                          });
                                        }}
                                        autoFocus
                                        aria-label="총 발주수량 수정"
                                      />
                                    ) : (
                                      <span className="poSavedSsCellWithPencil">
                                        <span>
                                          {po.total_quantity != null ? formatInt(po.total_quantity) : "–"}
                                        </span>
                                        <button
                                          type="button"
                                          className="poSavedSsPencilBtn poSavedPreventPoRowDbl"
                                          aria-label="총 발주수량 수정"
                                          onClick={() =>
                                            setSavedPoInline({
                                              orderId: String(po.id),
                                              lineId: null,
                                              field: "total_quantity",
                                              draft:
                                                po.total_quantity != null ? String(po.total_quantity) : "",
                                            })
                                          }
                                        >
                                          ✎
                                        </button>
                                      </span>
                                    )
                                  ) : (
                                    samePoDash
                                  )}
                                </td>
                                <td
                                  className={`poSavedSsColDate poSavedSsTd ${!head ? "poSavedSsTdSamePo" : ""} ${inboundBg}`}
                                >
                                  {(() => {
                                    const deliveryDisp = line
                                      ? line.delivery_available_date || po.delivery_available_date || "–"
                                      : po.delivery_available_date || "–";
                                    const deliveryDraft = line
                                      ? String(
                                          line.delivery_available_date || po.delivery_available_date || ""
                                        )
                                      : String(po.delivery_available_date || "");
                                    const inlineDel =
                                      String(savedPoInline?.orderId) === String(po.id) &&
                                      savedPoInline?.field === "delivery_available" &&
                                      (line
                                        ? String(savedPoInline?.lineId) === String(line.id)
                                        : savedPoInline?.lineId == null);
                                    if (inlineDel) {
                                      return (
                                        <div className="poExpectedInboundBody">
                                          <input
                                            type="text"
                                            className="poSavedSsInlineInput"
                                            disabled={Boolean(savedPoInline.tbd)}
                                            value={
                                              savedPoInline.tbd ? "" : String(savedPoInline.draft ?? "")
                                            }
                                            placeholder={PO_SAVED_DATE_PLACEHOLDER}
                                            onChange={(e) =>
                                              setSavedPoInline((s) =>
                                                s
                                                  ? {
                                                      ...s,
                                                      draft: e.target.value,
                                                      tbd: false,
                                                    }
                                                  : s
                                              )
                                            }
                                            onBlur={(e) => {
                                              const s = savedPoInlineRef.current;
                                              const tbd = Boolean(s?.field === "delivery_available" && s?.tbd);
                                              const v = tbd ? "" : e.currentTarget.value.trim();
                                              if (line) {
                                                void patchInboundLineField(po.id, line.id, {
                                                  delivery_available_date: v || null,
                                                });
                                              } else {
                                                void patchPurchaseOrderField(po.id, {
                                                  delivery_available_date: v || null,
                                                });
                                              }
                                            }}
                                            autoFocus
                                            aria-label="납품가능일 수정"
                                          />
                                          <label className="poExpectedInboundTbd">
                                            <input
                                              type="checkbox"
                                              checked={Boolean(savedPoInline.tbd)}
                                              onMouseDown={(e) => e.preventDefault()}
                                              onChange={(e) => {
                                                const tbd = e.target.checked;
                                                setSavedPoInline((s) =>
                                                  s && s.field === "delivery_available"
                                                    ? { ...s, tbd, draft: tbd ? "" : s.draft }
                                                    : s
                                                );
                                              }}
                                            />
                                            미정
                                          </label>
                                        </div>
                                      );
                                    }
                                    return (
                                      <span className="poSavedSsCellWithPencil">
                                        <span>{deliveryDisp}</span>
                                        <button
                                          type="button"
                                          className="poSavedSsPencilBtn poSavedPreventPoRowDbl"
                                          aria-label="납품가능일 수정"
                                          onClick={() => {
                                            const d0 =
                                              deliveryDraft ||
                                              (deliveryDisp !== "–" ? String(deliveryDisp) : "");
                                            const has = Boolean(String(d0 || "").trim());
                                            setSavedPoInline({
                                              orderId: String(po.id),
                                              lineId: line ? String(line.id) : null,
                                              field: "delivery_available",
                                              draft: has ? d0 : "",
                                              tbd: !has,
                                            });
                                          }}
                                        >
                                          ✎
                                        </button>
                                      </span>
                                    );
                                  })()}
                                </td>
                                <td
                                  className={`poSavedSsColDate poSavedSsTd ${!head ? "poSavedSsTdSamePo" : ""} ${inboundBg}`}
                                >
                                  {(() => {
                                    const expectedDisp = formatPoLineExpectedInboundDisplay(line, po);
                                    const expectedDraftBase = line
                                      ? String(line.expected_inbound_date || po.expected_inbound_date || "")
                                      : String(po.expected_inbound_date || "");
                                    const inlineExpected =
                                      String(savedPoInline?.orderId) === String(po.id) &&
                                      savedPoInline?.field === "expected_inbound" &&
                                      (line
                                        ? String(savedPoInline?.lineId) === String(line.id)
                                        : savedPoInline?.lineId == null);
                                    if (inlineExpected) {
                                      return (
                                        <div className="poExpectedInboundBody">
                                          <input
                                            type="text"
                                            className="poSavedSsInlineInput"
                                            disabled={Boolean(savedPoInline.tbd)}
                                            value={
                                              savedPoInline.tbd ? "" : String(savedPoInline.draft ?? "")
                                            }
                                            placeholder={PO_SAVED_DATE_PLACEHOLDER}
                                            onChange={(e) =>
                                              setSavedPoInline((s) =>
                                                s
                                                  ? {
                                                      ...s,
                                                      draft: e.target.value,
                                                      tbd: false,
                                                    }
                                                  : s
                                              )
                                            }
                                            onBlur={(e) => {
                                              const s = savedPoInlineRef.current;
                                              const tbd = Boolean(
                                                s?.field === "expected_inbound" && s?.tbd
                                              );
                                              const v = tbd ? "" : e.currentTarget.value.trim();
                                              if (line) {
                                                void patchInboundLineField(po.id, line.id, {
                                                  expected_inbound_date: v || null,
                                                });
                                              } else {
                                                void patchPurchaseOrderField(po.id, {
                                                  expected_inbound_date: v || null,
                                                });
                                              }
                                            }}
                                            autoFocus
                                            aria-label="입고예정일 수정"
                                          />
                                          <label className="poExpectedInboundTbd">
                                            <input
                                              type="checkbox"
                                              checked={Boolean(savedPoInline.tbd)}
                                              onMouseDown={(e) => e.preventDefault()}
                                              onChange={(e) => {
                                                const tbd = e.target.checked;
                                                setSavedPoInline((s) =>
                                                  s && s.field === "expected_inbound"
                                                    ? { ...s, tbd, draft: tbd ? "" : s.draft }
                                                    : s
                                                );
                                              }}
                                            />
                                            미정
                                          </label>
                                        </div>
                                      );
                                    }
                                    return (
                                      <span className="poSavedSsCellWithPencil">
                                        <span>{expectedDisp}</span>
                                        <button
                                          type="button"
                                          className="poSavedSsPencilBtn poSavedPreventPoRowDbl"
                                          aria-label="입고예정일 수정"
                                          onClick={() => {
                                            if (
                                              line &&
                                              String(line.inbound_status || "").trim().toUpperCase() === "O"
                                            ) {
                                              window.alert(PO_SAVED_EXPECTED_INBOUND_BLOCKED_O_MSG);
                                              return;
                                            }
                                            const e0 =
                                              expectedDraftBase ||
                                              (expectedDisp !== "–" && expectedDisp !== "미정"
                                                ? String(expectedDisp)
                                                : "");
                                            const has = Boolean(String(e0 || "").trim());
                                            setSavedPoInline({
                                              orderId: String(po.id),
                                              lineId: line ? String(line.id) : null,
                                              field: "expected_inbound",
                                              draft: has ? e0 : "",
                                              tbd: !has,
                                            });
                                          }}
                                        >
                                          ✎
                                        </button>
                                      </span>
                                    );
                                  })()}
                                </td>
                                <td
                                  className={`poSavedSsTd poSavedSsTdErpCol ${inboundBg}`}
                                  title={line ? `입고 ${Number(line.line_no) || "?"}차` : undefined}
                                >
                                  {formatSavedPoErpCol(po, line, inboundCnt)}
                                </td>
                                <td className={`poSavedSsColDate poSavedSsTd ${inboundBg}`}>
                                  {!line ? (
                                    <span className="poSavedSsCellWithPencil">
                                      <span>–</span>
                                      <button
                                        type="button"
                                        className="poSavedSsPencilBtn poSavedPreventPoRowDbl"
                                        aria-label="실제입고일 수정"
                                        onClick={async () => {
                                          const lineId = await ensureFirstInboundLine(po);
                                          if (!lineId) return;
                                          setSavedPoInline({
                                            orderId: String(po.id),
                                            lineId,
                                            field: "actual",
                                            draft: "",
                                          });
                                        }}
                                      >
                                        ✎
                                      </button>
                                    </span>
                                  ) : String(savedPoInline?.orderId) === String(po.id) &&
                                    String(savedPoInline?.lineId) === String(line.id) &&
                                    savedPoInline?.field === "actual" ? (
                                    <input
                                      type="text"
                                      className="poSavedSsInlineInput"
                                      value={savedPoInline.draft}
                                      placeholder={PO_SAVED_ACTUAL_INBOUND_PLACEHOLDER}
                                      onChange={(e) =>
                                        setSavedPoInline((s) => (s ? { ...s, draft: e.target.value } : s))
                                      }
                                      onBlur={(e) => {
                                        const parsed = normalizeActualInboundInput(e.currentTarget.value);
                                        const patch = { ...parsed };
                                        if (parsed.actual_inbound_date || String(parsed.actual_inbound_note || "").trim()) {
                                          patch.inbound_status = "O";
                                        }
                                        void patchInboundLineField(po.id, line.id, patch);
                                      }}
                                      autoFocus
                                    />
                                  ) : (
                                    <span className="poSavedSsCellWithPencil">
                                      <span>{formatPoInboundActualDisplay(line)}</span>
                                      <button
                                        type="button"
                                        className="poSavedSsPencilBtn poSavedPreventPoRowDbl"
                                        aria-label="실제입고일 수정"
                                        onClick={() =>
                                          setSavedPoInline({
                                            orderId: String(po.id),
                                            lineId: String(line.id),
                                            field: "actual",
                                            draft: String(line.actual_inbound_note || line.actual_inbound_date || ""),
                                          })
                                        }
                                      >
                                        ✎
                                      </button>
                                    </span>
                                  )}
                                </td>
                                <td className={`poSavedSsTd ${inboundBg}`}>
                                  {!line ? (
                                    <span className="poSavedSsCellWithPencil">
                                      <span>–</span>
                                      <button
                                        type="button"
                                        className="poSavedSsPencilBtn poSavedPreventPoRowDbl"
                                        aria-label="입고수량 수정"
                                        onClick={async () => {
                                          const lineId = await ensureFirstInboundLine(po);
                                          if (!lineId) return;
                                          setSavedPoInline({
                                            orderId: String(po.id),
                                            lineId,
                                            field: "qty",
                                            draft:
                                              po.total_quantity != null ? String(po.total_quantity) : "",
                                          });
                                        }}
                                      >
                                        ✎
                                      </button>
                                    </span>
                                  ) : String(savedPoInline?.orderId) === String(po.id) &&
                                    String(savedPoInline?.lineId) === String(line.id) &&
                                    savedPoInline?.field === "qty" ? (
                                    <input
                                      type="number"
                                      className="poSavedSsInlineInput poSavedSsInlineNumber"
                                      min={0}
                                      step="any"
                                      value={savedPoInline.draft}
                                      onChange={(e) =>
                                        setSavedPoInline((s) => (s ? { ...s, draft: e.target.value } : s))
                                      }
                                      onBlur={(e) => {
                                        const raw = e.currentTarget.value.replace(/,/g, "").trim();
                                        if (!raw) {
                                          void patchInboundLineField(po.id, line.id, { quantity: 0 });
                                          return;
                                        }
                                        if (Number.isNaN(Number(raw))) {
                                          setSavedPoInline(null);
                                          return;
                                        }
                                        void patchInboundLineField(po.id, line.id, { quantity: Number(raw) });
                                      }}
                                      autoFocus
                                    />
                                  ) : (
                                    <span className="poSavedSsCellWithPencil">
                                      <span>{line.quantity != null ? formatInt(line.quantity) : "–"}</span>
                                      <button
                                        type="button"
                                        className="poSavedSsPencilBtn poSavedPreventPoRowDbl"
                                        aria-label="입고수량 수정"
                                        onClick={() =>
                                          setSavedPoInline({
                                            orderId: String(po.id),
                                            lineId: String(line.id),
                                            field: "qty",
                                            draft: line.quantity != null ? String(line.quantity) : "",
                                          })
                                        }
                                      >
                                        ✎
                                      </button>
                                    </span>
                                  )}
                                </td>
                                <td className={`poSavedSsTd ${inboundBg}`}>
                                  {!line ? (
                                    <span className="poSavedSsCellWithPencil">
                                      <span>–</span>
                                      <button
                                        type="button"
                                        className="poSavedSsPencilBtn poSavedPreventPoRowDbl"
                                        aria-label="입고여부 수정"
                                        onClick={async () => {
                                          const lineId = await ensureFirstInboundLine(po);
                                          if (!lineId) return;
                                          setSavedPoInline({
                                            orderId: String(po.id),
                                            lineId,
                                            field: "status",
                                            draft: "X",
                                          });
                                        }}
                                      >
                                        ✎
                                      </button>
                                    </span>
                                  ) : String(savedPoInline?.orderId) === String(po.id) &&
                                    String(savedPoInline?.lineId) === String(line.id) &&
                                    savedPoInline?.field === "status" ? (
                                    <select
                                      className="poSavedSsInlineSelect"
                                      value={savedPoInline.draft || "X"}
                                      onChange={(e) => {
                                        const v = e.target.value;
                                        if (v === "O") {
                                          const hasActualDate = String(line.actual_inbound_date || "").trim();
                                          const hasActualNote = String(line.actual_inbound_note || "").trim();
                                          if (!hasActualDate && !String(hasActualNote || "").trim()) {
                                            window.alert(
                                              "입고 완료(O)로 변경하려면 먼저 실제 입고일 또는 텍스트를 입력해주세요."
                                            );
                                            return;
                                          }
                                        }
                                        if (v === "X") {
                                          void patchInboundLineField(po.id, line.id, {
                                            inbound_status: v,
                                            actual_inbound_date: null,
                                            actual_inbound_note: null,
                                          });
                                          return;
                                        }
                                        void patchInboundLineField(po.id, line.id, { inbound_status: v });
                                      }}
                                      onBlur={() =>
                                        setSavedPoInline((s) => (s?.field === "status" ? null : s))
                                      }
                                      autoFocus
                                    >
                                      <option value="O">O</option>
                                      <option value="X">X</option>
                                    </select>
                                  ) : (
                                    <span className="poSavedSsCellWithPencil">
                                      <span>{formatPoInboundStatus(line.inbound_status)}</span>
                                      <button
                                        type="button"
                                        className="poSavedSsPencilBtn poSavedPreventPoRowDbl"
                                        aria-label="입고여부 수정"
                                        onClick={() => {
                                          const st = String(line.inbound_status || "")
                                            .trim()
                                            .toUpperCase();
                                          setSavedPoInline({
                                            orderId: String(po.id),
                                            lineId: String(line.id),
                                            field: "status",
                                            draft: st === "O" || st === "X" ? st : "X",
                                          });
                                        }}
                                      >
                                        ✎
                                      </button>
                                    </span>
                                  )}
                                </td>
                                <td
                                  className={`poSavedSsTd poSavedSsMemoCol ${hasMemo ? "poSavedSsMemoColHasMemo" : ""} ${inboundBg}`}
                                >
                                  {!line ? (
                                    <span className="poSavedSsMemoCellInner">
                                      <span className="poSavedSsCellWithPencil">
                                        <span>–</span>
                                        <button
                                          type="button"
                                          className="poSavedSsPencilBtn poSavedPreventPoRowDbl"
                                          aria-label="비고 수정"
                                          onClick={async () => {
                                            const lineId = await ensureFirstInboundLine(po);
                                            if (!lineId) return;
                                            setSavedPoMemoModal({
                                              orderId: po.id,
                                              lineId,
                                              draft: "",
                                            });
                                          }}
                                        >
                                          ✎
                                        </button>
                                      </span>
                                    </span>
                                  ) : (
                                    <span className="poSavedSsMemoCellInner">
                                      {hasMemo ? (
                                        <button
                                          type="button"
                                          className="poSavedMemoViewLink poSavedPreventPoRowDbl"
                                          onClick={() =>
                                            setSavedPoMemoModal({
                                              orderId: po.id,
                                              lineId: line.id,
                                              draft: memoTrim,
                                            })
                                          }
                                        >
                                          메모 보기
                                        </button>
                                      ) : (
                                        <span className="poSavedSsCellWithPencil">
                                          <span>–</span>
                                          <button
                                            type="button"
                                            className="poSavedSsPencilBtn poSavedPreventPoRowDbl"
                                            aria-label="비고 수정"
                                            onClick={() =>
                                              setSavedPoMemoModal({
                                                orderId: po.id,
                                                lineId: line.id,
                                                draft: memoTrim,
                                              })
                                            }
                                          >
                                            ✎
                                          </button>
                                        </span>
                                      )}
                                    </span>
                                  )}
                                </td>
                                <td className={`poSavedSsTd poSavedSsSelectCol ${inboundBg}`}>
                                  {!line ? (
                                    <input
                                      type="checkbox"
                                      className="poSavedSsRowCheckbox poSavedPreventPoRowDbl"
                                      checked={Boolean(savedPoBulkSelected[`P|${po.id}`])}
                                      onChange={() => toggleSavedPoBulkKey(`P|${po.id}`)}
                                      aria-label="이 발주 삭제 대상에 포함 (입고 차수 없음)"
                                    />
                                  ) : (
                                    <input
                                      type="checkbox"
                                      className="poSavedSsRowCheckbox poSavedPreventPoRowDbl"
                                      checked={Boolean(savedPoBulkSelected[`L|${po.id}|${line.id}`])}
                                      onChange={() => toggleSavedPoBulkKey(`L|${po.id}|${line.id}`)}
                                      aria-label={`입고 차수 ${line.line_no ?? ""} 삭제 대상에 포함`}
                                    />
                                  )}
                                </td>
                              </tr>
                            );
                          })}
                          {sheetNewDraft ? (
                            <tr
                              key={`${po.id}-sheet-new-line`}
                              className="poSavedSsDataRow poSavedSsNewInboundSheetRow"
                              ref={(el) => {
                                const k = String(po.id);
                                if (el) savedPoNewLineRowRefs.current[k] = el;
                                else delete savedPoNewLineRowRefs.current[k];
                              }}
                            >
                              <td className="poSavedSsTd poSavedSsTdSamePo">
                                <span className="poSavedSsNewLineBadge">+ 새 차수</span>
                              </td>
                              <td className="poSavedSsColDate poSavedSsTd poSavedSsTdSamePo">
                                {samePoBlank}
                              </td>
                              <td className="poSavedSsColSku poSavedSsTd poSavedSsTdSamePo">{samePoBlank}</td>
                              <td className="poSavedSsTd poSavedSsTdSamePo">{samePoBlank}</td>
                              <td className="poSavedSsTd poSavedSsTdName poSavedSsTdSamePo">{samePoBlank}</td>
                              <td className="poSavedSsTd poSavedSsTdSamePo">{samePoBlank}</td>
                              <td className="poSavedSsTd poSavedSsTdSamePo">{samePoBlank}</td>
                              <td className="poSavedSsTd poSavedSsTdSamePo">{samePoDash}</td>
                              <td className={`poSavedSsColDate poSavedSsTd poSavedSsNewInboundCols`}>
                                <div className="poExpectedInboundBody">
                                  <input
                                    type="text"
                                    className="poSavedSsInlineInput poSavedSsNewLineInputWide"
                                    placeholder={PO_SAVED_DATE_PLACEHOLDER}
                                    value={sheetNewDraft.delivery_available_tbd ? "" : sheetNewDraft.delivery_available_date || ""}
                                    disabled={sheetNewDraft.delivery_available_tbd}
                                    onChange={(e) =>
                                      updateSavedPoNewLineDraft(po.id, {
                                        delivery_available_date: e.target.value,
                                        delivery_available_tbd: false,
                                      })
                                    }
                                    aria-label="납품가능일 (신규 차수)"
                                  />
                                  <label className="poExpectedInboundTbd">
                                    <input
                                      type="checkbox"
                                      checked={Boolean(sheetNewDraft.delivery_available_tbd)}
                                      onChange={(e) =>
                                        updateSavedPoNewLineDraft(po.id, {
                                          delivery_available_tbd: e.target.checked,
                                          delivery_available_date: e.target.checked
                                            ? ""
                                            : sheetNewDraft.delivery_available_date,
                                        })
                                      }
                                    />
                                    미정
                                  </label>
                                </div>
                              </td>
                              <td className={`poSavedSsColDate poSavedSsTd poSavedSsNewInboundCols`}>
                                <div className="poExpectedInboundBody">
                                  <input
                                    type="text"
                                    className="poSavedSsInlineInput poSavedSsNewLineInputWide"
                                    placeholder={PO_SAVED_DATE_PLACEHOLDER}
                                    value={
                                      sheetNewDraft.expected_inbound_tbd
                                        ? ""
                                        : sheetNewDraft.expected_inbound_date || ""
                                    }
                                    disabled={
                                      Boolean(sheetNewDraft.expected_inbound_tbd) ||
                                      String(sheetNewDraft.inbound_status || "X").toUpperCase() === "O"
                                    }
                                    onChange={(e) =>
                                      updateSavedPoNewLineDraft(po.id, {
                                        expected_inbound_date: e.target.value,
                                        expected_inbound_tbd: false,
                                      })
                                    }
                                    aria-label="입고예정일 (신규 차수)"
                                  />
                                  <label className="poExpectedInboundTbd">
                                    <input
                                      type="checkbox"
                                      checked={Boolean(sheetNewDraft.expected_inbound_tbd)}
                                      disabled={String(sheetNewDraft.inbound_status || "X").toUpperCase() === "O"}
                                      onChange={(e) =>
                                        updateSavedPoNewLineDraft(po.id, (cur) => ({
                                          expected_inbound_tbd: e.target.checked,
                                          expected_inbound_date: e.target.checked ? "" : cur.expected_inbound_date,
                                        }))
                                      }
                                    />
                                    미정
                                  </label>
                                </div>
                              </td>
                              <td className={`poSavedSsTd poSavedSsTdErpCol poSavedSsNewInboundCols`}>
                                <span className="poSavedSsNewLineHint">저장 시 부여</span>
                              </td>
                              <td className={`poSavedSsColDate poSavedSsTd poSavedSsNewInboundCols`}>
                                <input
                                  type="text"
                                  className="poSavedSsInlineInput poSavedSsNewLineInputWide"
                                  placeholder={PO_SAVED_ACTUAL_INBOUND_PLACEHOLDER}
                                  value={sheetNewDraft.actual_inbound_input || ""}
                                  onChange={(e) =>
                                    updateSavedPoNewLineDraft(po.id, {
                                      actual_inbound_input: e.target.value,
                                    })
                                  }
                                  aria-label="실제입고 (신규 차수)"
                                />
                              </td>
                              <td className={`poSavedSsTd poSavedSsNewInboundCols`}>
                                <input
                                  type="number"
                                  className="poSavedSsInlineInput poSavedSsInlineNumber"
                                  min={0}
                                  step="any"
                                  placeholder="수량"
                                  value={sheetNewDraft.quantity}
                                  onChange={(e) =>
                                    updateSavedPoNewLineDraft(po.id, { quantity: e.target.value })
                                  }
                                  aria-label="입고수량 (신규 차수)"
                                />
                              </td>
                              <td className={`poSavedSsTd poSavedSsNewInboundCols`}>
                                <select
                                  className="poSavedSsInlineSelect"
                                  value={String(sheetNewDraft.inbound_status || "X").toUpperCase()}
                                  onChange={(e) => {
                                    const v = e.target.value;
                                    updateSavedPoNewLineDraft(po.id, (cur) => ({
                                      inbound_status: v,
                                      expected_inbound_date: v === "O" ? "" : cur.expected_inbound_date,
                                      expected_inbound_tbd: v === "O" ? false : cur.expected_inbound_tbd,
                                    }));
                                  }}
                                  aria-label="입고여부 (신규 차수)"
                                >
                                  <option value="O">O</option>
                                  <option value="X">X</option>
                                </select>
                              </td>
                              <td className={`poSavedSsTd poSavedSsMemoCol poSavedSsNewInboundCols`}>
                                <div className="poSavedSsNewLineActions">
                                  <button
                                    type="button"
                                    className="primary poSavedSsNewLineSaveBtn"
                                    onClick={() =>
                                      void submitInboundLine(po.id, {
                                        fromSavedSheet: true,
                                        draft: sheetNewDraft,
                                      })
                                    }
                                  >
                                    저장
                                  </button>
                                  <button
                                    type="button"
                                    className="ghost poSavedSsNewLineCancelBtn"
                                    onClick={() => cancelSavedPoNewLineRow(po.id)}
                                  >
                                    취소
                                  </button>
                                </div>
                              </td>
                              <td className="poSavedSsTd poSavedSsSelectCol poSavedSsNewInboundCols" aria-hidden="true" />
                            </tr>
                          ) : null}
                        </tbody>
                      );
                    })}
                    </table>
                  </div>
                </div>
              </div>
                {editingPoId
                  ? (() => {
                      const po = purchaseOrders.find((p) => String(p.id) === String(editingPoId));
                      if (!po) return null;
                      const isEditing = true;
                      const d = poEditDraft;
                      if (!d) return null;
                      const displayLines = d ? d.lines : po.inbound_lines || [];
                      const inboundLineCount = displayLines.length;
                      const erpLabel = d ? d.erp_po_number || po.erp_po_number : po.erp_po_number;
                      const erpUnchanged =
                        String(d.erp_po_number || "").trim() === String(po.erp_po_number || "").trim();
                      return createPortal(
                        <div
                          className="poSavedPoEditModalBackdrop"
                          onClick={(e) => {
                            if (e.target === e.currentTarget) cancelPoEdit();
                          }}
                        >
                          <div
                            className="poSavedPoEditModal"
                            role="dialog"
                            aria-modal="true"
                            aria-labelledby="poSavedPoEditModalTitle"
                            onClick={(e) => e.stopPropagation()}
                          >
                            <div id="poSavedPoEditModalTitle" className="poSavedPoEditModalTitle">
                              발주 정보 수정
                            </div>
                          <article key={po.id} className="purchaseOrderCard">
                    <div className="poCardHeroBand">
                      <div className="poCardErpBand">
                        <div className="poCardHeadTop">
                          <div className="poCardPoBlock">
                            {isEditing && d ? (
                              <input
                                type="text"
                                className="poCardErpInput"
                                value={d.erp_po_number}
                                onChange={(e) =>
                                  setPoEditDraft((p) => (p ? { ...p, erp_po_number: e.target.value } : p))
                                }
                                placeholder="ERP PO 번호"
                                autoComplete="off"
                                aria-label="ERP PO 번호"
                              />
                            ) : (
                              <div className="poCardPo">{purchaseOrderErpForDisplay(po.erp_po_number)}</div>
                            )}
                          </div>
                          <div className="poCardHeadActions">
                            <>
                              <button type="button" className="primary" onClick={() => submitPoEditSave()}>
                                저장
                              </button>
                              <button type="button" className="ghost" onClick={() => cancelPoEdit()}>
                                취소
                              </button>
                              <button
                                type="button"
                                className="ghost poCardDeletePoBtn"
                                onClick={() => deletePurchaseOrder(po.id)}
                              >
                                삭제
                              </button>
                            </>
                          </div>
                        </div>
                        <div className="poCardMeta poCardMetaLg">
                          {isEditing && d ? (
                            <div className="poCardMetaEditRow">
                              <label className="poCardMetaField poCardMetaFieldStack">
                                <span className="poCardMetaLabel">발주일</span>
                                <div className="poExpectedInboundBody">
                                  <input
                                    type="text"
                                    className="poCardMetaInput"
                                    placeholder={PO_SAVED_DATE_PLACEHOLDER}
                                    disabled={isOrderDatePlannedNote(d.order_date_note)}
                                    value={isOrderDatePlannedNote(d.order_date_note) ? "" : d.order_date}
                                    onChange={(e) =>
                                      setPoEditDraft((p) =>
                                        p
                                          ? {
                                              ...p,
                                              order_date: e.target.value,
                                              order_date_note: e.target.value ? "" : p.order_date_note,
                                            }
                                          : p
                                      )
                                    }
                                  />
                                  <label className="poExpectedInboundTbd">
                                    <input
                                      type="checkbox"
                                      checked={isOrderDatePlannedNote(d.order_date_note)}
                                      onChange={(e) => {
                                        const planned = e.target.checked;
                                        setPoEditDraft((p) =>
                                          p
                                            ? {
                                                ...p,
                                                order_date: planned ? "" : p.order_date,
                                                order_date_note: planned ? "발주 예정" : "",
                                              }
                                            : p
                                        );
                                      }}
                                    />
                                    발주 예정
                                  </label>
                                </div>
                              </label>
                              <label className="poCardMetaField">
                                <span className="poCardMetaLabel">SKU</span>
                                <input
                                  type="text"
                                  className="poCardMetaInput"
                                  value={d.sku}
                                  onChange={(e) =>
                                    setPoEditDraft((p) =>
                                      p ? { ...p, sku: e.target.value, skuResolveHint: "" } : p
                                    )
                                  }
                                  onBlur={() => resolvePoEditSku()}
                                  autoComplete="off"
                                />
                              </label>
                            </div>
                          ) : (
                            <div className="poCardMetaLine">
                              <span className="poCardMetaItem">
                                <span className="poCardMetaK">발주일</span>{" "}
                                <span className="poCardMetaV">{formatPoOrderDateDisplay(po)}</span>
                              </span>
                              <span className="poCardMetaItem">
                                <span className="poCardMetaK">SKU</span>{" "}
                                <span className="poCardMetaV">{purchaseOrderSkuForDisplay(po.sku)}</span>
                              </span>
                            </div>
                          )}
                        </div>
                      </div>
                      {isEditing && d?.skuResolveHint ? (
                        <small className="poSkuHint poCardSkuHint">{d.skuResolveHint}</small>
                      ) : null}
                    </div>
                    <div className="poCardDetailWrap">
                    {isEditing && d ? (
                      <dl className="poCardDl poCardDlEdit poCardDlFields">
                        <div>
                          <dt>상품유형</dt>
                          <dd>
                            <PurchaseProductTypeField
                              preset={d.product_type_preset}
                              custom={d.product_type_custom}
                              onPresetChange={(v) =>
                                setPoEditDraft((p) => (p ? { ...p, product_type_preset: v } : p))
                              }
                              onCustomChange={(v) =>
                                setPoEditDraft((p) => (p ? { ...p, product_type_custom: v } : p))
                              }
                            />
                          </dd>
                        </div>
                        <div>
                          <dt>브랜드</dt>
                          <dd>
                            <input
                              type="text"
                              className="poCardDdInput"
                              value={d.brand}
                              onChange={(e) =>
                                setPoEditDraft((p) => (p ? { ...p, brand: e.target.value } : p))
                              }
                              autoComplete="off"
                            />
                          </dd>
                        </div>
                        <div>
                          <dt>상품명</dt>
                          <dd>
                            <input
                              type="text"
                              className="poCardDdInput"
                              value={d.product_name}
                              onChange={(e) =>
                                setPoEditDraft((p) => (p ? { ...p, product_name: e.target.value } : p))
                              }
                              autoComplete="off"
                            />
                          </dd>
                        </div>
                        <div>
                          <dt>제조사</dt>
                          <dd>
                            <input
                              type="text"
                              className="poCardDdInput"
                              value={d.manufacturer}
                              onChange={(e) =>
                                setPoEditDraft((p) => (p ? { ...p, manufacturer: e.target.value } : p))
                              }
                              autoComplete="off"
                            />
                          </dd>
                        </div>
                        <div>
                          <dt>총 발주수량</dt>
                          <dd>
                            <input
                              type="number"
                              className="poCardDdInput"
                              min={0}
                              step="any"
                              value={d.total_quantity}
                              onChange={(e) =>
                                setPoEditDraft((p) => (p ? { ...p, total_quantity: e.target.value } : p))
                              }
                            />
                          </dd>
                        </div>
                        <div>
                          <dt>납품가능일 (발주 공통)</dt>
                          <dd>
                            <input
                              type="text"
                              className="poCardDdInput"
                              placeholder={PO_SAVED_DATE_PLACEHOLDER}
                              value={d.delivery_available_date}
                              onChange={(e) =>
                                setPoEditDraft((p) =>
                                  p ? { ...p, delivery_available_date: e.target.value } : p
                                )
                              }
                            />
                            <span className="poInboundNewHint">
                              입고 차수별 납품가능일은 아래 표에서 입력합니다. 비우면 차수 값만 표시됩니다.
                            </span>
                          </dd>
                        </div>
                        <div>
                          <dt>입고예정일 (발주 공통)</dt>
                          <dd>
                            <div className="poExpectedInboundBody poCardExpectedInbound">
                              <input
                                type="text"
                                className="poCardDdInput poCardDdInputDate"
                                placeholder={PO_SAVED_DATE_PLACEHOLDER}
                                disabled={d.expected_inbound_tbd}
                                value={d.expected_inbound_tbd ? "" : d.expected_inbound_date}
                                onChange={(e) =>
                                  setPoEditDraft((p) =>
                                    p
                                      ? {
                                          ...p,
                                          expected_inbound_date: e.target.value,
                                          expected_inbound_tbd: false,
                                        }
                                      : p
                                  )
                                }
                              />
                              <label className="poExpectedInboundTbd">
                                <input
                                  type="checkbox"
                                  checked={d.expected_inbound_tbd}
                                  onChange={(e) => {
                                    const tbd = e.target.checked;
                                    setPoEditDraft((p) =>
                                      p
                                        ? {
                                            ...p,
                                            expected_inbound_tbd: tbd,
                                            expected_inbound_date: tbd ? "" : p.expected_inbound_date,
                                          }
                                        : p
                                    );
                                  }}
                                />
                                미정
                              </label>
                            </div>
                          </dd>
                        </div>
                      </dl>
                    ) : (
                      <dl className="poCardDl poCardDlFields">
                        <div>
                          <dt>상품유형</dt>
                          <dd>{po.product_type || "–"}</dd>
                        </div>
                        <div>
                          <dt>브랜드</dt>
                          <dd>{po.brand || "–"}</dd>
                        </div>
                        <div>
                          <dt>상품명</dt>
                          <dd>{po.product_name || "–"}</dd>
                        </div>
                        <div>
                          <dt>제조사</dt>
                          <dd>{po.manufacturer || "–"}</dd>
                        </div>
                        <div>
                          <dt>총 발주수량</dt>
                          <dd>{po.total_quantity != null ? formatInt(po.total_quantity) : "–"}</dd>
                        </div>
                        <div>
                          <dt>납품가능일</dt>
                          <dd>{po.delivery_available_date || "–"}</dd>
                        </div>
                        <div>
                          <dt>입고예정일</dt>
                          <dd>{po.expected_inbound_date || "미정"}</dd>
                        </div>
                      </dl>
                    )}
                    <div className="poInboundBlock">
                      <h4 className="poInboundTitle">입고 차수</h4>
                      <table className="poInboundTable">
                        <thead>
                          <tr>
                            <th>ERP PO 차수</th>
                            <th>납품가능일</th>
                            <th>입고예정일</th>
                            <th>실제입고일</th>
                            <th>입고수량</th>
                            <th>입고여부</th>
                          </tr>
                        </thead>
                        <tbody>
                          {!displayLines.length ? (
                            <tr>
                              <td colSpan={6} className="poInboundEmpty">
                                아직 등록된 입고 차수가 없습니다. 위 저장 표에서 해당 발주 행을 더블클릭하면 바로 아래에 새 행이 생깁니다.
                              </td>
                            </tr>
                          ) : (
                            displayLines.map((line) =>
                              isEditing && d ? (
                                <tr key={line.id}>
                                  <td
                                    className="poRefCell poInboundRefReadonly"
                                    title="ERP PO 차수는 직접 수정할 수 없습니다. ERP PO 번호를 바꾼 뒤 저장하면 코드가 맞춰 자동 반영됩니다."
                                  >
                                    {erpUnchanged
                                      ? formatSavedPoErpCol(po, line, inboundLineCount)
                                      : formatPoInboundRefPreview(d.erp_po_number, line, inboundLineCount)}
                                    {!erpUnchanged ? (
                                      <span className="poRefPreviewTag">미리보기</span>
                                    ) : null}
                                  </td>
                                  <td>
                                    <input
                                      type="text"
                                      className="poInboundCellControl"
                                      placeholder={PO_SAVED_DATE_PLACEHOLDER}
                                      aria-label={`납품가능일 – ${erpLabel}`}
                                      value={line.delivery_available_date || ""}
                                      onChange={(e) => {
                                        const v = e.target.value;
                                        setPoEditDraft((prev) =>
                                          prev
                                            ? {
                                                ...prev,
                                                lines: prev.lines.map((x) =>
                                                  x.id === line.id ? { ...x, delivery_available_date: v } : x
                                                ),
                                              }
                                            : prev
                                        );
                                      }}
                                    />
                                  </td>
                                  <td>
                                    <input
                                      type="text"
                                      className="poInboundCellControl"
                                      placeholder={PO_SAVED_DATE_PLACEHOLDER}
                                      aria-label={`입고예정일 – ${erpLabel}`}
                                      value={line.expected_inbound_date || ""}
                                      onChange={(e) => {
                                        const v = e.target.value;
                                        setPoEditDraft((prev) =>
                                          prev
                                            ? {
                                                ...prev,
                                                lines: prev.lines.map((x) =>
                                                  x.id === line.id ? { ...x, expected_inbound_date: v } : x
                                                ),
                                              }
                                            : prev
                                        );
                                      }}
                                    />
                                  </td>
                                  <td>
                                    <input
                                      type="text"
                                      className="poInboundCellControl"
                                      placeholder={PO_SAVED_ACTUAL_INBOUND_PLACEHOLDER}
                                      aria-label={`실제입고 – ${erpLabel}`}
                                      value={
                                        String(line.actual_inbound_note || "").trim() ||
                                        line.actual_inbound_date ||
                                        ""
                                      }
                                      onChange={(e) => {
                                        const v = e.target.value;
                                        const p = normalizeActualInboundInput(v);
                                        setPoEditDraft((prev) =>
                                          prev
                                            ? {
                                                ...prev,
                                                lines: prev.lines.map((x) =>
                                                  x.id === line.id
                                                    ? {
                                                        ...x,
                                                        actual_inbound_date: p.actual_inbound_date || "",
                                                        actual_inbound_note: p.actual_inbound_note || "",
                                                      }
                                                    : x
                                                ),
                                              }
                                            : prev
                                        );
                                      }}
                                    />
                                  </td>
                                  <td>
                                    <input
                                      type="number"
                                      className="poInboundCellControl"
                                      min={0}
                                      step="any"
                                      placeholder="입고수량"
                                      aria-label={`입고수량 – ${erpLabel}`}
                                      value={line.quantity}
                                      onChange={(e) => {
                                        const v = e.target.value;
                                        setPoEditDraft((prev) =>
                                          prev
                                            ? {
                                                ...prev,
                                                lines: prev.lines.map((x) =>
                                                  x.id === line.id ? { ...x, quantity: v } : x
                                                ),
                                              }
                                            : prev
                                        );
                                      }}
                                    />
                                  </td>
                                  <td>
                                    <select
                                      className="poInboundCellControl poInboundStatusSelect poInboundStatusSelectFull"
                                      aria-label={`입고여부 – ${erpLabel}`}
                                      value={line.inbound_status}
                                      onChange={(e) => {
                                        const v = e.target.value;
                                        setPoEditDraft((prev) =>
                                          prev
                                            ? {
                                                ...prev,
                                                lines: prev.lines.map((x) =>
                                                  x.id === line.id
                                                    ? {
                                                        ...x,
                                                        inbound_status: v,
                                                        expected_inbound_date:
                                                          v === "O" ? "" : x.expected_inbound_date,
                                                      }
                                                    : x
                                                ),
                                              }
                                            : prev
                                        );
                                      }}
                                    >
                                      <option value="O">O</option>
                                      <option value="X">X</option>
                                    </select>
                                  </td>
                                </tr>
                              ) : (
                                <tr key={line.id}>
                                  <td className="poRefCell">
                                    {formatSavedPoErpCol(po, line, inboundLineCount)}
                                  </td>
                                  <td>
                                    {line.delivery_available_date || po.delivery_available_date || "–"}
                                  </td>
                                  <td>{formatPoLineExpectedInboundDisplay(line, po)}</td>
                                  <td>{formatPoInboundActualDisplay(line)}</td>
                                  <td>{line.quantity != null ? formatInt(line.quantity) : "–"}</td>
                                  <td>{line.inbound_status}</td>
                                </tr>
                              )
                            )
                          )}
                        </tbody>
                      </table>
                    </div>
                    </div>
                  </article>
                          </div>
                        </div>,
                        document.body
                      );
                    })()
                  : null}
              </>
            )}
          </div>
          ) : null}
        </section>
      )}

      {savedPoMemoModal ? (
        <div className="poMemoModalBackdrop" onClick={() => void closeSavedPoMemoModal()}>
          <div className="poMemoModal" onClick={(e) => e.stopPropagation()}>
            <div className="poMemoModalHead">
              <div className="poMemoModalTitle">비고란</div>
              <button type="button" className="ghost poMemoModalClose" onClick={() => void closeSavedPoMemoModal()}>
                닫기
              </button>
            </div>
            <textarea
              className="poMemoModalTextarea"
              rows={8}
              value={savedPoMemoModal.draft}
              onChange={(e) => {
                const v = e.target.value;
                poMemoDraftRef.current = v;
                setSavedPoMemoModal((m) => (m ? { ...m, draft: v } : m));
                if (poMemoAutosaveTimerRef.current) {
                  clearTimeout(poMemoAutosaveTimerRef.current);
                }
                poMemoAutosaveTimerRef.current = window.setTimeout(() => {
                  const ids = poMemoModalIdsRef.current;
                  if (!ids) return;
                  void patchInboundLineField(ids.orderId, ids.lineId, {
                    line_memo: String(poMemoDraftRef.current || "").trim() || null,
                  });
                }, 550);
              }}
              placeholder="메모를 입력하세요"
              aria-label="비고란"
              autoFocus
            />
            <div className="poMemoModalActions">
              <button type="button" className="poMemoModalDeleteBtn" onClick={() => void deleteSavedPoMemo()}>
                삭제하기
              </button>
            </div>
          </div>
        </div>
      ) : null}

      {showCautionModal && (
        <div className="cautionModalBackdrop" onClick={() => setShowCautionModal(false)}>
          <div className="cautionModal" onClick={(e) => e.stopPropagation()}>
            <div className="cautionModalHeader">
              <div>
                <div className="cautionModalTitle">안내문</div>
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
              <div className="cautionSectionTitle">공통 (재고·출고·파일)</div>
              <ul className="cautionList">
                <li>
                  한국·해외 재고 탭에서 파일 업로드 후에는 재고 통합 실행을, 출고 탭에서는 출고 통합 실행을 눌러야 각
                  화면에 반영됩니다.
                </li>
                <li>재고 탭에서는 이미 올린 것과 같은 파일 이름은 다시 올라가지 않습니다.</li>
                <li>
                  재고 엑셀에 넣은 상품코드는 SKU 관리에서 그 국가로 먼저 등록되어 있어야 합니다. 하나라도 빠지면 그 파일
                  전체가 반영되지 않을 수 있습니다.
                </li>
                <li>재고 표의 상품코드에 마우스를 올리면, 등록되어 있는 바코드를 볼 수 있습니다.</li>
              </ul>
            </div>

            <div className="cautionSection">
              <div className="cautionSectionTitle">한국 재고</div>
              <ul className="cautionList">
                <li>
                  이지어드민에서 제공하는 현 재고 데이터는 엑셀로 열리기는 하지만 .xls 확장자로 제공됩니다. Excel
                  통합문서(.xlsx) 확장자로 변환하여 업로드 해주세요.
                </li>
                <li>
                  기준 날짜는 파일 이름에 들어 있는 날짜(예: 20250415)를 씁니다. 날짜가 파일명에 없으면 집계·비교가 어긋날 수
                  있습니다.
                </li>
                <li>재고 수량은 열 이름이 「가용재고」인 열만 읽습니다.</li>
                <li>날짜별로 올린 재고 스냅샷을 나란히 비교하는 화면입니다.</li>
              </ul>
            </div>

            <div className="cautionSection">
              <div className="cautionSectionTitle">해외 재고</div>
              <ul className="cautionList">
                <li>기준 날짜는 파일 안의 날짜 열(date)에서 가져옵니다.</li>
                <li>재고 수량은 열 이름이 「quantity」인 열만 읽습니다.</li>
                <li>
                  한국 현 재고와 비교는, 한국 시간 기준 오늘 날짜의 한국 재고가 있을 때만 켤 수 있습니다.
                </li>
              </ul>
            </div>

            <div className="cautionSection">
              <div className="cautionSectionTitle">재고 비교</div>
              <ul className="cautionList">
                <li>
                  「구분」은 같은 상품코드 안에서 서로 다른 행을 가리키는 레벨(L1 등)·창고(또는 파일의 분류 열)을
                  슬래시(/)로 이어 붙인 값입니다.
                </li>
                <li>선택한 날짜에 데이터가 없는 나라는 숫자 대신 –로 보입니다.</li>
                <li>다른 날짜나 예전 데이터를 임의로 끌어와 채우지 않습니다.</li>
                <li>반드시 같은 날짜의 데이터로만 비교합니다.</li>
              </ul>
            </div>

            <div className="cautionSection">
              <div className="cautionSectionTitle">출고 기록</div>
              <ul className="cautionList">
                <li>
                  엑셀에 「B2B 통합」「쿠팡 현황」 등 정해진 이름의 시트만 읽습니다. 시트 상단에 합계 행이 있어도
                  「품번(대표코드)」가 있는 헤더 행과 그 아래 데이터만 사용합니다.
                </li>
                <li>
                  품번(대표코드)는 SKU 관리에 등록된 한국 SKU와 일치해야 합니다. 마케팅 우선순위·구분·N월 합계·일자별
                  열은 엑셀 값을 그대로 쓰며, 브랜드·상품명은 DB 매핑 값으로 표시합니다.
                </li>
                <li>같은 날짜·상품 수치는 새 파일을 통합할 때 덮어씁니다(2026년 기준).</li>
                <li>
                  전체 출고 현황에서 시트(칩)를 바꿔 표를 보고, 상품별 출고 현황에서는 상품을 검색·선택하면 일자·월·시트별
                  차트와 시트×일자 표가 한 화면에 표시됩니다.
                </li>
              </ul>
            </div>

            <div className="cautionSection">
              <div className="cautionSectionTitle">발주 기록</div>
              <ul className="cautionList">
                <li>
                  새 발주에서 ERP 번호·상품코드를 비워도 저장할 수 있습니다. 비운 칸은 목록에서 보통 – 로 보입니다. 수량은
                  숫자로 적어 주세요.
                </li>
                <li>날짜는 칸에 직접 입력합니다. YYYY-MM-DD 형식을 권장합니다.</li>
                <li>
                  발주 예정이거나 납품일, 입고예정일이 미정인 경우 날짜는 비워 두고 체크박스에 체크하면 됩니다.
                </li>
                <li>
                  실제입고일에는 날짜를 YYYY-MM-DD 형식으로 입력하거나, 특수한 경우에는 예외 입고, 무상 입고 등을 입력해
                  주시면 됩니다.
                </li>
                <li>비고란에는 메모하고 싶은 내용을 자유롭게 적을 수 있습니다.</li>
                <li>저장하면 첫 입고 차수가 자동으로 생겨, 입고 관련 칸을 바로 쓸 수 있습니다.</li>
                <li>
                  저장된 발주 표에서 수정이 필요하면 연필 모양 버튼을 눌러 수정한 뒤, 화면 아무 곳이나 누르면 저장됩니다.
                </li>
                <li>입고 예정인 입고 건은 붉게 강조됩니다.</li>
                <li>비고에 메모가 있으면 노란색으로 강조됩니다.</li>
              </ul>
            </div>

            <div className="cautionSection">
              <div className="cautionSectionTitle">상품 매핑</div>
              <ul className="cautionList">
                <li>상품명이나 SKU로 검색하면, 국가별 상품명·상품코드를 한 번에 볼 수 있습니다.</li>
              </ul>
            </div>

            <div className="cautionSection">
              <div className="cautionSectionTitle">데이터 관리</div>
              <ul className="cautionList">
                <li>재고 탭에서 올린 파일·통합까지 끝난 파일에 관한 정보가 국가별로 보입니다.</li>
                <li>출고 탭에서 저장된 출고 엑셀은 「출고」 그룹에서 이름·용량 확인과 삭제, 출고 데이터만 초기화할 수 있습니다.</li>
                <li>파일 이름, 크기, 해당 파일의 재고 기준 날짜를 확인할 수 있습니다.</li>
                <li>
                  각 파일 옆 날짜는 기준일을 잡을 때 참고됩니다. 한국은 파일명 날짜도 중요하니, 여기 입력된 날짜와 파일명이
                  어긋나지 않게 맞춰 주세요.
                </li>
                <li>출고 파일에는 재고와 같은 기준일 입력란이 없습니다. 일자는 엑셀 시트·행에서 읽습니다.</li>
                <li>삭제하면 해당 파일에서 읽어 둔 재고·출고 저장 데이터는 완전히 삭제됩니다.</li>
                <li>국가별 데이터 초기화는 그 나라에 올린 재고 파일 데이터를 한꺼번에 삭제합니다.</li>
                <li>맨 위 전체 초기화는 모든 국가 데이터를 지웁니다. 되돌리기 어려우니 신중히 눌러 주세요.</li>
                <li>
                  여기서 지우는 것은 재고 파일·집계 쪽입니다. SKU 관리에 등록한 상품 매핑은 이 탭만으로는 지워지지 않습니다.
                </li>
              </ul>
            </div>

            <div className="cautionSection">
              <div className="cautionSectionTitle">SKU 관리</div>
              <ul className="cautionList">
                <li>
                  SKU 매핑 엑셀 템플릿(ZIP)으로 한국·미국·대만·홍콩 양식을 받을 수 있습니다. ZIP 안의 한국 파일에 구분
                  열이 포함됩니다. 하나의 파일 안에 여러 개의 시트를 읽을 수는 없으니, 되도록 시트를 더 추가하지는 말아
                  주세요.
                </li>
                <li>여러 개의 엑셀을 한 번에 선택해 올릴 수 있습니다.</li>
                <li>기존에 등록되어 있는 상품에 대한 정보가 업로드되면 기존 정보에서 갱신합니다.</li>
                <li>
                  같은 나라·같은 SKU가 서로 다른 상품으로 두 번 정의되면 그번 업로드 전체가 반영되지 않을 수 있습니다.
                </li>
                <li>엑셀 원본 파일은 보관하지 않고, 읽은 매핑 정보만 저장합니다.</li>
                <li>파일이 어렵다면 수기 작성 탭에서 직접 넣을 수 있습니다.</li>
              </ul>
            </div>
          </div>
        </div>
      )}

      {isSettingsScope && (
        <section className="settingsPane settingsCard">
          <div
            className={`settingsToolbar ${
              fileEntries.length === 0 && shipmentFileEntries.length === 0 ? "settingsToolbarWithNotice" : ""
            }`}
          >
            <button
              className="ghost settingsDangerButton"
              disabled={settingsMutating || (fileEntries.length === 0 && shipmentFileEntries.length === 0)}
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
              <details key={country} className="settingsGroup">
                <summary className="settingsHeader">
                  <span className="settingsHeaderLead">
                    <span className="settingsFileThumb" aria-hidden="true">
                      <FileText className="settingsFileThumbIcon" size={20} strokeWidth={2} />
                    </span>
                    <span>{countryLabel(country)}</span>
                  </span>
                  <span className="settingsHeaderRight">
                    <span>{formatInt(entries.length)}개 파일</span>
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
                      {country === "SHIPMENT" ? (
                        <div className="fileControl fileControlMuted">
                          <span>출고 엑셀(시트·행 단위 일자)</span>
                        </div>
                      ) : (
                        <div className="fileControl">
                          <span>날짜</span>
                          <input
                            type="date"
                            title={DATE_TEXT_INPUT_HINT}
                            value={toDateInputValue(entry.date)}
                            disabled={settingsMutating || !entry.dbFileId}
                            onChange={(e) => {
                              const next = e.target.value;
                              if (entry.dbFileId) {
                                void patchInventoryFileBaseDate(entry, next);
                              } else {
                                setFileEntries((prev) =>
                                  prev.map((x) => (x.id === entry.id ? { ...x, date: next } : x))
                                );
                              }
                            }}
                          />
                        </div>
                      )}
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
                <div className="skuUploadStage skuManageSinglePanel poOrderFileUploadStage">
                  <div className="settingsNotice poOrderFileUploadNotice">
                    <p className="poOrderFileUploadLead">
                      <strong>
                        한국·미국·대만·홍콩용 엑셀 템플릿을 받아 작성한 뒤, 파일들을 한꺼번에 업로드할 수 있습니다.
                      </strong>
                    </p>
                  </div>
                  <div className="poOrderFileTemplateRow">
                    <button
                      type="button"
                      className="poOrderFileTemplateBtn"
                      disabled={settingsMutating}
                      onClick={() => {
                        downloadSkuMappingCountryTemplatesZip().catch(() =>
                          window.alert("SKU 매핑 템플릿을 만드는 중 오류가 났습니다.")
                        );
                      }}
                    >
                      SKU 매핑 엑셀 템플릿 (ZIP)
                    </button>
                  </div>
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
                        <span className="skuManualKrFieldHead">바코드</span>
                        <div className="skuManualKrFieldBody">
                          <input
                            type="text"
                            value={manualMappingForm.barcode}
                            onChange={(e) =>
                              setManualMappingForm((prev) => ({ ...prev, barcode: e.target.value }))
                            }
                            placeholder="EAN·UPC 등 (선택)"
                            autoComplete="off"
                          />
                        </div>
                      </label>
                      <label className="skuManualKrField">
                        <span className="skuManualKrFieldHead">구분</span>
                        <div className="skuManualKrFieldBody">
                          <input
                            type="text"
                            value={manualMappingForm.segment}
                            onChange={(e) =>
                              setManualMappingForm((prev) => ({ ...prev, segment: e.target.value }))
                            }
                            placeholder="실제 단종만: 단종 또는 (X) 단종 (선택)"
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
          <section className="productMappingCard">
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
                상품명·SKU로 검색하면 조건에 맞는 매핑 카드가 아래에 표시됩니다.
              </div>
              <div className="productMappingSearchRow">
                <div className="searchWrap productMappingSearchWrap">
                  <SearchFieldIcon className="searchIcon" size={16} strokeWidth={2} />
                  <input
                    className="searchInput productMappingSearchInput"
                    type="text"
                    value={mappingSearchKeyword}
                    onChange={(e) => setMappingSearchKeyword(e.target.value)}
                    placeholder="상품코드 또는 상품명 검색"
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
                {productMappingCards.map((row) => {
                  const mappingItemDiscontinued =
                    String(row.segment || "").trim() === "단종" || row.segment_display === "단종";
                  return (
                  <article key={row._id} className="productMappingItemCard">
                    <div
                      className="productMappingAlignGrid"
                      title={`${row.brand || "–"} | ${row.kr_name || "상품명 없음"}${row.segment ? ` | 구분 ${row.segment}` : ""}${row.barcode ? ` | 바코드 ${row.barcode}` : ""}`}
                    >
                      <div className="productMappingGridHeadBand">
                        <div className="productMappingItemBrandPart productMappingGridHeadBrand">
                          {row.brand || "–"}
                        </div>
                        <span className="productMappingPipe productMappingGridHeadPipe" aria-hidden="true">
                          |
                        </span>
                        <div className="productMappingItemNamePart productMappingGridHeadName">
                          {row.kr_name || "상품명 없음"}
                        </div>
                        {mappingItemDiscontinued ? (
                          <div
                            className="productMappingGridHeadMeta productMappingDiscontinuedBadge"
                            title="DB 구분: 단종"
                          >
                            <Ban className="productMappingDiscontinuedIcon" size={18} strokeWidth={2} aria-hidden />
                            <span>단종</span>
                          </div>
                        ) : null}
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
                  );
                })}
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

