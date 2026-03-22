import io

import pandas as pd
from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import Response, StreamingResponse

from app.domains.reorder.dto.reorder_api_dto import ReorderPlanResponse
from app.domains.reorder.dto.reorder_dto import ReorderSettings
from app.domains.reorder.repositories.reorder_excel_repository import (
    create_dummy_data,
    to_excel_bytes,
)
from app.domains.reorder.services.reorder_service import calculate_reorder_plan
from app.shared.validation.reorder_validator import validate_input_frames

router = APIRouter(prefix="/api/reorder", tags=["reorder"])

SHEET_ALIASES = {
    "sales": {"sales", "sale", "판매", "판매데이터", "판매내역", "매출", "출고"},
    "inventory": {"inventory", "stock", "재고", "재고현황", "재고데이터"},
    "sku_policy": {"sku_policy", "policy", "master", "상품정책", "sku정책", "정책"},
}

COLUMN_ALIASES = {
    "sales": {
        "date": {"date", "일자", "날짜", "판매일", "기준일"},
        "brand": {"brand", "브랜드", "브랜드명"},
        "sku": {"sku", "상품코드", "품목코드", "제품코드", "itemcode", "코드"},
        "channel": {"channel", "채널", "플랫폼", "판매처", "몰"},
        "qty_sold": {
            "qty_sold",
            "판매수량",
            "판매량",
            "수량",
            "주문수량",
            "출고수량",
            "판매qty",
        },
    },
    "inventory": {
        "date": {"date", "일자", "날짜", "기준일"},
        "sku": {"sku", "상품코드", "품목코드", "제품코드", "itemcode", "코드"},
        "on_hand": {"on_hand", "현재고", "재고", "가용재고", "실재고"},
        "inbound_qty": {"inbound_qty", "입고예정", "입고예정수량", "발주잔량", "미입고수량"},
    },
    "sku_policy": {
        "sku": {"sku", "상품코드", "품목코드", "제품코드", "itemcode", "코드"},
        "lead_time_days": {"lead_time_days", "리드타임", "조달리드타임", "입고소요일", "leadtime"},
        "moq": {"moq", "최소주문수량", "최소발주수량", "최소수량"},
        "pack_size": {"pack_size", "박스단위", "포장단위", "입수"},
        "safety_stock": {"safety_stock", "안전재고", "버퍼재고"},
    },
}


def _normalize_key(value: str) -> str:
    return (
        str(value)
        .strip()
        .lower()
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
        .replace("(", "")
        .replace(")", "")
    )


def _match_sheet_name(raw_name: str) -> str | None:
    norm = _normalize_key(raw_name)
    for target, aliases in SHEET_ALIASES.items():
        if norm in {_normalize_key(alias) for alias in aliases}:
            return target
    return None


def _standardize_columns(df: pd.DataFrame, sheet_type: str) -> pd.DataFrame:
    alias_map = COLUMN_ALIASES[sheet_type]
    normalized_columns = {_normalize_key(col): col for col in df.columns}
    rename_map: dict[str, str] = {}

    for canonical, aliases in alias_map.items():
        candidates = {_normalize_key(canonical)} | {_normalize_key(alias) for alias in aliases}
        for candidate in candidates:
            original_col = normalized_columns.get(candidate)
            if original_col:
                rename_map[original_col] = canonical
                break

    standardized = df.rename(columns=rename_map)
    return standardized


def _extract_standard_sheets(excel_buffer: io.BytesIO) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    try:
        all_sheets = pd.read_excel(excel_buffer, sheet_name=None)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"엑셀 파싱 실패: {exc}") from exc

    if not all_sheets:
        raise HTTPException(status_code=400, detail="시트가 없는 엑셀 파일입니다.")

    matched: dict[str, pd.DataFrame] = {}
    unmatched: list[pd.DataFrame] = []

    for sheet_name, df in all_sheets.items():
        matched_name = _match_sheet_name(sheet_name)
        if matched_name:
            matched[matched_name] = _standardize_columns(df, matched_name)
        else:
            unmatched.append(df)

    # 시트명이 표준이 아닐 때도 컬럼 매칭률로 보정
    for sheet_type in ("sales", "inventory", "sku_policy"):
        if sheet_type in matched:
            continue
        best_df = None
        best_idx = -1
        best_score = -1
        required = set(COLUMN_ALIASES[sheet_type].keys())

        for idx, df in enumerate(unmatched):
            candidate = _standardize_columns(df, sheet_type)
            score = len(required.intersection(candidate.columns))
            if score > best_score:
                best_score = score
                best_df = candidate
                best_idx = idx

        # 최소한 절반 이상의 필수 컬럼이 매칭될 때만 자동 배정
        min_score = max(1, len(required) // 2)
        if best_df is not None and best_score >= min_score:
            matched[sheet_type] = best_df
            unmatched.pop(best_idx)

    missing_sheets = [key for key in ("sales", "inventory", "sku_policy") if key not in matched]
    if missing_sheets:
        existing_sheet_names = list(all_sheets.keys())
        sheet_hint = {
            key: sorted(list(SHEET_ALIASES[key])) for key in missing_sheets
        }
        raise HTTPException(
            status_code=400,
            detail=[
                f"필수 시트를 찾지 못했습니다: {missing_sheets}",
                f"업로드된 시트명: {existing_sheet_names}",
                f"인식 가능한 시트명 예시: {sheet_hint}",
            ],
        )

    return matched["sales"], matched["inventory"], matched["sku_policy"]


def _load_frames(
    use_dummy: bool, upload_file: UploadFile | None
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if use_dummy:
        return create_dummy_data()

    if upload_file is None:
        raise HTTPException(status_code=400, detail="엑셀 파일이 필요합니다.")

    raw = upload_file.file.read()
    if not raw:
        raise HTTPException(status_code=400, detail="업로드 파일이 비어 있습니다.")

    excel_buffer = io.BytesIO(raw)
    return _extract_standard_sheets(excel_buffer)


@router.get("/health")
def health_check() -> dict[str, str]:
    return {"status": "ok"}


@router.get("/template")
def download_template() -> Response:
    sales_df, inventory_df, policy_df = create_dummy_data()
    payload = to_excel_bytes(sales_df.head(0), inventory_df.head(0), policy_df.head(0))
    return Response(
        content=payload,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="inventory_template.xlsx"'},
    )


@router.post("/plan", response_model=ReorderPlanResponse)
def calculate_plan(
    file: UploadFile | None = File(default=None),
    use_dummy: bool = Form(default=True),
    w7: float = Form(default=0.5),
    w14: float = Form(default=0.3),
    w28: float = Form(default=0.2),
    target_cover_days: int = Form(default=21),
) -> ReorderPlanResponse:
    if any(x < 0 or x > 1 for x in (w7, w14, w28)):
        raise HTTPException(status_code=400, detail="가중치는 0~1 범위여야 합니다.")
    if target_cover_days < 1 or target_cover_days > 120:
        raise HTTPException(status_code=400, detail="커버 일수는 1~120 범위여야 합니다.")

    sales_df, inventory_df, policy_df = _load_frames(use_dummy=use_dummy, upload_file=file)
    errors = validate_input_frames(sales_df, inventory_df, policy_df)
    if errors:
        raise HTTPException(status_code=400, detail=errors)

    settings = ReorderSettings(
        w7=w7, w14=w14, w28=w28, target_cover_days=target_cover_days
    )
    result_df = calculate_reorder_plan(
        sales_df=sales_df,
        inventory_df=inventory_df,
        policy_df=policy_df,
        settings=settings,
    )

    rows = result_df.replace({float("inf"): None}).to_dict(orient="records")
    summary = {
        "total_sku": len(result_df),
        "reorder_count": int(result_df["need_reorder"].sum()),
    }
    return ReorderPlanResponse(summary=summary, rows=rows)


@router.post("/plan.xlsx")
def calculate_plan_excel(
    file: UploadFile | None = File(default=None),
    use_dummy: bool = Form(default=True),
    w7: float = Form(default=0.5),
    w14: float = Form(default=0.3),
    w28: float = Form(default=0.2),
    target_cover_days: int = Form(default=21),
) -> StreamingResponse:
    sales_df, inventory_df, policy_df = _load_frames(use_dummy=use_dummy, upload_file=file)
    errors = validate_input_frames(sales_df, inventory_df, policy_df)
    if errors:
        raise HTTPException(status_code=400, detail=errors)

    settings = ReorderSettings(
        w7=w7, w14=w14, w28=w28, target_cover_days=target_cover_days
    )
    result_df = calculate_reorder_plan(
        sales_df=sales_df,
        inventory_df=inventory_df,
        policy_df=policy_df,
        settings=settings,
    )

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="reorder_plan", index=False)
    buffer.seek(0)

    return StreamingResponse(
        buffer,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": 'attachment; filename="reorder_plan.xlsx"'},
    )

