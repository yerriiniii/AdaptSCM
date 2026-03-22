import io

import pandas as pd
import streamlit as st

from app.domains.reorder.dto.reorder_dto import ReorderSettings
from app.domains.reorder.repositories.reorder_excel_repository import (
    create_dummy_data,
    read_uploaded_excel,
    to_excel_bytes,
)
from app.domains.reorder.services.reorder_service import calculate_reorder_plan
from app.shared.validation.reorder_validator import validate_input_frames


def render_reorder_web() -> None:
    st.set_page_config(page_title="재고 발주 MVP", layout="wide")
    st.title("재고 발주 의사결정 MVP")
    st.caption("엑셀 업로드 또는 더미 데이터로 발주 시점/발주량 추천을 자동 계산합니다.")

    with st.sidebar:
        st.header("가중치 / 설정")
        w7 = st.slider("최근 7일 가중치", 0.0, 1.0, 0.5, 0.05)
        w14 = st.slider("최근 14일 가중치", 0.0, 1.0, 0.3, 0.05)
        w28 = st.slider("최근 28일 가중치", 0.0, 1.0, 0.2, 0.05)
        target_cover_days = st.number_input(
            "목표 재고 커버 일수", min_value=1, max_value=120, value=21
        )

        weight_sum = round(w7 + w14 + w28, 4)
        if abs(weight_sum - 1.0) > 1e-9:
            st.warning(f"가중치 합이 {weight_sum} 입니다. 계산은 진행되지만 1.0 권장.")

    st.subheader("1) 템플릿 다운로드")
    tmpl_sales, tmpl_inventory, tmpl_policy = create_dummy_data()
    st.download_button(
        label="샘플 엑셀 템플릿 다운로드",
        data=to_excel_bytes(
            tmpl_sales.head(0), tmpl_inventory.head(0), tmpl_policy.head(0)
        ),
        file_name="inventory_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.subheader("2) 데이터 입력")
    col1, col2 = st.columns([2, 1])
    with col1:
        uploaded = st.file_uploader("엑셀 업로드 (.xlsx)", type=["xlsx"])
    with col2:
        use_dummy = st.checkbox("더미 데이터 사용", value=True)

    sales_df = inventory_df = policy_df = None

    if uploaded is not None:
        try:
            sales_df, inventory_df, policy_df = read_uploaded_excel(uploaded)
            st.success("엑셀 로드 완료")
        except Exception as exc:
            st.error(f"엑셀 로드 실패: {exc}")
            return
    elif use_dummy:
        sales_df, inventory_df, policy_df = create_dummy_data()
        st.info("더미 데이터를 사용 중입니다.")
    else:
        st.stop()

    errors = validate_input_frames(sales_df, inventory_df, policy_df)
    if errors:
        st.error("입력 데이터 검증 실패")
        for err in errors:
            st.write(f"- {err}")
        st.stop()

    settings = ReorderSettings(
        w7=w7,
        w14=w14,
        w28=w28,
        target_cover_days=int(target_cover_days),
    )

    st.subheader("3) 발주 추천 계산")
    result_df = calculate_reorder_plan(
        sales_df=sales_df,
        inventory_df=inventory_df,
        policy_df=policy_df,
        settings=settings,
    )

    total_sku = len(result_df)
    reorder_count = int(result_df["need_reorder"].sum())
    st.metric("발주 필요 SKU 수", f"{reorder_count} / {total_sku}")

    show_only_reorder = st.checkbox("발주 필요 SKU만 보기", value=True)
    display_df = result_df[result_df["need_reorder"]] if show_only_reorder else result_df

    st.dataframe(
        display_df.style.format(
            {
                "weighted_daily_demand": "{:.2f}",
                "avg7": "{:.2f}",
                "avg14": "{:.2f}",
                "avg28": "{:.2f}",
                "days_of_supply": "{:.1f}",
                "reorder_point": "{:.1f}",
                "target_stock": "{:.1f}",
                "raw_order_qty": "{:.1f}",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine="openpyxl") as writer:
        result_df.to_excel(writer, sheet_name="reorder_plan", index=False)
    output_buffer.seek(0)

    st.download_button(
        label="발주 추천 결과 다운로드",
        data=output_buffer.read(),
        file_name="reorder_plan.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    with st.expander("계산식 설명"):
        st.markdown(
            """
            - `가중 일평균 판매량 = avg7*w7 + avg14*w14 + avg28*w28`
            - `재주문점(ROP) = 가중 일평균 판매량 * 리드타임 + 안전재고`
            - `목표재고 = 가중 일평균 판매량 * 목표 커버일수 + 안전재고`
            - `원시 발주량 = 목표재고 - (현재고 + 입고예정)`
            - `최종 발주량 = MOQ/박스단위 제약 반영 후 올림`
            """
        )

