"""
도서관 장서 수서 중복 체크 도우미
- 실행: streamlit run book_duplicate_checker.py
- 필요 패키지: pip install streamlit pandas openpyxl

[컬럼 규격]
알라딘 파일  : 서명 | 저자 | 출판사 | 정가 | 판매가 | ISBN13 | 출간일 | 분야
소장 목록    : 서명 | 저자 | 출판사 | 정가
"""

import io
import streamlit as st
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ────────────────────────────────────────────
# 컬럼 규격 (샘플 파일 기준 고정)
# ────────────────────────────────────────────
ALADDIN_COLS  = ["서명", "저자", "출판사", "정가", "판매가", "ISBN13", "출간일", "분야"]
LIBRARY_COLS  = ["서명", "저자", "출판사", "정가"]

# ────────────────────────────────────────────
# 페이지 설정
# ────────────────────────────────────────────
st.set_page_config(page_title="도서관 수서 중복 체크", page_icon="📚", layout="wide")
st.title("📚 도서관 장서 수서 중복 체크 도우미")
st.write("알라딘 구입 목록과 도서관 소장 목록을 대조합니다.")

# ────────────────────────────────────────────
# 유틸 함수
# ────────────────────────────────────────────
def clean_isbn(val) -> str:
    return str(val).strip().replace("-", "").replace(" ", "").removesuffix(".0")

def clean_text(val) -> str:
    return (
        str(val).strip().lower()
        .replace(" ", "").replace("-", "").replace(".", "")
        .replace(",", "").replace("·", "").replace("(", "").replace(")", "")
    )

def load_file(uploaded) -> pd.DataFrame | None:
    try:
        if uploaded.name.lower().endswith(".csv"):
            for enc in ["utf-8-sig", "cp949", "euc-kr", "utf-8"]:
                try:
                    uploaded.seek(0)
                    return pd.read_csv(uploaded, encoding=enc, dtype=str)
                except UnicodeDecodeError:
                    continue
            st.error("CSV 인코딩을 인식할 수 없습니다.")
            return None
        else:
            uploaded.seek(0)
            df = pd.read_excel(uploaded, dtype=str)
            return df.dropna(how="all").reset_index(drop=True)
    except Exception as e:
        st.error(f"파일 로드 오류: {e}")
        return None

def validate_columns(df: pd.DataFrame, required: list, file_label: str) -> bool:
    """필수 컬럼 존재 여부 검증. 없는 컬럼 목록을 에러로 표시."""
    missing = [c for c in required if c not in df.columns]
    if missing:
        st.error(
            f"**{file_label}** 파일에 아래 컬럼이 없습니다. "
            f"샘플 파일의 컬럼명을 확인해 주세요.\n\n"
            f"누락된 컬럼: `{'`, `'.join(missing)}`\n\n"
            f"현재 파일의 컬럼: `{'`, `'.join(df.columns.tolist())}`"
        )
        return False
    return True

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="수서중복체크결과")
        ws = writer.sheets["수서중복체크결과"]

        hdr_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        hdr_font = Font(bold=True, color="FFFFFF", size=10)
        thin = Side(style="thin")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for cell in ws[1]:
            cell.fill = hdr_fill
            cell.font = hdr_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
        ws.row_dimensions[1].height = 20

        dup_col_idx = next(
            (i for i, c in enumerate(ws[1], 1) if c.value == "중복여부"), None
        )
        yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            is_dup = dup_col_idx and "소장 중" in str(row[dup_col_idx - 1].value)
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(
                    horizontal="center" if cell.column == dup_col_idx else "left",
                    vertical="center",
                )
                if is_dup:
                    cell.fill = yellow

        for col in ws.columns:
            ltr = get_column_letter(col[0].column)
            w = max((len(str(c.value)) if c.value else 0) for c in col)
            ws.column_dimensions[ltr].width = max(10, min(w + 3, 50))

        ws.freeze_panes = "A2"
    return output.getvalue()

# ────────────────────────────────────────────
# 샘플 파일 안내
# ────────────────────────────────────────────
with st.expander("📋 업로드 파일 규격 및 샘플 다운로드", expanded=False):
    sc1, sc2 = st.columns(2)
    with sc1:
        st.markdown("**① 알라딘 구입 예정 목록**")
        st.markdown(
            "| 컬럼명 | 예시 |\n|--------|------|\n"
            "| 서명 | 파친코 |\n"
            "| 저자 | 이민진 |\n"
            "| 출판사 | 문학사상 |\n"
            "| 정가 | 18000 |\n"
            "| 판매가 | 16200 |\n"
            "| ISBN13 | 9788970129235 |\n"
            "| 출간일 | 2022-03-10 |\n"
            "| 분야 | 소설 |"
        )
    with sc2:
        st.markdown("**② 소장 목록**")
        st.markdown(
            "| 컬럼명 | 예시 |\n|--------|------|\n"
            "| 서명 | 채식주의자 |\n"
            "| 저자 | 한강 |\n"
            "| 출판사 | 창비 |\n"
            "| 정가 | 13000 |"
        )
    st.info("💡 컬럼명이 정확히 일치해야 합니다. 샘플 파일을 참고해 주세요.")

# ────────────────────────────────────────────
# 파일 업로드
# ────────────────────────────────────────────
st.divider()
col1, col2 = st.columns(2)
with col1:
    st.markdown("**① 알라딘 구입 예정 목록**")
    planned_file = st.file_uploader(
        "엑셀 파일 (.xlsx)",
        type=["xlsx"],
        key="planned",
        help="서명·저자·출판사·정가·판매가·ISBN13·출간일·분야 컬럼 필수",
    )
with col2:
    st.markdown("**② 도서관 현재 소장 목록**")
    library_file = st.file_uploader(
        "엑셀 또는 CSV (.xlsx / .csv)",
        type=["xlsx", "csv"],
        key="library",
        help="서명·저자·출판사·정가 컬럼 필수",
    )

# ────────────────────────────────────────────
# 두 파일이 모두 올라왔을 때
# ────────────────────────────────────────────
if planned_file and library_file:

    df_planned = load_file(planned_file)
    df_library = load_file(library_file)

    if df_planned is None or df_library is None:
        st.stop()

    # ── 컬럼 검증 (필수 컬럼 없으면 에러 표시 후 중단)
    ok_a = validate_columns(df_planned, ALADDIN_COLS,  "알라딘 구입 예정 목록")
    ok_l = validate_columns(df_library, LIBRARY_COLS,  "소장 목록")
    if not ok_a or not ok_l:
        st.stop()

    st.success(
        f"✅ 구입 예정 목록 **{len(df_planned):,}건** | 소장 목록 **{len(df_library):,}건** 로드 완료"
    )

    # ────────────────────────────────────────────
    # 대조 로직
    # 1단계: 서명 일치 확인
    # 2단계: 동명이서 의심 시 출판사까지 확인
    # ────────────────────────────────────────────

    # 소장 목록 — 서명별 출판사 목록 딕셔너리 { 정규화된_서명: [출판사, ...] }
    from collections import defaultdict
    title_to_publishers = defaultdict(list)
    for _, lib_row in df_library.iterrows():
        t = clean_text(lib_row["서명"])
        p = clean_text(lib_row["출판사"])
        if t and t != "nan":
            title_to_publishers[t].append(p)

    def check_row(row) -> tuple:
        title     = clean_text(row["서명"])
        publisher = clean_text(row["출판사"])

        if not title or title == "nan":
            return "⚠️ 확인 필요", "서명 정보 없음"

        matched_publishers = title_to_publishers.get(title)

        # 1단계: 서명 일치하는 책 없음 → 신규
        if not matched_publishers:
            return "✅ 신규(구입 가능)", "서명 불일치"

        # 서명 일치하는 책이 1권 → 소장 중
        if len(matched_publishers) == 1:
            return "❌ 소장 중(중복)", "서명 일치"

        # 2단계: 서명 일치 2권 이상 → 출판사까지 비교
        if publisher in matched_publishers:
            return "❌ 소장 중(중복)", "서명+출판사 일치"

        return "✅ 신규(구입 가능)", "동명이서(출판사 불일치)"

    results = df_planned.apply(check_row, axis=1, result_type="expand")
    df_planned["중복여부"] = results[0]
    df_planned["대조근거"] = results[1]

    # 출력 컬럼 순서: 중복여부·대조근거를 앞에, 나머지 원본 컬럼 전체 표시
    extra_cols = [c for c in df_planned.columns if c not in ("중복여부", "대조근거", "_isbn_clean")]
    output_cols = ["중복여부", "대조근거"] + extra_cols
    final_df = df_planned[output_cols]

    # ────────────────────────────────────────────
    # 결과 요약
    # ────────────────────────────────────────────
    total    = len(final_df)
    new_cnt  = int((final_df["중복여부"] == "✅ 신규(구입 가능)").sum())
    dup_cnt  = int((final_df["중복여부"] == "❌ 소장 중(중복)").sum())
    warn_cnt = int((final_df["중복여부"] == "⚠️ 확인 필요").sum())

    st.divider()
    st.subheader("📊 대조 결과 요약")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("전체",            f"{total:,}권")
    m2.metric("🟢 신규(구입 가능)", f"{new_cnt:,}권")
    m3.metric("🔴 소장 중(중복)",   f"{dup_cnt:,}권")
    m4.metric("⚠️ 확인 필요",      f"{warn_cnt:,}권")

    # ────────────────────────────────────────────
    # 결과 탭
    # ────────────────────────────────────────────
    tab1, tab2, tab3, tab4 = st.tabs([
        f"전체 ({total:,}권)",
        f"🟢 신규 ({new_cnt:,}권)",
        f"🔴 소장 중 ({dup_cnt:,}권)",
        f"⚠️ 확인 필요 ({warn_cnt:,}권)",
    ])

    def highlight(df):
        return df.style.apply(
            lambda row: [
                "background-color:#fff2cc;" if "소장 중" in str(row["중복여부"])
                else "background-color:#fce4d6;" if "확인 필요" in str(row["중복여부"])
                else ""
                for _ in row
            ], axis=1,
        )

    with tab1:
        st.dataframe(highlight(final_df), use_container_width=True, height=400)
    with tab2:
        d = final_df[final_df["중복여부"] == "✅ 신규(구입 가능)"]
        st.dataframe(d, use_container_width=True, height=400) if not d.empty else st.info("신규 도서가 없습니다.")
    with tab3:
        d = final_df[final_df["중복여부"] == "❌ 소장 중(중복)"]
        st.dataframe(d, use_container_width=True, height=400) if not d.empty else st.info("중복 도서가 없습니다.")
    with tab4:
        d = final_df[final_df["중복여부"] == "⚠️ 확인 필요"]
        if not d.empty:
            st.warning("서명 정보가 없어 자동 대조가 불가합니다. 수동으로 확인하세요.")
            st.dataframe(d, use_container_width=True, height=400)
        else:
            st.info("확인 필요 항목이 없습니다.")

    # ────────────────────────────────────────────
    # 다운로드
    # ────────────────────────────────────────────
    st.divider()
    st.subheader("💾 결과 다운로드")
    d1, d2, d3 = st.columns(3)
    with d1:
        st.download_button(
            "📥 전체 결과",
            data=to_excel_bytes(final_df),
            file_name="수서_중복체크_전체.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with d2:
        st.download_button(
            "🟢 신규 목록만",
            data=to_excel_bytes(final_df[final_df["중복여부"] == "✅ 신규(구입 가능)"]),
            file_name="수서_중복체크_신규.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with d3:
        st.download_button(
            "🔴 소장중 목록만",
            data=to_excel_bytes(final_df[final_df["중복여부"] == "❌ 소장 중(중복)"]),
            file_name="수서_중복체크_소장중.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

else:
    st.info("👆 두 파일을 모두 업로드하면 대조를 시작합니다.")
