"""
도서관 장서 수서 중복 체크 도우미
- 실행: streamlit run book_duplicate_checker.py
- 필요 패키지: pip install streamlit pandas openpyxl
"""

import io
import streamlit as st
import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

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
    """ISBN에서 하이픈·공백 제거, 소수점(.0) 제거."""
    return (
        str(val)
        .strip()
        .replace("-", "")
        .replace(" ", "")
        .removesuffix(".0")  # 엑셀이 숫자로 읽을 경우 대응
    )

def load_file(uploaded) -> pd.DataFrame | None:
    """xlsx / csv 자동 판별 로드. CSV는 인코딩 자동 감지."""
    try:
        if uploaded.name.lower().endswith(".csv"):
            for enc in ["utf-8-sig", "cp949", "euc-kr", "utf-8"]:
                try:
                    uploaded.seek(0)
                    return pd.read_csv(uploaded, encoding=enc, dtype=str)
                except UnicodeDecodeError:
                    continue
            st.error("CSV 인코딩을 인식할 수 없습니다. UTF-8 또는 EUC-KR로 저장 후 다시 시도하세요.")
            return None
        else:
            uploaded.seek(0)
            df = pd.read_excel(uploaded, dtype=str)
            df = df.dropna(how="all").reset_index(drop=True)
            return df
    except Exception as e:
        st.error(f"파일 로드 오류: {e}")
        return None

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """결과 DataFrame → 서식 있는 엑셀 바이트 (메모리 저장, 파일 저장 없음)."""
    output = io.BytesIO()  # ✅ 파일 저장 대신 메모리 버퍼 사용

    with pd.ExcelWriter(output, engine="openpyxl") as writer:  # ✅ with 문으로 자동 close
        df.to_excel(writer, index=False, sheet_name="수서중복체크결과")
        ws = writer.sheets["수서중복체크결과"]

        # 헤더 스타일
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

        # 비고 컬럼 위치 탐색
        note_col_idx = None
        for i, cell in enumerate(ws[1], 1):
            if cell.value == "비고":
                note_col_idx = i
                break

        # 데이터 행: 소장중 노란 강조
        yellow = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            is_dup = note_col_idx and "소장 중" in str(row[note_col_idx - 1].value)
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(
                    horizontal="center" if cell.column == note_col_idx else "left",
                    vertical="center",
                )
                if is_dup:
                    cell.fill = yellow

        # 열 너비 자동
        for col in ws.columns:
            ltr = get_column_letter(col[0].column)
            w = max((len(str(c.value)) if c.value else 0) for c in col)
            ws.column_dimensions[ltr].width = max(10, min(w + 3, 50))

        ws.freeze_panes = "A2"

    return output.getvalue()

# ────────────────────────────────────────────
# 파일 업로드
# ────────────────────────────────────────────
col1, col2 = st.columns(2)
with col1:
    st.markdown("**① 알라딘 구입 예정 목록**")
    planned_file = st.file_uploader(
        "엑셀 파일 (.xlsx)",
        type=["xlsx"],
        key="planned",
        help="알라딘에서 내려받은 파일 (도서명·저자·출판사·정가·판매가·ISBN13·출간일·주제/분야·비고)",
    )
with col2:
    st.markdown("**② 도서관 현재 소장 목록**")
    library_file = st.file_uploader(
        "엑셀 또는 CSV (.xlsx / .csv)",
        type=["xlsx", "csv"],
        key="library",
        help="DLS 등 도서관 관리 프로그램에서 내보낸 파일 (ISBN 컬럼 필수)",
    )

# ────────────────────────────────────────────
# 두 파일이 모두 올라왔을 때
# ────────────────────────────────────────────
if planned_file and library_file:

    df_planned = load_file(planned_file)
    df_library = load_file(library_file)

    if df_planned is None or df_library is None:
        st.stop()

    # ── 알라딘 ISBN13 컬럼 존재 확인
    if "ISBN13" not in df_planned.columns:
        st.error("알라딘 파일에 'ISBN13' 컬럼이 없습니다. 파일을 확인하세요.")
        st.stop()

    st.success(
        f"✅ 구입 예정 목록 **{len(df_planned):,}건** | 소장 목록 **{len(df_library):,}건** 로드 완료"
    )

    # ── 소장 목록 ISBN 컬럼 선택 (✅ selectbox 이후에만 df 접근)
    st.divider()
    lib_isbn_col = st.selectbox(
        "소장 목록 파일에서 **ISBN이 적힌 컬럼**을 선택하세요",
        options=df_library.columns.tolist(),
        index=next(
            (i for i, c in enumerate(df_library.columns) if "isbn" in c.lower()),
            0,
        ),
        help="컬럼명이 'ISBN', 'ISBN13', '바코드' 등으로 다를 수 있습니다.",
    )

    # ── 전처리
    df_library["_isbn_clean"] = df_library[lib_isbn_col].apply(clean_isbn)
    df_planned["_isbn_clean"] = df_planned["ISBN13"].apply(clean_isbn)

    existing_isbns = set(df_library["_isbn_clean"].unique()) - {"", "nan"}

    # ── 중복 체크
    def check_duplicate(isbn: str) -> str:
        if not isbn or isbn == "nan":
            return "⚠️ ISBN 없음"
        return "❌ 소장 중(중복)" if isbn in existing_isbns else "✅ 신규(구입 가능)"

    # ✅ 원본 '비고' 컬럼을 덮어쓰지 않고 '중복여부' 컬럼을 새로 추가
    df_planned["중복여부"] = df_planned["_isbn_clean"].apply(check_duplicate)

    # 출력용 컬럼 목록 (원본 컬럼 중 존재하는 것만 포함 → KeyError 방지)
    desired = ["도서명", "저자", "출판사", "정가", "판매가", "ISBN13", "출간일", "주제/분야", "비고"]
    output_cols = ["중복여부"] + [c for c in desired if c in df_planned.columns]
    final_df = df_planned[output_cols]

    # ── 요약 통계
    total = len(final_df)
    new_cnt = int((final_df["중복여부"] == "✅ 신규(구입 가능)").sum())
    dup_cnt = int((final_df["중복여부"] == "❌ 소장 중(중복)").sum())
    no_isbn = int((final_df["중복여부"] == "⚠️ ISBN 없음").sum())

    st.divider()
    st.subheader("📊 대조 결과 요약")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("전체", f"{total:,}권")
    m2.metric("🟢 신규(구입 가능)", f"{new_cnt:,}권")
    m3.metric("🔴 소장 중(중복)", f"{dup_cnt:,}권")
    m4.metric("⚠️ ISBN 없음", f"{no_isbn:,}권")

    # ── 결과 탭
    tab1, tab2, tab3 = st.tabs([
        f"전체 ({total:,}권)",
        f"🟢 신규 ({new_cnt:,}권)",
        f"🔴 소장 중 ({dup_cnt:,}권)",
    ])

    def highlight(df):
        return df.style.apply(
            lambda row: [
                "background-color:#fff2cc;" if "소장 중" in str(row["중복여부"]) else ""
                for _ in row
            ],
            axis=1,
        )

    with tab1:
        st.dataframe(highlight(final_df), use_container_width=True, height=400)
    with tab2:
        new_df = final_df[final_df["중복여부"] == "✅ 신규(구입 가능)"]
        st.dataframe(new_df, use_container_width=True, height=400) if not new_df.empty else st.info("신규 도서가 없습니다.")
    with tab3:
        dup_df = final_df[final_df["중복여부"] == "❌ 소장 중(중복)"]
        st.dataframe(dup_df, use_container_width=True, height=400) if not dup_df.empty else st.info("중복 도서가 없습니다.")

    # ── 다운로드
    st.divider()
    st.subheader("💾 결과 다운로드")
    d1, d2, d3 = st.columns(3)

    with d1:
        st.download_button(
            "📥 전체 결과",
            data=to_excel_bytes(final_df),               # ✅ BytesIO 메모리 변환
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
    with st.expander("📋 사용 방법"):
        st.markdown("""
1. **알라딘 구입 예정 목록** 엑셀 파일을 업로드하세요.
2. **도서관 소장 목록** 파일을 업로드하세요 (xlsx 또는 csv).
3. **ISBN 컬럼**을 선택하세요 (자동 감지되지만 확인 필요).
4. 결과를 확인하고 엑셀로 다운로드하세요.

> 💡 소장 목록의 ISBN 컬럼명이 'ISBN', 'ISBN13', '바코드' 등으로 다를 수 있으니 꼭 확인하세요.
        """)
