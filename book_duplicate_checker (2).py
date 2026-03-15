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
        .removesuffix(".0")
    )

def clean_text(val) -> str:
    """서명·저자 비교용 정규화: 소문자, 공백·특수문자 제거."""
    return (
        str(val)
        .strip()
        .lower()
        .replace(" ", "")
        .replace("-", "")
        .replace(".", "")
        .replace(",", "")
        .replace("·", "")
        .replace("(", "")
        .replace(")", "")
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
    """결과 DataFrame → 서식 있는 엑셀 바이트 (메모리 저장)."""
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

        dup_col_idx = None
        for i, cell in enumerate(ws[1], 1):
            if cell.value == "중복여부":
                dup_col_idx = i
                break

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

def auto_find(columns: list, keywords: list) -> int:
    """키워드로 컬럼 인덱스 자동 탐색. 없으면 0 반환."""
    for i, col in enumerate(columns):
        col_norm = col.strip().lower().replace(" ", "")
        for kw in keywords:
            if kw in col_norm:
                return i
    return 0

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
        help="DLS 등 도서관 관리 프로그램에서 내보낸 파일",
    )

# ────────────────────────────────────────────
# 두 파일이 모두 올라왔을 때
# ────────────────────────────────────────────
if planned_file and library_file:

    df_planned = load_file(planned_file)
    df_library = load_file(library_file)

    if df_planned is None or df_library is None:
        st.stop()

    # 알라딘 ISBN13 컬럼 자동 탐색 (컬럼명이 조금 달라도 대응)
    planned_isbn_col = next(
        (c for c in df_planned.columns if "isbn" in c.lower()),
        None,
    )
    if planned_isbn_col and planned_isbn_col != "ISBN13":
        df_planned = df_planned.rename(columns={planned_isbn_col: "ISBN13"})
        st.info(f"알라딘 파일의 ISBN 컬럼 `{planned_isbn_col}`을 `ISBN13`으로 인식했습니다.")
    elif not planned_isbn_col:
        st.warning("알라딘 파일에 ISBN 컬럼이 없습니다. 서명·저자 기준으로만 대조합니다.")
        df_planned["ISBN13"] = ""  # 빈 컬럼 추가 — 이후 로직에서 ISBN 없음으로 처리됨

    st.success(
        f"✅ 구입 예정 목록 **{len(df_planned):,}건** | 소장 목록 **{len(df_library):,}건** 로드 완료"
    )

    # ────────────────────────────────────────────
    # 대조 방법 선택
    # ────────────────────────────────────────────
    st.divider()
    st.subheader("🔧 대조 방법 설정")

    NONE_LABEL = "(없음 — 이 기준 사용 안 함)"
    lib_cols_with_none = [NONE_LABEL] + df_library.columns.tolist()

    method_col, info_col = st.columns([2, 1])

    with method_col:
        # ── ISBN 컬럼 선택
        isbn_auto = auto_find(df_library.columns.tolist(), ["isbn13", "isbn", "바코드"])
        # ISBN 컬럼이 자동 감지되면 해당 컬럼, 아니면 NONE_LABEL 선택
        has_isbn_auto = any("isbn" in c.lower() or "바코드" in c.lower() for c in df_library.columns)
        isbn_default = isbn_auto + 1 if has_isbn_auto else 0  # 0 = NONE_LABEL

        lib_isbn_col_raw = st.selectbox(
            "소장 목록 — ISBN 컬럼",
            options=lib_cols_with_none,
            index=isbn_default,
            help="ISBN 컬럼이 있으면 선택하세요. 없으면 '(없음)'을 선택하고 아래 서명·저자로 대조하세요.",
        )
        lib_isbn_col = None if lib_isbn_col_raw == NONE_LABEL else lib_isbn_col_raw

        # ── 서명 컬럼 선택
        title_auto = auto_find(df_library.columns.tolist(), ["서명", "도서명", "제목", "title"])
        lib_title_col_raw = st.selectbox(
            "소장 목록 — 서명(도서명) 컬럼",
            options=lib_cols_with_none,
            index=title_auto + 1,
            help="ISBN이 없을 때 서명으로 대조합니다. ISBN이 있어도 보조 기준으로 활용됩니다.",
        )
        lib_title_col = None if lib_title_col_raw == NONE_LABEL else lib_title_col_raw

        # ── 저자 컬럼 선택
        author_auto = auto_find(df_library.columns.tolist(), ["저자", "지은이", "author"])
        lib_author_col_raw = st.selectbox(
            "소장 목록 — 저자 컬럼",
            options=lib_cols_with_none,
            index=author_auto + 1,
            help="서명만으로 동명이서 오탐이 우려될 때 저자를 함께 대조합니다.",
        )
        lib_author_col = None if lib_author_col_raw == NONE_LABEL else lib_author_col_raw

    with info_col:
        # 현재 설정 기준 안내
        st.markdown("**현재 대조 기준**")
        if lib_isbn_col:
            st.success(f"① ISBN 대조 → `{lib_isbn_col}`")
        else:
            st.warning("① ISBN 대조 → 사용 안 함")

        if lib_title_col and lib_author_col:
            st.info(f"② 서명+저자 대조\n→ `{lib_title_col}` + `{lib_author_col}`")
        elif lib_title_col:
            st.info(f"② 서명 대조 → `{lib_title_col}`\n(저자 없음 — 동명이서 주의)")
        else:
            st.warning("② 서명 대조 → 사용 안 함")

        if not lib_isbn_col and not lib_title_col:
            st.error("대조 기준이 없습니다.\nISBN 또는 서명 컬럼을 선택하세요.")

    if not lib_isbn_col and not lib_title_col:
        st.stop()

    # ────────────────────────────────────────────
    # 전처리 — 소장 목록 대조 세트 구성
    # ────────────────────────────────────────────

    # ISBN 세트
    existing_isbns: set[str] = set()
    if lib_isbn_col:
        existing_isbns = set(
            df_library[lib_isbn_col].apply(clean_isbn).unique()
        ) - {"", "nan"}

    # 서명+저자 세트
    existing_title_author: set[tuple[str, str]] = set()
    existing_title_only: set[str] = set()
    if lib_title_col and lib_author_col:
        existing_title_author = set(
            zip(
                df_library[lib_title_col].apply(clean_text),
                df_library[lib_author_col].apply(clean_text),
            )
        )
    elif lib_title_col:
        existing_title_only = set(df_library[lib_title_col].apply(clean_text)) - {"", "nan"}

    # 알라딘 ISBN 정규화
    df_planned["_isbn_clean"] = df_planned["ISBN13"].apply(clean_isbn)

    # ────────────────────────────────────────────
    # 중복 체크 로직
    # ────────────────────────────────────────────
    def check_row(row) -> tuple[str, str]:
        """
        반환: (중복여부 라벨, 대조 근거)
        우선순위: ISBN → 서명+저자 → 서명만
        """
        isbn = row["_isbn_clean"]
        title = clean_text(row.get("도서명", ""))
        author = clean_text(row.get("저자", ""))

        # ① ISBN 대조
        if lib_isbn_col and isbn and isbn != "nan":
            if isbn in existing_isbns:
                return "❌ 소장 중(중복)", "ISBN 일치"
            # ISBN이 있고 소장 목록에 없으면 → 신규 확정 (오탐 방지)
            return "✅ 신규(구입 가능)", "ISBN 불일치"

        # ② ISBN 없는 경우 → 서명+저자 또는 서명만으로 대조
        if lib_title_col and lib_author_col and existing_title_author:
            if (title, author) in existing_title_author:
                return "❌ 소장 중(중복)", "서명+저자 일치"
            return "✅ 신규(구입 가능)", "서명+저자 불일치"

        if lib_title_col and existing_title_only:
            if title in existing_title_only:
                return "❌ 소장 중(중복)", "서명 일치"
            return "✅ 신규(구입 가능)", "서명 불일치"

        return "⚠️ 확인 필요", "대조 불가(ISBN·서명 모두 없음)"

    results = df_planned.apply(check_row, axis=1, result_type="expand")
    df_planned["중복여부"] = results[0]
    df_planned["대조근거"] = results[1]

    # 출력 컬럼 정리
    desired = ["도서명", "저자", "출판사", "정가", "판매가", "ISBN13", "출간일", "주제/분야", "비고"]
    output_cols = ["중복여부", "대조근거"] + [c for c in desired if c in df_planned.columns]
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
    m1.metric("전체", f"{total:,}권")
    m2.metric("🟢 신규(구입 가능)", f"{new_cnt:,}권")
    m3.metric("🔴 소장 중(중복)", f"{dup_cnt:,}권")
    m4.metric("⚠️ 확인 필요", f"{warn_cnt:,}권")

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
            ],
            axis=1,
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
            st.warning("아래 도서는 ISBN도 없고 서명 컬럼도 미선택이어서 자동 대조가 불가합니다. 수동으로 확인하세요.")
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
    with st.expander("📋 사용 방법"):
        st.markdown("""
1. **알라딘 구입 예정 목록** 엑셀 파일을 업로드하세요.
2. **도서관 소장 목록** 파일을 업로드하세요 (xlsx 또는 csv).
3. 소장 목록의 **ISBN / 서명 / 저자 컬럼**을 선택하세요.
   - ISBN 컬럼이 있으면 → ISBN으로 대조 (가장 정확)
   - ISBN 컬럼이 없으면 → 서명+저자로 대조
4. 결과를 확인하고 엑셀로 다운로드하세요.

| 대조 결과 | 의미 |
|-----------|------|
| ✅ 신규(구입 가능) | 소장 목록에 없음 |
| ❌ 소장 중(중복) | 이미 소장 중 |
| ⚠️ 확인 필요 | 대조 기준 정보가 부족해 수동 확인 필요 |

> 💡 **대조 우선순위**: ISBN이 있으면 ISBN만으로 판단합니다.
> ISBN이 없는 행에 한해 서명·저자로 보완 대조합니다.
        """)
