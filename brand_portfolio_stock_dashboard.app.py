import io
import os
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="브랜드 및 포트폴리오별 재고 현황 대시보드",
    page_icon="📦",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
    div[data-testid="stMetric"] {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        padding: 14px 18px;
        border-radius: 10px;
    }
    .section-title {
        font-size: 1.2rem;
        font-weight: 700;
        color: #1e293b;
        margin-top: 1rem;
        margin-bottom: 0.5rem;
        border-left: 4px solid #2563eb;
        padding-left: 10px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

DEFAULT_MASTER_FILE = "브랜드_제품코드_마스터.xlsx"
FALLBACK_MASTER_FILE = "통합 문서1.xlsx"

PORTFOLIO_OPTIONS = {
    "Acnes": [
        "클렌징",
        "스팟패치",
        "바디케어",
        "포 맨",
        "스킨케어",
        "더마 스킨케어",
        "미분류(확인필요)",
    ],
    "Lipcare": [
        "립아이스 프루티",
        "딥 모이스트",
        "매직컬러 / 틴티드컬러 라인",
        "메디케이티드",
        "소프트립스",
        "미분류(확인필요)",
    ],
    "HADALABO": [
        "Gokujyun (고쿠쥰)",
        "Shirojyun (시로쥰)",
        "Kiwamizu (키와미즈)",
        "Gokujyun premium (고쿠쥰 프리미엄)",
        "Tamagohada (타마고하다)",
        "미분류(확인필요)",
    ],
}

BRAND_CHOICES = [
    "Acnes",
    "Lipcare",
    "HADALABO",
    "Lotion",
    "Spray",
    "Roll on Lotion",
    "Melano",
    "Hand",
    "Hot&Cool",
    "부자재",
    "사은품",
    "미분류",
]


def classify_portfolio(brand: str, product_name: str) -> tuple[str, str, str]:
    name = str(product_name or "").strip()
    name_clean = re.sub(r"\s+", "", name).lower()

    if brand == "Acnes":
        if "포맨" in name_clean or "formen" in name_clean:
            return "포 맨", "✅ 정상매칭", "이미지(포 맨) 라인 일치"

        if "더마릴리프" in name_clean or "더마" in name_clean:
            if any(k in name_clean for k in ["클렌저", "워시", "폼"]):
                return "클렌징", "✅ 정상매칭", "이미지(클렌징-더마릴리프 모이스처 폼클렌저) 일치"
            if any(k in name_clean for k in ["리커버리", "밸런싱", "크림", "스킨"]):
                return "더마 스킨케어", "✅ 정상매칭", "이미지(더마 스킨케어) 라인 일치"
            return "미분류(확인필요)", "⚠️ 확인필요", "더마 라인 세부 카테고리 모호"

        if any(k in name_clean for k in ["바디워시", "바디미스트", "바디로션", "바디"]):
            if "바디로션" in name_clean:
                return "바디케어", "⚠️ 확인필요", "바디케어 추정(이미지 외 품목: 바디로션)"
            return "바디케어", "✅ 정상매칭", "이미지(바디케어) 라인 일치"

        if any(k in name_clean for k in ["스팟패치", "진정패치", "패치"]):
            if "토너패드" in name_clean or "시카패드" in name_clean or "올케어" in name_clean:
                if "토너패드" in name_clean:
                    return "스킨케어", "✅ 정상매칭", "이미지(스킨케어-3초진정 토너패드) 일치"
                return "스킨케어", "⚠️ 확인필요", "스킨케어 패드 추정(이미지 외 품목)"
            if "크림패치" in name_clean:
                return "스킨케어", "✅ 정상매칭", "이미지(스킨케어-스팟클리어 크림패치) 일치"
            if "화이트티트리" in name_clean:
                return "스팟패치", "⚠️ 확인필요", "스팟패치 추정(이미지 외 신규 라인: 화이트 티트리)"
            return "스팟패치", "✅ 정상매칭", "이미지(스팟패치) 라인 일치"

        if any(
            k in name_clean
            for k in [
                "세범밸런싱",
                "모이스처플루이드",
                "모이스처크림",
                "토너패드",
                "마스크",
                "선크림",
                "토너",
                "스팟크림",
            ]
        ):
            if "화이트티트리" in name_clean or "수딩마스크" in name_clean:
                return "스킨케어", "⚠️ 확인필요", "스킨케어 추정(이미지 외 품목/리뉴얼명)"
            return "스킨케어", "✅ 정상매칭", "이미지(스킨케어) 라인 일치"

        if any(
            k in name_clean
            for k in [
                "포밍워시",
                "클렌저",
                "크리미워시",
                "클렌징오일",
                "클렌징",
                "포하",
                "4ha",
            ]
        ):
            if "화이트티트리" in name_clean:
                return "클렌징", "⚠️ 확인필요", "클렌징 추정(이미지 외 신규 라인: 화이트 티트리)"
            return "클렌징", "✅ 정상매칭", "이미지(클렌징) 라인 일치"

        return "미분류(확인필요)", "⚠️ 확인필요", "포트폴리오 이미지 내 매칭 키워드 없음"

    if brand == "Lipcare":
        if "딥모이스트" in name_clean or "deepmoist" in name_clean:
            return "딥 모이스트", "✅ 정상매칭", "이미지(딥 모이스트) 라인 일치"

        if "메디케이티드" in name_clean or "medicated" in name_clean:
            return "메디케이티드", "✅ 정상매칭", "이미지(메디케이티드) 라인 일치"

        if "소프트립" in name_clean or "softlips" in name_clean:
            return "소프트립스", "✅ 정상매칭", "이미지(소프트립스) 라인 일치"

        if any(k in name_clean for k in ["매직컬러", "매직칼라", "틴티드컬러", "틴티드칼라"]):
            return "매직컬러 / 틴티드컬러 라인", "✅ 정상매칭", "이미지(매직컬러/틴티드컬러) 일치"

        if "립아이스" in name_clean or "lipice" in name_clean:
            if any(k in name_clean for k in ["프루티", "레몬", "스트로베리", "딸기", "그레이프", "애플"]):
                if "애플" in name_clean:
                    return "립아이스 프루티", "⚠️ 확인필요", "립아이스 프루티 추정(이미지 외 향: 애플)"
                return "립아이스 프루티", "✅ 정상매칭", "이미지(립아이스 프루티) 라인 일치"
            return "립아이스 프루티", "⚠️ 확인필요", "립아이스 세부 라인 확인 필요"

        return "미분류(확인필요)", "⚠️ 확인필요", "포트폴리오 이미지 외 라인(멜티크림/워터컬러/립퓨어 등)"

    if brand == "HADALABO":
        if "프리미엄" in name_clean or "premium" in name_clean:
            return "Gokujyun premium (고쿠쥰 프리미엄)", "✅ 정상매칭", "이미지(Gokujyun premium) 일치"

        if "시로쥰" in name_clean or "shirojyun" in name_clean:
            return "Shirojyun (시로쥰)", "✅ 정상매칭", "이미지(Shirojyun) 일치"

        if "키와미즈" in name_clean or "kiwamizu" in name_clean:
            return "Kiwamizu (키와미즈)", "✅ 정상매칭", "이미지(Kiwamizu) 일치"

        if "타마고하다" in name_clean or "tamagohada" in name_clean:
            return "Tamagohada (타마고하다)", "✅ 정상매칭", "이미지(Tamagohada) 일치"

        if "고쿠쥰" in name_clean or "gokujyun" in name_clean:
            if "탄력" in name_clean:
                return "Gokujyun (고쿠쥰)", "⚠️ 확인필요", "고쿠쥰 계열(탄력 라인: 이미지 외 세부라인)"
            return "Gokujyun (고쿠쥰)", "✅ 정상매칭", "이미지(Gokujyun) 라인 일치"

        return "미분류(확인필요)", "⚠️ 확인필요", "포트폴리오 이미지 외 라인(히아풀베리어 등)"

    return "-", "해당없음", "기타 브랜드"


def normalize_brand_name(val: str) -> str:
    s = str(val or "").strip()
    mapping = {
        "ACNES": "Acnes",
        "acnes": "Acnes",
        "MELANO": "Melano",
        "melano": "Melano",
        "hadalabo": "HADALABO",
        "Hadalabo": "HADALABO",
        "lipcare": "Lipcare",
        "LIPCARE": "Lipcare",
        "0": "미분류",
        "nan": "미분류",
        "": "미분류",
    }
    return mapping.get(s, s)


@st.cache_data
def load_default_master(primary_path: str = DEFAULT_MASTER_FILE, fallback_path: str = FALLBACK_MASTER_FILE) -> pd.DataFrame:
    target_path = primary_path if os.path.exists(primary_path) else (fallback_path if os.path.exists(fallback_path) else None)
    if target_path is None:
        return pd.DataFrame(columns=["제품코드", "브랜드", "제품군"])
    df = pd.read_excel(target_path, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    for col in ["제품코드", "브랜드", "제품군"]:
        if col not in df.columns:
            df[col] = ""
    df["제품코드"] = df["제품코드"].astype(str).str.strip()
    df["브랜드"] = df["브랜드"].apply(normalize_brand_name)
    df["제품군"] = df["제품군"].astype(str).str.strip()
    df = df[df["제품코드"].notna() & (df["제품코드"] != "") & (df["제품코드"] != "nan")]
    df = df.drop_duplicates(subset=["제품코드"], keep="first").reset_index(drop=True)
    return df[["제품코드", "브랜드", "제품군"]]


def generate_upload_template_bytes() -> bytes:
    sample_cols = [
        "순번",
        "센터",
        "셀",
        "상품",
        "상품명",
        "화주LOT",
        "유효일자",
        "입수량(BOX)",
        "정상수량_BOX",
        "정상수량_EA",
        "정상수량_환산",
        "불량수량_환산",
        "합계수량",
    ]
    sample_rows = [
        [1, "코람코센터", "EE-01-10-10", "ME00621PSB", "아크네스 퍼펙트 솔루션 바디워시 500ml", "LOT2026A", "2029-07-07", 12, 10, 0, 120, 0, 120],
        [2, "코람코센터", "EE-01-11-10", "KME90521MC3", "립아이스 매직칼라 딸기 #48 (New 2026)", "G04892", "2029-09-20", 48, 20, 0, 960, 0, 960],
        [3, "코람코센터", "EE-01-12-10", "ME90621PR4", "하다라보 고쿠쥰 프리미엄 로션 170ml", "H2601", "2029-05-01", 12, 15, 0, 180, 0, 180],
    ]
    df_template = pd.DataFrame(sample_rows, columns=sample_cols)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df_template.to_excel(writer, sheet_name="업로드 파일", index=False)
    return buf.getvalue()


def parse_inventory_excel(uploaded_file) -> tuple[pd.DataFrame, str]:
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    if "업로드 파일" in sheet_names:
        df = pd.read_excel(xls, sheet_name="업로드 파일")
        df.columns = [str(c).strip() for c in df.columns]
        required = {"상품", "상품명", "정상수량_환산"}
        if not required.issubset(set(df.columns)):
            raise ValueError(f"'업로드 파일' 시트에 필수 컬럼({required})이 누락되었습니다.")

        out = pd.DataFrame()
        out["센터"] = df["센터"].fillna("-") if "센터" in df.columns else "-"
        out["셀"] = df["셀"].fillna("-") if "셀" in df.columns else "-"
        out["상품"] = df["상품"].astype(str).str.strip()
        out["상품명"] = df["상품명"].astype(str).str.strip()
        out["화주LOT"] = df["화주LOT"].fillna("-").astype(str) if "화주LOT" in df.columns else "-"
        out["유효일자"] = df["유효일자"].fillna("-").astype(str) if "유효일자" in df.columns else "-"
        out["정상환산"] = pd.to_numeric(df["정상수량_환산"], errors="coerce").fillna(0)
        out["불량환산"] = (
            pd.to_numeric(df["불량수량_환산"], errors="coerce").fillna(0)
            if "불량수량_환산" in df.columns
            else 0
        )
        out["합계수량"] = (
            pd.to_numeric(df["합계수량"], errors="coerce").fillna(out["정상환산"] + out["불량환산"])
            if "합계수량" in df.columns
            else out["정상환산"] + out["불량환산"]
        )
        out = out[out["상품"].notna() & (out["상품"] != "") & (out["상품"] != "nan")].copy()
        return out, "표준 양식('업로드 파일' 시트) 로드 완료"

    if "sheet1" in sheet_names:
        df = pd.read_excel(xls, sheet_name="sheet1", header=1)
        df.columns = [str(c).strip() for c in df.columns]
        if "순번" in df.columns:
            df = df[df["순번"].notna()].copy()
        if "상품" not in df.columns or "환산" not in df.columns:
            raise ValueError("'sheet1' 시트의 헤더 구조가 올바르지 않습니다. 표준 '업로드 파일' 양식을 사용해 주세요.")

        out = pd.DataFrame()
        out["센터"] = df["센터"].fillna("-") if "센터" in df.columns else "-"
        out["셀"] = df["셀"].fillna("-") if "셀" in df.columns else "-"
        out["상품"] = df["상품"].astype(str).str.strip()
        out["상품명"] = df["상품명"].astype(str).str.strip() if "상품명" in df.columns else ""
        out["화주LOT"] = df["화주LOT"].fillna("-").astype(str) if "화주LOT" in df.columns else "-"
        out["유효일자"] = (
            pd.to_datetime(df["유효일자"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("-")
            if "유효일자" in df.columns
            else "-"
        )
        out["정상환산"] = pd.to_numeric(df["환산"], errors="coerce").fillna(0)
        out["불량환산"] = pd.to_numeric(df["환산.1"], errors="coerce").fillna(0) if "환산.1" in df.columns else 0
        out["합계수량"] = (
            pd.to_numeric(df["합계수량"], errors="coerce").fillna(out["정상환산"] + out["불량환산"])
            if "합계수량" in df.columns
            else out["정상환산"] + out["불량환산"]
        )
        out = out[out["상품"].notna() & (out["상품"] != "") & (out["상품"] != "nan")].copy()
        return out, "WMS 원본 양식('sheet1' 시트) 자동 변환 로드 완료"

    raise ValueError("엑셀 파일 내에 '업로드 파일' 시트(또는 'sheet1')가 존재하지 않습니다. 양식 틀을 확인해 주세요.")


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return buf.getvalue()


if "master_df" not in st.session_state:
    st.session_state["master_df"] = load_default_master()

if "portfolio_overrides" not in st.session_state:
    st.session_state["portfolio_overrides"] = {}

if "selected_brand" not in st.session_state:
    st.session_state["selected_brand"] = "Acnes"

st.title("📦 브랜드 및 포트폴리오별 통합 재고 현황 대시보드")
st.caption("브랜드_제품코드_마스터와 '업로드 파일' 양식의 재고 데이터를 결합하여 브랜드·포트폴리오별 재고를 실시간 분석합니다.")

with st.container(border=True):
    up_col1, up_col2 = st.columns([3, 1.2])
    with up_col1:
        uploaded_inv = st.file_uploader(
            "📂 재고 파일 업로드 ('업로드 파일' 시트 양식 또는 WMS 재고 엑셀)",
            type=["xlsx", "xls"],
            help="반드시 지정된 '업로드 파일' 시트 형식에 맞추어 업로드해 주세요.",
        )
    with up_col2:
        st.write("")
        st.download_button(
            label="📥 '업로드 파일' 표준 양식 다운로드",
            data=generate_upload_template_bytes(),
            file_name="업로드 파일.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

if uploaded_inv is None:
    st.info("👆 상단에서 재고 엑셀 파일을 업로드하면 브랜드별 재고 막대그래프와 포트폴리오 분석이 즉시 실행됩니다.")
    st.stop()

try:
    inv_df, load_msg = parse_inventory_excel(uploaded_inv)
    st.success(f"✅ {load_msg} (총 {len(inv_df):,}건 / 고유 품목 {inv_df['상품'].nunique():,}종)")
except Exception as e:
    st.error(f"❌ 파일 형식 오류: {e}")
    st.stop()

master_df = st.session_state["master_df"].copy()
master_codes = set(master_df["제품코드"].unique())
inv_unique_items = inv_df[["상품", "상품명"]].drop_duplicates(subset=["상품"])

new_items = inv_unique_items[~inv_unique_items["상품"].isin(master_codes)].copy()
if not new_items.empty:
    new_rows = pd.DataFrame(
        {
            "제품코드": new_items["상품"],
            "브랜드": "미분류",
            "제품군": new_items["상품명"],
        }
    )
    master_df = pd.concat([new_rows, master_df], ignore_index=True).drop_duplicates(subset=["제품코드"])
    st.session_state["master_df"] = master_df

active_master = master_df[master_df["제품코드"].isin(inv_unique_items["상품"])].copy()
active_master = active_master.sort_values(by=["브랜드", "제품코드"]).reset_index(drop=True)

unassigned_count = (active_master["브랜드"] == "미분류").sum()

with st.expander(
    f"🛠️ 코드별 브랜드 매핑 수정 (재고 보유 품목 {len(active_master)}종 | 미분류 {unassigned_count}종)",
    expanded=(unassigned_count > 0),
):
    st.markdown(
        "재고 파일에 포함된 **제품코드별 브랜드**를 아래 표에서 직접 수정할 수 있습니다. "
        "수정 즉시 그래프와 하단 내역에 자동 반영됩니다."
    )
    m_col1, m_col2 = st.columns([2, 1])
    with m_col1:
        search_master = st.text_input("🔍 제품코드 또는 제품명 검색", placeholder="예: ME00621, 아크네스, 립아이스...")
    with m_col2:
        only_unassigned = st.checkbox("미분류 코드만 보기", value=False)

    edit_target = active_master.copy()
    if search_master:
        mask = edit_target["제품코드"].str.contains(search_master, case=False, na=False) | edit_target[
            "제품군"
        ].str.contains(search_master, case=False, na=False)
        edit_target = edit_target[mask]
    if only_unassigned:
        edit_target = edit_target[edit_target["브랜드"] == "미분류"]

    edited_master = st.data_editor(
        edit_target,
        column_config={
            "제품코드": st.column_config.TextColumn("제품코드", disabled=True),
            "브랜드": st.column_config.SelectboxColumn("브랜드 (클릭하여 수정)", options=BRAND_CHOICES, required=True),
            "제품군": st.column_config.TextColumn("제품군(품목명)", disabled=True),
        },
        use_container_width=True,
        hide_index=True,
        height=280,
        key="brand_master_editor",
    )

    if not edited_master.equals(edit_target):
        update_map = dict(zip(edited_master["제품코드"], edited_master["브랜드"]))
        st.session_state["master_df"]["브랜드"] = st.session_state["master_df"].apply(
            lambda r: update_map.get(r["제품코드"], r["브랜드"]), axis=1
        )
        st.rerun()

    st.download_button(
        label="💾 수정된 '브랜드_제품코드_마스터.xlsx' 다운로드",
        data=to_excel_bytes(st.session_state["master_df"], sheet_name="Sheet1"),
        file_name=DEFAULT_MASTER_FILE,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

brand_map = dict(zip(st.session_state["master_df"]["제품코드"], st.session_state["master_df"]["브랜드"]))
inv_df["브랜드"] = inv_df["상품"].map(brand_map).fillna("미분류")

unique_prod = inv_df[["상품", "상품명", "브랜드"]].drop_duplicates(subset=["상품"]).copy()
port_results = unique_prod.apply(lambda r: classify_portfolio(r["브랜드"], r["상품명"]), axis=1)
unique_prod["자동_포트폴리오"] = [x[0] for x in port_results]
unique_prod["검수상태"] = [x[1] for x in port_results]
unique_prod["판별근거"] = [x[2] for x in port_results]

overrides = st.session_state["portfolio_overrides"]
unique_prod["최종_포트폴리오"] = unique_prod.apply(
    lambda r: overrides.get(r["상품"], r["자동_포트폴리오"]), axis=1
)
unique_prod["검수상태"] = unique_prod.apply(
    lambda r: "✏️ 사용자수정완료" if r["상품"] in overrides else r["검수상태"], axis=1
)

inv_merged = inv_df.merge(
    unique_prod[["상품", "최종_포트폴리오", "검수상태", "판별근거"]],
    on="상품",
    how="left",
)

st.markdown('<div class="section-title">📊 브랜드별 전체 재고 현황</div>', unsafe_allow_html=True)

ctrl_c1, ctrl_c2 = st.columns([2, 3])
with ctrl_c1:
    qty_metric = st.radio(
        "수량 기준 선택",
        options=["정상환산", "합계수량", "불량환산"],
        format_func=lambda x: {"정상환산": "정상재고 (환산 EA)", "합계수량": "총재고 (정상+불량 합계)", "불량환산": "불량재고 (환산 EA)"}[x],
        horizontal=True,
    )

k1, k2, k3, k4 = st.columns(4)
k1.metric("총 재고 수량 (EA)", f"{int(inv_merged[qty_metric].sum()):,}")
k2.metric("Acnes 재고", f"{int(inv_merged[inv_merged['브랜드'] == 'Acnes'][qty_metric].sum()):,}")
k3.metric("Lipcare 재고", f"{int(inv_merged[inv_merged['브랜드'] == 'Lipcare'][qty_metric].sum()):,}")
k4.metric("HADALABO 재고", f"{int(inv_merged[inv_merged['브랜드'] == 'HADALABO'][qty_metric].sum()):,}")

brand_summary = (
    inv_merged.groupby("브랜드", as_index=False)[qty_metric]
    .sum()
    .sort_values(by=qty_metric, ascending=False)
)

chart_brand_df = brand_summary.rename(columns={qty_metric: "재고수량(EA)"})
st.bar_chart(
    data=chart_brand_df,
    x="브랜드",
    y="재고수량(EA)",
    color="브랜드",
    height=400,
    use_container_width=True,
)

st.markdown('<div class="section-title">🔎 브랜드 선택 및 상세 재고 내역 조회</div>', unsafe_allow_html=True)

available_brands = brand_summary["브랜드"].tolist()
if st.session_state["selected_brand"] not in available_brands and available_brands:
    st.session_state["selected_brand"] = available_brands[0]

btn_cols = st.columns(min(len(available_brands), 6))
for idx, b_name in enumerate(available_brands):
    col_target = btn_cols[idx % len(btn_cols)]
    is_active = st.session_state["selected_brand"] == b_name
    b_qty = int(brand_summary.loc[brand_summary["브랜드"] == b_name, qty_metric].values[0])
    if col_target.button(
        f"{'✅ ' if is_active else ''}{b_name}\n({b_qty:,})",
        key=f"btn_brand_{b_name}",
        type="primary" if is_active else "secondary",
        use_container_width=True,
    ):
        st.session_state["selected_brand"] = b_name
        st.rerun()

sel_brand = st.session_state["selected_brand"]
brand_detail_df = inv_merged[inv_merged["브랜드"] == sel_brand].copy()

st.markdown(f"### 📌 [{sel_brand}] 상세 재고 및 포트폴리오 분석")

if sel_brand in PORTFOLIO_OPTIONS:
    p_col1, p_col2 = st.columns([1.2, 1.8])

    with p_col1:
        port_summary = (
            brand_detail_df.groupby("최종_포트폴리오", as_index=False)[qty_metric]
            .sum()
            .sort_values(by=qty_metric, ascending=False)
        )
        chart_port_df = port_summary.rename(columns={"최종_포트폴리오": "포트폴리오", qty_metric: "재고수량(EA)"})
        st.bar_chart(
            data=chart_port_df,
            x="포트폴리오",
            y="재고수량(EA)",
            color="포트폴리오",
            height=360,
            use_container_width=True,
        )

    with p_col2:
        prod_check_df = (
            brand_detail_df.groupby(["상품", "상품명", "최종_포트폴리오", "검수상태", "판별근거"], as_index=False)[
                qty_metric
            ]
            .sum()
            .sort_values(by=["검수상태", qty_metric], ascending=[True, False])
        )
        need_review_cnt = (prod_check_df["검수상태"] == "⚠️ 확인필요").sum()

        if need_review_cnt > 0:
            st.warning(
                f"⚠️ **{sel_brand}** 품목 중 포트폴리오 이미지와 100% 일치하지 않거나 신규 라인인 **{need_review_cnt}개 품목**이 감지되었습니다. 아래 표에서 포트폴리오를 직접 수정해 주세요."
            )
        else:
            st.success(f"✅ **{sel_brand}**의 모든 품목이 포트폴리오 이미지와 정상 매칭되었습니다. (필요 시 아래에서 변경 가능)")

        show_only_review = st.checkbox(
            "⚠️ 확인필요(미매칭/신규라인) 품목만 필터링하여 보기",
            value=(need_review_cnt > 0),
            key=f"chk_review_{sel_brand}",
        )
        display_check = (
            prod_check_df[prod_check_df["검수상태"] == "⚠️ 확인필요"].copy()
            if show_only_review
            else prod_check_df.copy()
        )

        edited_port = st.data_editor(
            display_check,
            column_config={
                "상품": st.column_config.TextColumn("제품코드", disabled=True),
                "상품명": st.column_config.TextColumn("제품명", disabled=True),
                "최종_포트폴리오": st.column_config.SelectboxColumn(
                    "포트폴리오 (클릭 수정)",
                    options=PORTFOLIO_OPTIONS[sel_brand],
                    required=True,
                ),
                "검수상태": st.column_config.TextColumn("더블체크 상태", disabled=True),
                "판별근거": st.column_config.TextColumn("매칭 사유", disabled=True),
                qty_metric: st.column_config.NumberColumn("재고수량", format="%d", disabled=True),
            },
            use_container_width=True,
            hide_index=True,
            height=300,
            key=f"port_editor_{sel_brand}",
        )

        if not edited_port.equals(display_check):
            for _, row in edited_port.iterrows():
                orig_val = display_check.loc[display_check["상품"] == row["상품"], "최종_포트폴리오"].values[0]
                if row["최종_포트폴리오"] != orig_val:
                    st.session_state["portfolio_overrides"][row["상품"]] = row["최종_포트폴리오"]
            st.rerun()

st.markdown(f"#### 📋 [{sel_brand}] 전체 재고 상세 내역 (센터 · LOT · 유효일자별)")

detail_cols = ["센터", "셀", "상품", "상품명", "최종_포트폴리오", "검수상태", "화주LOT", "유효일자", "정상환산", "불량환산", "합계수량"]
detail_view = brand_detail_df[detail_cols].sort_values(by=["최종_포트폴리오", "상품", "유효일자"])

st.dataframe(
    detail_view,
    use_container_width=True,
    hide_index=True,
    height=380,
)

st.download_button(
    label=f"📥 [{sel_brand}] 상세 재고내역 엑셀 다운로드",
    data=to_excel_bytes(detail_view, sheet_name=f"{sel_brand}_재고내역"),
    file_name=f"{sel_brand}_상세재고내역.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
