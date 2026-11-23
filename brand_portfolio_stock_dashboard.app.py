import io
import os
import re
from datetime import date
import openpyxl
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
    .block-container { padding-top: 1.1rem; padding-bottom: 2rem; }
    div[data-testid="stMetric"] {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        padding: 12px 16px;
        border-radius: 10px;
    }
    .section-title {
        font-size: 1.15rem;
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
NEW_SKU_TEMPLATE_FILE = "신규_제품코드_등록양식.xlsx"
EXCLUDED_KEYWORDS = ("겔", "베르가못", "플라스타", "플라스마")
GIFT_KEYWORDS = ("증정용", "증정", "사은품", "기프트", "gift", "샘플", "sample", "견본")

DAISO_ACNES_CODES = {
    "ME00621DCP",
    "ME00621DFC",
    "ME00621DSE",
    "ME00621DTF",
    "ME00621DTP",
    "ME00621DTT",
}

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
        "다이소아크네스": "Acnes",
        "다이소-아크네스": "Acnes",
        "0": "미분류",
        "nan": "미분류",
        "": "미분류",
    }
    return mapping.get(s, s)


@st.cache_data(show_spinner=False)
def load_default_master(
    primary_path: str = DEFAULT_MASTER_FILE, fallback_path: str = FALLBACK_MASTER_FILE
) -> pd.DataFrame:
    target_path = (
        primary_path
        if os.path.exists(primary_path)
        else (fallback_path if os.path.exists(fallback_path) else None)
    )
    if target_path is None:
        return pd.DataFrame(columns=["제품코드", "브랜드", "제품군"])
    df = pd.read_excel(target_path, sheet_name=0, engine="openpyxl")
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


def _find_col_index(row1: list, row2: list, target_name: str, fallback_idx: int) -> int:
    for idx, val in enumerate(row1):
        if val is not None and str(val).strip() == target_name:
            return idx
    for idx, val in enumerate(row2):
        if val is not None and str(val).strip() == target_name:
            return idx
    return fallback_idx


@st.cache_data(show_spinner=False)
def parse_warehouse_bytes(file_bytes: bytes, file_name: str) -> tuple[pd.DataFrame, dict]:
    fname = (file_name or "").lower()

    if fname.endswith(".csv"):
        df = pd.read_csv(io.BytesIO(file_bytes))
        df.columns = [str(c).strip() for c in df.columns]
        qty_col = (
            "정상수량_환산"
            if "정상수량_환산" in df.columns
            else ("환산" if "환산" in df.columns else None)
        )
        out = pd.DataFrame(
            {
                "센터": df["센터"].fillna("-").astype(str).str.strip() if "센터" in df.columns else "-",
                "셀": df["셀"].fillna("-").astype(str).str.strip() if "셀" in df.columns else "-",
                "상품": df["상품"].astype(str).str.strip() if "상품" in df.columns else "",
                "상품명": df["상품명"].fillna("").astype(str).str.strip() if "상품명" in df.columns else "",
                "화주LOT": df["화주LOT"].fillna("-").astype(str).str.strip() if "화주LOT" in df.columns else "-",
                "유효일자_raw": df["유효일자"] if "유효일자" in df.columns else None,
                "정상환산": pd.to_numeric(df[qty_col], errors="coerce").fillna(0.0) if qty_col else 0.0,
                "불량환산": pd.to_numeric(df["불량수량_환산"], errors="coerce").fillna(0.0)
                if "불량수량_환산" in df.columns
                else 0.0,
                "합계수량_raw": pd.to_numeric(df["합계수량"], errors="coerce")
                if "합계수량" in df.columns
                else None,
            }
        )
        raw_total_rows = len(out)
        source_label = "CSV 재고파일"
    else:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        sheet_names = wb.sheetnames
        sheet_name = (
            "재고파일"
            if "재고파일" in sheet_names
            else ("sheet1" if "sheet1" in sheet_names else sheet_names[0])
        )
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(values_only=True))
        wb.close()

        if len(rows) < 2:
            raise ValueError(f"'{sheet_name}' 시트에 데이터가 없습니다.")

        row1 = rows[0]
        row2 = rows[1]

        idx_seq = _find_col_index(row1, row2, "순번", 0)
        idx_center = _find_col_index(row1, row2, "센터", 1)
        idx_cell = _find_col_index(row1, row2, "셀", 2)
        idx_prod = _find_col_index(row1, row2, "상품", 3)
        idx_name = _find_col_index(row1, row2, "상품명", 4)
        idx_lot = _find_col_index(row1, row2, "화주LOT", 6)
        idx_exp = _find_col_index(row1, row2, "유효일자", 13)
        idx_total = _find_col_index(row1, row2, "합계수량", 31)

        idx_norm_conv = 27
        for idx, r1_val in enumerate(row1):
            if r1_val is not None and "정상수량" in str(r1_val).strip():
                for off in range(4):
                    c_pos = idx + off
                    if (
                        c_pos < len(row2)
                        and row2[c_pos] is not None
                        and "환산" in str(row2[c_pos]).strip()
                    ):
                        idx_norm_conv = c_pos
                        break
                break

        idx_bad_conv = 30
        for idx, r1_val in enumerate(row1):
            if r1_val is not None and "불량수량" in str(r1_val).strip():
                for off in range(4):
                    c_pos = idx + off
                    if (
                        c_pos < len(row2)
                        and row2[c_pos] is not None
                        and "환산" in str(row2[c_pos]).strip()
                    ):
                        idx_bad_conv = c_pos
                        break
                break

        has_sub_header = any(
            v is not None and str(v).strip() in {"BOX", "EA", "환산", "창고", "LOCATION"}
            for v in row2
        )
        start_idx = 2 if has_sub_header else 1

        records = []
        for r in rows[start_idx:]:
            if not r:
                continue
            seq_v = r[idx_seq] if idx_seq < len(r) else None
            prod_v = r[idx_prod] if idx_prod < len(r) else None
            if seq_v is None and (prod_v is None or str(prod_v).strip() == ""):
                continue
            p_str = str(prod_v or "").strip()
            if not p_str or p_str.lower() == "nan":
                continue

            records.append(
                {
                    "센터": str(r[idx_center] or "-").strip() if idx_center < len(r) else "-",
                    "셀": str(r[idx_cell] or "-").strip() if idx_cell < len(r) else "-",
                    "상품": p_str,
                    "상품명": str(r[idx_name] or "").strip() if idx_name < len(r) else "",
                    "화주LOT": str(r[idx_lot] or "-").strip() if idx_lot < len(r) else "-",
                    "유효일자_raw": r[idx_exp] if idx_exp < len(r) else None,
                    "정상환산": r[idx_norm_conv] if idx_norm_conv < len(r) else 0,
                    "불량환산": r[idx_bad_conv] if idx_bad_conv < len(r) else 0,
                    "합계수량_raw": r[idx_total] if idx_total < len(r) else None,
                }
            )

        out = pd.DataFrame(records)
        raw_total_rows = len(out)
        source_label = f"물류창고 '{sheet_name}' 시트"

    if out.empty:
        raise ValueError("업로드 파일에서 유효한 재고 데이터를 찾지 못했습니다.")

    out = out[
        out["상품"].notna() & (out["상품"] != "") & (out["상품"].str.lower() != "nan")
    ].copy()
    out["센터"] = out["센터"].replace({"": "-", "nan": "-", "None": "-"})
    out["셀"] = out["셀"].replace({"": "-", "nan": "-", "None": "-"})
    out["화주LOT"] = out["화주LOT"].replace({"": "-", "nan": "-", "None": "-"})

    out["정상환산"] = pd.to_numeric(out["정상환산"], errors="coerce").fillna(0.0)
    out["불량환산"] = pd.to_numeric(out["불량환산"], errors="coerce").fillna(0.0)
    tot_num = pd.to_numeric(out["합계수량_raw"], errors="coerce")
    out["합계수량"] = tot_num.fillna(out["정상환산"] + out["불량환산"])

    exp_dt = pd.to_datetime(out["유효일자_raw"], errors="coerce")
    raw_exp_str = (
        out["유효일자_raw"]
        .astype(str)
        .str.strip()
        .replace({"None": "-", "nan": "-", "NaT": "-", "": "-"})
    )
    out["유효일자"] = exp_dt.dt.strftime("%Y-%m-%d").fillna(raw_exp_str)

    excl_pattern = "|".join(re.escape(k) for k in EXCLUDED_KEYWORDS)
    excl_mask = out["상품명"].str.contains(excl_pattern, case=False, na=False)
    excluded_rows = int(excl_mask.sum())
    excluded_qty = float(out.loc[excl_mask, "합계수량"].sum())

    filtered = (
        out[~excl_mask]
        .drop(columns=["유효일자_raw", "합계수량_raw"])
        .reset_index(drop=True)
    )

    meta = {
        "source_label": source_label,
        "raw_total_rows": raw_total_rows,
        "excluded_rows": excluded_rows,
        "excluded_qty": excluded_qty,
        "final_rows": len(filtered),
        "unique_skus": int(filtered["상품"].nunique()),
    }
    return filtered, meta


@st.cache_data(show_spinner=False)
def parse_outbound_bytes(file_bytes: bytes, file_name: str) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    fname = (file_name or "").lower()
    if fname.endswith(".csv"):
        raw_df = pd.read_csv(io.BytesIO(file_bytes))
    else:
        xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
        sheet = xls.sheet_names[0]
        preview = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=10)
        header_row = 0
        code_tokens = {"상품코드", "제품코드", "품목코드", "상품", "제품", "품목", "ME코드"}
        qty_tokens = {"출고수량", "출고", "수량", "환산", "수주량", "합계", "출고량", "매출수량"}
        for r_idx in range(len(preview)):
            vals = {str(v).strip() for v in preview.iloc[r_idx].tolist() if pd.notna(v)}
            if (vals & code_tokens or any("코드" in x or "상품" in x for x in vals)) and (
                vals & qty_tokens or any("수량" in x or "출고" in x for x in vals)
            ):
                header_row = r_idx
                break
        raw_df = pd.read_excel(xls, sheet_name=sheet, header=header_row)

    df = raw_df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    code_col = next(
        (c for c in df.columns if c in {"상품코드", "제품코드", "품목코드", "상품", "제품", "품목", "SKU", "sku"}),
        None,
    )
    if code_col is None:
        code_col = next((c for c in df.columns if "코드" in c or "상품" in c or "품목" in c), df.columns[0])

    name_col = next((c for c in df.columns if c in {"상품명", "제품명", "품목명"}), None)
    if name_col is None:
        name_col = next((c for c in df.columns if "명" in c and "출고" not in c), None)

    div_col = next(
        (c for c in df.columns if c in {"구분", "출고구분", "유형", "거래구분", "매출구분", "채널"}),
        None,
    )
    if div_col is None:
        div_col = next((c for c in df.columns if "구분" in c), None)

    date_col = next(
        (c for c in df.columns if c in {"Date", "date", "일자", "날짜", "출고일", "출고일자", "매출일", "거래일"}),
        None,
    )
    if date_col is None:
        date_col = next(
            (c for c in df.columns if any(k in c.lower() for k in ["date", "일자", "날짜", "출고일"])),
            None,
        )

    qty_col = next(
        (c for c in df.columns if c in {"수량", "출고수량", "출고량", "환산", "수주수량"}),
        None,
    )
    if qty_col is None:
        qty_col = next(
            (
                c
                for c in df.columns
                if any(k in c for k in ["수량", "출고수량", "출고량", "환산", "출고"])
                and "일자" not in c
                and "날짜" not in c
                and "코드" not in c
                and "구분" not in c
            ),
            None,
        )
    if qty_col is None:
        num_cols = [c for c in df.columns if c != code_col and pd.api.types.is_numeric_dtype(df[c])]
        qty_col = num_cols[0] if num_cols else df.columns[-1]

    cleaned = pd.DataFrame(
        {
            "상품": df[code_col].astype(str).str.strip(),
            "출고_상품명": df[name_col].fillna("").astype(str).str.strip() if name_col else "",
            "구분": df[div_col].fillna("미구분").astype(str).str.strip() if div_col else "매출",
            "출고수량": pd.to_numeric(df[qty_col], errors="coerce").fillna(0.0),
        }
    )
    cleaned = cleaned[
        cleaned["상품"].notna()
        & (cleaned["상품"] != "")
        & (cleaned["상품"].str.lower() != "nan")
        & (~cleaned["상품"].isin(["상품코드", "제품코드", "품목코드", "상품", "합계", "총계"]))
    ].copy()

    dt_s = pd.to_datetime(df.loc[cleaned.index, date_col], errors="coerce") if date_col else pd.Series(pd.NaT, index=cleaned.index)
    cleaned["출고일자"] = dt_s

    sales_mask = cleaned["구분"].str.contains("매출", case=False, na=False)
    has_sales_div = bool(div_col and sales_mask.any())
    active_out = cleaned[sales_mask].copy() if has_sales_div else cleaned.copy()
    if active_out.empty:
        active_out = cleaned.copy()

    act_dt = active_out["출고일자"].dropna()
    if not act_dt.empty:
        unique_days = max(int(act_dt.dt.normalize().nunique()), 1)
        min_dt = act_dt.min()
        max_dt = act_dt.max()
        span_days = max(int((max_dt.normalize() - min_dt.normalize()).days) + 1, unique_days)
        denom_days = float(span_days)
        date_range_str = f"{min_dt.strftime('%Y-%m-%d')} ~ {max_dt.strftime('%Y-%m-%d')} ({span_days}일)"
    else:
        unique_days = 1
        span_days = 30
        denom_days = 30.0
        date_range_str = "날짜열 없음(30일 환산)"

    sku_outbound = (
        active_out.groupby("상품", as_index=False)
        .agg(
            누적_출고수량=("출고수량", "sum"),
            출고_발생일수=("출고일자", lambda s: int(s.dropna().dt.normalize().nunique()) if s.notna().any() else 1),
        )
    )
    sku_outbound["출고_발생일수"] = sku_outbound["출고_발생일수"].clip(lower=1)
    sku_outbound["일평균_출고량"] = (sku_outbound["누적_출고수량"] / denom_days).round(2)
    sku_outbound["월예상_출고량"] = (sku_outbound["일평균_출고량"] * 30.0).round(0)

    hist_df = active_out[active_out["출고일자"].notna()].copy()
    if not hist_df.empty:
        hist_df["연월"] = hist_df["출고일자"].dt.to_period("M").astype(str)
        monthly_hist = (
            hist_df.groupby(["상품", "연월"], as_index=False)["출고수량"]
            .sum()
            .sort_values(["상품", "연월"])
        )
    else:
        monthly_hist = pd.DataFrame(columns=["상품", "연월", "출고수량"])

    info = {
        "rows": len(active_out),
        "unique_skus": int(sku_outbound["상품"].nunique()),
        "total_outbound": float(active_out["출고수량"].sum()),
        "denom_days": int(span_days),
        "unique_days": unique_days,
        "date_range": date_range_str,
        "div_col": div_col or "구분열 없음(전체 집계)",
        "used_sales_filter": has_sales_div,
    }
    return sku_outbound, monthly_hist, info


@st.cache_data(show_spinner=False)
def classify_skus_cached(
    unique_sku_df: pd.DataFrame,
    master_records: tuple[tuple[str, str], ...],
    brand_overrides_tuple: tuple[tuple[str, str], ...],
    port_overrides_tuple: tuple[tuple[str, str], ...],
) -> pd.DataFrame:
    master_map = dict(master_records)
    brand_ov = dict(brand_overrides_tuple)
    port_ov = dict(port_overrides_tuple)

    brands = []
    auto_ports = []
    final_ports = []
    statuses = []
    reasons = []

    for row in unique_sku_df.itertuples(index=False):
        code = str(row.상품 or "").strip()
        name = str(row.상품명 or "").strip()
        name_clean = re.sub(r"\s+", "", name).lower()
        is_gift_name = any(k in name_clean for k in GIFT_KEYWORDS)

        if code in brand_ov:
            b = normalize_brand_name(brand_ov[code])
        elif is_gift_name:
            b = "사은품"
        elif code in DAISO_ACNES_CODES or ("아크네스화이트티트리" in name_clean and "다이소" in name_clean):
            b = "Acnes"
        else:
            m_brand = normalize_brand_name(master_map.get(code, "미분류"))
            if m_brand in ("미분류", "부자재", "사은품"):
                if "아크네스" in name_clean or "acnes" in name_clean:
                    b = "Acnes"
                elif "립" in name_clean or "lip" in name_clean:
                    b = "Lipcare"
                elif "하다라보" in name_clean or "hadalabo" in name_clean:
                    b = "HADALABO"
                else:
                    b = m_brand
            else:
                b = m_brand

        if b == "사은품":
            ap, st_val, rsn = "사은품", "✅ 정상매칭", "증정용/사은품 관련 단어 포함 → 별도 사은품 브랜드 분류"
        elif b == "Acnes":
            if "포맨" in name_clean or "formen" in name_clean:
                ap, st_val, rsn = "포 맨", "✅ 정상매칭", "Acnes 포맨 라인 일치"
            elif "더마릴리프" in name_clean or "더마" in name_clean:
                if any(k in name_clean for k in ["클렌저", "워시", "폼"]):
                    ap, st_val, rsn = "클렌징", "✅ 정상매칭", "Acnes 더마 클렌징 라인 일치"
                elif any(k in name_clean for k in ["리커버리", "밸런싱", "크림", "스킨"]):
                    ap, st_val, rsn = "더마 스킨케어", "✅ 정상매칭", "Acnes 더마 스킨케어 라인 일치"
                else:
                    ap, st_val, rsn = "미분류(확인필요)", "⚠️ 확인필요", "더마 세부 카테고리 모호"
            elif any(k in name_clean for k in ["바디워시", "바디미스트", "바디로션", "바디"]):
                ap, st_val, rsn = "바디케어", "✅ 정상매칭" if "바디로션" not in name_clean else "⚠️ 확인필요", "Acnes 바디케어 라인"
            elif any(k in name_clean for k in ["스팟패치", "진정패치", "패치"]):
                if "토너패드" in name_clean:
                    ap, st_val, rsn = "스킨케어", "✅ 정상매칭", "Acnes 스킨케어 토너패드 일치"
                elif "크림패치" in name_clean:
                    ap, st_val, rsn = "스킨케어", "⚠️ 확인필요", "스킨케어 패드 추정(크림패치)"
                elif "화이트티트리" in name_clean:
                    ap, st_val, rsn = "스팟패치", "⚠️ 확인필요", "화이트 티트리 스팟패치 추정"
                else:
                    ap, st_val, rsn = "스팟패치", "✅ 정상매칭", "Acnes 스팟패치 라인 일치"
            elif any(k in name_clean for k in ["세범밸런싱", "모이스처플루이드", "모이스처크림", "토너패드", "마스크", "선크림", "토너", "스팟크림", "크림", "스킨"]):
                ap, st_val, rsn = "스킨케어", "⚠️ 확인필요" if ("화이트티트리" in name_clean or "수딩마스크" in name_clean) else "✅ 정상매칭", "Acnes 스킨케어 라인"
            elif any(k in name_clean for k in ["포밍워시", "클렌저", "크리미워시", "클렌징오일", "클렌징", "포하", "4ha", "워시", "폼"]):
                ap, st_val, rsn = "클렌징", "⚠️ 확인필요" if "화이트티트리" in name_clean else "✅ 정상매칭", "Acnes 클렌징 라인"
            else:
                ap, st_val, rsn = "미분류(확인필요)", "⚠️ 확인필요", "Acnes 포트폴리오 키워드 미일치"
        elif b == "Lipcare":
            if "딥모이스트" in name_clean or "deepmoist" in name_clean:
                ap, st_val, rsn = "딥 모이스트", "✅ 정상매칭", "Lipcare 딥 모이스트 라인 일치"
            elif "메디케이티드" in name_clean or "medicated" in name_clean:
                ap, st_val, rsn = "메디케이티드", "✅ 정상매칭", "Lipcare 메디케이티드 라인 일치"
            elif "소프트립" in name_clean or "softlips" in name_clean:
                ap, st_val, rsn = "소프트립스", "✅ 정상매칭", "Lipcare 소프트립스 라인 일치"
            elif any(k in name_clean for k in ["매직컬러", "매직칼라", "틴티드컬러", "틴티드칼라"]):
                ap, st_val, rsn = "매직컬러 / 틴티드컬러 라인", "✅ 정상매칭", "Lipcare 매직/틴티드 컬러 일치"
            elif "립아이스" in name_clean or "lipice" in name_clean:
                ap, st_val, rsn = "립아이스 프루티", "✅ 정상매칭" if any(k in name_clean for k in ["프루티", "레몬", "스트로베리", "딸기", "그레이프"]) else "⚠️ 확인필요", "립아이스 프루티 라인"
            else:
                ap, st_val, rsn = "미분류(확인필요)", "⚠️ 확인필요", "Lipcare 세부 라인 확인 필요"
        elif b == "HADALABO":
            if "프리미엄" in name_clean or "premium" in name_clean:
                ap, st_val, rsn = "Gokujyun premium (고쿠쥰 프리미엄)", "✅ 정상매칭", "HADALABO 고쿠쥰 프리미엄 일치"
            elif "시로쥰" in name_clean or "shirojyun" in name_clean:
                ap, st_val, rsn = "Shirojyun (시로쥰)", "✅ 정상매칭", "HADALABO 시로쥰 일치"
            elif "키와미즈" in name_clean or "kiwamizu" in name_clean:
                ap, st_val, rsn = "Kiwamizu (키와미즈)", "✅ 정상매칭", "HADALABO 키와미즈 일치"
            elif "타마고하다" in name_clean or "tamagohada" in name_clean:
                ap, st_val, rsn = "Tamagohada (타마고하다)", "✅ 정상매칭", "HADALABO 타마고하다 일치"
            elif "고쿠쥰" in name_clean or "gokujyun" in name_clean:
                ap, st_val, rsn = "Gokujyun (고쿠쥰)", "⚠️ 확인필요" if "탄력" in name_clean else "✅ 정상매칭", "HADALABO 고쿠쥰 라인"
            else:
                ap, st_val, rsn = "미분류(확인필요)", "⚠️ 확인필요", "HADALABO 세부 라인 확인 필요"
        else:
            ap, st_val, rsn = "-", "해당없음", "기타/미분류 브랜드"

        if code in port_ov:
            fp = port_ov[code]
            st_val = "✏️ 사용자수정완료"
        else:
            fp = ap

        brands.append(b)
        auto_ports.append(ap)
        final_ports.append(fp)
        statuses.append(st_val)
        reasons.append(rsn)

    res = unique_sku_df[["상품", "상품명"]].copy()
    res["브랜드"] = brands
    res["자동_포트폴리오"] = auto_ports
    res["최종_포트폴리오"] = final_ports
    res["검수상태"] = statuses
    res["판별근거"] = reasons
    return res


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return buf.getvalue()


def get_template_bytes() -> bytes:
    if os.path.exists(NEW_SKU_TEMPLATE_FILE):
        with open(NEW_SKU_TEMPLATE_FILE, "rb") as f:
            return f.read()
    cols = ["제품코드", "브랜드", "제품명(제품군)", "포트폴리오", "입수량(BOX)", "수입리드타임(일)", "안전재고일수(일)", "비고(채널/규격/특이사항)"]
    sample = [["ME00999NEW", "Acnes", "아크네스 화이트 티트리 시카 세럼 30mL", "스킨케어", 24, 90, 30, "수입 리드타임 반영"]]
    return to_excel_bytes(pd.DataFrame(sample, columns=cols), sheet_name="신규제품_등록")


if "master_df" not in st.session_state:
    st.session_state["master_df"] = load_default_master()
if "custom_skus" not in st.session_state:
    st.session_state["custom_skus"] = {}
if "brand_overrides" not in st.session_state:
    st.session_state["brand_overrides"] = {}
if "portfolio_overrides" not in st.session_state:
    st.session_state["portfolio_overrides"] = {}
if "selected_brand" not in st.session_state:
    st.session_state["selected_brand"] = "Acnes"
if "active_tab" not in st.session_state:
    st.session_state["active_tab"] = "📊 브랜드·포트폴리오 현황"
if "selected_forecast_sku" not in st.session_state:
    st.session_state["selected_forecast_sku"] = ""

top_l, top_r = st.columns([1.35, 1.15])
with top_l:
    st.title("📦 브랜드 및 포트폴리오별 통합 재고 현황")
    st.caption("브랜드/포트폴리오 · 수입리드타임 기반 DOH·안전재고 팔로우업")

with top_r:
    with st.container(border=True):
        hdr_c1, hdr_c2 = st.columns([1.4, 1.0])
        with hdr_c1:
            st.markdown("**➕ 신규 제품코드 즉시 등록**")
        with hdr_c2:
            st.download_button(
                label="📄 신규등록 양식(.xlsx)",
                data=get_template_bytes(),
                file_name=NEW_SKU_TEMPLATE_FILE,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        with st.form("quick_new_sku_form", clear_on_submit=True):
            f_c1, f_c2, f_c3 = st.columns([1.1, 1.1, 1.3])
            with f_c1:
                new_code = st.text_input("제품코드*", placeholder="예: ME00999NEW")
            with f_c2:
                new_brand = st.selectbox("브랜드*", options=BRAND_CHOICES, index=0)
            with f_c3:
                new_name = st.text_input("제품명(제품군)", placeholder="예: 아크네스 화이트 티트리 세럼")

            f_c4, f_c5, f_c6 = st.columns([1.1, 1.0, 1.1])
            with f_c4:
                p_opts = PORTFOLIO_OPTIONS.get(new_brand, ["기본"])
                new_port = st.selectbox("포트폴리오", options=p_opts)
            with f_c5:
                new_lt = st.number_input("수입LT(일)", min_value=1, max_value=365, value=90, step=5)
            with f_c6:
                st.write("")
                submitted_new = st.form_submit_button("⚡ 즉시 연동", type="primary", use_container_width=True)

            if submitted_new:
                clean_c = str(new_code or "").strip()
                if not clean_c:
                    st.error("제품코드를 입력해주세요.")
                else:
                    clean_nm = str(new_name or "").strip() or f"신규등록품목 ({clean_c})"
                    st.session_state["custom_skus"][clean_c] = {
                        "상품명": clean_nm,
                        "브랜드": new_brand,
                        "포트폴리오": new_port,
                        "수입LT": int(new_lt),
                    }
                    st.session_state["brand_overrides"][clean_c] = new_brand
                    st.session_state["portfolio_overrides"][clean_c] = new_port
                    st.session_state["selected_forecast_sku"] = clean_c
                    st.success(f"신규 제품코드 '{clean_c}' ({new_brand})가 즉시 연동되었습니다.")
                    st.rerun()

up_c1, up_c2 = st.columns(2)
with up_c1:
    with st.container(border=True):
        uploaded_inv = st.file_uploader(
            "📂 1. 물류창고 재고파일 업로드 ('재고파일' 시트 양식)",
            type=["xlsx", "csv"],
            key="uploader_inv_main",
        )
with up_c2:
    with st.container(border=True):
        uploaded_out = st.file_uploader(
            "🚚 2. '일일출고 현황' 파일 업로드 (구분·Date·수량 기반 DOH·예상출고 연동)",
            type=["xlsx", "xls", "csv"],
            key="uploader_outbound_main",
        )

if uploaded_inv is None:
    st.info("👆 물류창고 재고파일을 업로드하면 재고 및 SCM 분석이 시작됩니다.")
    st.stop()

try:
    inv_df, meta = parse_warehouse_bytes(uploaded_inv.getvalue(), uploaded_inv.name)
except Exception as e:
    st.error(f"❌ 재고파일 파싱 오류: {e}")
    st.stop()

if st.session_state["custom_skus"]:
    existing_codes = set(inv_df["상품"].unique())
    extra_rows = []
    for c_code, c_info in st.session_state["custom_skus"].items():
        if c_code not in existing_codes:
            extra_rows.append(
                {
                    "센터": "신규등록(대기)",
                    "셀": "-",
                    "상품": c_code,
                    "상품명": c_info.get("상품명", c_code),
                    "화주LOT": "-",
                    "유효일자": "-",
                    "정상환산": 0.0,
                    "불량환산": 0.0,
                    "합계수량": 0.0,
                }
            )
    if extra_rows:
        inv_df = pd.concat([inv_df, pd.DataFrame(extra_rows)], ignore_index=True)

outbound_df = None
monthly_hist_df = pd.DataFrame(columns=["상품", "연월", "출고수량"])
outbound_info = None
if uploaded_out is not None:
    try:
        outbound_df, monthly_hist_df, outbound_info = parse_outbound_bytes(
            uploaded_out.getvalue(), uploaded_out.name
        )
    except Exception as e:
        st.warning(f"⚠️ '일일출고 현황' 파일 파싱 중 참고사항: {e}")

st.success(
    f"✅ {meta['source_label']} 로드 완료 | **분석 대상 {meta['final_rows']:,}건 (고유 SKU {meta['unique_skus']:,}종)**"
    + (
        f" | 🚚 일일출고({outbound_info['div_col']}: {'매출' if outbound_info['used_sales_filter'] else '전체'}) {outbound_info['rows']:,}건 ({outbound_info['date_range']}) 연동됨"
        if outbound_info
        else " | 💡 '일일출고 현황'을 올리면 구분(매출)·Date·수량 기반 DOH·월별 출고가 정밀 산출됩니다."
    )
)

master_df = st.session_state["master_df"]
master_map = dict(zip(master_df["제품코드"], master_df["브랜드"]))
for k_code, k_val in st.session_state["custom_skus"].items():
    master_map[k_code] = k_val.get("브랜드", "미분류")

unique_sku_df = inv_df[["상품", "상품명"]].drop_duplicates(subset=["상품"]).reset_index(drop=True)
for k_code, k_val in st.session_state["custom_skus"].items():
    mask_c = unique_sku_df["상품"] == k_code
    if mask_c.any() and k_val.get("상품명"):
        unique_sku_df.loc[mask_c, "상품명"] = k_val["상품명"]

sku_classified = classify_skus_cached(
    unique_sku_df,
    tuple(sorted(master_map.items())),
    tuple(sorted(st.session_state["brand_overrides"].items())),
    tuple(sorted(st.session_state["portfolio_overrides"].items())),
)

inv_merged = inv_df.merge(
    sku_classified[["상품", "브랜드", "최종_포트폴리오", "검수상태", "판별근거"]],
    on="상품",
    how="left",
)
inv_merged["브랜드"] = inv_merged["브랜드"].fillna("미분류")
inv_merged["최종_포트폴리오"] = inv_merged["최종_포트폴리오"].fillna("-")
inv_merged["검수상태"] = inv_merged["검수상태"].fillna("해당없음")

qty_metric = st.radio(
    "수량 기준 선택",
    options=["정상환산", "합계수량", "불량환산"],
    format_func=lambda x: {
        "정상환산": "정상재고 (환산 EA)",
        "합계수량": "총재고 (정상+불량 합계)",
        "불량환산": "불량재고 (환산 EA)",
    }[x],
    horizontal=True,
)

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("선택 기준 총 재고 (EA)", f"{int(inv_merged[qty_metric].sum()):,}")
k2.metric("Acnes", f"{int(inv_merged[inv_merged['브랜드'] == 'Acnes'][qty_metric].sum()):,}")
k3.metric("Lipcare", f"{int(inv_merged[inv_merged['브랜드'] == 'Lipcare'][qty_metric].sum()):,}")
k4.metric("HADALABO", f"{int(inv_merged[inv_merged['브랜드'] == 'HADALABO'][qty_metric].sum()):,}")
k5.metric("사은품", f"{int(inv_merged[inv_merged['브랜드'] == '사은품'][qty_metric].sum()):,}")

sku_qty = inv_merged.groupby("상품", as_index=False)[["정상환산", "불량환산", "합계수량"]].sum()
sku_view = sku_classified.merge(sku_qty, on="상품", how="left").fillna(0)

if outbound_df is not None and not outbound_df.empty:
    sku_view = sku_view.merge(outbound_df, on="상품", how="left")
    sku_view["누적_출고수량"] = sku_view["누적_출고수량"].fillna(0.0)
    sku_view["출고_발생일수"] = sku_view["출고_발생일수"].fillna(0).astype(int)
    sku_view["일평균_출고량"] = sku_view["일평균_출고량"].fillna(0.0)
    sku_view["월예상_출고량"] = sku_view["월예상_출고량"].fillna(0.0)
else:
    sku_view["누적_출고수량"] = 0.0
    sku_view["출고_발생일수"] = 0
    sku_view["일평균_출고량"] = 0.0
    sku_view["월예상_출고량"] = 0.0

v_col1, v_col2 = st.columns([1, 1])
with v_col1:
    if st.button(
        "📊 브랜드 · 포트폴리오 현황 & 매핑 수정",
        type="primary" if st.session_state["active_tab"] == "📊 브랜드·포트폴리오 현황" else "secondary",
        use_container_width=True,
    ):
        st.session_state["active_tab"] = "📊 브랜드·포트폴리오 현황"
        st.rerun()
with v_col2:
    if st.button(
        "🛡️ SCM 팔로우업: 커버리지(DOH) & 안전재고(Safety Stock) 알림",
        type="primary" if st.session_state["active_tab"] == "🛡️ SCM 팔로우업" else "secondary",
        use_container_width=True,
    ):
        st.session_state["active_tab"] = "🛡️ SCM 팔로우업"
        st.rerun()

if st.session_state["active_tab"] == "📊 브랜드·포트폴리오 현황":
    unassigned_sku_cnt = int((sku_view["브랜드"] == "미분류").sum())

    with st.expander(
        f"🛠️ SKU별 브랜드 매핑 수정 (고유 품목 {len(sku_view)}종 | 미분류 {unassigned_sku_cnt}종)",
        expanded=(unassigned_sku_cnt > 0),
    ):
        m_c1, m_c2 = st.columns([2, 1])
        with m_c1:
            q_search = st.text_input("🔍 제품코드 또는 제품명 검색", placeholder="예: ME00621DCP, 아크네스, 립, 하다라보...")
        with m_c2:
            q_only_un = st.checkbox("미분류 품목만 보기", value=False)

        edit_df = sku_view[["상품", "상품명", "브랜드"]].copy()
        if q_search:
            mask = edit_df["상품"].str.contains(q_search, case=False, na=False) | edit_df["상품명"].str.contains(
                q_search, case=False, na=False
            )
            edit_df = edit_df[mask]
        if q_only_un:
            edit_df = edit_df[edit_df["브랜드"] == "미분류"]

        edited_b = st.data_editor(
            edit_df,
            column_config={
                "상품": st.column_config.TextColumn("제품코드", disabled=True),
                "상품명": st.column_config.TextColumn("제품명", disabled=True),
                "브랜드": st.column_config.SelectboxColumn("브랜드 (클릭 수정)", options=BRAND_CHOICES, required=True),
            },
            use_container_width=True,
            hide_index=True,
            height=240,
            key="sku_brand_editor_v9",
        )

        changed_pairs = {}
        orig_b_map = dict(zip(edit_df["상품"], edit_df["브랜드"]))
        for r in edited_b.itertuples(index=False):
            if r.브랜드 != orig_b_map.get(r.상품):
                changed_pairs[r.상품] = r.브랜드
        if changed_pairs:
            st.session_state["brand_overrides"].update(changed_pairs)
            st.rerun()

        updated_master_export = sku_view[["상품", "브랜드", "상품명"]].rename(
            columns={"상품": "제품코드", "상품명": "제품군"}
        )
        st.download_button(
            label="💾 반영된 '브랜드_제품코드_마스터.xlsx' 다운로드",
            data=to_excel_bytes(updated_master_export, sheet_name="Sheet1"),
            file_name=DEFAULT_MASTER_FILE,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.markdown('<div class="section-title">📊 브랜드별 전체 재고 현황</div>', unsafe_allow_html=True)
    brand_sum = (
        inv_merged.groupby("브랜드", as_index=False)[qty_metric]
        .sum()
        .sort_values(by=qty_metric, ascending=False)
    )
    st.bar_chart(
        data=brand_sum.rename(columns={qty_metric: "재고수량(EA)"}),
        x="브랜드",
        y="재고수량(EA)",
        color="브랜드",
        height=360,
        use_container_width=True,
    )

    st.markdown('<div class="section-title">🔎 브랜드 선택 및 포트폴리오 · 상세 재고 조회</div>', unsafe_allow_html=True)
    brands_list = brand_sum["브랜드"].tolist()
    if st.session_state["selected_brand"] not in brands_list and brands_list:
        st.session_state["selected_brand"] = brands_list[0]

    b_cols = st.columns(min(len(brands_list), 6))
    for i, b_name in enumerate(brands_list):
        target_col = b_cols[i % len(b_cols)]
        b_val = int(brand_sum.loc[brand_sum["브랜드"] == b_name, qty_metric].values[0])
        is_sel = st.session_state["selected_brand"] == b_name
        if target_col.button(
            f"{'✅ ' if is_sel else ''}{b_name} ({b_val:,})",
            key=f"sel_b_{b_name}",
            type="primary" if is_sel else "secondary",
            use_container_width=True,
        ):
            st.session_state["selected_brand"] = b_name
            st.rerun()

    sel_brand = st.session_state["selected_brand"]
    brand_rows = inv_merged[inv_merged["브랜드"] == sel_brand]
    st.markdown(f"### 📌 [{sel_brand}] 상세 재고 및 포트폴리오 분석")

    if sel_brand in PORTFOLIO_OPTIONS:
        c_left, c_right = st.columns([1.15, 1.85])
        with c_left:
            p_sum = (
                brand_rows.groupby("최종_포트폴리오", as_index=False)[qty_metric]
                .sum()
                .sort_values(by=qty_metric, ascending=False)
            )
            st.bar_chart(
                data=p_sum.rename(columns={"최종_포트폴리오": "포트폴리오", qty_metric: "재고수량(EA)"}),
                x="포트폴리오",
                y="재고수량(EA)",
                color="포트폴리오",
                height=320,
                use_container_width=True,
            )

        with c_right:
            b_sku_df = sku_view[sku_view["브랜드"] == sel_brand][
                ["상품", "상품명", "최종_포트폴리오", "검수상태", "판별근거", qty_metric]
            ].sort_values(by=["검수상태", qty_metric], ascending=[True, False])

            need_cnt = int((b_sku_df["검수상태"] == "⚠️ 확인필요").sum())
            if need_cnt > 0:
                st.warning(f"⚠️ **{sel_brand}** 내 확인이 필요한 품목 **{need_cnt}종**이 있습니다. 아래 표에서 포트폴리오를 수정하세요.")
            else:
                st.success(f"✅ **{sel_brand}** 포트폴리오 분류가 완료되었습니다.")

            only_rev = st.checkbox("⚠️ 확인필요 품목만 보기", value=(need_cnt > 0), key=f"only_rev_{sel_brand}")
            show_p = b_sku_df[b_sku_df["검수상태"] == "⚠️ 확인필요"].copy() if only_rev else b_sku_df.copy()

            edited_p = st.data_editor(
                show_p,
                column_config={
                    "상품": st.column_config.TextColumn("제품코드", disabled=True),
                    "상품명": st.column_config.TextColumn("제품명", disabled=True),
                    "최종_포트폴리오": st.column_config.SelectboxColumn(
                        "포트폴리오 (클릭 수정)", options=PORTFOLIO_OPTIONS[sel_brand], required=True
                    ),
                    "검수상태": st.column_config.TextColumn("검수상태", disabled=True),
                    "판별근거": st.column_config.TextColumn("판별근거", disabled=True),
                    qty_metric: st.column_config.NumberColumn("재고수량", format="%d", disabled=True),
                },
                use_container_width=True,
                hide_index=True,
                height=260,
                key=f"p_edit_{sel_brand}",
            )

            orig_p_map = dict(zip(show_p["상품"], show_p["최종_포트폴리오"]))
            p_changes = {}
            for r in edited_p.itertuples(index=False):
                if r.최종_포트폴리오 != orig_p_map.get(r.상품):
                    p_changes[r.상품] = r.최종_포트폴리오
            if p_changes:
                st.session_state["portfolio_overrides"].update(p_changes)
                st.rerun()

    st.markdown(f"#### 📋 [{sel_brand}] 상세 재고 내역 (센터 · 셀 · LOT · 유효일자별)")
    detail_cols = ["센터", "셀", "상품", "상품명", "최종_포트폴리오", "검수상태", "화주LOT", "유효일자", "정상환산", "불량환산", "합계수량"]
    detail_df = brand_rows[detail_cols].sort_values(by=["최종_포트폴리오", "상품", "유효일자"])
    st.dataframe(detail_df, use_container_width=True, hide_index=True, height=340)

    me_s_col1, me_s_col2 = st.columns([1.8, 1.2])
    with me_s_col1:
        me_query = st.text_input(
            "🔎 ME코드 검색 (입력 시 해당 코드가 하단 '제품코드별 추후 월별 예상 출고량 및 예상 재고 추이' 선 그래프와 연동됩니다)",
            value=st.session_state.get("selected_forecast_sku", ""),
            placeholder="예: ME90421001, ME00621DCP",
            key="me_code_search_below_detail",
        )
    with me_s_col2:
        matched_me_codes = (
            [
                c
                for c in detail_df["상품"].drop_duplicates().tolist()
                if str(me_query or "").strip().upper() in str(c).upper()
            ]
            if str(me_query or "").strip()
            else detail_df["상품"].drop_duplicates().tolist()
        )
        if matched_me_codes and str(me_query or "").strip():
            st.session_state["selected_forecast_sku"] = matched_me_codes[0]
            st.caption(f"✅ 연동 코드: **{matched_me_codes[0]}** (총 {len(matched_me_codes)}건 일치)")
        elif str(me_query or "").strip():
            st.caption("⚠️ 일치하는 ME코드가 없습니다.")

    if str(me_query or "").strip() and matched_me_codes:
        filtered_detail_by_me = detail_df[detail_df["상품"].isin(matched_me_codes)]
        st.dataframe(filtered_detail_by_me, use_container_width=True, hide_index=True, height=180)

    st.download_button(
        label=f"📥 [{sel_brand}] 상세 재고내역 엑셀 다운로드",
        data=to_excel_bytes(detail_df, sheet_name=f"{sel_brand}_재고내역"),
        file_name=f"{sel_brand}_상세재고내역.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.markdown('<div class="section-title">🛡️ 출고/수주 대비 커버리지(DOH) 및 안전재고(Safety Stock) 알림</div>', unsafe_allow_html=True)

    if outbound_df is None:
        st.info(
            "💡 현재 '일일출고 현황' 파일이 미업로드 상태입니다. 상단 **2. '일일출고 현황' 파일 업로드**에 출고 데이터를 올리면 "
            "구분(매출)·Date·수량 기준으로 DOH와 안전재고 미달 알림이 계산됩니다."
        )

    p_c1, p_c2, p_c3 = st.columns(3)
    with p_c1:
        import_lt_days = st.slider(
            "외국계 화장품 수입입고 리드타임(일)",
            min_value=30,
            max_value=180,
            value=90,
            step=5,
        )
    with p_c2:
        safety_days = st.slider(
            "안전재고 기준 일수(일)",
            min_value=7,
            max_value=90,
            value=30,
            step=5,
        )
    with p_c3:
        danger_days = st.slider(
            "위험(결품·긴급발주) 커버리지 기준(일 미만)",
            min_value=15,
            max_value=120,
            value=60,
            step=5,
        )

    scm_df = sku_view.copy()
    scm_df["안전재고_수량(EA)"] = (scm_df["일평균_출고량"] * float(safety_days)).round(0)
    scm_df["수입발주점(ROP, EA)"] = (scm_df["일평균_출고량"] * float(import_lt_days + safety_days)).round(0)
    scm_df["커버리지(DOH, 일)"] = scm_df.apply(
        lambda r: round(r["정상환산"] / r["일평균_출고량"], 1)
        if r["일평균_출고량"] > 0
        else (999.0 if r["정상환산"] > 0 else 0.0),
        axis=1,
    )
    scm_df["안전재고_과부족(EA)"] = (scm_df["정상환산"] - scm_df["안전재고_수량(EA)"]).round(0)

    def _doh_status(r):
        if r["일평균_출고량"] <= 0:
            return "🟢 출고없음(재고확인)"
        doh = r["커버리지(DOH, 일)"]
        if doh < danger_days or r["정상환산"] < r["안전재고_수량(EA)"]:
            return "🚨 위험(수입리드타임 대비 부족)"
        if doh < (import_lt_days + safety_days):
            return "⚠️ 주의(수입발주 검토 필요)"
        return "✅ 안정(수입리드타임 확보)"

    scm_df["SCM_알림상태"] = scm_df.apply(_doh_status, axis=1)

    danger_cnt = int((scm_df["SCM_알림상태"].str.startswith("🚨")).sum())
    caution_cnt = int((scm_df["SCM_알림상태"].str.startswith("⚠️")).sum())
    safe_cnt = int((scm_df["SCM_알림상태"].str.startswith("✅")).sum())
    ss_short_cnt = int((scm_df["안전재고_과부족(EA)"] < 0).sum())

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("🚨 위험 품목(DOH < 위험기준)", f"{danger_cnt:,}종")
    m2.metric("⚠️ 주의 품목(발주검토 필요)", f"{caution_cnt:,}종")
    m3.metric("🛡️ 안전재고 미달 품목", f"{ss_short_cnt:,}종")
    m4.metric("✅ 안정 품목", f"{safe_cnt:,}종")

    st.markdown('<div class="section-title">📉 수입입고 리드타임 고려 커버리지(DOH) 위험 수준 그래프</div>', unsafe_allow_html=True)
    active_doh = scm_df[scm_df["일평균_출고량"] > 0].sort_values(by="커버리지(DOH, 일)", ascending=True).head(30).copy()
    if active_doh.empty:
        active_doh = scm_df.sort_values(by="커버리지(DOH, 일)", ascending=True).head(30).copy()
    active_doh["DOH_표시(일)"] = active_doh["커버리지(DOH, 일)"].clip(upper=365.0)
    active_doh["위험기준선(일)"] = float(danger_days)
    active_doh["수입리드타임선(일)"] = float(import_lt_days)
    st.line_chart(
        active_doh.set_index("상품")[["DOH_표시(일)", "위험기준선(일)", "수입리드타임선(일)"]],
        height=360,
        use_container_width=True,
    )

    filter_status = st.multiselect(
        "알림 상태 필터",
        options=["🚨 위험(수입리드타임 대비 부족)", "⚠️ 주의(수입발주 검토 필요)", "✅ 안정(수입리드타임 확보)", "🟢 출고없음(재고확인)"],
        default=["🚨 위험(수입리드타임 대비 부족)", "⚠️ 주의(수입발주 검토 필요)"],
    )
    filtered_scm = scm_df[scm_df["SCM_알림상태"].isin(filter_status)] if filter_status else scm_df
    filtered_scm = filtered_scm.sort_values(by=["커버리지(DOH, 일)", "안전재고_과부족(EA)"], ascending=[True, True])

    scm_cols = [
        "SCM_알림상태",
        "상품",
        "상품명",
        "브랜드",
        "정상환산",
        "일평균_출고량",
        "월예상_출고량",
        "커버리지(DOH, 일)",
        "안전재고_수량(EA)",
        "안전재고_과부족(EA)",
        "수입발주점(ROP, EA)",
    ]
    st.dataframe(filtered_scm[scm_cols], use_container_width=True, hide_index=True, height=320)

    st.download_button(
        label="📥 SCM 커버리지(DOH)·안전재고 알림 엑셀 다운로드",
        data=to_excel_bytes(filtered_scm[scm_cols], sheet_name="DOH_안전재고_알림"),
        file_name="SCM_커버리지_DOH_안전재고_알림.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown('<div class="section-title">📈 제품코드별 추후 월별 예상 출고량 및 예상 재고 추이(선 그래프)</div>', unsafe_allow_html=True)

sku_list = sku_view["상품"].tolist()
sku_label_map = {
    row.상품: f"{row.상품} | {row.브랜드} | {row.상품명}"
    for row in sku_view.itertuples(index=False)
}

default_sku = (
    st.session_state["selected_forecast_sku"]
    if st.session_state["selected_forecast_sku"] in sku_list
    else (sku_list[0] if sku_list else "")
)

fc_c1, fc_c2, fc_c3 = st.columns([2.2, 1.0, 1.0])
with fc_c1:
    picked_sku = st.selectbox(
        "제품코드 선택 (상세 재고 내역의 ME코드 검색과 연동됩니다)",
        options=sku_list,
        index=sku_list.index(default_sku) if default_sku in sku_list else 0,
        format_func=lambda c: sku_label_map.get(c, c),
    )
    st.session_state["selected_forecast_sku"] = picked_sku

with fc_c2:
    forecast_months = st.slider("예측 기간(개월)", min_value=3, max_value=12, value=6, step=1)

with fc_c3:
    growth_rate_pct = st.number_input("월 출고 성장률(%)", min_value=-50.0, max_value=100.0, value=0.0, step=5.0)

if picked_sku:
    sku_row = sku_view[sku_view["상품"] == picked_sku].iloc[0]
    cur_stock = float(sku_row["정상환산"])

    sku_m_hist = (
        monthly_hist_df[monthly_hist_df["상품"] == picked_sku].sort_values("연월").copy()
        if not monthly_hist_df.empty
        else pd.DataFrame(columns=["상품", "연월", "출고수량"])
    )

    if outbound_df is not None and not sku_m_hist.empty:
        base_monthly_out = float(sku_m_hist["출고수량"].mean())
        if outbound_info and outbound_info.get("denom_days", 30) < 25:
            daily_based = float(sku_row["일평균_출고량"]) * 30.0
            base_monthly_out = max(base_monthly_out, daily_based) if base_monthly_out > 0 else daily_based
    elif outbound_df is not None:
        base_monthly_out = float(sku_row["일평균_출고량"]) * 30.0
    else:
        base_monthly_out = float(sku_row["월예상_출고량"])

    today_dt = date.today()
    rows_timeline = []

    if not sku_m_hist.empty:
        for hr in sku_m_hist.itertuples(index=False):
            rows_timeline.append(
                {
                    "기간": f"{hr.연월} (실적)",
                    "월별 출고량(EA)": round(float(hr.출고수량), 0),
                    "예상 잔여재고(EA)": None,
                }
            )

    running_stock = cur_stock
    for m_idx in range(1, int(forecast_months) + 1):
        yr = today_dt.year + (today_dt.month + m_idx - 1) // 12
        mo = (today_dt.month + m_idx - 1) % 12 + 1
        m_label = f"{yr}-{mo:02d} (예상)"
        m_out = base_monthly_out * ((1.0 + float(growth_rate_pct) / 100.0) ** (m_idx - 1))
        running_stock = max(0.0, running_stock - m_out)

        rows_timeline.append(
            {
                "기간": m_label,
                "월별 출고량(EA)": round(m_out, 0),
                "예상 잔여재고(EA)": round(running_stock, 0),
            }
        )

    fc_chart_df = pd.DataFrame(rows_timeline).set_index("기간")

    fc_m1, fc_m2, fc_m3 = st.columns(3)
    fc_m1.metric("선택 제품코드", picked_sku)
    fc_m2.metric("현재 정상재고(EA)", f"{int(cur_stock):,}")
    fc_m3.metric("기준 월 출고량(EA)", f"{int(round(base_monthly_out)):,}")

    st.line_chart(fc_chart_df, height=340, use_container_width=True)
    st.dataframe(fc_chart_df.reset_index(), use_container_width=True, hide_index=True, height=240)
