import io
import os
import re
from datetime import date, timedelta
import altair as alt
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

FORCE_LIPCARE_CODES = {"ME90521DME", "ME90521WC4", "ME90521WC5"}

LOTION_CODE_MIGRATION = {
    "ME90421075": "ME90421001",
    "ME90421100": "ME90421002",
    "ME90421450": "ME90421003",
}

EXCLUDED_LOTION_DANGER_CODES = {
    "ME90421075",
    "ME90421001",
    "ME90421100",
    "ME90421002",
    "ME90421450",
    "ME90421003",
}


def map_lotion_code(code: str) -> str:
    c = str(code or "").strip()
    if c in LOTION_CODE_MIGRATION:
        return LOTION_CODE_MIGRATION[c]
    if c.startswith("ME") and len(c) >= 6:
        if c.endswith("1075"):
            return c[:-4] + "1001"
        if c.endswith("1100"):
            return c[:-4] + "1002"
        if c.endswith("1450"):
            return c[:-4] + "1003"
    return c


def is_migrated_lotion_code(code: str) -> bool:
    c = str(code or "").strip().upper()
    if c in EXCLUDED_LOTION_DANGER_CODES:
        return True
    if c.startswith("ME") and len(c) >= 6 and c.endswith(("1075", "1001", "1100", "1002", "1450", "1003")):
        return True
    return False


def _shift_years_safe(dt_obj: date, years_delta: int) -> date:
    target_year = dt_obj.year + years_delta
    try:
        return dt_obj.replace(year=target_year)
    except ValueError:
        # 2월 29일 등 윤년 보정
        return dt_obj.replace(year=target_year, month=2, day=28)


def evaluate_shelf_life(mfg_val, exp_val, ref_date: date | None = None, shelf_life_years: int = 3) -> dict:
    """
    유통기한은 제조일로부터 3년(기본) 기준.
    잔여일수 비율이 50% 미만이면 '유효일자 관리 품목'으로 분류.
    """
    today_d = ref_date or date.today()
    exp_ts = pd.to_datetime(exp_val, errors="coerce")
    if pd.isna(exp_ts):
        return {
            "잔여일수": None,
            "총수명일수": None,
            "잔여비율(%)": None,
            "유효일자_관리구분": "⚪ 일자미기재",
            "유효일자_관리대상": False,
        }

    exp_d = exp_ts.date()
    mfg_ts = pd.to_datetime(mfg_val, errors="coerce")
    has_valid_mfg = (
        pd.notna(mfg_ts)
        and mfg_ts.year >= 1950
        and mfg_ts.date() < exp_d
    )

    if has_valid_mfg:
        mfg_d = mfg_ts.date()
        total_days = max((exp_d - mfg_d).days, 1)
    else:
        mfg_d = _shift_years_safe(exp_d, -int(shelf_life_years))
        total_days = max((exp_d - mfg_d).days, 1)

    rem_days = (exp_d - today_d).days
    rem_ratio_pct = round((rem_days / float(total_days)) * 100.0, 1)

    if rem_days < 0:
        status_label = "🔴 만료경과(관리대상)"
        is_manage = True
    elif rem_ratio_pct < 50.0:
        status_label = "⏳ 유효일자 관리(50% 미만)"
        is_manage = True
    else:
        status_label = "✅ 정상(50% 이상)"
        is_manage = False

    return {
        "잔여일수": int(rem_days),
        "총수명일수": int(total_days),
        "잔여비율(%)": rem_ratio_pct,
        "유효일자_관리구분": status_label,
        "유효일자_관리대상": is_manage,
    }


PORTFOLIO_OPTIONS = {
    "Lotion": ["75ml", "100ml", "450ml", "미분류(확인필요)"],
    "Spray": ["쿨", "딥앤쿨", "익스트림", "미분류(확인필요)"],
    "HADALABO": ["고쿠쥰", "프리미엄", "탄력", "시로쥰", "키와미즈", "타마고하다", "미분류(확인필요)"],
    "Roll on Lotion": ["핫", "쿨", "핫&쿨 번들", "미분류(확인필요)"],
    "Hand": ["리치네일", "리치네일 리페어코트", "기획세트", "미분류(확인필요)"],
    "Melano": ["엔자임페이스 워시", "안티스팟 화이트닝에센스", "미분류(확인필요)"],
    "Acnes": ["다이소 출고 제품", "스킨케어", "클렌징", "스팟패치", "바디케어", "포 맨", "더마 스킨케어", "미분류(확인필요)"],
    "Lipcare": ["립아이스 프루티", "딥 모이스트", "매직컬러 / 틴티드컬러 라인", "메디케이티드", "소프트립스", "미분류(확인필요)"],
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
    "단종 예상 재고",
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
        qty_col = "정상수량_환산" if "정상수량_환산" in df.columns else ("환산" if "환산" in df.columns else None)
        out = pd.DataFrame(
            {
                "센터": df["센터"].fillna("-").astype(str).str.strip() if "센터" in df.columns else "-",
                "셀": df["셀"].fillna("-").astype(str).str.strip() if "셀" in df.columns else "-",
                "상품": df["상품"].astype(str).str.strip() if "상품" in df.columns else "",
                "상품명": df["상품명"].fillna("").astype(str).str.strip() if "상품명" in df.columns else "",
                "화주LOT": df["화주LOT"].fillna("-").astype(str).str.strip() if "화주LOT" in df.columns else "-",
                "제조일자_raw": df["제조일자"] if "제조일자" in df.columns else None,
                "유효일자_raw": df["유효일자"] if "유효일자" in df.columns else None,
                "정상환산": pd.to_numeric(df[qty_col], errors="coerce").fillna(0.0) if qty_col else 0.0,
                "불량환산": pd.to_numeric(df["불량수량_환산"], errors="coerce").fillna(0.0)
                if "불량수량_환산" in df.columns
                else 0.0,
                "합계수량_raw": pd.to_numeric(df["합계수량"], errors="coerce") if "합계수량" in df.columns else None,
            }
        )
        raw_total_rows = len(out)
        source_label = "CSV 재고파일"
    else:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), read_only=True, data_only=True)
        sheet_names = wb.sheetnames
        sheet_name = (
            "재고파일" if "재고파일" in sheet_names else ("sheet1" if "sheet1" in sheet_names else sheet_names[0])
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
        idx_mfg = _find_col_index(row1, row2, "제조일자", 12)
        idx_exp = _find_col_index(row1, row2, "유효일자", 13)
        idx_total = _find_col_index(row1, row2, "합계수량", 31)

        idx_norm_conv = 27
        for idx, r1_val in enumerate(row1):
            if r1_val is not None and "정상수량" in str(r1_val).strip():
                for off in range(4):
                    c_pos = idx + off
                    if c_pos < len(row2) and row2[c_pos] is not None and "환산" in str(row2[c_pos]).strip():
                        idx_norm_conv = c_pos
                        break
                break

        idx_bad_conv = 30
        for idx, r1_val in enumerate(row1):
            if r1_val is not None and "불량수량" in str(r1_val).strip():
                for off in range(4):
                    c_pos = idx + off
                    if c_pos < len(row2) and row2[c_pos] is not None and "환산" in str(row2[c_pos]).strip():
                        idx_bad_conv = c_pos
                        break
                break

        has_sub_header = any(
            v is not None and str(v).strip() in {"BOX", "EA", "환산", "창고", "LOCATION"} for v in row2
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
                    "제조일자_raw": r[idx_mfg] if idx_mfg < len(r) else None,
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

    out = out[out["상품"].notna() & (out["상품"] != "") & (out["상품"].str.lower() != "nan")].copy()
    out["센터"] = out["센터"].replace({"": "-", "nan": "-", "None": "-"})
    out["셀"] = out["셀"].replace({"": "-", "nan": "-", "None": "-"})
    out["화주LOT"] = out["화주LOT"].replace({"": "-", "nan": "-", "None": "-"})

    out["정상환산"] = pd.to_numeric(out["정상환산"], errors="coerce").fillna(0.0)
    out["불량환산"] = pd.to_numeric(out["불량환산"], errors="coerce").fillna(0.0)
    tot_num = pd.to_numeric(out["합계수량_raw"], errors="coerce")
    out["합계수량"] = tot_num.fillna(out["정상환산"] + out["불량환산"])

    mfg_dt = pd.to_datetime(out["제조일자_raw"], errors="coerce")
    raw_mfg_str = out["제조일자_raw"].astype(str).str.strip().replace({"None": "-", "nan": "-", "NaT": "-", "": "-"})
    out["제조일자"] = mfg_dt.dt.strftime("%Y-%m-%d").fillna(raw_mfg_str)

    exp_dt = pd.to_datetime(out["유효일자_raw"], errors="coerce")
    raw_exp_str = out["유효일자_raw"].astype(str).str.strip().replace({"None": "-", "nan": "-", "NaT": "-", "": "-"})
    out["유효일자"] = exp_dt.dt.strftime("%Y-%m-%d").fillna(raw_exp_str)

    excl_pattern = "|".join(re.escape(k) for k in EXCLUDED_KEYWORDS)
    excl_mask = out["상품명"].str.contains(excl_pattern, case=False, na=False)
    excluded_rows = int(excl_mask.sum())
    excluded_qty = float(out.loc[excl_mask, "합계수량"].sum())

    filtered = out[~excl_mask].drop(columns=["합계수량_raw"]).reset_index(drop=True)

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
            "원본상품": df[code_col].astype(str).str.strip(),
            "출고_상품명": df[name_col].fillna("").astype(str).str.strip() if name_col else "",
            "구분": df[div_col].fillna("미구분").astype(str).str.strip() if div_col else "매출",
            "출고수량": pd.to_numeric(df[qty_col], errors="coerce").fillna(0.0),
        }
    )
    cleaned = cleaned[
        cleaned["원본상품"].notna()
        & (cleaned["원본상품"] != "")
        & (cleaned["원본상품"].str.lower() != "nan")
        & (~cleaned["원본상품"].isin(["상품코드", "제품코드", "품목코드", "상품", "합계", "총계"]))
    ].copy()

    cleaned["상품"] = cleaned["원본상품"].apply(map_lotion_code)

    if date_col:
        dt_s = pd.to_datetime(df.loc[cleaned.index, date_col], errors="coerce")
    else:
        dt_s = pd.Series(pd.NaT, index=cleaned.index)
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

        ref_dt = max_dt.normalize()
        cutoff_6m = ref_dt - pd.DateOffset(months=6)
        recent_rows = active_out[(active_out["출고일자"] >= cutoff_6m) & (active_out["출고수량"] > 0)]
        active_6m_codes = set(recent_rows["상품"].unique()) | set(recent_rows["원본상품"].unique())
    else:
        unique_days = 1
        span_days = 30
        denom_days = 30.0
        date_range_str = "날짜열 없음(30일 환산)"
        recent_rows = active_out[active_out["출고수량"] > 0]
        active_6m_codes = set(recent_rows["상품"].unique()) | set(recent_rows["원본상품"].unique())

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
        "date_col": date_col or "Date/일자열 없음",
        "qty_col": qty_col,
        "active_6m_codes": tuple(sorted(active_6m_codes)),
        "has_outbound": True,
    }
    return sku_outbound, monthly_hist, info


@st.cache_data(show_spinner=False)
def parse_nondelivery_bytes(file_bytes: bytes) -> tuple[pd.DataFrame, dict]:
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    sheet_name = "미납내역(RAW)" if "미납내역(RAW)" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet_name)
    df.columns = [str(c).replace("\n", " ").strip() for c in df.columns]

    amt_col = next((c for c in df.columns if "미납금액" in c), None)
    qty_col = next((c for c in df.columns if "미납" in c and "수량" in c), None)
    reason_col = next((c for c in df.columns if "미납사유" in c), None)
    ch_col = next((c for c in df.columns if "CHANNEL" in c.upper() or "채널" in c), None)
    dt_col = next((c for c in df.columns if "납품일자" in c or "일자" in c), None)

    out = df.copy()
    out["미납금액_num"] = pd.to_numeric(out[amt_col], errors="coerce").fillna(0.0) if amt_col else 0.0
    out["미납수량_num"] = pd.to_numeric(out[qty_col], errors="coerce").fillna(0.0) if qty_col else 0.0
    out["미납사유_clean"] = out[reason_col].fillna("미기재").astype(str).str.strip() if reason_col else "미기재"
    out["CHANNEL_clean"] = out[ch_col].fillna("미구분").astype(str).str.strip() if ch_col else "미구분"

    if dt_col:
        dt_parsed = pd.to_datetime(out[dt_col], errors="coerce").dropna()
        months_cnt = max(int(dt_parsed.dt.to_period("M").nunique()), 1) if not dt_parsed.empty else 44
        period_str = (
            f"{dt_parsed.min().strftime('%Y-%m')} ~ {dt_parsed.max().strftime('%Y-%m')} ({months_cnt}개월)"
            if not dt_parsed.empty
            else "기간정보 없음(44개월 기준)"
        )
    else:
        months_cnt = 44
        period_str = "기간정보 없음(44개월 기준)"

    def _root_cause(s: str) -> str:
        if "수입" in s:
            return "수입 지연/이슈"
        if any(k in s for k in ["임가공", "덧방", "시험성적서", "입수변경", "불량이슈"]):
            return "임가공/생산/품질이슈"
        if any(k in s for k in ["재고", "단종"]):
            return "재고부족/단종"
        if any(k in s for k in ["납품 기준", "발주", "단가"]):
            return "발주/납품기준/단가이슈"
        if any(k in s for k in ["소터", "파손", "배송", "오이고", "거점", "오피킹", "오분배", "유통기한", "바코드", "표식지"]):
            return "물류/센터/출고이슈"
        if s == "미기재":
            return "미기재"
        return "기타"

    out["원인대분류"] = out["미납사유_clean"].apply(_root_cause)
    total_amt = float(out["미납금액_num"].sum())
    annualized_amt = (total_amt / float(months_cnt)) * 12.0

    meta = {
        "rows": len(out),
        "total_amt": total_amt,
        "total_qty": float(out["미납수량_num"].sum()),
        "months_cnt": months_cnt,
        "period_str": period_str,
        "annualized_amt": annualized_amt,
    }
    return out, meta


@st.cache_data(show_spinner=False)
def classify_skus_cached(
    unique_sku_df: pd.DataFrame,
    master_records: tuple[tuple[str, str], ...],
    brand_overrides_tuple: tuple[tuple[str, str], ...],
    port_overrides_tuple: tuple[tuple[str, str], ...],
    active_6m_codes_tuple: tuple[str, ...],
    has_outbound_data: bool,
) -> pd.DataFrame:
    master_map = dict(master_records)
    brand_ov = dict(brand_overrides_tuple)
    port_ov = dict(port_overrides_tuple)
    active_6m_set = set(active_6m_codes_tuple)

    brands = []
    auto_ports = []
    final_ports = []
    statuses = []
    reasons = []

    for row in unique_sku_df.itertuples(index=False):
        code = str(row.상품 or "").strip()
        code_upper = code.upper()
        name = str(row.상품명 or "").strip()
        name_clean = re.sub(r"\s+", "", name).lower()
        is_gift_name = any(k in name_clean for k in GIFT_KEYWORDS)

        if code in brand_ov:
            b = normalize_brand_name(brand_ov[code])
        elif "용기" in name_clean:
            b = "부자재"
        elif code_upper in FORCE_LIPCARE_CODES:
            b = "Lipcare"
        elif is_gift_name:
            b = "사은품"
        else:
            m_brand = normalize_brand_name(master_map.get(code, "미분류"))
            if m_brand in ("미분류", "부자재", "사은품"):
                if "아크네스" in name_clean or "acnes" in name_clean:
                    b = "Acnes"
                elif "립" in name_clean or "lip" in name_clean:
                    b = "Lipcare"
                elif "하다라보" in name_clean or "hadalabo" in name_clean:
                    b = "HADALABO"
                elif "멜라노" in name_clean or "melano" in name_clean:
                    b = "Melano"
                elif "롤온" in name_clean:
                    b = "Roll on Lotion"
                elif "핸드베일" in name_clean or "리치네일" in name_clean:
                    b = "Hand"
                elif "스프레이" in name_clean:
                    b = "Spray"
                elif "로션" in name_clean and "멘소래담" in name_clean:
                    b = "Lotion"
                else:
                    b = m_brand
            else:
                b = m_brand

            mapped_c = map_lotion_code(code)
            if (
                has_outbound_data
                and code_upper.startswith("ME")
                and b not in {"부자재", "사은품"}
                and (code not in active_6m_set)
                and (mapped_c not in active_6m_set)
            ):
                b = "단종 예상 재고"

        if b == "부자재":
            ap, st_val, rsn = "부자재", "✅ 정상매칭", "제품명에 '용기' 포함 → 부자재 자동 분류"
        elif b == "단종 예상 재고":
            ap, st_val, rsn = "단종 예상 재고", "⚠️ 단종예상", "ME코드 보유 품목 중 최근 6개월 출고 이력 없음"
        elif b == "사은품":
            ap, st_val, rsn = "사은품", "✅ 정상매칭", "증정용/사은품 키워드 포함 → 사은품 분류"
        elif b == "Lotion":
            if "75ml" in name_clean or code_upper.endswith(("1075", "1001", "175S")):
                ap, st_val, rsn = "75ml", "✅ 정상매칭", "Lotion 75ml 용량 분류"
            elif "100ml" in name_clean or code_upper.endswith(("1100", "1002", "110S")):
                ap, st_val, rsn = "100ml", "✅ 정상매칭", "Lotion 100ml 용량 분류"
            elif "450ml" in name_clean or code_upper.endswith(("1450", "1003")):
                ap, st_val, rsn = "450ml", "✅ 정상매칭", "Lotion 450ml 용량 분류"
            else:
                ap, st_val, rsn = "미분류(확인필요)", "⚠️ 확인필요", "Lotion 용량(75/100/450ml) 확인 필요"
        elif b == "Spray":
            if "딥앤쿨" in name_clean:
                ap, st_val, rsn = "딥앤쿨", "✅ 정상매칭", "Spray 딥앤쿨 라인"
            elif "익스트림" in name_clean:
                ap, st_val, rsn = "익스트림", "✅ 정상매칭", "Spray 익스트림 라인"
            elif "쿨" in name_clean or code_upper in {"ME0042118N", "ME00421185", "ME00421186", "ME00421175"}:
                ap, st_val, rsn = "쿨", "✅ 정상매칭", "Spray 쿨 라인"
            else:
                ap, st_val, rsn = "미분류(확인필요)", "⚠️ 확인필요", "Spray 세부 라인 확인 필요"
        elif b == "HADALABO":
            if "프리미엄" in name_clean or "premium" in name_clean:
                ap, st_val, rsn = "프리미엄", "✅ 정상매칭", "HADALABO 프리미엄 라인"
            elif "탄력" in name_clean:
                ap, st_val, rsn = "탄력", "✅ 정상매칭", "HADALABO 탄력 라인"
            elif "시로쥰" in name_clean or "shirojyun" in name_clean:
                ap, st_val, rsn = "시로쥰", "✅ 정상매칭", "HADALABO 시로쥰 라인"
            elif "키와미즈" in name_clean or "kiwamizu" in name_clean:
                ap, st_val, rsn = "키와미즈", "✅ 정상매칭", "HADALABO 키와미즈 라인"
            elif "타마고하다" in name_clean or "tamagohada" in name_clean:
                ap, st_val, rsn = "타마고하다", "✅ 정상매칭", "HADALABO 타마고하다 라인"
            elif "고쿠쥰" in name_clean or "gokujyun" in name_clean or "하다라보" in name_clean:
                ap, st_val, rsn = "고쿠쥰", "✅ 정상매칭", "HADALABO 고쿠쥰 기본 라인"
            else:
                ap, st_val, rsn = "미분류(확인필요)", "⚠️ 확인필요", "HADALABO 세부 라인 확인 필요"
        elif b == "Roll on Lotion":
            if "핫&쿨" in name_clean or "핫앤쿨" in name_clean:
                ap, st_val, rsn = "핫&쿨 번들", "✅ 정상매칭", "Roll on Lotion 핫&쿨 혼합 기획"
            elif "핫" in name_clean or code_upper in {"ME80421DRD", "ME80422DHE"}:
                ap, st_val, rsn = "핫", "✅ 정상매칭", "Roll on Lotion 핫 라인"
            elif "쿨" in name_clean or code_upper in {"ME80421DCH", "ME80422DHD", "ME80421DRC"}:
                ap, st_val, rsn = "쿨", "✅ 정상매칭", "Roll on Lotion 쿨 라인"
            else:
                ap, st_val, rsn = "미분류(확인필요)", "⚠️ 확인필요", "Roll on Lotion 핫/쿨 구분 확인 필요"
        elif b == "Hand":
            if "기획세트" in name_clean or "기획" in name_clean or "&" in name_clean:
                ap, st_val, rsn = "기획세트", "✅ 정상매칭", "Hand 리치네일&리페어코트 기획세트"
            elif "리페어코트" in name_clean:
                ap, st_val, rsn = "리치네일 리페어코트", "✅ 정상매칭", "Hand 리치네일 리페어코트 단품"
            elif "리치네일" in name_clean:
                ap, st_val, rsn = "리치네일", "✅ 정상매칭", "Hand 리치네일 단품"
            else:
                ap, st_val, rsn = "미분류(확인필요)", "⚠️ 확인필요", "Hand 세부 라인 확인 필요"
        elif b == "Melano":
            if "에센스" in name_clean or "안티스팟" in name_clean or "화이트닝" in name_clean:
                ap, st_val, rsn = "안티스팟 화이트닝에센스", "✅ 정상매칭", "Melano 안티스팟 화이트닝에센스"
            elif "워시" in name_clean or "엔자임" in name_clean or code_upper in {"ME90621MAF", "ME90621MAM", "ME90621MAK"}:
                ap, st_val, rsn = "엔자임페이스 워시", "✅ 정상매칭", "Melano 엔자임페이스 워시"
            else:
                ap, st_val, rsn = "미분류(확인필요)", "⚠️ 확인필요", "Melano 세부 라인 확인 필요"
        elif b == "Acnes":
            if "티트리" in name_clean:
                ap, st_val, rsn = "다이소 출고 제품", "✅ 정상매칭", "Acnes 제품명에 '티트리' 포함 → 다이소 출고 제품"
            elif "마스크" in name_clean or "패드" in name_clean:
                ap, st_val, rsn = "스킨케어", "✅ 정상매칭", "Acnes 제품명에 '마스크' 또는 '패드' 포함 → 스킨케어"
            elif "포맨" in name_clean or "formen" in name_clean:
                ap, st_val, rsn = "포 맨", "✅ 정상매칭", "Acnes 포맨 라인 일치"
            elif "더마릴리프" in name_clean or "더마" in name_clean:
                if any(k in name_clean for k in ["클렌저", "워시", "폼"]):
                    ap, st_val, rsn = "클렌징", "✅ 정상매칭", "Acnes 더마 클렌징 라인 일치"
                elif any(k in name_clean for k in ["리커버리", "밸런싱", "크림", "스킨"]):
                    ap, st_val, rsn = "더마 스킨케어", "✅ 정상매칭", "Acnes 더마 스킨케어 라인 일치"
                else:
                    ap, st_val, rsn = "미분류(확인필요)", "⚠️ 확인필요", "더마 세부 카테고리 모호"
            elif any(k in name_clean for k in ["바디워시", "바디미스트", "바디로션", "바디"]):
                ap, st_val, rsn = "바디케어", "✅ 정상매칭", "Acnes 바디케어 라인"
            elif any(k in name_clean for k in ["스팟패치", "진정패치", "패치"]):
                ap, st_val, rsn = "스팟패치", "✅ 정상매칭", "Acnes 스팟패치 라인 일치"
            elif any(k in name_clean for k in ["세범밸런싱", "모이스처플루이드", "모이스처크림", "선크림", "토너", "스팟크림", "스킨"]):
                ap, st_val, rsn = "스킨케어", "✅ 정상매칭", "Acnes 스킨케어 라인"
            elif any(k in name_clean for k in ["포밍워시", "클렌저", "크리미워시", "클렌징오일", "클렌징", "포하", "4ha"]):
                ap, st_val, rsn = "클렌징", "✅ 정상매칭", "Acnes 클렌징 라인"
            else:
                ap, st_val, rsn = "미분류(확인필요)", "⚠️ 확인필요", "Acnes 포트폴리오 키워드 미일치"
        elif b == "Lipcare":
            if "딥모이스트" in name_clean or "deepmoist" in name_clean or code_upper == "ME90521DME":
                ap, st_val, rsn = "딥 모이스트", "✅ 정상매칭", "Lipcare 딥 모이스트 라인 일치"
            elif "메디케이티드" in name_clean or "메디케이트" in name_clean or "medicated" in name_clean:
                ap, st_val, rsn = "메디케이티드", "✅ 정상매칭", "Lipcare 메디케이티드 라인 일치"
            elif "소프트립" in name_clean or "softlips" in name_clean:
                ap, st_val, rsn = "소프트립스", "✅ 정상매칭", "Lipcare 소프트립스 라인 일치"
            elif (
                any(k in name_clean for k in ["매직컬러", "매직칼라", "틴티드컬러", "틴티드칼라", "워터컬러"])
                or code_upper in {"ME90521WC4", "ME90521WC5"}
            ):
                ap, st_val, rsn = "매직컬러 / 틴티드컬러 라인", "✅ 정상매칭", "Lipcare 컬러/틴티드 라인 일치"
            elif "립아이스" in name_clean or "lipice" in name_clean:
                ap, st_val, rsn = "립아이스 프루티", "✅ 정상매칭", "립아이스 프루티 라인"
            else:
                ap, st_val, rsn = "미분류(확인필요)", "⚠️ 확인필요", "Lipcare 세부 라인 확인 필요"
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
    cols = [
        "제품코드",
        "브랜드",
        "제품명(제품군)",
        "포트폴리오",
        "입수량(BOX)",
        "수입리드타임(일)",
        "안전재고일수(일)",
        "비고(채널/규격/특이사항)",
    ]
    sample = [["ME00999NEW", "Acnes", "아크네스 화이트 티트리 시카 세럼 30mL (다이소)", "다이소 출고 제품", 24, 90, 30, "수입 리드타임 반영"]]
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
    st.session_state["selected_brand"] = "Lotion"
if "active_tab" not in st.session_state:
    st.session_state["active_tab"] = "📊 브랜드·포트폴리오 현황"
if "selected_forecast_sku" not in st.session_state:
    st.session_state["selected_forecast_sku"] = ""

top_l, top_r = st.columns([1.35, 1.15])
with top_l:
    st.title("📦 브랜드 및 포트폴리오별 통합 재고 현황")
    st.caption("브랜드/포트폴리오 세분화 · 유효일자(제조일+3년, 잔여 50% 미만) 관리 · 수입리드타임 기반 DOH·ROP·회전율·미납률 관리")

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

up_c1, up_c2, up_c3 = st.columns(3)
with up_c1:
    with st.container(border=True):
        uploaded_inv = st.file_uploader(
            "📂 1. 물류창고 재고파일 업로드",
            type=["xlsx", "csv"],
            key="uploader_inv_main",
        )
with up_c2:
    with st.container(border=True):
        uploaded_out = st.file_uploader(
            "🚚 2. '일일출고 현황' 파일 업로드",
            type=["xlsx", "xls", "csv"],
            key="uploader_outbound_main",
        )
with up_c3:
    with st.container(border=True):
        uploaded_nd = st.file_uploader(
            "📉 3. '미납내역' 파일 업로드 (목표 미납률 산출)",
            type=["xlsx", "xls"],
            key="uploader_nondelivery_main",
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
                    "제조일자": "-",
                    "유효일자": "-",
                    "제조일자_raw": None,
                    "유효일자_raw": None,
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
        outbound_df, monthly_hist_df, outbound_info = parse_outbound_bytes(uploaded_out.getvalue(), uploaded_out.name)
    except Exception as e:
        st.warning(f"⚠️ '일일출고 현황' 파일 파싱 중 참고사항: {e}")

nd_df = None
nd_info = None
if uploaded_nd is not None:
    try:
        nd_df, nd_info = parse_nondelivery_bytes(uploaded_nd.getvalue())
    except Exception as e:
        st.warning(f"⚠️ '미납내역' 파일 파싱 중 참고사항: {e}")

st.success(
    f"✅ {meta['source_label']} 로드 완료 | 전체 {meta['raw_total_rows']:,}건 중 제외({', '.join(EXCLUDED_KEYWORDS)}) {meta['excluded_rows']:,}건 "
    f"({int(meta['excluded_qty']):,} EA) 반영 → **분석 대상 {meta['final_rows']:,}건 (고유 SKU {meta['unique_skus']:,}종)**"
    + (
        f" | 🚚 일일출고 {outbound_info['rows']:,}건 ({outbound_info['date_range']})"
        if outbound_info
        else ""
    )
    + (
        f" | 📉 미납내역 {nd_info['rows']:,}건 (누적 {nd_info['total_amt']:,.0f}원, {nd_info['period_str']})"
        if nd_info
        else ""
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

active_6m_tuple = outbound_info["active_6m_codes"] if outbound_info else ()
has_outbound_flag = bool(outbound_info["has_outbound"]) if outbound_info else False

sku_classified = classify_skus_cached(
    unique_sku_df,
    tuple(sorted(master_map.items())),
    tuple(sorted(st.session_state["brand_overrides"].items())),
    tuple(sorted(st.session_state["portfolio_overrides"].items())),
    active_6m_tuple,
    has_outbound_flag,
)

inv_merged = inv_df.merge(
    sku_classified[["상품", "브랜드", "최종_포트폴리오", "검수상태", "판별근거"]],
    on="상품",
    how="left",
)
inv_merged["브랜드"] = inv_merged["브랜드"].fillna("미분류")
inv_merged["최종_포트폴리오"] = inv_merged["최종_포트폴리오"].fillna("-")
inv_merged["검수상태"] = inv_merged["검수상태"].fillna("해당없음")

# 유통기한(제조일로부터 3년) 기준 잔여일수 비율 50% 미만 관리 품목 산출
today_ref = date.today()
shelf_eval_rows = [
    evaluate_shelf_life(r.제조일자_raw, r.유효일자_raw, ref_date=today_ref, shelf_life_years=3)
    for r in inv_merged.itertuples(index=False)
]
shelf_eval_df = pd.DataFrame(shelf_eval_rows, index=inv_merged.index)
inv_merged = pd.concat([inv_merged, shelf_eval_df], axis=1)

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

exp_manage_qty = int(inv_merged[inv_merged["유효일자_관리대상"]][qty_metric].sum())

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("선택 기준 총 재고 (EA)", f"{int(inv_merged[qty_metric].sum()):,}")
k2.metric("Acnes", f"{int(inv_merged[inv_merged['브랜드'] == 'Acnes'][qty_metric].sum()):,}")
k3.metric("Lipcare", f"{int(inv_merged[inv_merged['브랜드'] == 'Lipcare'][qty_metric].sum()):,}")
k4.metric("HADALABO", f"{int(inv_merged[inv_merged['브랜드'] == 'HADALABO'][qty_metric].sum()):,}")
k5.metric("⏳ 유효일자 관리(50% 미만)", f"{exp_manage_qty:,}")

sku_qty = inv_merged.groupby("상품", as_index=False)[["정상환산", "불량환산", "합계수량"]].sum()
sku_view = sku_classified.merge(sku_qty, on="상품", how="left").fillna(0)
sku_view["출고매핑코드"] = sku_view["상품"].apply(map_lotion_code)

if outbound_df is not None and not outbound_df.empty:
    sku_view = sku_view.merge(outbound_df, left_on="출고매핑코드", right_on="상품", how="left", suffixes=("", "_out"))
    if "상품_out" in sku_view.columns:
        sku_view = sku_view.drop(columns=["상품_out"])
    sku_view["누적_출고수량"] = sku_view["누적_출고수량"].fillna(0.0)
    sku_view["출고_발생일수"] = sku_view["출고_발생일수"].fillna(0).astype(int)
    sku_view["일평균_출고량"] = sku_view["일평균_출고량"].fillna(0.0)
    sku_view["월예상_출고량"] = sku_view["월예상_출고량"].fillna(0.0)
else:
    sku_view["누적_출고수량"] = 0.0
    sku_view["출고_발생일수"] = 0
    sku_view["일평균_출고량"] = (sku_view["정상환산"] / 30.0).round(2)
    sku_view["월예상_출고량"] = sku_view["정상환산"].round(0)

# 3개 전용 칸(브랜드·포트폴리오 현황 / SCM 팔로우업 / 유효일자 관리)
v_col1, v_col2, v_col3 = st.columns([1, 1, 1])
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
        "🛡️ SCM 팔로우업: DOH & 안전재고 알림",
        type="primary" if st.session_state["active_tab"] == "🛡️ SCM 팔로우업" else "secondary",
        use_container_width=True,
    ):
        st.session_state["active_tab"] = "🛡️ SCM 팔로우업"
        st.rerun()
with v_col3:
    if st.button(
        "⏳ 유효일자 관리 (잔여 50% 미만)",
        type="primary" if st.session_state["active_tab"] == "⏳ 유효일자 관리" else "secondary",
        use_container_width=True,
    ):
        st.session_state["active_tab"] = "⏳ 유효일자 관리"
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
            key="sku_brand_editor_v13",
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
            p_chart_df = p_sum.rename(columns={"최종_포트폴리오": "포트폴리오", qty_metric: "재고수량(EA)"})

            if sel_brand == "Melano":
                maf_qty = int(brand_rows[brand_rows["상품"] == "ME90621MAF"][qty_metric].sum())
                mam_qty = int(brand_rows[brand_rows["상품"] == "ME90621MAM"][qty_metric].sum())
                mak_qty = int(brand_rows[brand_rows["상품"] == "ME90621MAK"][qty_metric].sum())

                p_chart_df["maf (디에이치, 메이크)"] = p_chart_df["포트폴리오"].apply(
                    lambda x: f"{maf_qty:,} EA" if x == "엔자임페이스 워시" else "-"
                )
                p_chart_df["mam (마트)"] = p_chart_df["포트폴리오"].apply(
                    lambda x: f"{mam_qty:,} EA" if x == "엔자임페이스 워시" else "-"
                )
                p_chart_df["mak (수입상태)"] = p_chart_df["포트폴리오"].apply(
                    lambda x: f"{mak_qty:,} EA" if x == "엔자임페이스 워시" else "-"
                )

                melano_chart = (
                    alt.Chart(p_chart_df)
                    .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
                    .encode(
                        x=alt.X("포트폴리오:N", sort="-y", title="포트폴리오"),
                        y=alt.Y("재고수량(EA):Q", title="재고수량(EA)"),
                        color=alt.Color("포트폴리오:N", legend=None),
                        tooltip=[
                            alt.Tooltip("포트폴리오:N", title="포트폴리오"),
                            alt.Tooltip("재고수량(EA):Q", title="합계 재고(EA)", format=","),
                            alt.Tooltip("maf (디에이치, 메이크):N", title="maf (디에이치, 메이크)"),
                            alt.Tooltip("mam (마트):N", title="mam (마트)"),
                            alt.Tooltip("mak (수입상태):N", title="mak (수입상태)"),
                        ],
                    )
                    .properties(height=320)
                )
                st.altair_chart(melano_chart, use_container_width=True)
            else:
                st.bar_chart(
                    data=p_chart_df,
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
    detail_cols = [
        "센터",
        "셀",
        "상품",
        "상품명",
        "최종_포트폴리오",
        "검수상태",
        "화주LOT",
        "제조일자",
        "유효일자",
        "잔여일수",
        "잔여비율(%)",
        "유효일자_관리구분",
        "정상환산",
        "불량환산",
        "합계수량",
    ]
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
        forecast_months = st.slider("향후 예측 기간(최대 6개월)", min_value=1, max_value=6, value=6, step=1)

    with fc_c3:
        growth_rate_pct = st.number_input("월 출고 성장률(%)", min_value=-50.0, max_value=100.0, value=0.0, step=5.0)

    if picked_sku:
        sku_row = sku_view[sku_view["상품"] == picked_sku].iloc[0]
        cur_stock = float(sku_row["정상환산"])
        mapped_code = map_lotion_code(picked_sku)

        sku_m_hist = (
            monthly_hist_df[monthly_hist_df["상품"] == mapped_code].sort_values("연월").copy()
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
        fc_m1.metric(
            "선택 제품코드",
            f"{picked_sku}" + (f" (출고합산:{mapped_code})" if mapped_code != picked_sku else ""),
        )
        fc_m2.metric("현재 정상재고(EA)", f"{int(cur_stock):,}")
        fc_m3.metric("기준 월 출고량(EA)", f"{int(round(base_monthly_out)):,}")

        st.line_chart(fc_chart_df, height=340, use_container_width=True)

        fc_table_df = fc_chart_df.reset_index()

        def _highlight_zero_stock_row(row):
            rem = row.get("예상 잔여재고(EA)")
            if pd.notna(rem) and float(rem) <= 0:
                return ["background-color: #fee2e2; color: #b91c1c; font-weight: 700;"] * len(row)
            return [""] * len(row)

        styled_fc = fc_table_df.style.apply(_highlight_zero_stock_row, axis=1).format(
            {
                "월별 출고량(EA)": lambda v: f"{int(v):,}" if pd.notna(v) else "-",
                "예상 잔여재고(EA)": lambda v: f"{int(v):,}" if pd.notna(v) else "-",
            }
        )
        st.dataframe(styled_fc, use_container_width=True, hide_index=True, height=260)

elif st.session_state["active_tab"] == "🛡️ SCM 팔로우업":
    st.markdown('<div class="section-title">🛡️ 출고/수주 대비 커버리지(DOH) 및 안전재고(Safety Stock) 알림</div>', unsafe_allow_html=True)

    if outbound_df is None:
        st.info(
            "💡 현재 '일일출고 현황' 파일이 미업로드 상태입니다. 상단 **2. '일일출고 현황' 파일 업로드**에 출고 데이터를 올리면 "
            "실출고 기준으로 DOH, ROP 충족률, 연간 재고회전율(ITO), 예상 결품일이 산출됩니다."
        )

    p_c1, p_c2, p_c3 = st.columns(3)
    with p_c1:
        import_lt_days = st.slider(
            "수입입고 리드타임(Lead Time, 일)",
            min_value=30,
            max_value=180,
            value=90,
            step=5,
        )
    with p_c2:
        safety_days = st.slider(
            "안전재고 기준 일수(Safety Stock Days, 일)",
            min_value=7,
            max_value=90,
            value=30,
            step=5,
        )
    with p_c3:
        danger_days = st.slider(
            "위험 커버리지 기준(일 미만)",
            min_value=15,
            max_value=120,
            value=60,
            step=5,
        )

    scm_df = sku_view.copy()
    scm_df["안전재고_수량(EA)"] = (scm_df["일평균_출고량"] * float(safety_days)).round(0)
    scm_df["발주점(ROP, EA)"] = (scm_df["일평균_출고량"] * float(import_lt_days + safety_days)).round(0)
    scm_df["커버리지(DOH, 일)"] = scm_df.apply(
        lambda r: round(r["정상환산"] / r["일평균_출고량"], 1)
        if r["일평균_출고량"] > 0
        else (999.0 if r["정상환산"] > 0 else 0.0),
        axis=1,
    )
    scm_df["안전재고_과부족(EA)"] = (scm_df["정상환산"] - scm_df["안전재고_수량(EA)"]).round(0)
    scm_df["ROP_충족률(%)"] = scm_df.apply(
        lambda r: round((r["정상환산"] / r["발주점(ROP, EA)"]) * 100.0, 1)
        if r["발주점(ROP, EA)"] > 0
        else (100.0 if r["정상환산"] > 0 else 0.0),
        axis=1,
    )
    scm_df["연간_재고회전율(ITO, 회)"] = scm_df.apply(
        lambda r: round((r["일평균_출고량"] * 365.0) / r["정상환산"], 2)
        if r["정상환산"] > 0 and r["일평균_출고량"] > 0
        else 0.0,
        axis=1,
    )
    scm_df["안전재고_여유일수(일)"] = (scm_df["커버리지(DOH, 일)"] - float(safety_days)).round(1)

    base_today = date.today()

    def _stockout_date_str(r):
        if r["일평균_출고량"] <= 0:
            return "-"
        if r["정상환산"] <= 0:
            return "즉시결품"
        days_left = int(r["정상환산"] // r["일평균_출고량"])
        if days_left > 730:
            return "2년 이상"
        return (base_today + timedelta(days=days_left)).strftime("%Y-%m-%d")

    scm_df["예상_결품일"] = scm_df.apply(_stockout_date_str, axis=1)

    def _doh_status(r):
        if is_migrated_lotion_code(r["상품"]):
            return "⚪ 로션코드변경_제외"
        if r["일평균_출고량"] <= 0:
            return "🟢 출고없음(재고확인)"
        doh = r["커버리지(DOH, 일)"]
        if doh < danger_days or r["정상환산"] < r["안전재고_수량(EA)"]:
            return "🚨 위험(수입리드타임 대비 부족)"
        if doh < (import_lt_days + safety_days):
            return "⚠️ 주의(발주검토 필요)"
        return "✅ 안정(리드타임 확보)"

    scm_df["SCM_알림상태"] = scm_df.apply(_doh_status, axis=1)

    danger_cnt = int((scm_df["SCM_알림상태"] == "🚨 위험(수입리드타임 대비 부족)").sum())
    caution_cnt = int((scm_df["SCM_알림상태"] == "⚠️ 주의(발주검토 필요)").sum())
    ss_short_cnt = int(
        ((~scm_df["상품"].apply(is_migrated_lotion_code)) & (scm_df["안전재고_과부족(EA)"] < 0)).sum()
    )
    avg_ito = (
        round(float(scm_df[scm_df["연간_재고회전율(ITO, 회)"] > 0]["연간_재고회전율(ITO, 회)"].mean()), 2)
        if (scm_df["연간_재고회전율(ITO, 회)"] > 0).any()
        else 0.0
    )
    avg_rop_fill = (
        round(float(scm_df[scm_df["일평균_출고량"] > 0]["ROP_충족률(%)"].mean()), 1)
        if (scm_df["일평균_출고량"] > 0).any()
        else 0.0
    )

    m1, m2, m3, m4, m5 = st.columns(5)
    m1.metric("🚨 위험 품목(로션변경 제외)", f"{danger_cnt:,}종")
    m2.metric("⚠️ 주의 품목(발주검토 필요)", f"{caution_cnt:,}종")
    m3.metric("🛡️ 안전재고 미달 품목", f"{ss_short_cnt:,}종")
    m4.metric("🎯 평균 ROP 충족률", f"{avg_rop_fill:.1f}%")
    m5.metric("🔄 평균 재고회전율(ITO)", f"{avg_ito:.2f}회/년")

    st.markdown('<div class="section-title">📉 수입입고 리드타임 고려 커버리지(DOH) 위험 수준 그래프</div>', unsafe_allow_html=True)

    f_col1, f_col2, f_col3 = st.columns([1.4, 1.4, 1.2])
    with f_col1:
        filter_status = st.multiselect(
            "알림 상태 필터",
            options=[
                "🚨 위험(수입리드타임 대비 부족)",
                "⚠️ 주의(발주검토 필요)",
                "✅ 안정(리드타임 확보)",
                "🟢 출고없음(재고확인)",
                "⚪ 로션코드변경_제외",
            ],
            default=["🚨 위험(수입리드타임 대비 부족)", "⚠️ 주의(발주검토 필요)"],
            key="scm_alert_filter_top",
        )
    with f_col2:
        code_candidates = (
            scm_df[scm_df["SCM_알림상태"].isin(filter_status)]["상품"].tolist()
            if filter_status
            else scm_df["상품"].tolist()
        )
        if not code_candidates:
            code_candidates = scm_df["상품"].tolist()
        picked_me_graph = st.selectbox(
            "ME코드(상품코드) 선택 (선택 코드 1개 그래프 조회)",
            options=["전체(상위 25개)"] + code_candidates,
            index=0,
            key="scm_single_me_selector",
        )
    with f_col3:
        chart_metric_mode = st.selectbox(
            "보조 지표 보기",
            options=["DOH vs 리드타임/위험선", "ROP 충족률(%) & 안전재고 여유일수"],
            index=0,
            key="scm_metric_mode_selector",
        )

    filtered_scm = scm_df[scm_df["SCM_알림상태"].isin(filter_status)] if filter_status else scm_df
    filtered_scm = filtered_scm.sort_values(by=["커버리지(DOH, 일)", "안전재고_과부족(EA)"], ascending=[True, True])

    if picked_me_graph == "전체(상위 25개)":
        g_df = filtered_scm[filtered_scm["일평균_출고량"] > 0].head(25).copy()
        if g_df.empty:
            g_df = filtered_scm.head(25).copy()
        g_df["DOH_표시(일)"] = g_df["커버리지(DOH, 일)"].clip(upper=365.0)
        g_df["위험기준선(일)"] = float(danger_days)
        g_df["수입리드타임선(일)"] = float(import_lt_days)

        if chart_metric_mode == "DOH vs 리드타임/위험선":
            melted = g_df.melt(
                id_vars=["상품", "상품명", "SCM_알림상태"],
                value_vars=["DOH_표시(일)", "위험기준선(일)", "수입리드타임선(일)"],
                var_name="구분",
                value_name="일수",
            )
            doh_chart = (
                alt.Chart(melted)
                .mark_line(point=True)
                .encode(
                    x=alt.X("상품:N", sort=g_df["상품"].tolist(), title="ME코드(상품코드)"),
                    y=alt.Y("일수:Q", title="일수(Days)"),
                    color=alt.Color("구분:N", title="지표 구분"),
                    tooltip=[
                        alt.Tooltip("상품:N", title="ME코드"),
                        alt.Tooltip("상품명:N", title="제품명"),
                        alt.Tooltip("SCM_알림상태:N", title="알림상태"),
                        alt.Tooltip("구분:N", title="구분"),
                        alt.Tooltip("일수:Q", title="일수", format=".1f"),
                    ],
                )
                .properties(height=340)
            )
            st.altair_chart(doh_chart, use_container_width=True)
        else:
            rop_chart = (
                alt.Chart(g_df)
                .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
                .encode(
                    x=alt.X("상품:N", sort=g_df["상품"].tolist(), title="ME코드(상품코드)"),
                    y=alt.Y("ROP_충족률(%):Q", title="ROP 충족률(%)"),
                    color=alt.Color("SCM_알림상태:N", title="알림상태"),
                    tooltip=[
                        alt.Tooltip("상품:N", title="ME코드"),
                        alt.Tooltip("상품명:N", title="제품명"),
                        alt.Tooltip("ROP_충족률(%):Q", title="ROP 충족률(%)", format=".1f"),
                        alt.Tooltip("안전재고_여유일수(일):Q", title="안전재고 여유일수(일)", format=".1f"),
                        alt.Tooltip("연간_재고회전율(ITO, 회):Q", title="연간 재고회전율(ITO)", format=".2f"),
                    ],
                )
                .properties(height=340)
            )
            st.altair_chart(rop_chart, use_container_width=True)
    else:
        one_row = scm_df[scm_df["상품"] == picked_me_graph].iloc[0]
        s_c1, s_c2, s_c3, s_c4 = st.columns(4)
        s_c1.metric("커버리지(DOH)", f"{one_row['커버리지(DOH, 일)']} 일")
        s_c2.metric("ROP 충족률", f"{one_row['ROP_충족률(%)']} %")
        s_c3.metric("연간 재고회전율(ITO)", f"{one_row['연간_재고회전율(ITO, 회)']} 회")
        s_c4.metric("예상 결품일", f"{one_row['예상_결품일']}")

        one_df = pd.DataFrame(
            {
                "항목": ["현재 DOH(일)", "안전재고 기준(일)", "위험기준선(일)", "수입리드타임(일)"],
                "일수": [
                    float(one_row["커버리지(DOH, 일)"]),
                    float(safety_days),
                    float(danger_days),
                    float(import_lt_days),
                ],
            }
        )
        single_chart = (
            alt.Chart(one_df)
            .mark_bar(cornerRadiusTopLeft=5, cornerRadiusTopRight=5)
            .encode(
                x=alt.X("항목:N", sort=None, title="기준 항목"),
                y=alt.Y("일수:Q", title="일수(Days)"),
                color=alt.Color("항목:N", legend=None),
                tooltip=[
                    alt.Tooltip("항목:N", title="항목"),
                    alt.Tooltip("일수:Q", title="일수(Days)", format=".1f"),
                ],
            )
            .properties(height=300)
        )
        st.altair_chart(single_chart, use_container_width=True)

    st.markdown('<div class="section-title">📋 SCM 재고 건전성 및 발주·결품 관리 지표</div>', unsafe_allow_html=True)
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
        "안전재고_여유일수(일)",
        "발주점(ROP, EA)",
        "ROP_충족률(%)",
        "연간_재고회전율(ITO, 회)",
        "예상_결품일",
    ]
    st.dataframe(filtered_scm[scm_cols], use_container_width=True, hide_index=True, height=340)

    st.markdown('<div class="section-title">📉 미납내역 분석 및 전체 매출액 대비 목표 미납률 설정</div>', unsafe_allow_html=True)

    nd_c1, nd_c2 = st.columns([1.4, 1.0])
    with nd_c1:
        annual_revenue_krw = st.number_input(
            "연간 전체 매출액(원) 입력 (미납률 산출 기준)",
            min_value=100_000_000.0,
            max_value=10_000_000_000_000.0,
            value=10_000_000_000.0,
            step=500_000_000.0,
            format="%.0f",
            key="scm_annual_revenue_input",
        )
    with nd_c2:
        reduction_pct = st.slider(
            "미납 감축 목표 비율(%)",
            min_value=10.0,
            max_value=70.0,
            value=37.5,
            step=2.5,
            key="scm_target_reduction_pct",
        )

    if nd_df is not None and nd_info is not None:
        ann_nd_amt = float(nd_info["annualized_amt"])
        actual_rate_pct = (ann_nd_amt / float(annual_revenue_krw)) * 100.0
        target_rate_pct = actual_rate_pct * (1.0 - float(reduction_pct) / 100.0)
        target_ann_amt = ann_nd_amt * (1.0 - float(reduction_pct) / 100.0)

        r1, r2, r3, r4 = st.columns(4)
        r1.metric("누적 미납금액", f"{nd_info['total_amt']:,.0f}원")
        r2.metric("연환산 미납금액", f"{ann_nd_amt:,.0f}원/년")
        r3.metric("매출 대비 실적 미납률", f"{actual_rate_pct:.2f}%")
        r4.metric("매출 대비 목표 미납률", f"{target_rate_pct:.2f}% (목표액 {target_ann_amt:,.0f}원)")

        cause_tbl = (
            nd_df.groupby("원인대분류", as_index=False)
            .agg(
                건수=("미납금액_num", "count"),
                미납수량=("미납수량_num", "sum"),
                미납금액=("미납금액_num", "sum"),
            )
            .sort_values(by="미납금액", ascending=False)
        )
        cause_tbl["금액비중(%)"] = (cause_tbl["미납금액"] / max(nd_info["total_amt"], 1.0) * 100.0).round(2)
        st.dataframe(cause_tbl, use_container_width=True, hide_index=True, height=240)
    else:
        st.info("💡 '미납내역' 엑셀 파일을 업로드하면 원인별 미납금액과 전체 매출액 대비 실적·목표 미납률이 자동 산출됩니다.")

    st.download_button(
        label="📥 SCM 커버리지(DOH)·안전재고·회전율 지표 엑셀 다운로드",
        data=to_excel_bytes(filtered_scm[scm_cols], sheet_name="SCM_지표_알림"),
        file_name="SCM_커버리지_DOH_안전재고_지표.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    # 전용 칸: 유효일자 관리 (유통기한 제조일로부터 3년 기준, 잔여일수 비율 50% 미만)
    st.markdown('<div class="section-title">⏳ 유효일자 관리 전용 칸 (유통기한 3년 기준 · 잔여일수 50% 미만)</div>', unsafe_allow_html=True)
    st.caption(
        f"기준일: **{today_ref.strftime('%Y-%m-%d')}** | 유통기한: **제조일로부터 3년** | "
        "잔여일수 비율이 **50% 미만**인 LOT/재고를 별도 관리 대상으로 분류합니다."
    )

    exp_target_df = inv_merged[inv_merged["유효일자_관리대상"]].copy()
    expired_df = exp_target_df[exp_target_df["잔여일수"] < 0]
    under_50_df = exp_target_df[exp_target_df["잔여일수"] >= 0]

    e1, e2, e3, e4 = st.columns(4)
    e1.metric("유효일자 관리 대상 LOT", f"{len(exp_target_df):,}건")
    e2.metric("관리 대상 재고수량", f"{int(exp_target_df[qty_metric].sum()):,} EA")
    e3.metric("잔여 50% 미만(유효)", f"{len(under_50_df):,}건 ({int(under_50_df[qty_metric].sum()):,} EA)")
    e4.metric("만료 경과 LOT", f"{len(expired_df):,}건 ({int(expired_df[qty_metric].sum()):,} EA)")

    if not exp_target_df.empty:
        exp_brand_sum = (
            exp_target_df.groupby(["브랜드", "유효일자_관리구분"], as_index=False)[qty_metric]
            .sum()
            .sort_values(by=qty_metric, ascending=False)
        )
        exp_chart = (
            alt.Chart(exp_brand_sum)
            .mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
            .encode(
                x=alt.X("브랜드:N", sort="-y", title="브랜드"),
                y=alt.Y(f"{qty_metric}:Q", title="관리 대상 재고수량(EA)"),
                color=alt.Color("유효일자_관리구분:N", title="관리 구분"),
                tooltip=[
                    alt.Tooltip("브랜드:N", title="브랜드"),
                    alt.Tooltip("유효일자_관리구분:N", title="관리 구분"),
                    alt.Tooltip(f"{qty_metric}:Q", title="재고수량(EA)", format=","),
                ],
            )
            .properties(height=300)
        )
        st.altair_chart(exp_chart, use_container_width=True)

        exp_cols = [
            "유효일자_관리구분",
            "브랜드",
            "최종_포트폴리오",
            "상품",
            "상품명",
            "센터",
            "셀",
            "화주LOT",
            "제조일자",
            "유효일자",
            "잔여일수",
            "총수명일수",
            "잔여비율(%)",
            "정상환산",
            "불량환산",
            "합계수량",
        ]
        exp_sorted = exp_target_df[exp_cols].sort_values(by=["잔여일수", "잔여비율(%)"], ascending=[True, True])
        st.dataframe(exp_sorted, use_container_width=True, hide_index=True, height=360)

        st.download_button(
            label="📥 유효일자 관리 품목(잔여 50% 미만) 엑셀 다운로드",
            data=to_excel_bytes(exp_sorted, sheet_name="유효일자_관리_50미만"),
            file_name="유효일자_관리_품목_50미만.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.success("✅ 현재 기준일 기준으로 유통기한 잔여일수 50% 미만인 유효일자 관리 대상 LOT가 없습니다.")
