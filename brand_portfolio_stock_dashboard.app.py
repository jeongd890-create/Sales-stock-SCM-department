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
        return dt_obj.replace(year=target_year, month=2, day=28)


def evaluate_shelf_life(mfg_val, exp_val, ref_date: date | None = None, shelf_life_years: int = 3) -> dict:
    """
    유통기한은 제조일로부터 3년(기본) 기준.
    유효일자가 지난 경우 '유효일자 만료', 잔여일수 비율 50% 미만인 경우 '유효일자 관리 필요'로 분류.
    """
    today_d = ref_date or date.today()
    exp_ts = pd.to_datetime(exp_val, errors="coerce")
    if pd.isna(exp_ts):
        return {
            "잔여일수": None,
            "총수명일수": None,
            "잔여비율(%)": None,
            "유효일자_관리구분": "일자미기재",
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
        status_label = "유효일자 만료"
        is_manage = True
    elif rem_ratio_pct < 50.0:
        status_label = "유효일자 관리 필요"
        is_manage = True
    else:
        status_label = "정상(50% 이상)"
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
        active_6m_codes =
