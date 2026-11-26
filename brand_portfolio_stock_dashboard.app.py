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
