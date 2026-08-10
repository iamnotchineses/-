import html
import re
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

st.set_page_config(page_title="메카 매출 대시보드", page_icon="📊", layout="wide")

APP_DIR = Path(__file__).parent


# -----------------------------
# 데이터 소스 = parquet (엑셀 업로드 없음)
#   같은 폴더의 *.parquet 을 '컬럼 구성'으로 자동 판별한다(파일명 무관).
#     매출   : 주문번호 · 쇼핑몰 · 출고날짜 를 모두 가진 파일 (여러 개면 전부 합산)
#     재고   : 라인명 · 입고이력 · 가용수량 을 모두 가진 파일
#     이미지 : (라인명|모델명|상품명) + 이미지/URL 컬럼을 가진 파일
#              → 이미지 parquet 이 있으면 '이미지*.xlsx' 는 읽지 않는다(로딩 속도).
# -----------------------------
SALES_KEYS = {"주문번호", "쇼핑몰", "출고날짜"}
STOCK_KEYS = {"라인명", "입고이력", "가용수량"}
IMG_NAME_KEYS = ("라인명", "모델명", "상품명")   # 이미지 매핑 parquet 의 이름 컬럼 후보

# 재고 스냅샷 기준일. 재고 parquet 의 '기준일' 값으로 로드 시 갱신된다.
#   ('N일전' 입고이력을 실제 날짜로 되돌릴 때의 기준 — 오늘 날짜로 계산하면 파일이 오래될수록 어긋난다)
STOCK_BASE_DATE = pd.Timestamp.now().normalize()


def _parquet_cols(path) -> set:
    """parquet 의 컬럼명만 빠르게 읽는다(데이터는 안 읽음). 실패 시 빈 set."""
    try:
        import pyarrow.parquet as pq
        return set(pq.read_schema(str(path)).names)
    except Exception:
        try:
            return set(pd.read_parquet(path).columns)
        except Exception:
            return set()


def _is_image_cols(cols: set) -> bool:
    """이미지 매핑 parquet 판별: 이름 컬럼 + 이미지/URL 컬럼."""
    has_name = any(c in cols for c in IMG_NAME_KEYS)
    has_url = any(("이미지" in str(c)) or ("url" in str(c).lower()) for c in cols)
    return has_name and has_url


def scan_parquet_files(folder) -> tuple[list, list, list]:
    """폴더의 *.parquet → (매출, 재고, 이미지 매핑) 파일 목록."""
    sales, stock, image = [], [], []
    for p in sorted(Path(folder).glob("*.parquet"), key=lambda x: x.name):
        cols = _parquet_cols(p)
        if SALES_KEYS <= cols:
            sales.append(p)
        elif STOCK_KEYS <= cols:
            stock.append(p)
        elif _is_image_cols(cols):
            image.append(p)
    return sales, stock, image


img_map: dict = {}   # 라인명(또는 모델명) → 이미지 URL


def _img_keys(name) -> set:
    """이미지 매칭 키 후보: 원본 + 끝의 사이즈 '(...)' 제거형."""
    s = "" if name is None else str(name).strip()
    return {k for k in (s, re.sub(r"\s*\([^()]*\)\s*$", "", s).strip()) if k}


def find_image_file():
    """앱 폴더의 이미지 매핑 엑셀('이미지*.xlsx' 등)을 찾는다(최신 우선)."""
    for pat in ("이미지*.xlsx", "이미지*.xls", "image*.xlsx", "images*.xlsx"):
        c = sorted(APP_DIR.glob(pat), key=lambda p: (p.stat().st_mtime, p.name), reverse=True)
        if c:
            return c[0]
    return None


def _rows_to_img_map(pairs) -> dict:
    """(이름, URL) 나열 → {키: URL}. URL 이 아닌 행/헤더행은 건너뛴다."""
    m = {}
    for name, url in pairs:
        u = str(url).strip() if url is not None else ""
        if (name is None) or (not u.lower().startswith(("http://", "https://"))):
            continue
        for k in _img_keys(name):
            m[k] = u
    return m


@st.cache_data(show_spinner=False)
def load_image_map(sig: tuple) -> dict:
    """이미지 매핑 로드. parquet(이름 컬럼 + 이미지/URL 컬럼) 과 엑셀 둘 다 지원.
    엑셀은 시트명에 '이미지'가 있으면 그 시트, 없으면 첫 시트의 A열=이름 · B열=URL."""
    out = {}
    for path, _mt in sig:
        p = Path(path)
        try:
            if p.suffix.lower() == ".parquet":
                d = pd.read_parquet(p)
                d.columns = [clean_col_name(c) for c in d.columns]
                ncol = next((c for c in IMG_NAME_KEYS if c in d.columns), d.columns[0])
                ucol = next((c for c in d.columns
                             if ("이미지" in str(c)) or ("url" in str(c).lower())), None)
                if ucol is None:
                    continue
                out.update(_rows_to_img_map(zip(d[ncol], d[ucol])))
            else:
                xls = pd.ExcelFile(p)
                target = next((s for s in xls.sheet_names
                               if "이미지" in str(s) or "image" in str(s).lower()),
                              xls.sheet_names[0])
                d = pd.read_excel(xls, sheet_name=target, header=None, usecols=[0, 1])
                out.update(_rows_to_img_map(zip(d.iloc[:, 0], d.iloc[:, 1])))
        except Exception:
            continue
    return out


def _sigs(paths) -> tuple:
    """캐시 키: (경로, 수정시각). 파일이 바뀌면 자동으로 다시 읽는다."""
    out = []
    for p in paths:
        try:
            out.append((str(p), Path(p).stat().st_mtime))
        except OSError:
            out.append((str(p), 0.0))
    return tuple(out)

# -----------------------------
# Style
# -----------------------------
st.markdown(
    """
    <style>
    .main .block-container {padding-top: 1.5rem; padding-bottom: 2rem;}
    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #e8eef5;
        padding: 14px 14px;
        border-radius: 16px;
        box-shadow: 0 4px 16px rgba(15, 23, 42, 0.04);
    }
    div[data-testid="stMetric"] label {color:#64748b; font-size:0.85rem;}
    div[data-testid="stMetricValue"] {
        font-size: 1.3rem;
        font-weight: 700;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    div[data-testid="stMetricValue"] > div {
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }
    .section-title {
        font-size: 1.15rem;
        font-weight: 800;
        margin: 1.2rem 0 .4rem 0;
        color: #0f172a;
    }
    .hint {
        color: #64748b;
        font-size: .9rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# -----------------------------
# Helpers
# -----------------------------
def clean_col_name(col: str) -> str:
    return re.sub(r"\s+", " ", str(col).replace("\n", " ")).strip()


def find_col(df: pd.DataFrame, candidates: list[str], fallback_contains: str | None = None) -> str | None:
    cols = list(df.columns)
    for c in candidates:
        if c in cols:
            return c
    if fallback_contains:
        for c in cols:
            if fallback_contains in c:
                return c
    return None


def to_number(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series(dtype=float)
    return pd.to_numeric(
        s.astype(str).str.replace(",", "", regex=False).str.replace("%", "", regex=False),
        errors="coerce",
    ).fillna(0)


def _parse_date_flexible(series: pd.Series) -> pd.Series:
    """날짜 안전 파싱. 엑셀 serial(정수)·YYYYMMDD·다양한 포맷 문자열·datetime 객체 혼재 모두 처리.
    숫자를 무조건 epoch(나노초)로 보던 문제(45292/20240715/0 → 1970-01 → FW69) 방지."""
    if series is None or len(series) == 0:
        return pd.Series(pd.NaT, index=getattr(series, "index", None), dtype="datetime64[ns]")
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce")
    out = pd.Series(pd.NaT, index=series.index, dtype="datetime64[ns]")
    # 1) 이미 날짜/시간 객체 (openpyxl datetime 등)
    is_dt = series.apply(lambda v: (not isinstance(v, str)) and hasattr(v, "year"))
    if is_dt.any():
        out.loc[is_dt] = pd.to_datetime(series[is_dt], errors="coerce")
    # 2) 숫자: YYYYMMDD(8자리) vs 엑셀 serial 구분
    num = pd.to_numeric(series.where(~is_dt), errors="coerce")
    ymd = num.between(19000101, 21001231)
    if ymd.any():
        out.loc[ymd] = pd.to_datetime(
            num[ymd].round().astype("int64").astype(str), format="%Y%m%d", errors="coerce")
    ser = num.between(20000, 60000) & ~ymd  # 약 1954~2064년 엑셀 날짜 serial
    if ser.any():
        out.loc[ser] = pd.to_datetime(num[ser], unit="D", origin="1899-12-30", errors="coerce")
    # 3) 나머지 문자열 (2024-07-15, 2024.7.15, 2024/07/15, 시간 포함 등 포맷 혼재 가능)
    _txt = series.astype(str).str.strip()
    rest = (~is_dt) & num.isna() & _txt.ne("") & ~_txt.str.lower().isin(["nan", "nat", "none"])
    if rest.any():
        try:
            out.loc[rest] = pd.to_datetime(_txt[rest], errors="coerce", format="mixed")
        except (ValueError, TypeError):
            out.loc[rest] = pd.to_datetime(_txt[rest], errors="coerce")
    return out


def money(v) -> str:
    try:
        return f"{float(v):,.0f}원"
    except Exception:
        return "0원"


def num(v) -> str:
    try:
        return f"{float(v):,.0f}"
    except Exception:
        return "0"


def eok(v) -> str:
    """금액 표기: 1억 이상이면 'X.X억', 그 미만이면 전체 숫자(콤마)."""
    try:
        v = float(v)
    except Exception:
        return "0"
    if pd.isna(v):
        return "-"
    if abs(v) >= 1e8:
        return f"{v / 1e8:.1f}억"
    return f"{v:,.0f}"


def to_line(model) -> str:
    """모델명에서 끝의 사이즈 '(...)' 를 떼어 라인명으로 변환.
    예: 'COHBU M26388 ALI BLANC/BLEU CIEL (XL)' -> 'COHBU M26388 ALI BLANC/BLEU CIEL'
        'K100979-001 (44)' -> 'K100979-001'
    """
    return re.sub(r"\s*\([^()]*\)\s*$", "", str(model)).strip()


line_map: dict = {}  # 모델명(정규화) → 라인명. 로드 시 재고 parquet 의 라인명으로 채움.


def _norm_model(s) -> str:
    return re.sub(r"\s+", " ", str(s)).strip()


def _line_of(model) -> str:
    """모델명 → 라인명. line_map(재고 parquet) 우선, 없으면 끝의 사이즈 '(...)' 제거."""
    nm = _norm_model(model)
    if nm in line_map:
        return line_map[nm]
    nm2 = _norm_model(to_line(model))
    if nm2 in line_map:
        return line_map[nm2]
    return to_line(model)


def pct(v) -> str:
    try:
        if pd.isna(v) or np.isinf(v):
            return "-"
        return f"{float(v):,.1f}%"
    except Exception:
        return "-"


def growth_pct(v) -> str:
    """Format growth-rate columns like ▲ 18.0% / ▼ 15.0% with one decimal."""
    try:
        if pd.isna(v) or np.isinf(v):
            return "-"
        value = float(v)
        if value > 0:
            return f"▲ {abs(value):,.1f}%"
        if value < 0:
            return f"▼ {abs(value):,.1f}%"
        return "0.0%"
    except Exception:
        return "-"


def add_rate(df: pd.DataFrame, current_col: str, prev_col: str, out_col: str = "YoY 신장률") -> pd.DataFrame:
    prev = df[prev_col].replace(0, np.nan)
    df[out_col] = ((df[current_col] - df[prev_col]) / prev.abs()) * 100
    return df


def sort_desc(df: pd.DataFrame, by: str) -> pd.DataFrame:
    if by in df.columns:
        return df.sort_values(by=by, ascending=False, na_position="last")
    return df


# 표/차트의 '총매출' 라벨 접두어 — 본문에서 모드(주간/월간)에 따라 재설정됨
INTERVAL_LABEL = "주간"


def format_table(df: pd.DataFrame):
    """표시용 (DataFrame, column_config) 반환.
    금액/수량은 '숫자'로 유지해 콤마 표시(localized)와 숫자 정렬이 둘 다 되게 한다.
    비율(수익률/비중)은 숫자+'%' 포맷, 증감률(신장률)만 ▲▼ 문자열, Rank는 문자열.
    """
    out = df.copy()
    # 수량을 최종판매가(총매출) 바로 왼쪽으로 이동
    if "수량" in out.columns and "최종판매가" in out.columns:
        cols = list(out.columns)
        cols.remove("수량")
        cols.insert(cols.index("최종판매가"), "수량")
        out = out[cols]
    rename_map = {}
    if "최종판매가" in out.columns:
        rename_map["최종판매가"] = "총매출"
    if "수익원(실배송비)" in out.columns:
        rename_map["수익원(실배송비)"] = "수익원"
    if rename_map:
        out = out.rename(columns=rename_map)
    money_keywords = ["판매가", "매출", "수익원", "원가", "증감", "객단가", "수수료액", "배송비", "신장액"]
    pct_keywords = ["률", "율", "비중", "Rate"]
    colcfg = {}
    for c in out.columns:
        cs = str(c)
        if c == "Rank":
            out[c] = pd.to_numeric(out[c], errors="coerce").apply(lambda x: "-" if pd.isna(x) else f"{x:,.0f}")
        elif "신장률" in cs or "신장율" in cs:
            out[c] = pd.to_numeric(out[c], errors="coerce").apply(growth_pct)
        elif any(k in cs for k in pct_keywords):
            out[c] = pd.to_numeric(out[c], errors="coerce").replace([np.inf, -np.inf], np.nan)
            colcfg[c] = st.column_config.NumberColumn(cs, format="%.1f%%")
        elif any(k in cs for k in money_keywords) or re.fullmatch(r"\d{4}년", cs) or re.fullmatch(r"\d{1,2}월\s?\d{1,2}주차", cs):
            out[c] = pd.to_numeric(out[c], errors="coerce").round(0).astype("Int64")
            colcfg[c] = st.column_config.NumberColumn(cs, format="localized")
        elif cs in ("수량", "주문수"):
            out[c] = pd.to_numeric(out[c], errors="coerce").round(0).astype("Int64")
            colcfg[c] = st.column_config.NumberColumn(cs, format="localized")
        else:
            out[c] = out[c].replace({None: "-", np.nan: "-"})
    return out, colcfg


def aggregate(df: pd.DataFrame, group_cols: list[str], metric_cols: dict) -> pd.DataFrame:
    agg_spec = {}
    for out, col in metric_cols.items():
        if col and col in df.columns:
            agg_spec[out] = (col, "sum")
    result = df.groupby(group_cols, dropna=False).agg(**agg_spec).reset_index()
    if "최종판매가" in result.columns and "수량" in result.columns:
        result["객단가"] = np.where(result["수량"] != 0, result["최종판매가"] / result["수량"], 0)
    if "수익원(실배송비)" in result.columns and "최종판매가" in result.columns:
        result["수익률"] = np.where(result["최종판매가"] != 0, result["수익원(실배송비)"] / result["최종판매가"] * 100, 0)
    total = result["최종판매가"].sum() if "최종판매가" in result.columns else 0
    if total != 0 and "최종판매가" in result.columns:
        result["매출비중"] = result["최종판매가"] / total * 100
    return sort_desc(result, "최종판매가")



def rank_table(df: pd.DataFrame, name_col: str) -> pd.DataFrame:
    """Add rank and replace the displayed name with rank order prefix."""
    out = df.copy().reset_index(drop=True)
    out.insert(0, "Rank", np.arange(1, len(out) + 1))
    if name_col in out.columns:
        clean_name = out[name_col].astype(str).str.replace(r"^\s*\d+\s*[\.\)\-_/]*\s*", "", regex=True)
        out[name_col] = out["Rank"].astype(str) + ". " + clean_name
    return out


def top_sales_table(df: pd.DataFrame, group_cols: list[str], topn: int = 30, sort_by: str = "최종판매가") -> pd.DataFrame:
    table = aggregate(df, group_cols, metric_cols)
    if sort_by in table.columns:
        table = table.sort_values(sort_by, ascending=False, na_position="last")
    table = table.head(topn).reset_index(drop=True)
    table.insert(0, "Rank", np.arange(1, len(table) + 1))
    return table

def yoy_by_group(df: pd.DataFrame, group_col: str, base_year: int, metric_col: str) -> pd.DataFrame:
    prev_year = base_year - 1
    temp = df[df["연도"].isin([prev_year, base_year])]
    pivot = temp.pivot_table(index=group_col, columns="연도", values=metric_col, aggfunc="sum", fill_value=0).reset_index()
    if prev_year not in pivot.columns:
        pivot[prev_year] = 0
    if base_year not in pivot.columns:
        pivot[base_year] = 0
    pivot = pivot.rename(columns={prev_year: f"{prev_year}년", base_year: f"{base_year}년"})
    pivot["YoY 신장액"] = pivot[f"{base_year}년"] - pivot[f"{prev_year}년"]
    pivot = add_rate(pivot, f"{base_year}년", f"{prev_year}년")
    return sort_desc(pivot, f"{base_year}년")


def trend_by_group(df: pd.DataFrame, group_col: str, metric_col: str, topn=None) -> pd.DataFrame:
    """그룹별 연도 추이: 데이터에 존재하는 모든 연도를 열로 펼친다. 최신 연도 기준 내림차순."""
    years = sorted(int(y) for y in df["연도"].dropna().unique())
    pivot = df.pivot_table(
        index=group_col, columns="연도", values=metric_col, aggfunc="sum", fill_value=0
    ).reset_index()
    for y in years:
        if y not in pivot.columns:
            pivot[y] = 0
    pivot = pivot.rename(columns={y: f"{y}년" for y in years})
    year_cols = [f"{y}년" for y in years]
    pivot = pivot[[group_col] + year_cols]
    if year_cols:
        pivot = pivot.sort_values(year_cols[-1], ascending=False, na_position="last")
    pivot = pivot.reset_index(drop=True)
    if topn:
        pivot = pivot.head(topn)
    return pivot


def wow_by_group(df: pd.DataFrame, group_col: str, metric_col: str, week_order: list, topn=None) -> pd.DataFrame:
    """그룹별 주차 추이: week_order(시간순 주차 라벨)대로 열을 펼친다. 최신 주차 기준 내림차순."""
    pivot = df.pivot_table(
        index=group_col, columns="주차", values=metric_col, aggfunc="sum", fill_value=0
    ).reset_index()
    week_cols = [w for w in week_order if w in pivot.columns]
    pivot = pivot[[group_col] + week_cols]
    if week_cols:
        pivot = pivot.sort_values(week_cols[-1], ascending=False, na_position="last")
    pivot = pivot.reset_index(drop=True)
    if topn:
        pivot = pivot.head(topn)
    return pivot


def _prep_sales(d: pd.DataFrame) -> pd.DataFrame:
    """parquet 매출 원본에 엑셀 전처리(process_excel)와 같은 규칙을 적용한다.
    쇼핑몰명 통일(_D_RENAME_MAP) · 제외 대상 행 삭제 · 공식/병행 분류."""
    d = d.copy()
    # 출고날짜는 parquet 에서 문자열로 들어오므로 여기서 datetime 으로 변환
    #   (반품행의 출고날짜 보정 시 문자열 컬럼에 날짜를 넣으면 pandas 가 거부한다)
    for _dc in ("출고날짜", "출고일", "판매일자", "주문일자"):
        if _dc in d.columns:
            d[_dc] = _parse_date_flexible(d[_dc])
    if "쇼핑몰" in d.columns:
        d["쇼핑몰"] = d["쇼핑몰"].astype(str).str.strip().replace(_D_RENAME_MAP)
        _kill = d["쇼핑몰"].apply(lambda s: any(k in s for k in _DELETE_D_KEYWORDS))
        d = d[~_kill]
    if "모델명" in d.columns:
        d = d[~d["모델명"].astype(str).str.strip().isin(_DELETE_H_VALUES)]
    if {"브랜드", "대카테고리", "모델명"} <= set(d.columns):
        d["공식/병행"] = [_classify_official(b, g, h) for b, g, h
                        in zip(d["브랜드"].astype(str), d["대카테고리"].astype(str), d["모델명"].astype(str))]
    return d.reset_index(drop=True)


@st.cache_data(show_spinner="매출 데이터 로딩 중…")
def load_sales_parquet(sigs: tuple, stock_sigs: tuple = ()) -> pd.DataFrame:
    """매출 parquet(여러 연도) → 최종 DataFrame. 파일이 여러 개면 전부 합산한다.
    라인명 부여까지 이 캐시 안에서 끝낸다(130만 행 map 을 재실행마다 반복하지 않도록).
    ※ 라인명은 재고 parquet 의 line_map 에 의존하므로 stock_sigs 도 캐시 키에 포함한다.
      (호출 전에 line_map 이 채워져 있어야 한다 — 재고를 먼저 로드할 것)"""
    frames = []
    for path, _mt in sigs:
        one = pd.read_parquet(path)
        one.columns = [clean_col_name(c) for c in one.columns]
        frames.append(one)
    if not frames:
        return pd.DataFrame()
    raw = pd.concat(frames, ignore_index=True) if len(frames) > 1 else frames[0]
    df = _finalize_df(_prep_sales(raw))
    # 라인명 부여(베스트 상품을 라인명으로 취합) — 고유 모델명에만 함수 적용 후 map
    if "모델명" in df.columns:
        _uniq = pd.Index(df["모델명"].astype(str).unique())
        df["라인명"] = df["모델명"].astype(str).map({m: _line_of(m) for m in _uniq})
    else:
        df["라인명"] = ""
    return df


@st.cache_data(show_spinner="재고 원가 추정 중…", max_entries=3)
def attach_stock_cost_cached(_sdf: pd.DataFrame, _sales: pd.DataFrame, base, key: tuple) -> pd.DataFrame:
    """attach_stock_cost 캐시 래퍼. 결과는 브랜드/필터와 무관하고 parquet 이 바뀔 때만 달라진다.
    ※ _sdf/_sales 는 언더스코어 인자라 해싱하지 않는다(760MB DataFrame 해싱 방지).
      캐시 키는 key(매출·재고 파일의 경로+수정시각) + base(재고 스냅샷 기준일)."""
    return attach_stock_cost(_sdf, _sales, base)


@st.cache_data(show_spinner="출고 raw 불러오는 중…", max_entries=8)
def load_raw_by_models(sigs: tuple, models: tuple) -> pd.DataFrame:
    """검색된 모델명들의 매출 parquet '원본 행'을 전처리 없이 그대로 읽는다.
    (_finalize_df 는 메모리 절약을 위해 주문번호·출고원가·카테고리 등을 버리므로
     raw 표시는 parquet 을 다시 읽어서 원본 컬럼을 전부 보여준다.)"""
    ms = [str(m) for m in models]
    if not ms:
        return pd.DataFrame()
    frames = []
    for path, _mt in sigs:
        one = None
        try:
            import pyarrow.parquet as pq
            one = pq.read_table(str(path), filters=[("모델명", "in", ms)]).to_pandas()
        except Exception:
            try:
                _d = pd.read_parquet(path)
                one = _d[_d["모델명"].astype(str).isin(set(ms))]
            except Exception:
                continue
        if one is not None and len(one):
            frames.append(one)
    if not frames:
        return pd.DataFrame()
    raw = pd.concat(frames, ignore_index=True)
    if "출고날짜" in raw.columns:
        raw = raw.sort_values("출고날짜", ascending=False, kind="mergesort")
    return raw.reset_index(drop=True)


def _finalize_df(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [clean_col_name(c) for c in df.columns]

    # Detect columns
    date_col = find_col(df, ["출고날짜", "출고일", "판매일자", "주문일자"], "날짜")
    qty_col = find_col(df, ["수량", "판매수량"], "수량")
    gross_col = find_col(df, ["매출가", "총매출액", "매출액"], "매출")
    net_col = find_col(df, ["최종판매가", "순매출액", "실매출액"], "최종")
    profit_col = find_col(df, ["수익원(실배송비)", "수익원 실배송비", "수익원", "공헌이익"], "수익원")
    cost_col = find_col(df, ["원가총액", "출고원가"], "원가")
    mall_col = find_col(df, ["쇼핑몰", "몰", "채널"], "쇼핑몰")
    brand_col = find_col(df, ["브랜드", "브랜드명"], "브랜드")
    # 대분류: 위에서 AA열을 '대분류'로 보존했다. G열 '대카테고리'(브랜드패션 등)는 절대 쓰지 않는다.
    if "대분류" in df.columns:
        category_col = "대분류"
    else:
        category_col = find_col(df, ["카테고리", "분류"], "분류")
    model_col = find_col(df, ["모델명", "상품명", "품목명", "상품코드"], "모델")
    order_col = find_col(df, ["주문번호", "주문ID", "주문코드"], "주문")
    note_col = find_col(df, ["비고", "상태", "구분"], "비고")

    # 공식/병행 구분 컬럼: 헤더명 우선, 없으면 값이 공식/병행 으로만 이루어진 컬럼을 자동 탐색
    official_col = find_col(df, ["공식/병행", "공식병행", "공식여부"])
    if official_col is None:
        for _c in df.columns:
            _vals = set(df[_c].dropna().astype(str).str.strip().unique())
            if _vals and _vals <= {"공식", "병행"}:
                official_col = _c
                break

    # Standardize important columns
    if date_col is None:
        raise ValueError("날짜 컬럼을 찾지 못했습니다. '출고날짜' 또는 날짜가 포함된 컬럼이 필요합니다.")
    df["날짜"] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=["날짜"]).copy()
    # 반품(수량<0): 비고의 '반품(YYYY-MM-DD)' 실제 반품일을 출고날짜/분석날짜로 사용 (출고일 혼동 방지)
    if qty_col and note_col and qty_col in df.columns and note_col in df.columns:
        _ret = pd.to_numeric(df[qty_col], errors="coerce").fillna(0) < 0
        if _ret.any():
            _rd = pd.to_datetime(
                df.loc[_ret, note_col].astype(str).str.extract(r"(\d{4}[-./]\d{1,2}[-./]\d{1,2})", expand=False),
                errors="coerce")
            df.loc[_ret, "날짜"] = _rd.fillna(df.loc[_ret, "날짜"])
            if date_col and date_col in df.columns:
                df.loc[_ret, date_col] = df.loc[_ret, "날짜"]
    df["연도"] = df["날짜"].dt.year.astype(int)
    df["월"] = df["날짜"].dt.month.astype(int)
    df["연월"] = df["날짜"].dt.to_period("M").astype(str)
    # (주간/월간 기간 파생은 아래 수량 표준화 후에 수행 — 반품=수량 음수 제외 위해)
    df["요일순"] = df["날짜"].dt.weekday.astype(int)
    df["요일라벨"] = df["요일순"].map({0: "월", 1: "화", 2: "수", 3: "목", 4: "금", 5: "토", 6: "일"})

    for col in [qty_col, gross_col, net_col, profit_col, cost_col]:
        if col and col in df.columns:
            df[col] = to_number(df[col])

    if qty_col and qty_col != "수량":
        df["수량"] = df[qty_col]
    elif "수량" not in df.columns:
        df["수량"] = 0

    # (주간/월간/연간 기간 파생[주시작·주차]은 본문에서 '합쳐진 전체 df' 기준으로 수행한다.
    #  멀티파일 업로드 시 파일마다 따로 잡혀 라벨/모드가 틀어지는 문제 방지)

    if gross_col and gross_col != "매출가":
        df["매출가"] = df[gross_col]
    elif "매출가" not in df.columns:
        df["매출가"] = 0

    if net_col and net_col != "최종판매가":
        df["최종판매가"] = df[net_col]
    elif "최종판매가" not in df.columns:
        df["최종판매가"] = df["매출가"]

    # 매출가 컬럼이 비어있는(거의 0인) export 에서는 최종판매가를 매출가로 사용한다.
    if df["매출가"].abs().sum() == 0 or (df["매출가"] != 0).mean() < 0.05:
        df["매출가"] = df["최종판매가"]

    if profit_col and profit_col != "수익원(실배송비)":
        df["수익원(실배송비)"] = df[profit_col]
    elif "수익원(실배송비)" not in df.columns:
        df["수익원(실배송비)"] = 0

    if cost_col and cost_col != "원가총액":
        df["원가총액"] = df[cost_col]
    elif "원가총액" not in df.columns:
        df["원가총액"] = 0

    for std, col in {
        "쇼핑몰": mall_col,
        "브랜드": brand_col,
        "대분류": category_col,
        "공식/병행": official_col,
        "모델명": model_col,
        "주문번호": order_col,
        "비고": note_col,
    }.items():
        if col and col in df.columns:
            df[std] = df[col].fillna("미분류").astype(str)
        elif std not in df.columns:
            df[std] = "미분류"
        else:
            df[std] = df[std].fillna("미분류").astype(str)

    # Normalize text values
    for c in ["쇼핑몰", "브랜드", "대분류", "공식/병행", "모델명", "비고"]:
        df[c] = df[c].replace({"nan": "미분류", "None": "미분류", "": "미분류"})

    # 대분류 정규화: 세부 카테고리(드레스/모자/상의 등)를 _CATEGORY_MAP 으로 8개 대분류에 통합.
    #   8개(시계/주얼리/가방/지갑/의류/신발/소품/용품) 밖이면 '미분류'. 원본은 진단용 보존.
    try:
        df["대분류_원본"] = df["대분류"].astype(str)
        df["대분류"] = df["대분류"].map(_classify_to8)
    except Exception:
        pass

    # 사은품/쇼핑백 제외 (매출 분석 공통 기준) — 캐시 구간에서 한 번만 수행
    df = df[~df["브랜드"].astype(str).str.strip().isin(_GIFT_BRANDS)]

    # 대시보드에서 쓰지 않는 컬럼 제거 (85만행 × 20여 컬럼 → 메모리 절반 이하)
    _KEEP = ["날짜", "연도", "쇼핑몰", "브랜드", "대분류", "대분류_원본", "모델명", "비고",
             "수량", "최종판매가", "수익원(실배송비)", "원가총액"]
    df = df[[c for c in _KEEP if c in df.columns]]
    return df.reset_index(drop=True)


_STOCK_CAT_DEBUG = {}  # 재고가 분류로 읽은 컬럼/매핑률 진단 기록


def _max_days_ago(text):
    """문자열의 'N일전' 들 중 가장 큰 N(가장 오래된 입고경과일) 반환. 여러 입고건/사이즈 대비. 없으면 NaN."""
    nums = re.findall(r"(\d+)\s*일\s*전", str(text))
    return max(int(n) for n in nums) if nums else np.nan


def _inbound_qty(text):
    """'N일전/M' 또는 'N일전M개' 들의 M(입고수량) 합 = 누적 입고량. 없으면 0."""
    pairs = re.findall(r"\d+\s*일\s*전\s*/?\s*([\d,]+)", str(text))
    return sum(int(m.replace(",", "")) for m in pairs) if pairs else 0


def _parse_inbound_events(text):
    """'N일전/M'(또는 'N일전M개') → [(N=경과일, M=수량), ...]. FIFO 입고 큐 구성용."""
    return [(int(n), int(m.replace(",", "")))
            for n, m in re.findall(r"(\d+)\s*일\s*전\s*/?\s*([\d,]+)", str(text))]


@st.cache_data(show_spinner=False)
def load_stock_parquet(sigs: tuple) -> pd.DataFrame:
    """재고 parquet → 표준 재고 DataFrame.
    입력 컬럼: 라인명·브랜드·대카테고리·카테고리·종류·성별·모델명·수량·가용수량·
              입고이력('N일전/수량')·총입고량·입고경과일·회전율·기준일
    ※ parquet 에는 원가가 없으므로 원가평균/총원가는 attach_stock_cost() 에서
      판매 실적의 실제 출고원가로 '추정' 해 채운다."""
    frames = []
    for path, _mt in sigs:
        one = pd.read_parquet(path)
        one.columns = [clean_col_name(c) for c in one.columns]
        frames.append(one)
    if not frames:
        return pd.DataFrame()
    sdf = pd.concat(frames, ignore_index=True) if len(frames) > 1 else frames[0]

    out = pd.DataFrame(index=range(len(sdf)))
    for c in ("라인명", "브랜드", "모델명", "대카테고리", "카테고리", "종류", "성별"):
        out[c] = sdf[c].astype(str).str.strip().values if c in sdf.columns else ""
    if "모델명" not in sdf.columns:
        out["모델명"] = out["라인명"]
    out["수량"] = to_number(sdf["수량"]).values if "수량" in sdf.columns else 0
    out["가용수량"] = to_number(sdf["가용수량"]).values if "가용수량" in sdf.columns else np.nan

    # 대분류: 후보(카테고리/종류) 중 8개 대분류 매핑률이 높은 컬럼 채택 (매출과 동일 규칙)
    _cands = [c for c in ("카테고리", "종류", "소분류", "품목") if c in sdf.columns]
    _best_c, _best_r, _rep = None, -1.0, {}
    for c in _cands:
        mp = sdf[c].astype(str).str.strip().map(_classify_to8)
        r = float((mp != "미분류").mean()) if len(mp) else 0.0
        _rep[c] = round(r, 3)
        if r > _best_r:
            _best_r, _best_c = r, c
    _raw_cat = (sdf[_best_c].astype(str).str.strip() if (_best_c and _best_r > 0)
                else pd.Series(out["카테고리"].values))
    out["대분류"] = _raw_cat.map(_classify_to8).values
    out["대분류_원본"] = _raw_cat.values
    _STOCK_CAT_DEBUG.clear()
    _STOCK_CAT_DEBUG.update({"selected": str(_best_c), "rate": round(_best_r, 3),
                             "candidates": _rep, "all_cols": [str(c) for c in sdf.columns]})

    # 기준일(재고 스냅샷 날짜) — 'N일전' 을 실제 날짜로 되돌리는 기준
    base = pd.to_datetime(sdf["기준일"], errors="coerce").max() if "기준일" in sdf.columns else pd.NaT
    if pd.isna(base):
        base = pd.Timestamp.now().normalize()
    out["기준일"] = base

    hist = sdf["입고이력"].astype(str) if "입고이력" in sdf.columns else None
    if "입고경과일" in sdf.columns:
        _el = pd.to_numeric(sdf["입고경과일"], errors="coerce")
    elif hist is not None:
        _el = hist.map(_max_days_ago)
    else:
        _el = pd.Series(np.nan, index=sdf.index)
    out["입고경과일행"] = _el.values                       # 가장 오래된 입고 경과일
    if "총입고량" in sdf.columns:
        out["입고수량합행"] = to_number(sdf["총입고량"]).values
    else:
        out["입고수량합행"] = hist.map(_inbound_qty).values if hist is not None else 0
    out["입고이벤트"] = (hist.map(_parse_inbound_events).values if hist is not None
                     else [[] for _ in range(len(out))])
    # 입고일자 = 기준일 − 가장 오래된 입고 경과일 (입고연도 판정용)
    out["입고일자"] = [base - pd.Timedelta(days=int(v)) if pd.notna(v) else pd.NaT for v in _el]
    out["입고원가이벤트"] = [[] for _ in range(len(out))]   # 원가 추정 후 attach_stock_cost 에서 채움
    out["원가평균"] = np.nan
    out["총원가"] = 0.0
    out["공식/병행"] = [_classify_official(b, d, m) for b, d, m
                     in zip(out["브랜드"], out["대카테고리"], out["모델명"])]
    out = out[~(out["라인명"].isin(["", "nan", "None"]) & out["모델명"].isin(["", "nan", "None"]))]
    return out.reset_index(drop=True)


def attach_stock_cost(sdf: pd.DataFrame, sales: pd.DataFrame, base) -> pd.DataFrame:
    """재고 parquet 에 원가가 없으므로 판매 실적의 실제 출고원가로 개당원가를 추정한다.
    우선순위: 같은 모델명(사이즈까지 동일) → 같은 라인명 → 같은 브랜드×대분류 중앙값.
    총원가 = 재고수량 × 추정 개당원가. (재고금액·입고연도별 입고원가는 모두 '추정치')"""
    if sdf is None or sdf.empty or sales is None or sales.empty:
        return sdf
    need = {"모델명", "수량", "원가총액"}
    if not need <= set(sales.columns):
        return sdf
    s = sales.copy()
    s["수량"] = pd.to_numeric(s["수량"], errors="coerce").fillna(0)
    s["원가총액"] = pd.to_numeric(s["원가총액"], errors="coerce").fillna(0)
    s = s[(s["수량"] > 0) & (s["원가총액"] > 0)]
    if s.empty:
        return sdf

    def _unit(keys):
        gp = s.groupby(keys)
        return (gp["원가총액"].sum() / gp["수량"].sum().replace(0, np.nan))

    by_model = _unit(s["모델명"].astype(str).str.strip())
    by_line = (_unit(s["라인명"].astype(str).str.strip()) if "라인명" in s.columns
               else pd.Series(dtype=float))
    _u = s["원가총액"] / s["수량"]
    by_bc = (_u.groupby([s["브랜드"].astype(str), s["대분류"].astype(str)]).median()
             if {"브랜드", "대분류"} <= set(s.columns) else pd.Series(dtype=float))

    unit = sdf["모델명"].astype(str).str.strip().map(by_model)
    if len(by_line):
        unit = unit.fillna(sdf["라인명"].astype(str).str.strip().map(by_line))
    if len(by_bc):
        _bc = pd.Series([by_bc.get((str(b), str(c)), np.nan)
                         for b, c in zip(sdf["브랜드"], sdf["대분류"])], index=sdf.index)
        unit = unit.fillna(_bc)

    sdf = sdf.copy()
    sdf["원가평균"] = unit.values
    sdf["총원가"] = (pd.to_numeric(sdf["수량"], errors="coerce").fillna(0)
                  * unit.fillna(0).values).round(0)
    # 입고건별 입고원가(추정) = 입고건 수량 × 추정 개당원가
    ev_cost = []
    for evs, u in zip(sdf["입고이벤트"], unit.values):
        if (not isinstance(evs, list)) or (not evs) or pd.isna(u):
            ev_cost.append([])
            continue
        ev_cost.append([(base - pd.Timedelta(days=int(n)), float(u)) for n, _q in evs])
    sdf["입고원가이벤트"] = ev_cost
    return sdf


def line_map_from_stock(stock_df: pd.DataFrame) -> dict:
    """재고 DataFrame 에서 모델명→라인명 매핑 추출."""
    if stock_df is None or stock_df.empty:
        return {}
    m: dict[str, str] = {}
    for md, ln in zip(stock_df["모델명"], stock_df["라인명"]):
        k, v = _norm_model(md), str(ln).strip()
        if k and v and k.lower() not in ("nan", "none", "") and v.lower() not in ("nan", "none", ""):
            m.setdefault(k, v)
    return m


# ============================================================
# 매출 데이터 공통 규칙 (엑셀 전처리 process_excel.py 와 동일)
#   쇼핑몰명 통일 · 제외 대상 행 · 공식/병행 분류 · 8개 대분류 매핑
# ============================================================
_GIFT_BRANDS = ["쇼핑백", "사은품"]
_DELETE_H_VALUES = {"파슬AS", "쿠팡그로스 재고손실보상", "쿠팡그로스 기타정산"}
_DELETE_D_KEYWORDS = ["방송", "홈방", "나린인터", "태그바이"]
_D_RENAME_MAP = {
    "KREAM": "크림 주식회사", "카카오톡선물하기_디젤": "카카오톡선물하기",
    "카카오톡선물하기_병행": "카카오톡선물하기", "카카오톡선물하기_공식": "카카오톡선물하기",
    "에이블리(블리블리)": "에이블리", "에이블리(치페)": "에이블리",
    "무신사_블리블리": "무신사", "Wconcept(뷰티)": "Wconcept",
    "29CM(티켓투더문)": "29CM(공식)", "29CM(디젤)": "29CM(공식)",
    "카카오스타일 (치페)": "카카오스타일 (지그재그)",
    "카카오스타일 (티켓투더문)": "카카오스타일 (지그재그)",
    "카카오스타일 (블리블리)": "카카오스타일 (지그재그)",
}
_OFFICIAL_F_ONLY = {
    "블리블리", "헤브블루", "미스그린", "치페", "파슬", "아르마니", "티켓투더문",
    "아르마니익스체인지", "울프1834", "인도솔", "썬젤리", "스카겐", "미니쿄모", "스케쳐스",
}
_CATEGORY_MAP = {
    "가방": "가방", "귀걸이": "주얼리", "드레스": "의류", "라이터": "용품", "마사지볼": "용품",
    "모자": "소품", "목걸이": "주얼리", "문구": "용품", "반지": "주얼리", "밴드": "시계",
    "벨트": "소품", "상의": "의류", "시계": "시계", "신발": "신발", "아우터": "의류",
    "잡화ACC": "소품", "지갑": "지갑", "침낭": "용품", "키링&키홀더": "소품", "팔찌": "주얼리",
    "폼롤러": "용품", "하의": "의류", "핸드폰케이스": "소품", "홈데코": "용품", "우산": "소품",
    "옷걸이": "용품", "에어팟케이스": "용품", "언더웨어": "의류", "바디케어": "용품",
    "쇼핑백": "용품", "향수": "용품", "스킨케어": "용품", "거치대": "시계", "인솔": "용품",
    "쥬얼리보관함": "주얼리", "와인더": "시계", "시계보관함": "시계", "완구": "용품",
    "손난로": "용품", "참": "주얼리", "보온주머니": "용품", "생활잡화": "용품",
    "스포츠용품": "용품", "스윔웨어": "용품", "수납용품": "용품", "브로치": "소품",
    "케이블": "시계", "생활용품": "용품", "욕실용품": "용품", "슬립웨어": "의류",
    "아이메이크업": "용품", "립메이크업": "용품", "베이스메이크업": "용품", "뷰티소품": "용품",
    "클렌징": "용품", "선케어": "용품", "헤어케어": "용품", "주방용품": "용품",
}

# 대분류는 이 8개만 허용. _CATEGORY_MAP 으로 매핑 후 이 밖이면 '미분류'.
_ALLOWED_CATS = {"시계", "주얼리", "가방", "지갑", "의류", "신발", "소품", "용품"}

# 정확매칭(_CATEGORY_MAP)으로 안 잡히는 세부 명칭(반팔티/스니커즈 등)용 키워드 규칙.
#   더 구체적인 대분류를 앞에 둬 충돌 최소화(용품은 광범위해서 맨 뒤). 카테고리 글자에 키워드 포함 시 매칭.
_CAT_KEYWORDS = [
    ("시계", ["시계", "손목시계", "워치", "와인더"]),
    ("주얼리", ["목걸이", "귀걸이", "귀고리", "반지", "팔찌", "발찌", "브로치", "펜던트",
              "이어링", "이어커프", "네크리스", "뱅글", "앵클릿", "주얼리", "쥬얼리"]),
    ("지갑", ["지갑", "장지갑", "반지갑", "카드지갑", "카드케이스", "머니클립", "월렛", "코인케이스", "카드홀더"]),
    ("가방", ["가방", "백팩", "토트", "크로스", "숄더", "클러치", "파우치", "더플",
             "에코백", "메신저", "힙색", "보스턴", "버킷백", "호보", "쇼퍼백"]),
    ("신발", ["신발", "슈즈", "스니커", "운동화", "로퍼", "구두", "부츠", "샌들", "슬리퍼",
             "슬라이드", "뮬", "모카신", "더비", "옥스포드", "워커", "펌프스", "힐"]),
    ("소품", ["모자", "볼캡", "비니", "버킷햇", "페도라", "벨트", "스카프", "머플러", "목도리",
             "장갑", "양말", "삭스", "넥타이", "보타이", "손수건", "행커치프", "헤어밴드", "헤어핀",
             "머리끈", "머리띠", "집게핀", "선글라스", "안경테", "아이웨어", "키링", "키홀더"]),
    ("의류", ["반팔", "긴팔", "민소매", "나시", "티셔츠", "셔츠", "남방", "블라우스", "맨투맨",
             "후드", "후디", "니트", "스웨터", "가디건", "베스트", "조끼", "원피스", "드레스",
             "점프수트", "스커트", "치마", "팬츠", "바지", "슬랙스", "데님", "청바지", "조거",
             "트레이닝", "레깅스", "쇼츠", "반바지", "코트", "패딩", "다운", "점퍼", "자켓",
             "재킷", "블루종", "아노락", "아우터", "상의", "하의", "정장", "수트", "셋업",
             "폴로", "크롭", "슬립", "잠옷", "파자마", "언더웨어", "속옷", "브라", "드로즈",
             "러닝", "내의", "수영복", "스윔", "래쉬가드", "카라티"]),
    ("용품", ["향수", "퍼퓸", "코롱", "디퓨저", "캔들", "향초", "바디", "핸드크림", "풋크림",
             "로션", "크림", "세럼", "앰플", "에센스", "토너", "스킨", "클렌징", "마스크팩",
             "립밤", "마스카라", "섀도", "립스틱", "틴트", "쿠션팩트", "파운데이션", "컨실러",
             "선크림", "선스틱", "미스트", "헤어오일", "샴푸", "트리트먼트", "욕실", "주방",
             "텀블러", "보틀", "머그", "문구", "노트", "완구", "인형", "수납", "정리함",
             "폼롤러", "마사지", "우산", "양산", "담요", "블랭킷", "쿠션", "홈데코", "거치대",
             "충전기", "케이블", "보조배터리", "에어팟", "폰케이스", "그립톡", "스마트톡"]),
]


def _classify_to8(raw) -> str:
    """세부 카테고리를 8개 대분류로. 정확매칭(_CATEGORY_MAP) → 키워드 부분매칭 → 미분류."""
    s = str(raw).strip()
    if not s or s in ("nan", "None", "미분류"):
        return "미분류"
    v = _CATEGORY_MAP.get(s)
    if v in _ALLOWED_CATS:
        return v
    if s in _ALLOWED_CATS:          # 이미 대분류
        return s
    z = s.replace(" ", "")
    for cat, kws in _CAT_KEYWORDS:
        for kw in kws:
            if kw in z:
                return cat
    return "미분류"
def _classify_official(brand, daecat, cat) -> str:
    """브랜드 + 대카테고리 + 카테고리로 공식/병행 분류 (판매·재고 공통 로직)."""
    f = str(brand).strip(); g = str(daecat).strip(); h = str(cat).strip()
    if f in _OFFICIAL_F_ONLY:                               return "공식"
    if f == "마이클코어스" and g == "시계쥬얼리":          return "공식"
    if f == "디젤" and g == "시계쥬얼리":                  return "공식"
    if f == "라코스테" and g == "브랜드패션":              return "공식"
    if f == "토리버치" and h.startswith("TBW"):            return "공식"
    if f == "비비안웨스트우드" and h.startswith("VV"):     return "공식"
    return "병행"


# -----------------------------
# UI
# -----------------------------
st.title("🏷️ 브랜드 매출 대시보드")
st.caption("브랜드 1개 선택 → 연도·분기별 추이 · 쇼핑몰별 성과 · 카테고리/상품 비중")

_sales_paths, _stock_paths, _img_paths = scan_parquet_files(APP_DIR)
# 이미지 매핑은 parquet 우선. parquet 이 있으면 느린 엑셀(이미지*.xlsx)은 읽지 않는다.
_img_xlsx = find_image_file()
_img_xlsx_skipped = None
if _img_xlsx is not None:
    if _img_paths:
        _img_xlsx_skipped = _img_xlsx
    else:
        _img_paths = _img_paths + [_img_xlsx]

with st.sidebar:
    st.header("데이터")
    if _sales_paths:
        st.caption("📄 매출: " + " · ".join(f"**{p.name}**" for p in _sales_paths))
    else:
        st.caption("📄 매출 parquet 없음")
    if _stock_paths:
        st.caption("📦 재고: " + " · ".join(f"**{p.name}**" for p in _stock_paths))
    else:
        st.caption("📦 재고 parquet 없음 — 재고 섹션이 표시되지 않습니다")
    if _img_paths:
        st.caption("🖼 이미지: " + " · ".join(f"**{p.name}**" for p in _img_paths))
        if _img_xlsx_skipped is not None:
            st.caption(f"   ↳ parquet 우선 — `{_img_xlsx_skipped.name}` 은 읽지 않음(로딩 속도)")
    else:
        st.caption("🖼 이미지 없음 — 같은 폴더에 이미지 parquet(모델명/라인명 + 이미지URL) 또는 "
                   "'이미지.xlsx'(A열 이름 · B열 URL)를 두면 상품 사진이 표시됩니다")
    _up = st.file_uploader("데이터 추가 업로드 (선택)", type=["parquet", "xlsx", "xls"],
                           accept_multiple_files=True, key="pq_upl")
    if _up:
        _tmp = Path(tempfile.gettempdir()) / "_brand_parquet"
        _tmp.mkdir(parents=True, exist_ok=True)
        for _uf in _up:
            _fp = _tmp / _uf.name
            _fp.write_bytes(_uf.getvalue())
            if _fp.suffix.lower() in (".xlsx", ".xls") and _fp.name not in {q.name for q in _img_paths}:
                _img_paths = _img_paths + [_fp]     # 엑셀 업로드 = 이미지 매핑으로 취급
        _s2, _k2, _i2 = scan_parquet_files(_tmp)
        _sales_paths = _sales_paths + [p for p in _s2 if p.name not in {q.name for q in _sales_paths}]
        _stock_paths = _stock_paths + [p for p in _k2 if p.name not in {q.name for q in _stock_paths}]
        _img_paths = _img_paths + [p for p in _i2 if p.name not in {q.name for q in _img_paths}]
    st.caption("앱과 같은 폴더의 파일을 컬럼 구성으로 자동 판별합니다 "
               "(매출: 주문번호·쇼핑몰·출고날짜 / 재고: 라인명·입고이력·가용수량 / "
               "이미지: 라인명+이미지URL).")

try:
    if not _sales_paths:
        st.error("매출 parquet 파일을 찾지 못했습니다. "
                 f"'{APP_DIR}' 폴더에 매출 parquet(주문번호·쇼핑몰·출고날짜 포함)을 두거나 "
                 "왼쪽에서 업로드하세요.")
        st.stop()

    # 재고를 먼저 로드한다 — 매출의 '라인명' 이 재고에서 만든 line_map 에 의존하기 때문.
    img_map = load_image_map(_sigs(_img_paths)) if _img_paths else {}
    stock_df = load_stock_parquet(_sigs(_stock_paths)) if _stock_paths else pd.DataFrame()
    if not stock_df.empty:
        line_map.update(line_map_from_stock(stock_df))     # 재고의 모델명→라인명
        _b = pd.to_datetime(stock_df["기준일"], errors="coerce").max()
        if pd.notna(_b):
            STOCK_BASE_DATE = pd.Timestamp(_b).normalize()

    # 매출 로드 + 라인명 부여 (둘 다 캐시 안에서 처리 — 재실행마다 130만 행을 다시 훑지 않는다)
    df = load_sales_parquet(_sigs(_sales_paths), _sigs(_stock_paths))
    if df.empty:
        st.error("매출 데이터가 비어 있습니다.")
        st.stop()

    # 재고 원가(parquet 에 없음) → 판매 실적의 실제 출고원가로 추정 (캐시: parquet 이 바뀔 때만 재계산)
    if not stock_df.empty:
        stock_df = attach_stock_cost_cached(
            stock_df, df, STOCK_BASE_DATE,
            _sigs(_sales_paths) + _sigs(_stock_paths),
        )

    if len(_sales_paths) > 1:
        st.success(f"📎 매출 {len(_sales_paths)}개 파일 병합 — "
                   + ", ".join(p.name for p in _sales_paths)
                   + f" · 총 {len(df):,}행", icon="✅")
except Exception as e:
    st.error(f"데이터를 읽는 중 오류가 발생했습니다: {e}")
    st.stop()


def make_product_display(prod_df: pd.DataFrame, extra_cols: list, img_width: str = "small"):
    """집계된 상품 표(모델명 포함)에 라인명→이미지 매칭하여 (표시용 df, column_config) 반환.
    img_map 이 비어있으면 이미지 컬럼 없이 그대로 표시. img_width: small/medium/large."""
    out = prod_df.copy().reset_index(drop=True)
    if "라인명" not in out.columns:
        out["라인명"] = out["모델명"].apply(to_line) if "모델명" in out.columns else ""
    show_img = bool(img_map)
    if show_img:
        out["이미지"] = out["라인명"].map(img_map).fillna("")
    out.insert(0, "Rank", np.arange(1, len(out) + 1))
    cols = ["Rank"] + (["이미지"] if show_img else []) + extra_cols
    disp, fcfg = format_table(out[cols])
    colcfg = {"Rank": st.column_config.TextColumn("#"), **fcfg}
    if show_img:
        colcfg["이미지"] = st.column_config.ImageColumn("이미지", width=img_width)
    return disp, colcfg


def _img_html(line_name: str, px: int, radius: int = 10, font: int = 11) -> str:
    """카드 왼쪽 썸네일. img_map 이 비어 있으면 빈 문자열(이미지 영역 자체가 사라짐)."""
    if not img_map:
        return ""
    url = img_map.get(str(line_name).strip(), "")
    if url:
        return (f'<img src="{html.escape(url, quote=True)}" '
                f'style="width:{px}px;height:{px}px;object-fit:cover;border-radius:{radius}px;'
                f'flex:0 0 auto;background:#f1f5f9;border:1px solid #eef2f7;">')
    return (f'<div style="width:{px}px;height:{px}px;border-radius:{radius}px;background:#f1f5f9;'
            f'flex:0 0 auto;display:flex;align-items:center;justify-content:center;'
            f'color:#cbd5e1;font-size:{font}px;">no img</div>')


def product_cards_html(prod_df: pd.DataFrame, n: int = 10, img_px: int = 80, start: int = 1, step: int = 1) -> str:
    """상품을 카드형 HTML로 렌더 (순위는 인라인 #N).
    start/step: 순위 = start + i*step (좌우 교차 배치 시 step=2 등)."""
    rows = prod_df.head(n).reset_index(drop=True)
    has_cat = "대분류" in rows.columns
    cards = []
    for i, r in rows.iterrows():
        rank = start + i * step
        brand = html.escape(str(r.get("브랜드", "")))
        cat = html.escape(str(r.get("대분류", ""))) if has_cat else ""
        season = ""
        if "입고연도" in rows.columns:
            _sv = str(r.get("입고연도", "")).strip()
            if _sv and _sv not in ("nan", "미상", "None"):
                season = html.escape(f"{_sv}입고")
        _parts = [f"#{rank}", brand] + ([cat] if cat else []) + ([season] if season else [])
        meta = " · ".join(_parts)
        line_name = str(r.get("라인명", "")).strip() or to_line(str(r.get("모델명", "")))
        model = html.escape(line_name if len(line_name) <= 38 else line_name[:37] + "…")
        rate = r.get("수익률", float("nan"))
        rate_s = f"{rate:.1f}%" if pd.notna(rate) and np.isfinite(rate) else "-"
        qty = r.get("수량", float("nan"))
        qty_s = f" · {int(qty):,}개" if pd.notna(qty) else ""
        sales_s = eok(r.get("최종판매가", 0))
        profit_s = eok(r.get("수익원(실배송비)", 0))
        cards.append(
            f'<div style="display:flex;gap:10px;align-items:center;padding:8px 8px;'
            f'border-bottom:1px solid #eef2f7;">{_img_html(line_name, img_px, 8, 10)}'
            f'<div style="min-width:0;flex:1;">'
            f'<div style="font-size:11px;color:#94a3b8;">{meta}</div>'
            f'<div style="font-size:13px;font-weight:600;color:#0f172a;white-space:nowrap;'
            f'overflow:hidden;text-overflow:ellipsis;">{model}</div>'
            f'<div style="font-size:13px;color:#0f172a;">{sales_s} '
            f'<span style="color:#64748b;">· 수익 {profit_s} · {rate_s}{qty_s}</span></div>'
            f'</div></div>'
        )
    return (
        '<div style="border:1px solid #e8eef5;border-radius:12px;overflow:hidden;'
        'box-shadow:0 2px 8px rgba(15,23,42,0.03);">' + "".join(cards) + "</div>"
    )


def stock_cards_html(stock_rows: pd.DataFrame, n: int = 10, img_px: int = 130, start: int = 1, step: int = 1) -> str:
    """재고 상품 카드형 HTML (판매 카드와 동일 레이아웃). 매출/수익률 대신 재고수량·총원가 표시."""
    rows = stock_rows.head(n).reset_index(drop=True)
    has_cat = "대분류" in rows.columns
    cards = []
    for i, r in rows.iterrows():
        rank = start + i * step
        brand = html.escape(str(r.get("브랜드", "")))
        cat = html.escape(str(r.get("대분류", ""))) if has_cat else ""
        season = ""
        if "입고연도" in rows.columns:
            _sv = str(r.get("입고연도", "")).strip()
            if _sv and _sv not in ("nan", "미상", "None"):
                season = html.escape(f"{_sv}입고")
        _parts = [f"#{rank}", brand] + ([cat] if cat else []) + ([season] if season else [])
        meta = " · ".join(_parts)
        line_name = str(r.get("라인명", "")).strip() or to_line(str(r.get("모델명", "")))
        model = html.escape(line_name if len(line_name) <= 38 else line_name[:37] + "…")
        try:
            qty_i = int(float(r.get("재고수량", 0)))
        except Exception:
            qty_i = 0
        cost_s = eok(r.get("총원가", 0))
        el_s = ""
        _el = r.get("입고경과일")
        try:
            if _el is not None and not pd.isna(_el):
                el_s = f' · <span style="color:#b45309;">입고경과 {int(_el)}일</span>'
        except Exception:
            el_s = ""
        cards.append(
            f'<div style="display:flex;gap:12px;align-items:center;padding:10px 10px;'
            f'border-bottom:1px solid #eef2f7;">{_img_html(line_name, img_px)}'
            f'<div style="min-width:0;flex:1;">'
            f'<div style="font-size:11px;color:#94a3b8;">{meta}</div>'
            f'<div style="font-size:14px;font-weight:600;color:#0f172a;white-space:nowrap;'
            f'overflow:hidden;text-overflow:ellipsis;">{model}</div>'
            f'<div style="font-size:13px;color:#0f172a;">재고 {qty_i:,}개 '
            f'<span style="color:#64748b;">· 원가 {cost_s}</span>{el_s}</div>'
            f'</div></div>'
        )
    return (
        '<div style="border:1px solid #e8eef5;border-radius:12px;overflow:hidden;'
        'box-shadow:0 2px 8px rgba(15,23,42,0.03);">' + "".join(cards) + "</div>"
    )


def metric_cards_html(df: pd.DataFrame, value_fn, n: int = 10, img_px: int = 130,
                      start: int = 1, step: int = 1) -> str:
    """범용 카드(재고 카드와 동일 레이아웃). value_fn(row)이 값줄 HTML을 반환."""
    rows = df.head(n).reset_index(drop=True)
    has_cat = "대분류" in rows.columns
    cards = []
    for i, r in rows.iterrows():
        rank = start + i * step
        brand = html.escape(str(r.get("브랜드", "")))
        cat = html.escape(str(r.get("대분류", ""))) if has_cat else ""
        season = ""
        if "입고연도" in rows.columns:
            _sv = str(r.get("입고연도", "")).strip()
            if _sv and _sv not in ("nan", "미상", "None"):
                season = html.escape(f"{_sv}입고")
        _parts = [f"#{rank}"] + ([brand] if brand else []) + ([cat] if cat else []) + ([season] if season else [])
        meta = " · ".join(_parts)
        line_name = str(r.get("라인명", "")).strip() or to_line(str(r.get("모델명", "")))
        model = html.escape(line_name if len(line_name) <= 38 else line_name[:37] + "…")
        cards.append(
            f'<div style="display:flex;gap:12px;align-items:center;padding:10px 10px;'
            f'border-bottom:1px solid #eef2f7;">{_img_html(line_name, img_px)}'
            f'<div style="min-width:0;flex:1;">'
            f'<div style="font-size:11px;color:#94a3b8;">{meta}</div>'
            f'<div style="font-size:14px;font-weight:600;color:#0f172a;white-space:nowrap;'
            f'overflow:hidden;text-overflow:ellipsis;">{model}</div>'
            f'{value_fn(r)}</div></div>'
        )
    return ('<div style="border:1px solid #e8eef5;border-radius:12px;overflow:hidden;'
            'box-shadow:0 2px 8px rgba(15,23,42,0.03);">' + "".join(cards) + "</div>")


VIBRANT_COLORS = [
    "#2563eb", "#f59e0b", "#10b981", "#ec4899", "#8b5cf6",
    "#06b6d4", "#ef4444", "#eab308", "#6366f1", "#14b8a6",
]


def share_donut(agg_df: pd.DataFrame, name_col: str, value_col: str, title: str, cmap: dict | None = None):
    """비중 도넛(생기있는 색). 슬라이스 라벨 = 이름 + 비중%만(크게). 값이 양수인 항목만."""
    d = agg_df.copy()
    d = d[pd.to_numeric(d[value_col], errors="coerce").fillna(0) > 0]
    names = d[name_col].astype(str).tolist()
    vals = list(d[value_col])
    if cmap:
        colors = [cmap.get(n, "#9aa5b1") for n in names]
    else:
        colors = [VIBRANT_COLORS[i % len(VIBRANT_COLORS)] for i in range(len(names))]
    fig = go.Figure(
        go.Pie(
            labels=names, values=vals, hole=0.5,
            textinfo="label+percent", textposition="inside",
            insidetextorientation="horizontal", textfont=dict(size=15),
            marker=dict(colors=colors, line=dict(color="#ffffff", width=2)),
            sort=False, direction="clockwise",
        )
    )
    fig.update_layout(title=title, showlegend=False, margin=dict(t=50, b=10, l=10, r=10),
                      font=dict(size=14))
    return fig


# =============================================================
# 브랜드 매출 대시보드 — 본문
#   브랜드 1개 선택 → 연도·분기별 추이 · 쇼핑몰별 · 카테고리/상품 비중
#   (공식/병행 구분 없음, 목표 없음)
# =============================================================

# ----- 입고연도 부여: 재고 '입고일자' 연도 기준 -----
#   (시즌(SS/FW) 구분은 제거됨 — 재고 관련 집계는 모두 '입고연도' 로 묶는다)
def _in_year_one(d) -> str:
    """단일 날짜 → 입고연도 라벨('2024'). NaT → '미상'."""
    d = pd.Timestamp(d)
    return "미상" if pd.isna(d) else f"{int(d.year)}"


def _in_year(dates) -> pd.Series:
    """datetime → 입고연도 라벨 Series. NaT → '미상'."""
    d = pd.to_datetime(dates, errors="coerce")
    lab = pd.Series("미상", index=d.index, dtype=object)
    ok = d.notna()
    if ok.any():
        lab.loc[d.index[ok.values]] = d[ok].dt.year.astype(int).astype(str).values
    return lab


def _yr_int(lab) -> int:
    """'2024' → 2024, '미상' → -1."""
    try:
        return int(str(lab))
    except Exception:
        return -1

# 입고일자 = '재고 파일의 입고일자'를 모델명→라인명 순으로 매칭.
#   한 상품이 여러 번 입고됐으면 '최초 입고일'(min)을 기준으로 본다.
#   (최근 입고 기준으로 바꾸려면 아래 .min() 두 곳을 .max() 로)
def _instock_maps(sdf_stock):
    if not isinstance(sdf_stock, pd.DataFrame) or sdf_stock.empty or "입고일자" not in sdf_stock.columns:
        return {}, {}
    s = sdf_stock.copy()
    s["입고일자"] = _parse_date_flexible(s["입고일자"])
    s = s.dropna(subset=["입고일자"])
    # 오파싱(1970/2069 등) 방어: 합리적 연도만 사용
    s = s[s["입고일자"].dt.year.between(2010, 2035)]
    if s.empty:
        return {}, {}
    by_model = s.groupby(s["모델명"].astype(str).str.strip())["입고일자"].min()
    by_line = s.groupby(s["라인명"].astype(str).str.strip())["입고일자"].min()
    return by_model.to_dict(), by_line.to_dict()


def _events_by_model(sdf_stock):
    """재고 입고이벤트를 모델명별 [(입고일, 수량)] 리스트(입고일 오름차순)로. FIFO 큐용."""
    if (not isinstance(sdf_stock, pd.DataFrame) or sdf_stock.empty
            or "입고이벤트" not in sdf_stock.columns or "모델명" not in sdf_stock.columns):
        return {}
    today = STOCK_BASE_DATE          # 재고 스냅샷 기준일('N일전'의 기준)
    d = {}
    for m, evs in zip(sdf_stock["모델명"].astype(str).str.strip(), sdf_stock["입고이벤트"]):
        if not isinstance(evs, list) or not evs:
            continue
        bucket = d.setdefault(m, [])
        for n_days, qty in evs:
            try:
                bucket.append((today - pd.Timedelta(days=int(n_days)), int(qty)))
            except Exception:
                continue
    for m in d:
        d[m].sort(key=lambda x: x[0])  # 입고일 오름차순(오래된 것 먼저)
    return d


_stock_for_brand = stock_df if "stock_df" in dir() else pd.DataFrame()
MIN_INBOUND_YEAR = 2024  # 매출이 24년부터 시작 → 23년 이전 입고분은 재고 집계에서 제외


# ----- 사이드바: 브랜드 선택 (선택 전에는 아무것도 그리지 않는다) -----
with st.sidebar:
    st.header("브랜드")
    _brand_tot = df.groupby("브랜드")["최종판매가"].sum().sort_values(ascending=False)
    brand_list = [b for b in _brand_tot.index.tolist()
                  if str(b).strip() not in ("미분류", "nan", "")]
    if not brand_list:
        st.error("브랜드 데이터가 없습니다.")
        st.stop()
    sel_brand = st.selectbox(
        "브랜드 검색", brand_list, index=None,
        placeholder="브랜드명을 입력해 검색하세요",
        format_func=lambda b: f"{b}  ·  {eok(_brand_tot.get(b, 0))}",
    )
    st.caption(f"전체 {len(brand_list)}개 브랜드")

    _line_q = st.text_input("라인명 검색 (출고 raw)", "",
                            placeholder="라인명·모델명 일부 입력",
                            key="line_raw_q")
    st.caption("입력하면 대시보드 맨 위에 해당 상품의 출고 raw 가 표시됩니다.")

# 브랜드 미선택 = 빈 화면(안내만). 무거운 계산·차트는 전부 건너뛴다.
if not sel_brand:
    st.info("왼쪽 사이드바에서 **브랜드를 검색**하면 그 브랜드의 연도·분기별 추이 · 쇼핑몰별 성과 · "
            "카테고리/상품 비중 · 재고가 한 번에 표시됩니다.", icon="🔎")
    st.caption(f"매출 {len(df):,}행 · 브랜드 {len(brand_list)}개 로드됨"
               + (f" · 재고 {len(_stock_for_brand):,}행" if len(_stock_for_brand) else ""))
    st.stop()

# ----- 선택한 브랜드만 추림 (전체가 아니라 브랜드 단위 → 전환이 즉시 반응) -----
_bstock = (_stock_for_brand[_stock_for_brand["브랜드"].astype(str).str.strip() == str(sel_brand).strip()]
           if ("브랜드" in _stock_for_brand.columns and not _stock_for_brand.empty)
           else _stock_for_brand)
df = df[df["브랜드"] == sel_brand].reset_index(drop=True)
_m_map, _l_map = _instock_maps(_bstock)
_ev_by_model = _events_by_model(_bstock)

# ----- 사이드바: 필터 -----
with st.sidebar:
    st.divider()
    st.header("필터")
    bdf0 = df.copy()

    _years_all = sorted(int(y) for y in bdf0["날짜"].dt.year.dropna().unique())
    sel_years = st.multiselect("연도 (판매·연도별 분석용)", _years_all,
                               default=_years_all[-3:] if _years_all else [])

    def _msa(label, options):
        opts = sorted([x for x in options if pd.notna(x)])
        return st.multiselect(label, opts, default=opts)

    sel_malls = _msa("쇼핑몰", bdf0["쇼핑몰"].unique())
    sel_cats = _msa("대분류", bdf0["대분류"].unique())
    sel_notes = _msa("비고", bdf0["비고"].unique())
    include_returns = st.checkbox("반품/음수 포함", value=True)

# g: 브랜드 + (쇼핑몰/대분류/비고/반품) 필터 — 전체 판매기간
g = bdf0[
    bdf0["쇼핑몰"].isin(sel_malls)
    & bdf0["대분류"].isin(sel_cats)
    & bdf0["비고"].isin(sel_notes)
].copy()
if not include_returns:
    g = g[(g["수량"] >= 0) & (g["최종판매가"] >= 0)].copy()
if g.empty:
    st.warning("필터 조건에 해당하는 데이터가 없습니다.")
    st.stop()

# f: 판매 상세 (라인 단위 합산)
f = g.copy()

# 판매(출고)연도 — 연도별 집계·구성 차트용
g["판매연도"] = g["날짜"].dt.year.astype("Int64")
f["판매연도"] = f["날짜"].dt.year.astype("Int64")
recent_years = sorted(int(y) for y in g["판매연도"].dropna().unique())[-6:]

# 대분류 색맵 (도넛·연도별구성 stacked 공유) — 이 브랜드 매출 비중 큰 순
_cat_order_all = g.groupby("대분류")["최종판매가"].sum().sort_values(ascending=False).index.tolist()
_CAT_CMAP = {c: VIBRANT_COLORS[i % len(VIBRANT_COLORS)] for i, c in enumerate(_cat_order_all)}

metric_cols = {
    "최종판매가": "최종판매가",
    "수량": "수량",
    "수익원(실배송비)": "수익원(실배송비)",
}

# ----- 브랜드 헤더 -----
st.markdown(
    f"<div style='font-size:1.6rem;font-weight:800;color:#0f172a;margin:.2rem 0 .2rem;'>🏷️ {html.escape(str(sel_brand))}</div>",
    unsafe_allow_html=True,
)

# ----- KPI -----
tot_sales = float(f["최종판매가"].sum())
tot_qty = float(f["수량"].sum())
tot_profit = float(f["수익원(실배송비)"].sum())
avg_price = tot_sales / tot_qty if tot_qty else 0
profit_rate = tot_profit / tot_sales * 100 if tot_sales else 0

k = st.columns(5)
k[0].metric("총매출", eok(tot_sales))
k[1].metric("수량", num(tot_qty))
k[2].metric("객단가", eok(avg_price))
k[3].metric("수익", eok(tot_profit))   # 수익률은 옆 카드에 별도 표시(중복 제거)
k[4].metric("수익률", pct(profit_rate))

st.markdown(
    f"<div class='hint'>{len(f):,}행 · "
    f"기간 {f['날짜'].min().date()} ~ {f['날짜'].max().date()} · 사은품/쇼핑백 제외</div>",
    unsafe_allow_html=True,
)

# =============================================================
# 0) 라인명 검색 → 출고 raw  (사이드바 '라인명 검색' 에 입력했을 때만)
# =============================================================
_kw = str(_line_q or "").strip()
if _kw:
    st.markdown(f"<div class='section-title'>🔎 출고 raw — '{html.escape(_kw)}'</div>",
                unsafe_allow_html=True)
    _hit = g[g["라인명"].astype(str).str.contains(_kw, case=False, regex=False, na=False)
             | g["모델명"].astype(str).str.contains(_kw, case=False, regex=False, na=False)]
    if _hit.empty:
        st.warning(f"'{_kw}' 와(과) 일치하는 라인명·모델명이 **{sel_brand}** 에 없습니다. "
                   "(사이드바 필터에 걸려 빠졌을 수도 있습니다)", icon="🔍")
    else:
        _lines = (_hit.groupby("라인명")["최종판매가"].sum()
                  .sort_values(ascending=False).index.tolist())
        # ---- 매칭 라인 썸네일 (매출순 · 최대 12개) ----
        _thumbs = []
        for _ln in _lines[:12]:
            _thumbs.append(
                f'<div style="width:104px;flex:0 0 auto;text-align:center;">'
                f'{_img_html(_ln, 104, 10, 11)}'
                f'<div style="font-size:11px;color:#475569;margin-top:4px;line-height:1.25;'
                f'word-break:break-all;">{html.escape(str(_ln))}</div></div>')
        if _thumbs:
            st.markdown('<div style="display:flex;gap:10px;flex-wrap:wrap;margin:.2rem 0 .6rem;">'
                        + "".join(_thumbs) + "</div>", unsafe_allow_html=True)
            if len(_lines) > 12:
                st.caption(f"…외 {len(_lines) - 12}개 라인 (썸네일은 매출순 12개까지)")

        # ---- 원본 parquet 에서 그대로 읽어온 출고 raw ----
        _models = tuple(sorted(set(_hit["모델명"].astype(str))))
        _raw = load_raw_by_models(_sigs(_sales_paths), _models)
        if not _raw.empty and "브랜드" in _raw.columns:
            _raw = _raw[_raw["브랜드"].astype(str).str.strip() == str(sel_brand).strip()]
        if _raw.empty:
            st.caption("원본 parquet 에서 해당 모델의 행을 찾지 못했습니다.")
        else:
            _n_show = 5000
            st.markdown(f"**출고 raw** · 라인 {len(_lines)}개 · 모델(사이즈) {len(_models)}개 · "
                        f"**{len(_raw):,}행**"
                        + (f" (최근 {_n_show:,}행만 표시)" if len(_raw) > _n_show else ""))
            _numfmt = {c: st.column_config.NumberColumn(c, format="localized")
                       for c in ("수량", "출고원가", "최종판매가", "수익원(실배송비)")
                       if c in _raw.columns}
            st.dataframe(_raw.head(_n_show), hide_index=True, use_container_width=True,
                         height=min(80 + len(_raw) * 35, 560), column_config=_numfmt)
            st.caption("매출 parquet 원본 그대로 (전처리·반품일 보정 없음) · 출고날짜 내림차순 · "
                       "사이드바 필터(쇼핑몰·대분류·비고)로 걸러진 모델만 대상입니다.")
    st.divider()

# =============================================================
# 1) 연도별 매출 · 수익 (판매 = 출고일 기준)
# =============================================================
st.markdown(f"<div class='section-title'>연도별 매출 · 수익 — {sel_brand}</div>", unsafe_allow_html=True)
st.caption("판매(출고)일 기준 · 반품 차감.")

with st.expander("🔧 재고 입고 진단 — 이상값 확인용"):
    st.write(f"입고이력이 있는 모델: **{len(_ev_by_model):,}개**")
    if _m_map:
        _keys = list(_m_map.keys())[:15]
        _smp = pd.DataFrame({"모델명": _keys,
                             "입고일자(파싱결과)": [pd.Timestamp(_m_map[k]).date() for k in _keys]})
        st.dataframe(_smp, hide_index=True, use_container_width=True)
    else:
        st.caption("재고에서 입고일자를 못 읽었습니다 (재고 미업로드 / 입고이력 컬럼 없음 / 형식 문제).")
    st.markdown("---")
    st.write("**재고 입고이력 파싱 진단 (이 브랜드)**")
    _sdbg = _bstock
    if isinstance(_sdbg, pd.DataFrame) and not _sdbg.empty:
        if "입고이벤트" in _sdbg.columns:
            st.write("입고이벤트(경과일/수량) 샘플:", [e for e in _sdbg["입고이벤트"].head(5) if e][:5])
        if "원가평균" in _sdbg.columns:
            _cs = _sdbg["원가평균"].dropna()
            st.write(f"추정 개당원가 매칭: {len(_cs):,} / {len(_sdbg):,}행 "
                     f"({(len(_cs) / len(_sdbg) * 100 if len(_sdbg) else 0):.0f}%)")
        st.caption(f"재고 스냅샷 기준일: {pd.Timestamp(STOCK_BASE_DATE).date()}")
        st.caption(f"재고 전체 컬럼({_stock_for_brand.shape[1]}개): "
                   f"{[str(c) for c in _stock_for_brand.columns.tolist()]}")
    else:
        st.caption("이 브랜드 재고가 없습니다.")

_ya = (g.dropna(subset=["판매연도"]).groupby("판매연도")
         .agg(매출=("최종판매가", "sum"), 수량=("수량", "sum"),
              수익=("수익원(실배송비)", "sum")).reset_index().sort_values("판매연도"))
_ya["연도"] = _ya["판매연도"].astype(int).astype(str)
_ya["라벨"] = _ya["매출"].apply(eok)

yc1, yc2 = st.columns([1.15, 1.5])
with yc1:
    fig_y = px.bar(_ya, x="연도", y="매출", text="라벨", color="연도",
                   title=f"{sel_brand} 연도별 매출",
                   labels={"매출": "매출", "연도": "연도"},
                   color_discrete_sequence=VIBRANT_COLORS)
    fig_y.update_traces(textposition="outside", textangle=0, cliponaxis=False)
    fig_y.update_layout(xaxis_type="category", showlegend=False, margin=dict(t=54, b=10),
                        uniformtext_minsize=9, uniformtext_mode="hide")
    if len(_ya):
        _ymax = float(_ya["매출"].max())
        _ymin = float(_ya["매출"].min())
        fig_y.update_yaxes(range=[min(0, _ymin) * 1.1, _ymax * 1.18 if _ymax > 0 else _ymax * 0.8])
    st.plotly_chart(fig_y, use_container_width=True)

with yc2:
    yt = pd.DataFrame({"연도": _ya["연도"].values})
    yt["합계매출"] = _ya["매출"].round(0).astype("int64").values
    yt["수량"] = _ya["수량"].round(0).astype("int64").values
    yt["수익"] = _ya["수익"].round(0).astype("int64").values
    _mv = _ya["매출"].values.astype(float)
    yt["수익률"] = np.where(_mv != 0, _ya["수익"].values.astype(float) / np.where(_mv == 0, np.nan, _mv) * 100, np.nan)
    yt["전년比"] = (_ya["매출"].pct_change() * 100).apply(growth_pct).values
    st.markdown("**연도별 매출 · 수익**")
    _ycfg = {c: st.column_config.NumberColumn(c, format="localized") for c in ("합계매출", "수량", "수익")}
    _ycfg["수익률"] = st.column_config.NumberColumn("수익률", format="%.1f%%")
    st.dataframe(yt, hide_index=True, use_container_width=True, height=60 + len(yt) * 36,
                 column_config=_ycfg)

# ----- 입고연도별 재고 소진 : 입고원가(입고수량 × 추정 개당원가) vs 현재 재고원가(추정) -----
_bsea = _bstock
_in_cost, _cur_cost = {}, {}
if (not _bsea.empty) and ("입고이벤트" in _bsea.columns) and ("입고원가이벤트" in _bsea.columns):
    _t0b = STOCK_BASE_DATE
    for _qev, _cev in zip(_bsea["입고이벤트"], _bsea["입고원가이벤트"]):
        if not isinstance(_qev, list):
            continue
        # 입고연도별 개당원가 (입고건 날짜 → 연도)
        _c_by_y = {}
        for _d, _c in (_cev if isinstance(_cev, list) else []):
            _c_by_y[_in_year_one(_d)] = float(_c)
        # 입고연도별 수량 × 그 연도 개당원가
        for _n, _qq in _qev:
            _y = _in_year_one(_t0b - pd.Timedelta(days=int(_n)))
            if _yr_int(_y) < MIN_INBOUND_YEAR:
                continue
            _in_cost[_y] = _in_cost.get(_y, 0.0) + float(_qq) * _c_by_y.get(_y, 0.0)
if (not _bsea.empty) and ("총원가" in _bsea.columns) and ("입고일자" in _bsea.columns):
    _amt2 = pd.to_numeric(_bsea["총원가"], errors="coerce").fillna(0.0)
    _y2 = _in_year(_bsea["입고일자"])
    for _a, _ylab in zip(_amt2, _y2):
        if _yr_int(_ylab) >= MIN_INBOUND_YEAR:
            _cur_cost[_ylab] = _cur_cost.get(_ylab, 0.0) + float(_a)
_years_sorted = sorted({y for y in (set(_in_cost) | set(_cur_cost)) if _yr_int(y) >= MIN_INBOUND_YEAR},
                       key=_yr_int)
if _years_sorted:
    st2 = pd.DataFrame({"입고연도": _years_sorted})
    st2["총입고원가"] = [int(round(_in_cost.get(y, 0.0))) for y in _years_sorted]
    st2["현재총원가"] = [int(round(_cur_cost.get(y, 0.0))) for y in _years_sorted]
    st2["소진율%"] = [min(max((_in_cost.get(y, 0) - _cur_cost.get(y, 0)) / _in_cost.get(y, 0) * 100, 0.0), 100.0)
                   if _in_cost.get(y, 0) else np.nan for y in _years_sorted]
    st.markdown("**입고연도별 재고 소진** (입고원가 vs 현재 재고원가)")
    _s2cfg = {c: st.column_config.NumberColumn(c, format="localized") for c in ("총입고원가", "현재총원가")}
    _s2cfg["소진율%"] = st.column_config.NumberColumn("소진율%", format="%.1f%%")
    st.dataframe(st2, hide_index=True, use_container_width=True, height=60 + len(st2) * 36,
                 column_config=_s2cfg)
    _tot_in = float(sum(_in_cost.values()))
    _cost_col = next((c for c in ("원가총액", "원가") if c in g.columns), None)
    _recov = float(g["수익원(실배송비)"].sum()) + (float(pd.to_numeric(g[_cost_col], errors="coerce").fillna(0).sum())
                                                if _cost_col else 0.0)
    _recov_s = f"{_recov / _tot_in * 100:.1f}%" if _tot_in else "-"
    st.caption("총입고원가 = 입고수량 × 추정 개당원가 · 현재총원가 = 남은 재고수량 × 추정 개당원가 · "
               "소진율 = (총입고원가−현재총원가)÷총입고원가. "
               f"이 브랜드 원가회수율 = (판매수익+판매원가)÷총입고원가 = **{_recov_s}**. "
               "⚠ 재고 parquet 에 원가가 없어 개당원가는 판매 실적의 실제 출고원가로 추정한 값입니다.")
else:
    st.caption("※ 입고연도별 재고 소진표는 재고 parquet(입고이력·수량)이 있어야 표시됩니다.")


# =============================================================
# 1-2) 분기별 판매 추이 (판매=출고일 기준)
# =============================================================
st.markdown(f"<div class='section-title'>분기별 판매 추이 — {sel_brand}</div>", unsafe_allow_html=True)
st.caption("판매(출고)일 기준 · 분기별 매출. 필터(쇼핑몰·대분류·비고)만 적용되며 전체 판매기간을 봅니다.")

q = g.copy()
_qp = q["날짜"].dt.to_period("Q")
q["분기"] = _qp.astype(str)                       # 2024Q1
q["분기정렬"] = _qp.astype("int64")
q["분기표시"] = q["분기"].str.replace("Q", " Q", regex=False)
q["분기연도"] = q["날짜"].dt.year.astype(int)
q["분기No"] = q["날짜"].dt.quarter.astype(int)
qa = (q.groupby(["분기표시"])
        .agg(매출=("최종판매가", "sum"), 수량=("수량", "sum"),
             정렬=("분기정렬", "first"), 연도=("분기연도", "first"))
        .reset_index().sort_values("정렬"))
qa["라벨"] = qa["매출"].apply(eok)
qa["연도"] = qa["연도"].astype(int).astype(str)

qc1, qc2 = st.columns([1.7, 1])
with qc1:
    fig_q = px.bar(qa, x="분기표시", y="매출", text="라벨", color="연도",
                   title=f"{sel_brand} 분기별 매출(판매일 기준 · 연도별 색)",
                   labels={"매출": "매출", "분기표시": "분기"},
                   color_discrete_sequence=VIBRANT_COLORS)
    fig_q.update_traces(textposition="outside", textangle=0, cliponaxis=False)
    fig_q.update_layout(xaxis_type="category", margin=dict(t=54, b=10),
                        uniformtext_minsize=9, uniformtext_mode="hide", legend_title_text="연도")
    fig_q.update_xaxes(categoryorder="array", categoryarray=qa["분기표시"].tolist(),
                       tickangle=-30, title_text="")
    fig_q.update_yaxes(title_text="")
    if len(qa):
        _qm = float(qa["매출"].max()); _qn = float(qa["매출"].min())
        fig_q.update_yaxes(range=[min(0, _qn) * 1.1, _qm * 1.18 if _qm > 0 else _qm * 0.8])
    st.plotly_chart(fig_q, use_container_width=True)

with qc2:
    # 분기 × 연도 (전년 동분기 비교)
    qpv = q.pivot_table(index="분기No", columns="분기연도", values="최종판매가",
                        aggfunc="sum", fill_value=0).sort_index()
    if not qpv.empty:
        qd = pd.DataFrame({"분기": [f"Q{int(i)}" for i in qpv.index]})
        for _yy in qpv.columns:
            qd[str(int(_yy))] = qpv[_yy].round(0).astype("int64").values
        st.markdown("**분기 × 연도 매출**")
        _qcfg = {str(int(c)): st.column_config.NumberColumn(str(int(c)), format="localized")
                 for c in qpv.columns}
        st.dataframe(qd, hide_index=True, use_container_width=True, height=60 + len(qd) * 36,
                     column_config=_qcfg)
        _yrs = sorted(int(c) for c in qpv.columns)
        if len(_yrs) >= 2:
            _a, _b = _yrs[-2], _yrs[-1]
            _ga = float(qpv[_a].sum()); _gb = float(qpv[_b].sum())
            _gr = (_gb - _ga) / abs(_ga) * 100 if _ga else np.nan
            st.caption(f"{_b}년 {eok(_gb)} · 전년 {eok(_ga)} → **{growth_pct(_gr)}**")

# =============================================================
# 2) 쇼핑몰별 — 이 브랜드가 어디서 잘 나가나
# =============================================================
st.markdown(f"<div class='section-title'>쇼핑몰별 — {sel_brand}</div>", unsafe_allow_html=True)

mc1, mc2 = st.columns([1, 1.3])
with mc1:
    mb = aggregate(f, ["쇼핑몰"], metric_cols).head(10).copy()
    if not mb.empty:
        mb["라벨"] = mb["매출비중"].apply(lambda x: f"{x:.1f}%")
        fig_m = px.bar(mb, x="쇼핑몰", y="최종판매가", text="라벨",
                       title=f"{sel_brand} 쇼핑몰 TOP 10 (매출 비중)",
                       labels={"최종판매가": "매출"})
        fig_m.update_traces(textposition="inside", insidetextanchor="middle", textangle=0,
                            textfont=dict(size=14, color="#ffffff"), marker_color="#2563eb",
                            cliponaxis=False)
        fig_m.update_layout(xaxis_type="category", margin=dict(t=54, b=10),
                            uniformtext_minsize=10, uniformtext_mode="hide")
        fig_m.update_xaxes(categoryorder="total descending", tickangle=-30, title_text="",
                           tickfont=dict(size=11))
        fig_m.update_yaxes(title_text="")
        st.plotly_chart(fig_m, use_container_width=True)
with mc2:
    mall_t = aggregate(f, ["쇼핑몰"], metric_cols).reset_index(drop=True)
    mall_t = mall_t[["쇼핑몰", "수량", "최종판매가", "객단가", "수익률", "매출비중"]].head(50)
    mall_t = rank_table(mall_t, "쇼핑몰")
    _ft, _fc = format_table(mall_t)
    st.dataframe(_ft, hide_index=True, use_container_width=True, height=430,
                 column_config={**_fc, "Rank": st.column_config.TextColumn("#")})

# 연도별 실적 + 쇼핑몰 × 연도 (판매일 기준 · 사이드바 '연도' 필터 적용)
_gy = g.copy()
_gy["연도"] = _gy["날짜"].dt.year
_yrs = [int(y) for y in (sel_years or []) if y in set(_gy["연도"].dropna().astype(int))]
if not _yrs:
    _yrs = sorted(int(y) for y in _gy["연도"].dropna().unique())[-3:]
_gy = _gy[_gy["연도"].isin(_yrs)]
if not _gy.empty:
    st.markdown("**연도별 쇼핑몰 실적** (연도별 3분할 · 판매일 기준)")
    _T = float(_gy["최종판매가"].sum())
    _P = float(_gy["수익원(실배송비)"].sum())
    _Q = float(_gy["수량"].sum())
    _PR = (_P / _T * 100) if _T else 0.0
    st.markdown(f"합계({' · '.join(str(y) for y in _yrs)}) · "
                f"**총매출 {eok(_T)}** · 수익 {eok(_P)} · 수익률 {_PR:.1f}% · 수량 {int(_Q):,}개")
    _gall = g.copy()
    _gall["연도"] = _gall["날짜"].dt.year
    for _box, _yv in zip(st.columns(len(_yrs)), _yrs):
        with _box:
            _cur = _gall[_gall["연도"] == _yv]
            _prev = _gall[_gall["연도"] == _yv - 1]
            _tot = float(_cur["최종판매가"].sum())
            _prof = float(_cur["수익원(실배송비)"].sum())
            _pr = (_prof / _tot * 100) if _tot else 0.0
            _ptot = float(_prev["최종판매가"].sum())
            _yoy = ((_tot - _ptot) / _ptot * 100) if _ptot else np.nan
            st.markdown(f"**{_yv}년** · 총매출 {eok(_tot)}")
            st.caption(f"수익률 {_pr:.1f}% · 전년比 {growth_pct(_yoy)}")
            _a = _cur.groupby("쇼핑몰").agg(매출=("최종판매가", "sum"), 수량=("수량", "sum"),
                                          수익=("수익원(실배송비)", "sum"))
            if _a.empty:
                st.caption("데이터 없음")
                continue
            _a["객단가"] = np.where(_a["수량"] != 0, _a["매출"] / _a["수량"], 0)
            _a["수익률"] = np.where(_a["매출"] != 0, _a["수익"] / _a["매출"] * 100, 0)
            _pm = _prev.groupby("쇼핑몰")["최종판매가"].sum()
            _pv = pd.Series(_a.index.map(_pm), index=_a.index).astype(float)
            _a["전년비"] = np.where(_pv.notna() & (_pv != 0), (_a["매출"] - _pv) / _pv * 100, np.nan)
            _a = _a.sort_values("매출", ascending=False).head(20).reset_index()
            _d = pd.DataFrame({
                "쇼핑몰": [f"{i}. {v}" for i, v in enumerate(_a["쇼핑몰"], 1)],
                "매출": _a["매출"].round(0).astype("int64").values,
                "객단가": _a["객단가"].round(0).astype("int64").values,
                "수익률": _a["수익률"].round(1).values,
                "전년비": pd.Series(_a["전년비"]).apply(growth_pct).values,
            })
            st.dataframe(_d, hide_index=True, use_container_width=True,
                         height=min(60 + len(_d) * 36, 620),
                         column_config={
                             "매출": st.column_config.NumberColumn("매출", format="localized"),
                             "객단가": st.column_config.NumberColumn("객단가", format="localized"),
                             "수익률": st.column_config.NumberColumn("수익률", format="%.1f%%")})

# =============================================================
# 3) 카테고리(대분류) 비중·매출
# =============================================================
st.markdown(f"<div class='section-title'>카테고리(대분류) · 매출 기준 — {sel_brand}</div>", unsafe_allow_html=True)
cat_t = aggregate(f, ["대분류"], metric_cols).reset_index(drop=True)
cc1, cc2 = st.columns([1.3, 1])
with cc1:
    st.plotly_chart(share_donut(cat_t, "대분류", "최종판매가", f"{sel_brand} 카테고리 비중 (매출)", cmap=_CAT_CMAP),
                    use_container_width=True)
with cc2:
    ct = cat_t.copy()
    ct.insert(0, "Rank", np.arange(1, len(ct) + 1))
    ct = ct[["Rank", "대분류", "수량", "최종판매가", "객단가", "수익률", "매출비중"]]
    _ftc, _fcc = format_table(ct)
    st.dataframe(_ftc, hide_index=True, use_container_width=True, height=60 + len(ct) * 36,
                 column_config={**_fcc, "Rank": st.column_config.TextColumn("#")})

# 8개 대분류로 매핑 안 된(미분류) 원본 카테고리 표시 — 매핑 추가 참고용
if "대분류_원본" in f.columns and (f["대분류"] == "미분류").any():
    _miss = (f.loc[f["대분류"] == "미분류", "대분류_원본"].astype(str)
             .value_counts().head(12).index.tolist())
    _miss = [m for m in _miss if m not in ("미분류", "nan", "")]
    if _miss:
        st.caption(f"⚠️ 8개 대분류로 매핑 안 된 항목 → '미분류'로 모음: {', '.join(_miss)}  ·  어느 대분류로 넣을지 알려주면 추가합니다.")

# 대분류 × 연도 구성 (최근 연도, 누적 · 판매일 기준)
_ry = [str(y) for y in recent_years]
if len(_ry) >= 2:
    gm = g[g["판매연도"].astype("Int64").astype(str).isin(_ry)].copy()
    gm["연도"] = gm["판매연도"].astype(int).astype(str)
    cs = gm.groupby(["연도", "대분류"], as_index=False)["최종판매가"].sum()
    cs["연도"] = pd.Categorical(cs["연도"], categories=_ry, ordered=True)
    cs = cs.sort_values("연도")
    fig_cs = px.bar(cs, x="연도", y="최종판매가", color="대분류", barmode="stack",
                    title=f"{sel_brand} 연도별 카테고리 구성",
                    category_orders={"연도": _ry, "대분류": _cat_order_all},
                    color_discrete_map=_CAT_CMAP,
                    labels={"최종판매가": "매출"})
    fig_cs.update_layout(xaxis_type="category", margin=dict(t=54, b=10), legend_title_text="대분류")
    st.plotly_chart(fig_cs, use_container_width=True)

# =============================================================
# 4) 브랜드 TOP 10 상품 (3분할: 매출액 / 수익 / 수익률)
# =============================================================
st.markdown(f"<div class='section-title'>{sel_brand} TOP 10 상품</div>", unsafe_allow_html=True)
prod = aggregate(f, ["브랜드", "대분류", "라인명"], metric_cols).reset_index(drop=True)
if prod.empty:
    st.caption("표시할 상품이 없습니다.")
else:
    _by_sales = prod.sort_values("최종판매가", ascending=False).head(10).reset_index(drop=True)
    _by_profit = prod.sort_values("수익원(실배송비)", ascending=False).head(10).reset_index(drop=True)
    _by_rate = (prod[(prod["최종판매가"] > 0) & (prod["수량"] >= 5)].sort_values("수익률", ascending=False)
                .head(10).reset_index(drop=True))
    _t1, _t2, _t3 = st.columns(3)
    with _t1:
        st.markdown("**매출액순**")
        st.markdown(product_cards_html(_by_sales, n=10, img_px=72), unsafe_allow_html=True)
    with _t2:
        st.markdown("**수익순**")
        st.markdown(product_cards_html(_by_profit, n=10, img_px=72), unsafe_allow_html=True)
    with _t3:
        st.markdown("**수익률순** (5개 이상 판매)")
        st.markdown(product_cards_html(_by_rate, n=10, img_px=72), unsafe_allow_html=True)

# =============================================================
# =============================================================
# 6) 자동 요약
# =============================================================
st.markdown("<div class='section-title'>자동 요약</div>", unsafe_allow_html=True)
_mt = aggregate(f, ["쇼핑몰"], metric_cols).head(1)
_ct = aggregate(f, ["대분류"], metric_cols).head(1)
_lt = aggregate(f, ["라인명"], metric_cols).head(1)
sm = []
sm.append(f"- **{sel_brand}** 합계 매출 **{eok(tot_sales)}** · 수량 **{num(tot_qty)}개** · "
          f"객단가 **{eok(avg_price)}** · 수익률 **{pct(profit_rate)}**.")
if len(_ya) >= 2:
    _y1, _y0 = _ya.iloc[-1], _ya.iloc[-2]
    _gr = (float(_y1["매출"]) - float(_y0["매출"])) / abs(float(_y0["매출"])) * 100 if float(_y0["매출"]) else np.nan
    sm.append(f"- **{_y1['연도']}년** 매출 **{eok(float(_y1['매출']))}**, "
              f"전년({_y0['연도']}년) **{eok(float(_y0['매출']))}** 대비 **{growth_pct(_gr)}**.")
elif len(_ya) == 1:
    sm.append(f"- **{_ya.iloc[0]['연도']}년** 매출 **{eok(float(_ya.iloc[0]['매출']))}** (전년 데이터 없음).")
if not _mt.empty:
    sm.append(f"- 쇼핑몰 1위 **{_mt.iloc[0]['쇼핑몰']}** · {eok(_mt.iloc[0]['최종판매가'])} "
              f"(비중 {pct(_mt.iloc[0].get('매출비중', np.nan))}).")
if not _ct.empty:
    sm.append(f"- 카테고리 1위 **{_ct.iloc[0]['대분류']}** · {eok(_ct.iloc[0]['최종판매가'])}.")
if not _lt.empty:
    sm.append(f"- 베스트 라인 **{_lt.iloc[0]['라인명']}** · {eok(_lt.iloc[0]['최종판매가'])}.")
st.markdown("\n".join(sm))

# =============================================================
# 7) 재고 (이 브랜드 · 재고 파일이 있을 때만)
# =============================================================
if "stock_df" in dir() and isinstance(stock_df, pd.DataFrame) and not stock_df.empty:
    bstock = stock_df[stock_df["브랜드"].astype(str).str.strip() == str(sel_brand).strip()].copy().reset_index(drop=True)
    st.markdown(f"<div class='section-title'>📦 재고 — {sel_brand}</div>", unsafe_allow_html=True)
    if bstock.empty:
        st.caption(f"재고 파일에서 '{sel_brand}' 브랜드를 찾지 못했습니다. (판매 데이터와 재고의 브랜드 표기가 다를 수 있음)")
    else:
        bstock["수량"] = pd.to_numeric(bstock["수량"], errors="coerce").fillna(0)
        bstock["총원가"] = pd.to_numeric(bstock["총원가"], errors="coerce").fillna(0)
        # 대분류: 재고 자체 분류(매출과 동일 _CATEGORY_MAP 적용됨) 우선, '미분류'만 판매 라인매핑으로 보강
        if "대분류" not in bstock.columns:
            bstock["대분류"] = "미분류"
        bstock["대분류"] = (bstock["대분류"].astype(str).str.strip()
                          .replace({"": "미분류", "nan": "미분류", "None": "미분류"}))
        _mask = bstock["대분류"] == "미분류"
        if _mask.any():
            # 보강이 필요한 라인만 대상으로 (라인명, 대분류) 빈도 1위를 구한다
            _need = set(bstock.loc[_mask, "라인명"].astype(str).str.strip())
            _sub = df[df["라인명"].astype(str).str.strip().isin(_need)]
            _line2cat = (_sub.groupby(["라인명", "대분류"], dropna=False).size()
                         .reset_index(name="_n").sort_values("_n", ascending=False)
                         .drop_duplicates("라인명").set_index("라인명")["대분류"])
            _bf = bstock.loc[_mask, "라인명"].astype(str).str.strip().map(_line2cat)
            _bf = _bf.where(_bf.isin(_ALLOWED_CATS))
            bstock.loc[_mask, "대분류"] = _bf.fillna("미분류").values
        # 입고연도: 재고 입고일자 기준
        bstock["입고연도"] = _in_year(bstock["입고일자"]) if "입고일자" in bstock.columns else "미상"

        b1, b2, b3 = st.columns(3)
        b1.metric("총 재고수량", f"{int(bstock['수량'].sum()):,}개")
        b2.metric("총 재고원가", eok(bstock["총원가"].sum()))
        b3.metric("라인 수", f"{bstock['라인명'].nunique():,}")
        # 8개로 매핑 안 된(미분류) 재고 카테고리 표시 — 키워드 추가 참고용
        _src_col = "대분류_원본" if "대분류_원본" in bstock.columns else "카테고리"
        if _src_col in bstock.columns and (bstock["대분류"] == "미분류").any():
            _ms = (bstock.loc[bstock["대분류"] == "미분류", _src_col].astype(str)
                   .value_counts().head(15).index.tolist())
            _ms = [m for m in _ms if m not in ("미분류", "nan", "")]
            if _ms:
                st.caption(f"⚠️ 8개 대분류로 매핑 안 된 재고 카테고리: {', '.join(_ms)}  ·  어느 대분류인지 알려주면 추가합니다.")
        with st.expander("🔧 재고 분류 진단 — 어느 컬럼을 분류로 읽었나"):
            if _STOCK_CAT_DEBUG:
                st.write(f"분류로 채택한 컬럼: **{_STOCK_CAT_DEBUG.get('selected')}** "
                         f"(8개 매핑률 {_STOCK_CAT_DEBUG.get('rate')})")
                st.write("후보 컬럼별 매핑률:", _STOCK_CAT_DEBUG.get("candidates"))
                st.caption(f"재고 전체 컬럼: {', '.join(_STOCK_CAT_DEBUG.get('all_cols', []))}")
            if "대분류_원본" in bstock.columns:
                _src = (bstock.groupby("대분류_원본")["대분류"].first()
                        .reset_index().rename(columns={"대분류_원본": "재고 원본값", "대분류": "→ 매핑"}))
                st.dataframe(_src, hide_index=True, use_container_width=True,
                             height=min(60 + len(_src) * 32, 360))

        # ---- 회전율 · 완판 분석 (라인 단위) ----
        st.markdown("**회전율 · 완판 분석** (라인 단위)")
        _tdy = STOCK_BASE_DATE
        # 라인별 입고량 + 첫입고일 — 재고 입고이벤트에서
        _inb_rows = []
        if "입고이벤트" in bstock.columns:
            for _ln, _evs in zip(bstock["라인명"].astype(str).str.strip(), bstock["입고이벤트"]):
                if not isinstance(_evs, list):
                    continue
                for _n, _qq in _evs:
                    _din = _tdy - pd.Timedelta(days=int(_n))
                    if _din.year < MIN_INBOUND_YEAR:
                        continue
                    _inb_rows.append((_ln, int(_qq), _din))
        if not _inb_rows:
            st.caption("입고이력('N일전/수량')이 없어 회전율을 계산할 수 없습니다.")
        else:
            _inb_ls = (pd.DataFrame(_inb_rows, columns=["라인명", "입고량", "입고일"])
                       .groupby("라인명").agg(입고량=("입고량", "sum"),
                                             첫입고일=("입고일", "min")).reset_index())
            _gk = g["라인명"].astype(str).str.strip()
            _sal_ls = (g.assign(_ln=_gk).groupby("_ln").agg(
                판매량=("수량", "sum"), 매출=("최종판매가", "sum"), 수익=("수익원(실배송비)", "sum")
            ).reset_index().rename(columns={"_ln": "라인명"}))
            _last_sale = g.assign(_ln=_gk).groupby("_ln")["날짜"].max()
            _ln_cat = bstock.groupby(bstock["라인명"].astype(str).str.strip())["대분류"].first()
            _rt = _inb_ls.merge(_sal_ls, on="라인명", how="left")
            _rt["판매량"] = _rt["판매량"].fillna(0)
            for _cc in ("매출", "수익"):
                if _cc in _rt.columns:
                    _rt[_cc] = _rt[_cc].fillna(0)
            _rt = _rt[_rt["입고량"] > 0].copy()
            _rt["회전율"] = _rt["판매량"] / _rt["입고량"] * 100
            _rt["현재고"] = (_rt["입고량"] - _rt["판매량"]).clip(lower=0)
            _rt["마지막판매"] = _rt["라인명"].map(_last_sale)
            _rt["완판기간"] = (_rt["마지막판매"] - _rt["첫입고일"]).dt.days
            _rt["입고경과일"] = (_tdy - _rt["첫입고일"]).dt.days
            _rt["수익률"] = np.where(_rt.get("매출", 0) > 0, _rt.get("수익", 0) / _rt["매출"].replace(0, np.nan) * 100, np.nan)
            _rt["대분류"] = _rt["라인명"].map(_ln_cat)

            def _turn_val(r):
                _el = (f' · 입고 {int(r["입고경과일"])}일' if pd.notna(r.get("입고경과일")) else '')
                return (f'<div style="font-size:13px;color:#0f172a;">회전율 '
                        f'<span style="color:#10b981;font-weight:700;">{r["회전율"]:.1f}%</span> '
                        f'<span style="color:#64748b;">· 입고 {int(r["입고량"]):,} · 현재고 {int(r["현재고"]):,} '
                        f'· 판매 {int(round(r["판매량"])):,}{_el}</span></div>')

            def _sellout_val(r):
                _fi = pd.Timestamp(r["첫입고일"]).strftime("%Y.%m.%d") if pd.notna(r["첫입고일"]) else "?"
                _ls = pd.Timestamp(r["마지막판매"]).strftime("%Y.%m.%d") if pd.notna(r["마지막판매"]) else "?"
                _pr = f' · 수익률 {r["수익률"]:.1f}%' if pd.notna(r.get("수익률")) else ''
                return (f'<div style="font-size:13px;color:#0f172a;">'
                        f'<span style="color:#b45309;font-weight:700;">완판 {int(r["완판기간"])}일</span> '
                        f'<span style="color:#64748b;">· 입고 {int(r["입고량"]):,} · 판매 {int(round(r["판매량"])):,}{_pr}</span></div>'
                        f'<div style="font-size:12px;color:#64748b;">최초입고 {_fi} → 최종판매 {_ls}</div>')

            # 3분할: 회전율 낮은 TOP50 / 회전율 높은 TOP50 / 완판 TOP50
            # (완판 판정: 부동소수 오차로 100.0이 미완판으로 새는 것 방지 → 99.95% 이상은 완판)
            _SOLD = _rt["회전율"] >= 99.95
            _live = _rt[~_SOLD]
            # 회전율 낮은: 입고 30일 이내(아직 팔릴 시간 부족)는 제외
            _tp_low = (_live[_live["입고경과일"] > 30].sort_values("회전율")
                       .head(50).reset_index(drop=True))
            _tp_high = _live.sort_values("회전율", ascending=False).head(50).reset_index(drop=True)
            _sd = (_rt[_SOLD & _rt["완판기간"].notna() & (_rt["완판기간"] >= 0)]
                   .sort_values("완판기간").head(50).reset_index(drop=True))

            _c1, _c2, _c3 = st.columns(3)
            with _c1:
                st.markdown("**회전율 낮은 TOP50** (입고 30일↑ · 안 팔림)")
                if _tp_low.empty:
                    st.caption("대상 없음")
                else:
                    st.markdown(metric_cards_html(_tp_low, _turn_val, n=len(_tp_low), img_px=92, start=1, step=1),
                                unsafe_allow_html=True)
            with _c2:
                st.markdown("**회전율 높은 TOP50** (완판 100% 제외)")
                if _tp_high.empty:
                    st.caption("대상 없음")
                else:
                    st.markdown(metric_cards_html(_tp_high, _turn_val, n=len(_tp_high), img_px=92, start=1, step=1),
                                unsafe_allow_html=True)
            with _c3:
                st.markdown("**완판 TOP50** (100% · 빨리 완판순)")
                if _sd.empty:
                    st.caption("완판(100% 회전) 상품 없음")
                else:
                    st.markdown(metric_cards_html(_sd, _sellout_val, n=len(_sd), img_px=92, start=1, step=1),
                                unsafe_allow_html=True)
            st.caption("회전율 = 라인 총판매량 ÷ 라인 총입고량. "
                       "완판=100% 소진(99.95%↑), 완판기간=최초 입고일~최종 판매일.")

        # 총재고 순위 (회전율 밑 · 판매 카드 스타일 · 총원가순)
        st.markdown("**총재고 순위**")
        _has_elapsed = "입고경과일행" in bstock.columns and bstock["입고경과일행"].notna().any()
        _agg_kw = dict(총원가=("총원가", "sum"), 재고수량=("수량", "sum"),
                       브랜드=("브랜드", "first"), 대분류=("대분류", "first"), 입고연도=("입고연도", "first"))
        if _has_elapsed:
            _agg_kw["입고경과일"] = ("입고경과일행", "max")  # 모든 사이즈 중 가장 오래된(=일수 최대)
        bt = (bstock.groupby("라인명", dropna=False).agg(**_agg_kw)
              .reset_index().sort_values("총원가", ascending=False).head(10).reset_index(drop=True))
        if bt.empty:
            st.caption("표시할 재고가 없습니다.")
        else:
            _bl = bt.iloc[0::2]
            _br = bt.iloc[1::2]
            kc1, kc2 = st.columns(2)
            with kc1:
                st.markdown(stock_cards_html(_bl, n=len(_bl), img_px=130, start=1, step=2),
                            unsafe_allow_html=True)
            with kc2:
                st.markdown(stock_cards_html(_br, n=len(_br), img_px=130, start=2, step=2),
                            unsafe_allow_html=True)

        sc1, sc2 = st.columns(2)
        # 카테고리(대분류)별 재고 비중 — 원가 기준
        with sc1:
            st.markdown("**카테고리별 재고 비중 (원가)**")
            catg = (bstock.groupby("대분류", dropna=False)["총원가"].sum()
                    .reset_index().sort_values("총원가", ascending=False))
            catg = catg[catg["총원가"] > 0]
            if catg.empty:
                st.caption("재고원가 데이터가 없습니다.")
            else:
                st.plotly_chart(share_donut(catg, "대분류", "총원가", "카테고리별 재고원가"),
                                use_container_width=True)
                _ct = catg["총원가"].sum()
                cdisp = pd.DataFrame({"대분류": catg["대분류"].values,
                                      "재고원가": catg["총원가"].round(0).astype("int64").values,
                                      "비중": (catg["총원가"] / _ct * 100).round(1).values})
                st.dataframe(cdisp, hide_index=True, use_container_width=True,
                             height=min(60 + len(cdisp) * 36, 320),
                             column_config={"재고원가": st.column_config.NumberColumn("재고원가", format="localized"),
                                            "비중": st.column_config.NumberColumn("비중", format="%.1f%%")})
        # 입고연도별 재고 비중 — 원가 기준
        with sc2:
            st.markdown("**입고연도별 재고 비중 (원가)**")
            seg = bstock.groupby("입고연도", dropna=False)["총원가"].sum().reset_index()
            seg = seg[seg["총원가"] > 0]
            if seg.empty:
                st.caption("재고원가 데이터가 없습니다.")
            else:
                seg["정렬"] = seg["입고연도"].map(lambda x: _yr_int(x) if _yr_int(x) > 0 else 9999)
                seg = seg.sort_values("정렬")
                seg["타입"] = np.where(seg["입고연도"].astype(str) == "미상", "미상", "입고")
                seg["라벨"] = seg["총원가"].apply(eok)
                fig_sg = px.bar(seg, x="입고연도", y="총원가", color="타입", text="라벨",
                                category_orders={"입고연도": seg["입고연도"].astype(str).tolist()},
                                color_discrete_map={"입고": "#6366f1", "미상": "#94a3b8"},
                                title="입고연도별 재고원가", labels={"총원가": "재고원가"})
                fig_sg.update_traces(textposition="outside", textangle=0, cliponaxis=False)
                fig_sg.update_layout(xaxis_type="category", margin=dict(t=54, b=10),
                                     legend_title_text="", uniformtext_minsize=8, uniformtext_mode="hide")
                fig_sg.update_xaxes(title_text="")
                fig_sg.update_yaxes(title_text="")
                st.plotly_chart(fig_sg, use_container_width=True)
                _st = seg["총원가"].sum()
                sdisp = pd.DataFrame({"입고연도": seg["입고연도"].astype(str).values,
                                      "재고원가": seg["총원가"].round(0).astype("int64").values,
                                      "비중": (seg["총원가"] / _st * 100).round(1).values})
                st.dataframe(sdisp, hide_index=True, use_container_width=True,
                             height=min(60 + len(sdisp) * 36, 320),
                             column_config={"재고원가": st.column_config.NumberColumn("재고원가", format="localized"),
                                            "비중": st.column_config.NumberColumn("비중", format="%.1f%%")})
