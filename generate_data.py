# -*- coding: utf-8 -*-
"""
리턴프리/카카오 채널 대시보드 데이터 생성 스크립트

사용법:
    python generate_data.py

입력:
    - RETURNFREE_DB_PATH: 리턴프리 SQLite DB 경로 (rentals_리턴프리 테이블)
    - KAKAO_XLSX_PATH   : 카카오 대여내역 엑셀 경로 (있으면 갱신, 없으면 스킵)

출력:
    - output/assets/data.js  (index.html이 <script>로 읽어들이는 전역 변수 DASHBOARD_DATA)

주의:
    - 매출은 VAT 제외 기준: 총결제요금 / 1.1
    - 하이패스요금은 별도로 청구되는 요금이며 매출에서 제외하지 않음 (총결제요금에 이미 포함된 그대로 사용)
"""

import sqlite3
import json
import glob
import os
import sys
from datetime import datetime, timedelta

import pandas as pd

# ------------------------------------------------------------------
# 경로 설정 (환경에 맞게 수정)
# ------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def find_latest_file(folder, pattern="*.xlsx"):
    """폴더 안에서 패턴에 맞는 파일 중 가장 최근에 수정된 파일의 경로를 반환.
    폴더가 없거나 매칭되는 파일이 없으면 None."""
    if not folder or not os.path.isdir(folder):
        return None
    files = glob.glob(os.path.join(folder, pattern))
    if not files:
        return None
    return max(files, key=os.path.getmtime)


RETURNFREE_DB_PATH = os.environ.get(
    "RETURNFREE_DB_PATH",
    r"C:\Users\USER\Desktop\수경\3. 데이터\대여내역\리턴프리\rental_data_returnfree_24~.db",
)
KAKAO_DB_PATH = os.environ.get(
    "KAKAO_DB_PATH",
    r"C:\Users\USER\Desktop\수경\3. 데이터\대여내역\KM\rental_data_km_24~.db",
)
RESERVATION_CONTROL_DB_PATH = os.environ.get(
    "RESERVATION_CONTROL_DB_PATH",
    r"C:\Users\USER\Desktop\수경\3. 데이터\예약관제\리턴프리\rental_data_예약관제_리턴프리_23~.db",
)
# 스테이션 등록 파일: 폴더 내 가장 최근 수정된 파일을 자동으로 사용
STATION_REGISTRY_FOLDER = os.environ.get(
    "STATION_REGISTRY_FOLDER",
    r"C:\Users\USER\Desktop\수경\3. 데이터\스팟\리턴프리(스테이션)",
)
STATION_REGISTRY_XLSX_PATH = os.environ.get(
    "STATION_REGISTRY_XLSX_PATH",
    "",  # 비워두면 STATION_REGISTRY_FOLDER에서 최신 파일을 자동 탐색
) or find_latest_file(STATION_REGISTRY_FOLDER)
# 주유충전비 안분 파일: 월별로 파일이 여러 개 쌓여있는 폴더 전체를 패턴 매칭으로 읽어들임
FUEL_PL_GLOB = os.environ.get(
    "FUEL_PL_GLOB",
    r"C:\Users\USER\Desktop\수경\3. 데이터\주유손익\*주유충전비*.xlsx",
)
OUTPUT_DIR = os.path.join(BASE_DIR, "output", "assets")
OUTPUT_JS = os.path.join(OUTPUT_DIR, "data.js")
TEMPLATE_HTML = os.path.join(BASE_DIR, "template", "index_template.html")
TEMPLATE_CHARTJS = os.path.join(BASE_DIR, "template", "chart.umd.min.js")
OUTPUT_HTML = os.path.join(BASE_DIR, "output", "index.html")


FIRST_USE_SHEET_PATH = os.path.join(BASE_DIR, "data", "first_use_sheet.csv")
# 첫이용 마스터 파일: 폴더 내 가장 최근 수정된 파일을 자동으로 사용
FIRST_USE_MASTER_FOLDER = os.environ.get(
    "FIRST_USE_MASTER_FOLDER",
    r"C:\Users\USER\Desktop\수경\3. 데이터\첫이용\master",
)
FIRST_USE_MASTER_XLSX_PATH = os.environ.get(
    "FIRST_USE_MASTER_XLSX_PATH",
    "",  # 비워두면 FIRST_USE_MASTER_FOLDER에서 최신 파일을 자동 탐색
) or find_latest_file(FIRST_USE_MASTER_FOLDER)

# 구글시트 직접 연동은 사내 그룹 공유 설정(도메인 인증 필요)이라 인증 없는 요청으로는 불가능하여 사용하지 않음.
# 대신 회원별 "생애 첫 이용" 이벤트 원본 마스터 파일(1행=1명)을 직접 집계해서 사용한다.


def load_first_use_master():
    """회원별 첫이용 이벤트 원본 마스터 엑셀(1행 = 회원 1명의 생애 첫 이용 건)을 읽어
    date(첫이용일자) 단위로 집계한 DataFrame(date, first_use)을 반환한다.
    파일이 없으면 None."""
    if not os.path.exists(FIRST_USE_MASTER_XLSX_PATH):
        return None
    m = pd.read_excel(FIRST_USE_MASTER_XLSX_PATH, usecols=["운행 시작"])
    m["운행 시작"] = pd.to_datetime(m["운행 시작"], errors="coerce")
    m = m.dropna(subset=["운행 시작"])
    m["date"] = m["운행 시작"].dt.date
    daily = m.groupby("date").size().reset_index(name="first_use")
    print(f"[첫이용 마스터] {FIRST_USE_MASTER_XLSX_PATH} 사용 ({len(m)}건, {daily['date'].min()}~{daily['date'].max()})")
    return daily


def _parse_wide_first_use_csv(text: str):
    """(레거시) 구글시트를 그대로 CSV로 내보낸 '넓은' 포맷(일자 헤더 + RF전체이용건(YYYY)/RF첫이용건(YYYY) 행)을
    date,total,first_use 긴 포맷 DataFrame으로 변환한다. 마스터 파일이 없을 때만 사용."""
    import csv
    import io
    import re

    reader = csv.reader(io.StringIO(text))
    rows = list(reader)
    if not rows:
        return None

    header = rows[0]
    date_cols = []  # (col_index, month, day)
    for c, cell in enumerate(header):
        m = re.match(r"^(\d{1,2})/(\d{1,2})", str(cell).strip())
        if m:
            date_cols.append((c, int(m.group(1)), int(m.group(2))))

    total_rows, first_rows = {}, {}
    for r, row in enumerate(rows[1:], start=1):
        label = str(row[0]).strip() if row else ""
        m = re.match(r"^RF전체이용건\((\d{4})\)$", label)
        if m:
            total_rows[m.group(1)] = row
            continue
        m = re.match(r"^RF첫이용건\((\d{4})\)$", label)
        if m:
            first_rows[m.group(1)] = row

    years = sorted(set(total_rows) & set(first_rows))
    if not years or not date_cols:
        return None

    records = []
    for year in years:
        trow, frow = total_rows[year], first_rows[year]
        for c, month, day in date_cols:
            if c >= len(trow):
                continue
            tval = str(trow[c]).replace(",", "").strip()
            fval = str(frow[c]).replace(",", "").strip() if c < len(frow) else ""
            if tval == "" or not tval.lstrip("-").isdigit():
                continue
            try:
                date_str = f"{int(year):04d}-{month:02d}-{day:02d}"
            except ValueError:
                continue
            first_use = int(fval) if fval.isdigit() else None
            records.append({"date": date_str, "total": int(tval), "first_use": first_use})

    if not records:
        return None
    return pd.DataFrame.from_records(records)


def load_first_use_sheet():
    """첫이용 데이터를 로드한다.
    1) 회원별 첫이용 이벤트 마스터 파일(data/master_리턴프리_첫이용_운영본.xlsx)이 있으면 최우선 사용 (가장 정확)
    2) 없으면 로컬 백업 CSV(data/first_use_sheet.csv, 예전 구글시트 내보내기 형식)로 대체
    3) 둘 다 없으면 None (이 경우 DB 기반 추정치로 대체됨)"""
    master = load_first_use_master()
    if master is not None:
        # total(전체이용건)은 반환자유 DB 쪽 데이터와 병합 시점에 맞춰줘야 하므로 여기선 first_use만 채우고
        # total은 호출부에서 DB 기반 전체 건수와 merge해서 채운다.
        return master.rename(columns={"first_use": "first_use"}).assign(total=None)

    if os.path.exists(FIRST_USE_SHEET_PATH):
        df = pd.read_csv(FIRST_USE_SHEET_PATH)
        print(f"[첫이용 시트] 로컬 백업 CSV 사용 ({len(df)}행)")
        df["date"] = pd.to_datetime(df["date"]).dt.date
        df = df.dropna(subset=["total", "first_use"])
        return df

    return None



TARGET_REGIONS = ["서울", "경기", "인천", "천안.아산"]


def classify_region(raw_region: str) -> str:
    """스테이션 등록 파일의 '지역'(시/도 + 시군구) 문자열을 4개 관심 권역 + 기타로 분류."""
    if not isinstance(raw_region, str):
        return "기타"
    if raw_region.startswith("서울"):
        return "서울"
    if raw_region.startswith("경기"):
        return "경기"
    if raw_region.startswith("인천"):
        return "인천"
    if "천안" in raw_region or "아산" in raw_region:
        return "천안.아산"
    return "기타"


def strip_station_prefix(name: str) -> str:
    """'(운영종료)', '(8/31 운영종료)' 같은 상태 접두어를 제거해 기준 스테이션명으로 변환."""
    import re

    return re.sub(r"^(\([^)]*\))+", "", str(name)).strip()


def load_station_region_map():
    """스테이션 등록 파일에서 '리턴프리 스팟 ID -> 4개 권역(또는 기타)' 매핑을 만든다.
    파일이 없으면 None을 반환하고, 이 경우 지역 이동 차트는 생략된다."""
    if not STATION_REGISTRY_XLSX_PATH or not os.path.exists(STATION_REGISTRY_XLSX_PATH):
        return None
    reg = pd.read_excel(STATION_REGISTRY_XLSX_PATH, usecols=["리턴프리 스팟 ID", "지역"])
    reg["region_group"] = reg["지역"].apply(classify_region)
    print(f"[스테이션 등록] {STATION_REGISTRY_XLSX_PATH} 사용 ({len(reg)}행)")
    # 같은 스팟 ID가 여러 행(이력)에 있을 수 있어 마지막 값 사용
    return reg.groupby("리턴프리 스팟 ID")["region_group"].last().to_dict()


def load_fuel_pl():
    """주유충전비 안분 파일(들)을 읽어 BM='Returnfree' 건에 대해 Adjusted_Fuel_Cost와
    주유손익을 계산한 DataFrame을 반환한다. 파일이 없으면 None.

    Adjusted_Fuel_Cost 산식 (사용자 확정):
      1) Trip_Status에 '취소' 포함되거나 '미사용완료'/'운행시작지연'인 건 제외
      2) (Trip_End_Time 기준 월, 차종)별로 거리합/주유비합을 구해 km당주유비 계산
         - 이 km당주유비는 전체 BM(왕복+편도+카카오) 데이터를 다 사용해서 산정 (표본을 늘려 정확도 확보)
      3) 각 건의 Adjusted_Fuel_Cost = Trip_Distance × 그 달·그 차종의 km당주유비
      4) 그 다음 BM='Returnfree'만 추려서 리턴프리 대시보드용으로 사용
    """
    paths = sorted(glob.glob(FUEL_PL_GLOB))
    if not paths:
        return None

    frames = []
    for p in paths:
        try:
            frames.append(pd.read_excel(p))
        except Exception as e:
            print(f"[경고] 주유손익 파일 로드 실패: {p} ({e})")
    if not frames:
        return None

    raw = pd.concat(frames, ignore_index=True)
    if "Reservation_ID" in raw.columns:
        raw = raw.drop_duplicates(subset="Reservation_ID")

    exclude = (
        raw["Trip_Status"].astype(str).str.contains("취소", na=False)
        | raw["Trip_Status"].isin(["미사용완료", "운행시작지연"])
    )
    base = raw[~exclude].copy()
    base["Trip_End_Time"] = pd.to_datetime(base["Trip_End_Time"], errors="coerce")
    base = base.dropna(subset=["Trip_End_Time"])
    base["ym"] = base["Trip_End_Time"].dt.to_period("M").astype(str)
    base["Trip_Distance"] = pd.to_numeric(base["Trip_Distance"], errors="coerce").fillna(0)
    base["Total_Fuel_Cost_Excl_VAT"] = pd.to_numeric(base["Total_Fuel_Cost_Excl_VAT"], errors="coerce").fillna(0)

    rate = (
        base.groupby(["ym", "Vehicle_Type"])
        .agg(거리합=("Trip_Distance", "sum"), 주유비합=("Total_Fuel_Cost_Excl_VAT", "sum"))
    )
    rate = rate[rate["거리합"] > 0]
    rate["km당주유비"] = rate["주유비합"] / rate["거리합"]

    base = base.merge(rate["km당주유비"], on=["ym", "Vehicle_Type"], how="left")
    base["Adjusted_Fuel_Cost"] = base["Trip_Distance"] * base["km당주유비"]
    base["Net_Driving_Fee_Excl_VAT"] = pd.to_numeric(base["Net_Driving_Fee_Excl_VAT"], errors="coerce").fillna(0)
    base["주유손익"] = base["Net_Driving_Fee_Excl_VAT"] - base["Adjusted_Fuel_Cost"]

    rf = base[base["BM"] == "Returnfree"].copy()
    rf = rf.dropna(subset=["km당주유비"])  # 해당 월·차종 표본이 없어 요율을 못 구한 건 제외

    trip_start = pd.to_datetime(rf["Trip_Start_Time"], errors="coerce")
    hour = trip_start.dt.hour
    rf["daytype"] = ((hour >= 21) | (hour < 5)).map({True: "심야", False: "주간"})

    print(f"[주유손익] {len(paths)}개 파일, Returnfree {len(rf):,}건 집계 완료")
    return rf


def week_month_labels(year_week_pairs):
    """(year, week) 쌍 리스트를 받아 ['N주'] 또는 월이 바뀌는 시점엔 ['N주','M월'] 형태의
    라벨 리스트로 변환한다 (Chart.js 멀티라인 틱 라벨용)."""
    from datetime import date as _date

    labels = []
    last_month = None
    for y, w in year_week_pairs:
        try:
            month = _date.fromisocalendar(int(y), int(w), 1).month
        except Exception:
            month = None
        if month != last_month:
            labels.append([f"{int(w)}주", f"{month}월" if month else ""])
            last_month = month
        else:
            labels.append(f"{int(w)}주")
    return labels


RETURNFREE_2026_GOALS = {
    "2026-01": 317340400,
    "2026-02": 450623490,
    "2026-03": 575635450,
    "2026-04": 545644400,
    "2026-05": 607345840,
    "2026-06": 676012800,
    "2026-07": 676895400,
    "2026-08": 633688800,
    "2026-09": 614356800,
    "2026-10": 576289200,
    "2026-11": 535744200,
    "2026-12": 482534400,
}


def load_returnfree(db_path: str) -> dict:
    if not os.path.exists(db_path):
        print(f"[경고] 리턴프리 DB를 찾을 수 없습니다: {db_path}")
        return None

    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query(
        '''
        SELECT 운행시작일, 운행종료일, 회원ID, 차량번호, 총결제요금, 하이패스요금, 운행거리, 추가주행요금, 출발스테이션번호, 도착스테이션번호, 쿠폰명, "할인(관리자쿠폰)" as 할인관리자쿠폰, 운행요금, 사용포인트, "할인(스테이션)" as 할인스테이션
        FROM rentals_리턴프리
        ''',
        conn,
    )
    conn.close()

    df["운행시작일"] = pd.to_datetime(df["운행시작일"], format="mixed", errors="coerce")
    df = df.dropna(subset=["운행시작일"])
    df["date"] = df["운행시작일"].dt.date
    df["ym"] = df["운행시작일"].dt.to_period("M").astype(str)
    df["매출"] = df["총결제요금"] / 1.1

    hour = df["운행시작일"].dt.hour
    df["구분"] = ((hour >= 21) | (hour < 5)).map({True: "심야", False: "주간"})

    def nd_agg(source_df, group_col, value_col=None):
        if value_col is None:
            g = source_df.groupby([group_col, "구분"]).size().unstack(fill_value=0)
        else:
            g = source_df.groupby([group_col, "구분"])[value_col].sum().unstack(fill_value=0)
        g = g.reset_index()
        if "심야" not in g.columns:
            g["심야"] = 0
        if "주간" not in g.columns:
            g["주간"] = 0
        g["합계"] = g["심야"] + g["주간"]
        g["심야비중"] = (g["심야"] / g["합계"] * 100).round(1)
        g["주간비중"] = (g["주간"] / g["합계"] * 100).round(1)
        return g.sort_values(group_col)

    def to_nd_dict(cnt_g, rev_g):
        return {
            "night_cnt": cnt_g["심야"].astype(int).tolist(),
            "day_cnt": cnt_g["주간"].astype(int).tolist(),
            "total_cnt": cnt_g["합계"].astype(int).tolist(),
            "night_pct": cnt_g["심야비중"].tolist(),
            "day_pct": cnt_g["주간비중"].tolist(),
            "night_rev": rev_g["심야"].round(0).astype(int).tolist(),
            "day_rev": rev_g["주간"].round(0).astype(int).tolist(),
            "total_rev": rev_g["합계"].round(0).astype(int).tolist(),
            "night_rev_pct": rev_g["심야비중"].tolist(),
            "day_rev_pct": rev_g["주간비중"].tolist(),
        }

    monthly = (
        df.groupby("ym")
        .agg(건수=("매출", "count"), 매출=("매출", "sum"), 운행거리=("운행거리", "sum"), 총결제요금=("총결제요금", "sum"))
        .reset_index()
    )
    monthly["매출"] = monthly["매출"].round(0)
    monthly["목표대비매출"] = (monthly["총결제요금"] / 1.1).round(0)

    monthly_nd_cnt = nd_agg(df, "ym")
    monthly_nd_rev = nd_agg(df, "ym", "매출")

    daily = (
        df.groupby("date")
        .agg(건수=("매출", "count"), 매출=("매출", "sum"), 운행거리=("운행거리", "sum"))
        .reset_index()
    )
    daily["매출"] = daily["매출"].round(0)

    daily_nd_cnt = nd_agg(df, "date")
    daily_nd_rev = nd_agg(df, "date", "매출")
    daily_nd_cnt_recent = daily_nd_cnt.tail(60)
    daily_nd_rev_recent = daily_nd_rev.tail(60)

    # 주차별 심야/주간 트렌드 (심야: 21:00~04:59 출발, 주간: 그 외) - 연도별로 분리
    daily_recent = daily.tail(60)

    wk = df.copy()
    iso = wk["운행시작일"].dt.isocalendar()
    wk["iso_year"] = iso["year"]
    wk["iso_week"] = iso["week"]
    wk["yw"] = wk["iso_year"].astype(str) + "-" + wk["iso_week"].astype(str).str.zfill(2)

    wg_cnt = nd_agg(wk, "yw")
    wg_rev = nd_agg(wk, "yw", "매출")
    yw_meta = wk.drop_duplicates("yw")[["yw", "iso_year", "iso_week"]].sort_values("yw")
    wg_cnt = wg_cnt.merge(yw_meta, on="yw").sort_values(["iso_year", "iso_week"])
    wg_rev = wg_rev.merge(yw_meta, on="yw").sort_values(["iso_year", "iso_week"])

    weekly_by_year = {}
    for yr in sorted(wg_cnt["iso_year"].unique()):
        c = wg_cnt[wg_cnt["iso_year"] == yr]
        r = wg_rev[wg_rev["iso_year"] == yr]
        d = to_nd_dict(c, r)
        d["labels"] = week_month_labels([(yr, w) for w in c["iso_week"]])
        weekly_by_year[str(int(yr))] = d

    # 전체 연도 통합 뷰 (연속 주차)
    d_all = to_nd_dict(wg_cnt, wg_rev)
    d_all["labels"] = week_month_labels(list(zip(wg_cnt["iso_year"], wg_cnt["iso_week"])))
    weekly_by_year["전체"] = d_all

    # ---------------------------------------------------------------
    # 주간 트렌드 상세지표 (전년동기 vs 이번 해)
    # ---------------------------------------------------------------
    yd = df.copy()
    yd["운행종료일"] = pd.to_datetime(yd["운행종료일"], format="mixed", errors="coerce")
    yd = yd.dropna(subset=["운행종료일"])
    yd["이용분"] = (yd["운행종료일"] - yd["운행시작일"]).dt.total_seconds() / 60
    yd = yd[(yd["이용분"] > 0) & (yd["이용분"] < 24 * 60 * 14)]  # 이상치(14일 초과) 제외
    yd["매출"] = yd["총결제요금"] / 1.1
    yd["운행거리"] = pd.to_numeric(yd["운행거리"], errors="coerce")
    yd["쿠폰"] = yd["쿠폰명"].fillna("쿠폰미사용")
    yd["할인금액"] = pd.to_numeric(yd["할인관리자쿠폰"], errors="coerce").fillna(0)
    yd["운행요금"] = pd.to_numeric(yd["운행요금"], errors="coerce").fillna(0)
    yd["사용포인트"] = pd.to_numeric(yd["사용포인트"], errors="coerce").fillna(0)
    yd["할인스테이션"] = pd.to_numeric(yd["할인스테이션"], errors="coerce").fillna(0)
    yd["추가주행매출"] = pd.to_numeric(yd["추가주행요금"], errors="coerce") / 1.1
    iso2 = yd["운행시작일"].dt.isocalendar()
    yd["iso_year"] = iso2["year"].astype(int)
    yd["iso_week"] = iso2["week"].astype(int)
    yd["daytype"] = yd["운행시작일"].dt.weekday.map(lambda x: "주말" if x >= 5 else "주중")
    _hour = yd["운행시작일"].dt.hour
    yd["hourtype"] = ((_hour >= 21) | (_hour < 5)).map({True: "심야", False: "주간"})

    first_use_idx = yd.groupby("회원ID")["운행시작일"].idxmin()
    yd["첫이용"] = False
    yd.loc[first_use_idx, "첫이용"] = True

    # 전년동기 비교는 "몇 번째 주차인지"가 아니라 "같은 요일 기준으로 364일(52주) 전"으로 맞춘다.
    # (ISO 주차 번호만으로 매칭하면 53주짜리 해가 끼는 해엔 뒤로 갈수록 요일이 밀림)
    # orig_year는 실제 달력 연도(라벨/연도 구분용)로 그대로 두고, iso_week만 정렬용으로 덮어쓴다.
    weekly_yoy_years = sorted(yd["iso_year"].unique().tolist())
    cmp_years = weekly_yoy_years[-2:] if len(weekly_yoy_years) >= 2 else weekly_yoy_years
    cur_year = cmp_years[-1]
    prev_year = cmp_years[0] if len(cmp_years) > 1 else None

    yd["orig_year"] = yd["iso_year"]
    yd["cmp_date"] = yd["운행시작일"].dt.normalize()
    if prev_year is not None:
        _prev_mask = yd["iso_year"] == prev_year
        yd.loc[_prev_mask, "cmp_date"] = yd.loc[_prev_mask, "cmp_date"] + pd.Timedelta(days=364)
    yd["iso_week"] = yd["cmp_date"].dt.isocalendar()["week"].astype(int)
    yd["iso_year"] = yd["orig_year"]  # 그룹핑용 라벨은 실제 달력연도 그대로 유지

    # 차량수는 "그 주에 리턴프리로 1회 이상 운행된 고유 차량수"를 사용한다 (요일/시간대 세부 필터와 무관하게
    # 항상 주 전체 기준의 값을 분모로 사용해, 부분집합을 합쳤을 때 전체와 어긋나지 않도록 한다).
    total_vehicle_by_week = (
        yd.groupby(["iso_year", "iso_week"])["차량번호"].nunique().reset_index(name="전체차량수")
    )

    AVAILABLE_MIN = {
        "전체": 10080, "주중": 7200, "주말": 2880, "주간": 6720, "심야": 3360,
        "주중+주간": 5 * 16 * 60, "주중+심야": 5 * 8 * 60,
        "주말+주간": 2 * 16 * 60, "주말+심야": 2 * 8 * 60,
    }

    def compute_metric_block(sub):
        g = sub.groupby(["iso_year", "iso_week"]).agg(
            매출=("매출", "sum"),
            추가주행매출합=("추가주행매출", "sum"),
            이용분합=("이용분", "sum"),
            거리합=("운행거리", "sum"),
            건수=("매출", "count"),
            전체이용자수=("회원ID", "nunique"),
            첫이용자수=("첫이용", "sum"),
        ).reset_index().sort_values(["iso_year", "iso_week"])
        g = g.merge(total_vehicle_by_week, on=["iso_year", "iso_week"], how="left")
        g["차량수"] = g["전체차량수"]
        g["대당매출"] = (g["매출"] / g["전체차량수"]).round(0)
        g["평균이용시간"] = (g["이용분합"] / g["건수"]).round(1)
        g["평균이동거리"] = (g["거리합"] / g["건수"]).round(1)
        g["분당매출"] = (g["매출"] / g["이용분합"]).round(0)
        # 재이용율 = 전주 이용자수 / 금주 이용자수 * 100 (같은 시간구간 내 연속 비교)
        g["재이용율"] = (g["전체이용자수"].shift(1) / g["전체이용자수"] * 100).round(1)
        return g

    metric_map = {
        "차량수": ("차량수", "대수"),
        "대당매출": ("대당매출", "매출"),
        "가동률": ("가동률", "%"),
        "이용자수": ("전체이용자수", "명"),
        "첫이용자": ("첫이용자수", "명"),
        "재이용율": ("재이용율", "%"),
        "평균이용시간": ("평균이용시간", "분"),
        "평균이동거리": ("평균이동거리", "km"),
        "추가주행매출": ("추가주행매출합", "매출"),
        "분당매출": ("분당매출", "매출"),
    }

    weekly_yoy = {"years": [str(y) for y in cmp_years], "daytypes": {}}
    _ref_year = cmp_years[-1] if cmp_years else datetime.now().year
    weekly_yoy["week_labels"] = week_month_labels([(_ref_year, w) for w in range(1, 54)])
    ALL_DTYPES = ["전체", "주중", "주말", "주간", "심야", "주중+주간", "주중+심야", "주말+주간", "주말+심야"]

    for dtype in ALL_DTYPES:
        if dtype == "전체":
            sub = yd
        elif dtype in ("주중", "주말"):
            sub = yd[yd["daytype"] == dtype]
        elif dtype in ("주간", "심야"):
            sub = yd[yd["hourtype"] == dtype]
        else:
            wd, ht = dtype.split("+")
            sub = yd[(yd["daytype"] == wd) & (yd["hourtype"] == ht)]
        g = compute_metric_block(sub)
        g["가동률"] = (g["이용분합"] / (g["차량수"] * AVAILABLE_MIN[dtype]) * 100).round(1)

        metrics_out = {}
        for m_name, (col, unit) in metric_map.items():
            per_year = {}
            for y in cmp_years:
                yrow = g[g["iso_year"] == y].sort_values("iso_week")
                weeks = yrow["iso_week"].tolist()
                vals = yrow[col].tolist()
                per_year[str(y)] = {"weeks": weeks, "values": [None if pd.isna(v) else (round(v, 1) if isinstance(v, float) else v) for v in vals]}
            metrics_out[m_name] = {"unit": unit, "by_year": per_year}
        weekly_yoy["daytypes"][dtype] = metrics_out

    # ---------------------------------------------------------------
    # 이용시간구간별 분포 (전년동기 vs 이번 해 주차별)
    # ---------------------------------------------------------------
    DURATION_BUCKETS = ["1시간 이하", "1시간 초과~2시간 이하", "2시간 초과~3시간 이하", "3시간 초과"]

    def classify_duration(minutes):
        if minutes <= 60:
            return DURATION_BUCKETS[0]
        if minutes <= 120:
            return DURATION_BUCKETS[1]
        if minutes <= 180:
            return DURATION_BUCKETS[2]
        return DURATION_BUCKETS[3]

    yd["이용시간구간"] = yd["이용분"].apply(classify_duration)
    dur_g = (
        yd.groupby(["iso_year", "iso_week", "이용시간구간"]).size().unstack(fill_value=0)
    )
    for b in DURATION_BUCKETS:
        if b not in dur_g.columns:
            dur_g[b] = 0
    dur_g = dur_g.reset_index()

    duration_by_year = {}
    for y in cmp_years:
        yrow = dur_g[dur_g["iso_year"] == y].sort_values("iso_week")
        weeks = yrow["iso_week"].tolist()
        duration_by_year[str(y)] = {
            "weeks": weeks,
            "buckets": {b: yrow[b].astype(int).tolist() for b in DURATION_BUCKETS},
        }
    weekly_yoy["duration_buckets"] = {"labels": DURATION_BUCKETS, "by_year": duration_by_year}

    # 첫이용자는 수기 관리 시트(구글시트)가 있으면 그 값을 우선 사용 (재이용율은 위 이용자수 기반 공식 사용)
    first_use_sheet = load_first_use_sheet()
    if first_use_sheet is not None:
        fs = first_use_sheet.copy()
        fs["date_dt"] = pd.to_datetime(fs["date"])
        fs["orig_year"] = fs["date_dt"].dt.isocalendar()["year"].astype(int)
        fs["cmp_date"] = fs["date_dt"]
        if prev_year is not None:
            _pm = fs["orig_year"] == prev_year
            fs.loc[_pm, "cmp_date"] = fs.loc[_pm, "cmp_date"] + pd.Timedelta(days=364)
        fs["iso_year"] = fs["orig_year"]
        fs["iso_week"] = fs["cmp_date"].dt.isocalendar()["week"].astype(int)
        fs["daytype"] = fs["date_dt"].dt.weekday.map(lambda x: "주말" if x >= 5 else "주중")

        for dtype in ["전체", "주중", "주말"]:
            fsub = fs if dtype == "전체" else fs[fs["daytype"] == dtype]
            fg = fsub.groupby(["iso_year", "iso_week"]).agg(
                total=("total", "sum"), first_use=("first_use", "sum")
            ).reset_index()

            per_year = {}
            for y in cmp_years:
                yrow = fg[fg["iso_year"] == y].sort_values("iso_week")
                weeks = yrow["iso_week"].tolist()
                vals = yrow["first_use"].tolist()
                per_year[str(y)] = {"weeks": weeks, "values": [round(v, 1) if isinstance(v, float) else v for v in vals]}
            unit = weekly_yoy["daytypes"][dtype]["첫이용자"]["unit"]
            weekly_yoy["daytypes"][dtype]["첫이용자"] = {"unit": unit, "by_year": per_year, "source": "master"}

    # ---------------------------------------------------------------
    # 건수·매출 트렌드 위에 띄울 주간 실적 요약 카드 (주차 선택 가능, 전주/전년동기 대비)
    # ---------------------------------------------------------------
    def series_by_year(frame, col):
        out = {}
        for y in cmp_years:
            rows = frame[frame["iso_year"] == y].sort_values("iso_week")
            vals = rows[col].tolist()
            out[str(y)] = {
                "weeks": rows["iso_week"].tolist(),
                "values": [round(v, 1) if isinstance(v, float) else v for v in vals],
            }
        return out

    g_cnt_rev = (
        yd.groupby(["iso_year", "iso_week"])
        .agg(건수=("매출", "count"), 매출=("매출", "sum"))
        .reset_index()
        .sort_values(["iso_year", "iso_week"])
    )
    users_g = (
        yd.groupby(["iso_year", "iso_week"])["회원ID"].nunique().reset_index(name="이용자수_raw")
    )
    g_cnt_rev = g_cnt_rev.merge(users_g, on=["iso_year", "iso_week"], how="left")
    g_cnt_rev["인당매출"] = (g_cnt_rev["매출"] / g_cnt_rev["이용자수_raw"]).round(0)
    g_cnt_rev["건당매출"] = (g_cnt_rev["매출"] / g_cnt_rev["건수"]).round(0)

    day_g = (
        yd[yd["hourtype"] == "주간"].groupby(["iso_year", "iso_week"]).size()
        .reset_index(name="건수").sort_values(["iso_year", "iso_week"])
    )
    night_g = (
        yd[yd["hourtype"] == "심야"].groupby(["iso_year", "iso_week"]).size()
        .reset_index(name="건수").sort_values(["iso_year", "iso_week"])
    )

    kpi_metrics = {
        "건수": {"unit": "건", "by_year": series_by_year(g_cnt_rev, "건수")},
        "주간건수": {"unit": "건", "by_year": series_by_year(day_g, "건수")},
        "심야건수": {"unit": "건", "by_year": series_by_year(night_g, "건수")},
        "총매출": {"unit": "매출", "by_year": series_by_year(g_cnt_rev, "매출")},
    }
    for m_name in ["이용자수"]:
        m = weekly_yoy["daytypes"]["전체"].get(m_name)
        if m:
            kpi_metrics[m_name] = {"unit": m["unit"], "by_year": m["by_year"]}
    kpi_metrics.update({
        "건당매출": {"unit": "매출", "by_year": series_by_year(g_cnt_rev, "건당매출")},
        "인당매출": {"unit": "매출", "by_year": series_by_year(g_cnt_rev, "인당매출")},
    })
    for m_name in ["평균이용시간", "평균이동거리", "분당매출"]:
        m = weekly_yoy["daytypes"]["전체"].get(m_name)
        if m:
            kpi_metrics[m_name] = {"unit": m["unit"], "by_year": m["by_year"]}

    # ---------------------------------------------------------------
    # 주간 실적 요약의 월별 버전 (같은 지표 세트를 월 단위로)
    # ---------------------------------------------------------------
    yd["ym_year"] = yd["운행시작일"].dt.year
    yd["ym_month"] = yd["운행시작일"].dt.month

    def series_by_year_month(frame, col):
        out = {}
        years_m = sorted(frame["ym_year"].unique().tolist())
        for y in years_m:
            rows = frame[frame["ym_year"] == y].sort_values("ym_month")
            vals = rows[col].tolist()
            out[str(y)] = {
                "weeks": rows["ym_month"].tolist(),  # 프론트에서 재사용하는 키(주차/월 공통) - 월에서는 '월'을 담음
                "values": [round(v, 1) if isinstance(v, float) else v for v in vals],
            }
        return out

    g_cnt_rev_m = (
        yd.groupby(["ym_year", "ym_month"])
        .agg(건수=("매출", "count"), 매출=("매출", "sum"))
        .reset_index()
        .sort_values(["ym_year", "ym_month"])
    )
    users_g_m = yd.groupby(["ym_year", "ym_month"])["회원ID"].nunique().reset_index(name="이용자수_raw")
    g_cnt_rev_m = g_cnt_rev_m.merge(users_g_m, on=["ym_year", "ym_month"], how="left")
    g_cnt_rev_m["인당매출"] = (g_cnt_rev_m["매출"] / g_cnt_rev_m["이용자수_raw"]).round(0)
    g_cnt_rev_m["건당매출"] = (g_cnt_rev_m["매출"] / g_cnt_rev_m["건수"]).round(0)

    day_g_m = (
        yd[yd["hourtype"] == "주간"].groupby(["ym_year", "ym_month"]).size()
        .reset_index(name="건수").sort_values(["ym_year", "ym_month"])
    )
    night_g_m = (
        yd[yd["hourtype"] == "심야"].groupby(["ym_year", "ym_month"]).size()
        .reset_index(name="건수").sort_values(["ym_year", "ym_month"])
    )
    misc_g_m = (
        yd.groupby(["ym_year", "ym_month"])
        .agg(이용분합=("이용분", "sum"), 거리합=("운행거리", "sum"), 건수=("매출", "count"), 매출=("매출", "sum"))
        .reset_index()
    )
    misc_g_m["평균이용시간"] = (misc_g_m["이용분합"] / misc_g_m["건수"]).round(0)
    misc_g_m["평균이동거리"] = (misc_g_m["거리합"] / misc_g_m["건수"]).round(1)
    misc_g_m["분당매출"] = (misc_g_m["매출"] / misc_g_m["이용분합"]).round(0)

    kpi_metrics_monthly = {
        "건수": {"unit": "건", "by_year": series_by_year_month(g_cnt_rev_m, "건수")},
        "주간건수": {"unit": "건", "by_year": series_by_year_month(day_g_m, "건수")},
        "심야건수": {"unit": "건", "by_year": series_by_year_month(night_g_m, "건수")},
        "총매출": {"unit": "매출", "by_year": series_by_year_month(g_cnt_rev_m, "매출")},
        "이용자수": {"unit": "명", "by_year": series_by_year_month(g_cnt_rev_m, "이용자수_raw")},
        "건당매출": {"unit": "매출", "by_year": series_by_year_month(g_cnt_rev_m, "건당매출")},
        "인당매출": {"unit": "매출", "by_year": series_by_year_month(g_cnt_rev_m, "인당매출")},
        "평균이용시간": {"unit": "분", "by_year": series_by_year_month(misc_g_m, "평균이용시간")},
        "평균이동거리": {"unit": "km", "by_year": series_by_year_month(misc_g_m, "평균이동거리")},
        "분당매출": {"unit": "매출", "by_year": series_by_year_month(misc_g_m, "분당매출")},
    }

    # 주차 선택 드롭다운에 쓸 실제 날짜 범위 (해당 연도 기준 월~일)
    kpi_week_ranges = {}
    for w in sorted(set(g_cnt_rev[g_cnt_rev["iso_year"] == cur_year]["iso_week"])):
        start = datetime.fromisocalendar(int(cur_year), int(w), 1).date()
        end = start + timedelta(days=6)
        kpi_week_ranges[int(w)] = f"{start.isoformat()} ~ {end.isoformat()}"

    # ---------------------------------------------------------------
    # 요일 x 시간대 출발건수 히트맵
    # ---------------------------------------------------------------
    def build_heatmap(sub):
        h = sub.copy()
        h["weekday"] = h["운행시작일"].dt.weekday  # 0=월 ... 6=일
        h["hour"] = h["운행시작일"].dt.hour
        pivot = h.groupby(["weekday", "hour"]).size().unstack(fill_value=0)
        pivot = pivot.reindex(index=range(7), columns=range(24), fill_value=0)
        return pivot.values.tolist()

    _latest_dt = df["운행시작일"].max()
    heatmaps = {
        "최근 12주": build_heatmap(df[df["운행시작일"] >= _latest_dt - pd.Timedelta(weeks=12)]),
        "최근 4주": build_heatmap(df[df["운행시작일"] >= _latest_dt - pd.Timedelta(weeks=4)]),
    }

    # ---------------------------------------------------------------
    # 지역(서울/경기/인천/천안.아산) 출발·반납 건수
    # ---------------------------------------------------------------
    regional_moves = None
    regional_flows_out = None
    region_map = load_station_region_map()
    if region_map is not None:
        rg = df.copy()
        rg["출발지역"] = rg["출발스테이션번호"].map(region_map).fillna("기타")
        rg["도착지역"] = rg["도착스테이션번호"].map(region_map).fillna("기타")

        def region_counts(sub):
            dep = sub["출발지역"].value_counts()
            ret = sub["도착지역"].value_counts()
            return {
                r: {"출발": int(dep.get(r, 0)), "반납": int(ret.get(r, 0))}
                for r in TARGET_REGIONS
            }

        FLOW_NODES = TARGET_REGIONS + ["기타"]

        def region_flows(sub):
            pair_counts = sub.groupby(["출발지역", "도착지역"]).size()
            flows = []
            for src in FLOW_NODES:
                for dst in FLOW_NODES:
                    v = int(pair_counts.get((src, dst), 0))
                    if v > 0:
                        flows.append({"from": src, "to": dst, "value": v})
            return flows

        regional_moves = {
            "전체 기간": region_counts(rg),
            "최근 12주": region_counts(rg[rg["운행시작일"] >= _latest_dt - pd.Timedelta(weeks=12)]),
            "최근 4주": region_counts(rg[rg["운행시작일"] >= _latest_dt - pd.Timedelta(weeks=4)]),
        }
        regional_flows_out = {
            "전체 기간": region_flows(rg),
            "최근 12주": region_flows(rg[rg["운행시작일"] >= _latest_dt - pd.Timedelta(weeks=12)]),
            "최근 4주": region_flows(rg[rg["운행시작일"] >= _latest_dt - pd.Timedelta(weeks=4)]),
        }

    # ---------------------------------------------------------------
    # 쿠폰 성과 테이블 (주차/월 선택, 사용건수는 전주/전월 대비 %)
    # ---------------------------------------------------------------
    cp = yd.copy()
    iso_cp = cp["운행시작일"].dt.isocalendar()
    cp["yw"] = iso_cp["year"].astype(str) + "-W" + iso_cp["week"].astype(int).astype(str).str.zfill(2)
    cp["ym"] = cp["운행시작일"].dt.to_period("M").astype(str)

    top_coupons = (
        cp["쿠폰"].value_counts().index.tolist()
    )
    # '쿠폰미사용'은 항상 포함, 나머지는 사용건수 상위 24개만
    top_coupons = ["쿠폰미사용"] + [c for c in top_coupons if c != "쿠폰미사용"][:24]

    def coupon_period_table(period_col):
        periods = sorted(cp[period_col].unique().tolist())
        g = (
            cp[cp["쿠폰"].isin(top_coupons)]
            .groupby([period_col, "쿠폰"])
            .agg(
                건수=("매출", "count"),
                할인합=("할인금액", "sum"),
                매출합=("매출", "sum"),
                이용분합=("이용분", "sum"),
                거리합=("운행거리", "sum"),
                운행요금합=("운행요금", "sum"),
                포인트합=("사용포인트", "sum"),
                할인스테이션합=("할인스테이션", "sum"),
            )
            .reset_index()
        )
        coupons_out = {}
        for name in top_coupons:
            sub = g[g["쿠폰"] == name].drop(columns="쿠폰").set_index(period_col)
            sub = sub.reindex(periods, fill_value=0)
            cnt = sub["건수"].tolist()
            base = sub["운행요금합"] + sub["포인트합"] + sub["할인스테이션합"] + sub["할인합"]
            coupons_out[name] = {
                "cnt": [int(v) for v in cnt],
                "discount_per": [round(sub["할인합"].iloc[i] / cnt[i], 0) if cnt[i] else 0 for i in range(len(periods))],
                "discount_rate": [round(sub["할인합"].iloc[i] / base.iloc[i] * 100, 1) if base.iloc[i] else 0 for i in range(len(periods))],
                "revenue_per": [round(sub["매출합"].iloc[i] / cnt[i], 0) if cnt[i] else 0 for i in range(len(periods))],
                "revenue_per_min": [round(sub["매출합"].iloc[i] / sub["이용분합"].iloc[i], 0) if sub["이용분합"].iloc[i] else 0 for i in range(len(periods))],
                "avg_min": [round(sub["이용분합"].iloc[i] / cnt[i], 0) if cnt[i] else 0 for i in range(len(periods))],
                "avg_km": [round(sub["거리합"].iloc[i] / cnt[i], 1) if cnt[i] else 0 for i in range(len(periods))],
            }
        return {"periods": periods, "coupons": coupons_out}

    weekly_table = coupon_period_table("yw")
    monthly_table = coupon_period_table("ym")

    weekly_period_ranges = {}
    for yw in weekly_table["periods"]:
        yr, wk = yw.split("-W")
        start = datetime.fromisocalendar(int(yr), int(wk), 1).date()
        end = start + timedelta(days=6)
        weekly_period_ranges[yw] = f"{start.isoformat()} ~ {end.isoformat()}"

    coupon_summary = {
        "weekly": {**weekly_table, "period_ranges": weekly_period_ranges},
        "monthly": monthly_table,
    }

    # ---------------------------------------------------------------
    # 주유손익 (BM=Returnfree, Adjusted_Fuel_Cost 기반)
    # ---------------------------------------------------------------
    fuel_pl = None
    fpl = load_fuel_pl()
    if fpl is not None and len(fpl) > 0:
        def fuel_agg(g):
            driving = g["Net_Driving_Fee_Excl_VAT"].sum()
            fuel = g["Adjusted_Fuel_Cost"].sum()
            dist = g["Trip_Distance"].sum()
            cnt = len(g)
            pl = driving - fuel
            return {
                "cnt": int(cnt),
                "driving_fee": round(driving),
                "adjusted_fuel_cost": round(fuel),
                "pl": round(pl),
                "rate": round(pl / driving * 100, 1) if driving else None,
                "per_km": round(pl / dist, 1) if dist else None,
                "per_trip": round(pl / cnt) if cnt else None,
            }

        def fuel_breakdown(g):
            by_fuel_type = [{"name": name, **fuel_agg(sub)} for name, sub in g.groupby("Fuel_Type")]
            by_fuel_type.sort(key=lambda x: -x["cnt"])
            by_vehicle_type = [{"name": name, **fuel_agg(sub)} for name, sub in g.groupby("Vehicle_Type")]
            by_vehicle_type.sort(key=lambda x: -x["cnt"])
            by_daytype = {name: fuel_agg(sub) for name, sub in g.groupby("daytype")}
            return {
                "summary": fuel_agg(g),
                "by_fuel_type": by_fuel_type,
                "by_vehicle_type": by_vehicle_type,
                "by_daytype": by_daytype,
            }

        def trip_level_stats(g):
            total = len(g)
            loss_rows = g[g["주유손익"] < 0]
            profit_rows = g[g["주유손익"] >= 0]
            loss = len(loss_rows)
            scatter = [
                {"dist": round(row["Trip_Distance"], 1), "pl": round(row["주유손익"]), "profit": bool(row["주유손익"] >= 0)}
                for _, row in g.iterrows()
            ]
            return {
                "cnt_total": total,
                "cnt_loss": loss,
                "loss_rate": round(loss / total * 100, 1) if total else None,
                "loss_sum": round(loss_rows["주유손익"].sum()) if loss else 0,
                "profit_sum": round(profit_rows["주유손익"].sum()) if len(profit_rows) else 0,
                "scatter": scatter,
            }

        def customer_level_stats(g):
            cust = g.groupby("Customer_ID").agg(pl=("주유손익", "sum"), trips=("주유손익", "count")).reset_index()
            total_cust = len(cust)
            loss_cust = cust[cust["pl"] < 0]
            profit_cust = cust[cust["pl"] >= 0]
            loss_cnt = len(loss_cust)
            scatter = [
                {"trips": int(r.trips), "pl": round(r.pl), "profit": bool(r.pl >= 0)}
                for r in cust.itertuples()
            ]
            return {
                "customer_cnt": total_cust,
                "loss_customer_cnt": loss_cnt,
                "loss_customer_rate": round(loss_cnt / total_cust * 100, 1) if total_cust else None,
                "loss_sum": round(loss_cust["pl"].sum()) if loss_cnt else 0,
                "profit_sum": round(profit_cust["pl"].sum()) if len(profit_cust) else 0,
                "scatter": scatter,
            }

        fuel_months = sorted(fpl["ym"].unique().tolist())
        fuel_by_month = {ym: fuel_breakdown(fpl[fpl["ym"] == ym]) for ym in fuel_months}
        fuel_by_month["전체"] = fuel_breakdown(fpl)

        fuel_trip_level = {ym: trip_level_stats(fpl[fpl["ym"] == ym]) for ym in fuel_months}
        fuel_trip_level["전체"] = trip_level_stats(fpl)

        fuel_customer_level = {ym: customer_level_stats(fpl[fpl["ym"] == ym]) for ym in fuel_months}
        fuel_customer_level["전체"] = customer_level_stats(fpl)

        fuel_monthly_trend = {"months": fuel_months, "pl": [], "rate": [], "per_km": [], "per_trip": []}
        for ym in fuel_months:
            a = fuel_agg(fpl[fpl["ym"] == ym])
            fuel_monthly_trend["pl"].append(a["pl"])
            fuel_monthly_trend["rate"].append(a["rate"])
            fuel_monthly_trend["per_km"].append(a["per_km"])
            fuel_monthly_trend["per_trip"].append(a["per_trip"])

        fuel_pl = {
            "months": fuel_months,
            "by_month": fuel_by_month,
            "trip_level": fuel_trip_level,
            "customer_level": fuel_customer_level,
            "monthly_trend": fuel_monthly_trend,
        }

    latest_ym = monthly["ym"].max()
    this_month = df[df["ym"] == latest_ym]
    latest_date = df["date"].max()
    month_start = latest_date.replace(day=1)

    summary = {
        "total_cnt": int(len(df)),
        "total_rev": round(df["매출"].sum()),
        "avg_fare": round(df["매출"].sum() / len(df)) if len(df) else 0,
        "latest_ym": latest_ym,
        "latest_month_cnt": int(len(this_month)),
        "latest_month_rev": round(this_month["매출"].sum()),
        "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }

    # 2026년 목표 대비 실적 (실적 = 총결제요금 / 1.1)
    goals = []
    for ym, goal in RETURNFREE_2026_GOALS.items():
        row = monthly[monthly["ym"] == ym]
        actual = float(row["목표대비매출"].iloc[0]) if len(row) else None
        goals.append({
            "ym": ym,
            "goal": goal,
            "actual": round(actual) if actual is not None else None,
            "rate": round(actual / goal * 100, 1) if actual is not None and goal else None,
        })
    latest_goal_row = next((g for g in goals if g["ym"] == latest_ym), None)
    if latest_goal_row is not None:
        latest_goal_row["data_from"] = month_start.strftime("%m/%d")
        latest_goal_row["data_to"] = latest_date.strftime("%m/%d")

    return {
        "monthly_labels": monthly["ym"].tolist(),
        "monthly_rev": monthly["매출"].astype(int).tolist(),
        "monthly_cnt": monthly["건수"].astype(int).tolist(),
        "monthly_nd": to_nd_dict(monthly_nd_cnt, monthly_nd_rev),
        "daily_labels": [str(d) for d in daily_recent["date"].tolist()],
        "daily_rev": daily_recent["매출"].astype(int).tolist(),
        "daily_cnt": daily_recent["건수"].astype(int).tolist(),
        "daily_nd": to_nd_dict(daily_nd_cnt_recent, daily_nd_rev_recent),
        "weekly_by_year": weekly_by_year,
        "weekly_yoy": weekly_yoy,
        "kpi_metrics": kpi_metrics,
        "kpi_metrics_monthly": kpi_metrics_monthly,
        "kpi_week_ranges": kpi_week_ranges,
        "heatmaps": heatmaps,
        "regional_moves": regional_moves,
        "regional_flows": regional_flows_out,
        "coupon_summary": coupon_summary,
        "fuel_pl": fuel_pl,
        "summary": summary,
        "goals_2026": goals,
        "latest_goal": latest_goal_row,
    }


def load_kakao(db_path: str):
    if not db_path or not os.path.exists(db_path):
        return None

    conn = sqlite3.connect(db_path)
    df = pd.read_sql_query(
        """
        SELECT 운행시작일, 운행종료일, 예약번호, 출발스테이션, 도착스테이션, 운행거리,
               이용요금, 시간초과요금, "할인(시동OFF)" as 할인시동오프,
               "패널티(지역이탈반납)" as 패널티지역이탈, "패널티(기타)" as 패널티기타
        FROM rentals_KM
        """,
        conn,
    )
    conn.close()

    # 왕복/편도 구분: 출발-도착 스테이션 동일 여부로 재계산 (원본 서비스구분 컬럼은 참고만)
    df["구분"] = df.apply(
        lambda r: "왕복" if r["출발스테이션"] == r["도착스테이션"] else "편도", axis=1
    )
    df["운행시작일"] = pd.to_datetime(df["운행시작일"], format="mixed", errors="coerce")
    df["운행종료일"] = pd.to_datetime(df["운행종료일"], format="mixed", errors="coerce")
    df = df.dropna(subset=["운행시작일"])
    df["date"] = df["운행시작일"].dt.date
    for col in ["이용요금", "시간초과요금", "할인시동오프", "패널티지역이탈", "패널티기타", "운행거리"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    # 매출 = (이용요금 + 시간초과요금 + 할인(시동OFF) + 패널티(지역이탈반납) + 패널티(기타)) / 1.1
    df["매출"] = (
        df["이용요금"] + df["시간초과요금"] + df["할인시동오프"] + df["패널티지역이탈"] + df["패널티기타"]
    ) / 1.1

    # 예약관제 DB에서 연락처(회원 식별용) 조인 - 카카오(KM) 예약건도 이 DB에 같이 들어있음
    if os.path.exists(RESERVATION_CONTROL_DB_PATH):
        try:
            rconn = sqlite3.connect(RESERVATION_CONTROL_DB_PATH)
            res = pd.read_sql_query(
                'SELECT "예약 번호" as 예약번호, 연락처 FROM rentals_예약관제_리턴프리', rconn
            )
            rconn.close()
            res = res.drop_duplicates(subset="예약번호")
            df = df.merge(res, on="예약번호", how="left")
        except Exception as e:
            print(f"[경고] 예약관제 DB 조인 실패: {e}")
            df["연락처"] = None
    else:
        df["연락처"] = None

    # 이용시간(분), 심야(21시~04시59분)/주간 구분
    df["이용분"] = (df["운행종료일"] - df["운행시작일"]).dt.total_seconds() / 60
    hour = df["운행시작일"].dt.hour
    df["hourtype"] = ((hour >= 21) | (hour < 5)).map({True: "심야", False: "주간"})

    # ---------------------------------------------------------------
    # (기존) 일별 왕복/편도 매출 트렌드
    # ---------------------------------------------------------------
    daily = (
        df.groupby(["date", "구분"])
        .agg(건수=("매출", "count"), 매출=("매출", "sum"))
        .reset_index()
    )
    dates = sorted(df["date"].unique())
    result = {"labels": [str(d) for d in dates], "roundtrip": [], "oneway": [], "roundtrip_cnt": [], "oneway_cnt": []}
    for d in dates:
        rt = daily[(daily["date"] == d) & (daily["구분"] == "왕복")]
        ow = daily[(daily["date"] == d) & (daily["구분"] == "편도")]
        result["roundtrip"].append(round(rt["매출"].sum()) if len(rt) else 0)
        result["oneway"].append(round(ow["매출"].sum()) if len(ow) else 0)
        result["roundtrip_cnt"].append(int(rt["건수"].sum()) if len(rt) else 0)
        result["oneway_cnt"].append(int(ow["건수"].sum()) if len(ow) else 0)

    result["summary"] = {
        "total_cnt": int(len(df)),
        "roundtrip_cnt": int((df["구분"] == "왕복").sum()),
        "oneway_cnt": int((df["구분"] == "편도").sum()),
        "total_rev": round(df["매출"].sum()),
    }

    # ---------------------------------------------------------------
    # (신규) 리턴프리와 동일한 방식의 실적 요약 카드용 kpi_metrics (주차/월)
    # ---------------------------------------------------------------
    weekly_years = sorted(df["운행시작일"].dt.isocalendar().year.unique().tolist())
    cmp_years = weekly_years[-2:] if len(weekly_years) >= 2 else weekly_years
    cur_year = cmp_years[-1]
    prev_year = cmp_years[0] if len(cmp_years) > 1 else None

    iso = df["운행시작일"].dt.isocalendar()
    df["orig_year"] = iso["year"].astype(int)
    df["cmp_date"] = df["운행시작일"].dt.normalize()
    if prev_year is not None:
        _pm = df["orig_year"] == prev_year
        df.loc[_pm, "cmp_date"] = df.loc[_pm, "cmp_date"] + pd.Timedelta(days=364)
    df["iso_week"] = df["cmp_date"].dt.isocalendar()["week"].astype(int)
    df["iso_year"] = df["orig_year"]
    df["ym_year"] = df["운행시작일"].dt.year
    df["ym_month"] = df["운행시작일"].dt.month

    def series_by_year_week(frame, col):
        out = {}
        for y in cmp_years:
            rows = frame[frame["iso_year"] == y].sort_values("iso_week")
            vals = rows[col].tolist()
            out[str(y)] = {"weeks": rows["iso_week"].tolist(), "values": [round(v, 1) if isinstance(v, float) else v for v in vals]}
        return out

    def series_by_year_month(frame, col):
        out = {}
        for y in sorted(frame["ym_year"].unique().tolist()):
            rows = frame[frame["ym_year"] == y].sort_values("ym_month")
            vals = rows[col].tolist()
            out[str(y)] = {"weeks": rows["ym_month"].tolist(), "values": [round(v, 1) if isinstance(v, float) else v for v in vals]}
        return out

    def build_kpi(period_cols, series_fn):
        g_cnt_rev = df.groupby(period_cols).agg(건수=("매출", "count"), 매출=("매출", "sum")).reset_index()
        users_g = df.groupby(period_cols)["연락처"].nunique().reset_index(name="이용자수_raw")
        g_cnt_rev = g_cnt_rev.merge(users_g, on=period_cols, how="left")
        g_cnt_rev["인당매출"] = (g_cnt_rev["매출"] / g_cnt_rev["이용자수_raw"]).round(0)
        g_cnt_rev["건당매출"] = (g_cnt_rev["매출"] / g_cnt_rev["건수"]).round(0)

        day_g = df[df["hourtype"] == "주간"].groupby(period_cols).size().reset_index(name="건수")
        night_g = df[df["hourtype"] == "심야"].groupby(period_cols).size().reset_index(name="건수")

        dur_df = df.dropna(subset=["이용분"])
        misc_g = (
            dur_df.groupby(period_cols)
            .agg(이용분합=("이용분", "sum"), 거리합=("운행거리", "sum"), 건수=("매출", "count"), 매출=("매출", "sum"))
            .reset_index()
        )
        misc_g["평균이용시간"] = (misc_g["이용분합"] / misc_g["건수"]).round(0)
        misc_g["평균이동거리"] = (misc_g["거리합"] / misc_g["건수"]).round(1)
        misc_g["분당매출"] = (misc_g["매출"] / misc_g["이용분합"]).round(0)

        return {
            "건수": {"unit": "건", "by_year": series_fn(g_cnt_rev, "건수")},
            "주간건수": {"unit": "건", "by_year": series_fn(day_g, "건수")},
            "심야건수": {"unit": "건", "by_year": series_fn(night_g, "건수")},
            "총매출": {"unit": "매출", "by_year": series_fn(g_cnt_rev, "매출")},
            "이용자수": {"unit": "명", "by_year": series_fn(g_cnt_rev, "이용자수_raw")},
            "건당매출": {"unit": "매출", "by_year": series_fn(g_cnt_rev, "건당매출")},
            "인당매출": {"unit": "매출", "by_year": series_fn(g_cnt_rev, "인당매출")},
            "평균이용시간": {"unit": "분", "by_year": series_fn(misc_g, "평균이용시간")},
            "평균이동거리": {"unit": "km", "by_year": series_fn(misc_g, "평균이동거리")},
            "분당매출": {"unit": "매출", "by_year": series_fn(misc_g, "분당매출")},
        }

    result["kpi_metrics"] = build_kpi(["iso_year", "iso_week"], series_by_year_week)
    result["kpi_metrics_monthly"] = build_kpi(["ym_year", "ym_month"], series_by_year_month)

    kpi_week_ranges = {}
    for w in sorted(set(df[df["iso_year"] == cur_year]["iso_week"])):
        start = datetime.fromisocalendar(int(cur_year), int(w), 1).date()
        end = start + timedelta(days=6)
        kpi_week_ranges[int(w)] = f"{start.isoformat()} ~ {end.isoformat()}"
    result["kpi_week_ranges"] = kpi_week_ranges

    return result


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    data = {}

    rf = load_returnfree(RETURNFREE_DB_PATH)
    if rf is not None:
        data["returnfree"] = rf
        print(f"[리턴프리] {rf['summary']['total_cnt']:,}건 집계 완료 (최신월 {rf['summary']['latest_ym']})")
    else:
        print("[리턴프리] 데이터 없음 - 이전 data.js 유지 또는 스킵")

    kk = load_kakao(KAKAO_DB_PATH)
    if kk is not None:
        data["kakao"] = kk
        print(f"[카카오] {kk['summary']['total_cnt']:,}건 집계 완료")
    else:
        print("[카카오] 파일 없음 - 스킵 (KAKAO_XLSX_PATH 환경변수로 지정 가능)")

    data["generated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    data_js = "const DASHBOARD_DATA = " + json.dumps(data, ensure_ascii=False) + ";\n"

    with open(OUTPUT_JS, "w", encoding="utf-8") as f:
        f.write(data_js)

    # index.html은 assets 폴더 없이도 단독으로 열리도록 chart.js와 데이터를 인라인으로 삽입
    if os.path.exists(TEMPLATE_HTML) and os.path.exists(TEMPLATE_CHARTJS):
        with open(TEMPLATE_HTML, encoding="utf-8") as f:
            html = f.read()
        with open(TEMPLATE_CHARTJS, encoding="utf-8") as f:
            chartjs = f.read()

        html = html.replace(
            '<script src="assets/chart.umd.min.js"></script>', f"<script>{chartjs}</script>"
        )
        html = html.replace('<script src="assets/data.js"></script>', f"<script>{data_js}</script>")

        with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"완료: {OUTPUT_HTML} (단독 실행 가능)")
    else:
        print(f"완료: {OUTPUT_JS} (템플릿 파일이 없어 index.html은 갱신하지 않음)")


if __name__ == "__main__":
    main()
