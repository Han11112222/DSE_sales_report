import io
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import matplotlib as mpl
import plotly.express as px
import streamlit as st

# ─────────────────────────────────────────────────────────
# 기본 설정
# ─────────────────────────────────────────────────────────
def set_korean_font():
    ttf = Path(__file__).parent / "NanumGothic-Regular.ttf"
    if ttf.exists():
        try:
            mpl.font_manager.fontManager.addfont(str(ttf))
            mpl.rcParams["font.family"] = "NanumGothic"
            mpl.rcParams["axes.unicode_minus"] = False
        except Exception:
            pass

set_korean_font()
st.set_page_config(page_title="판매량 현황 분석", layout="wide")

DEFAULT_SALES_XLSX = "판매량(계획_실적).xlsx"

# 엑셀 헤더 → 분석 그룹 매핑 (요청하신 대로 '기타'로 통합)
USE_COL_TO_GROUP: Dict[str, str] = {
    "취사용": "가정용",
    "개별난방용": "가정용",
    "중앙난방용": "가정용",
    "자가열전용": "가정용",
    
    "일반용": "영업용",
    
    "업무난방용": "업무용",
    "냉방용": "업무용",
    "주한미군": "업무용",
    
    "산업용": "산업용",
    
    "수송용(CNG)": "기타",
    "수송용(BIO)": "기타",
    
    "열병합용": "기타",
    "열병합용1": "기타",
    "열병합용2": "기타",
    
    "연료전지용": "기타",
    "열전용설비용": "기타",
}

GROUP_OPTIONS: List[str] = [
    "총량", "가정용", "영업용", "업무용", "산업용", "기타"
]

# ─────────────────────────────────────────────────────────
# 공통 데이터 유틸
# ─────────────────────────────────────────────────────────
def center_style(styler):
    """모든 표 숫자 가운데 정렬용 공통 스타일."""
    styler = styler.set_properties(**{"text-align": "center"})
    styler = styler.set_table_styles([dict(selector="th", props=[("text-align", "center")])])
    return styler

def _clean_base(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "Unnamed: 0" in out.columns:
        out = out.drop(columns=["Unnamed: 0"])
    out["연"] = pd.to_numeric(out["연"], errors="coerce").astype("Int64")
    out["월"] = pd.to_numeric(out["월"], errors="coerce").astype("Int64")
    return out

def keyword_group(col: str) -> Optional[str]:
    c = str(col)
    if any(k in c for k in ["수송용", "열병합", "연료전지", "열전용"]): return "기타"
    if c in ["산업용"]: return "산업용"
    if c in ["일반용"]: return "영업용"
    if any(k in c for k in ["취사용", "난방용", "자가열"]): return "가정용"
    if any(k in c for k in ["업무", "냉방", "주한미군"]): return "업무용"
    return None

def make_long(plan_df: pd.DataFrame, actual_df: pd.DataFrame) -> pd.DataFrame:
    plan_df = _clean_base(plan_df)
    actual_df = _clean_base(actual_df)

    records = []
    for label, df in [("계획", plan_df), ("실적", actual_df)]:
        for col in df.columns:
            if col in ["연", "월"]: continue

            group = USE_COL_TO_GROUP.get(col)
            if group is None: group = keyword_group(col)
            if group is None: continue

            base = df[["연", "월"]].copy()
            base["그룹"] = group
            base["용도"] = col
            base["계획/실적"] = label
            base["값"] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
            records.append(base)

    if not records:
        return pd.DataFrame(columns=["연", "월", "그룹", "용도", "계획/실적", "값"])

    long_df = pd.concat(records, ignore_index=True)
    long_df = long_df.dropna(subset=["연", "월"])
    long_df["연"] = long_df["연"].astype(int)
    long_df["월"] = long_df["월"].astype(int)
    
    # ★ 2022년 ~ 2026년 데이터만 필터링
    long_df = long_df[long_df["연"].isin([2022, 2023, 2024, 2025, 2026])]
    
    return long_df

def load_all_sheets(excel_bytes: bytes) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(io.BytesIO(excel_bytes), engine="openpyxl")
    needed = ["계획_부피", "실적_부피", "계획_열량", "실적_열량"]
    out: Dict[str, pd.DataFrame] = {}
    for name in needed:
        if name in xls.sheet_names:
            out[name] = xls.parse(name)
    return out

def build_long_dict(sheets: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    long_dict: Dict[str, pd.DataFrame] = {}
    if ("계획_부피" in sheets) and ("실적_부피" in sheets):
        long_dict["부피"] = make_long(sheets["계획_부피"], sheets["실적_부피"])
    if ("계획_열량" in sheets) and ("실적_열량" in sheets):
        long_dict["열량"] = make_long(sheets["계획_열량"], sheets["실적_열량"])
    return long_dict

# ─────────────────────────────────────────────────────────
# 기존 화면 복구 : 월별 추이 그래프 (사진과 동일한 UI)
# ─────────────────────────────────────────────────────────
def monthly_trend_section(long_df: pd.DataFrame, unit_label: str, key_prefix: str = ""):
    if long_df.empty:
        st.info("2022~2026년 데이터가 없습니다.")
        return

    years_all = sorted(long_df["연"].unique().tolist())
    default_year = years_all[-1] if years_all else 2026

    # 1. 상단 기준 선택기 (사진과 동일한 구성)
    c1, c2, c3 = st.columns([1.2, 1.2, 1.6])
    with c1:
        sel_year = st.selectbox("기준 연도", options=years_all, index=years_all.index(default_year), key=f"{key_prefix}year")
    with c2:
        sel_month = st.selectbox("기준 월", options=list(range(1, 13)), index=11, key=f"{key_prefix}month") # 12월 기본
    with c3:
        st.markdown("<div style='padding-top:28px;font-size:14px;color:#666;'>집계 기준: <b>연 누적</b></div>", unsafe_allow_html=True)
    
    st.markdown(f"<div style='margin-top:-4px;font-size:13px;color:#666;'>선택 기준: <b>{sel_year}년 {sel_month}월 · 연 누적</b></div>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    # 2. 연도 다중 선택 (그래프용)
    default_plot_years = [y for y in [2024, 2025, 2026] if y in years_all]
    sel_years = st.multiselect("연도 선택(그래프)", options=years_all, default=default_plot_years, key=f"{key_prefix}years")

    # 3. 그룹 선택 (버튼형)
    try:
        sel_group = st.segmented_control("그룹 선택", GROUP_OPTIONS, selection_mode="single", default="영업용", key=f"{key_prefix}group")
    except Exception:
        sel_group = st.radio("그룹 선택", GROUP_OPTIONS, index=GROUP_OPTIONS.index("영업용"), horizontal=True, key=f"{key_prefix}group_radio")

    if not sel_years:
        st.info("표시할 연도를 한 개 이상 선택해 주세요.")
        return

    # 데이터 필터링
    base = long_df[long_df["연"].isin(sel_years)].copy()
    base = base[base["월"] <= sel_month]

    if sel_group != "총량":
        base = base[base["그룹"] == sel_group]

    plot_df = base.groupby(["연", "월", "계획/실적"], as_index=False)["값"].sum().sort_values(["연", "월", "계획/실적"])

    if plot_df.empty:
        st.info("선택 조건에 해당하는 데이터가 없습니다.")
        return

    # 그래프 라벨 생성 (사진의 범례와 일치하도록)
    group_label = sel_group if sel_group != "총량" else "총량"
    plot_df["라벨"] = plot_df["연"].astype(str) + f"년 · {group_label}"

    # 4. 라인 그래프
    fig = px.line(
        plot_df,
        x="월",
        y="값",
        color="라벨",
        line_dash="계획/실적",
        category_orders={"계획/실적": ["실적", "계획"]},
        line_dash_map={"실적": "solid", "계획": "dash"},
        markers=True,
    )
    fig.update_layout(
        xaxis=dict(dtick=1, title="월"),
        yaxis=dict(title=f"판매량 ({unit_label})"),
        margin=dict(l=10, r=10, t=60, b=10),
        legend=dict(title="연도 / 구분", orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hovermode="x unified"
    )
    st.plotly_chart(fig, use_container_width=True)

    # 5. 하단 데이터 표
    st.markdown("##### 🔢 월별 상세 데이터표")
    plot_df["표_컬럼"] = plot_df["연"].astype(str) + "년 " + plot_df["계획/실적"]
    table = plot_df.pivot_table(index="월", columns="표_컬럼", values="값", aggfunc="sum").sort_index().fillna(0.0)
    
    total_row = table.sum(numeric_only=True)
    table.loc["합계"] = total_row
    table = table.reset_index()
    
    numeric_cols = [c for c in table.columns if c != "월"]
    styled = center_style(table.style.format({col: "{:,.0f}" for col in numeric_cols}))
    st.dataframe(styled, use_container_width=True, hide_index=True)


# ─────────────────────────────────────────────────────────
# 메인 앱 구동부
# ─────────────────────────────────────────────────────────
def main():
    st.title("📊 판매량 현황 분석")

    with st.sidebar:
        st.header("📂 데이터 불러오기")
        src = st.radio("데이터 소스", ["레포 파일 사용", "엑셀 업로드(.xlsx)"], index=0, key="sales_src")
        excel_bytes = None
        base_info = ""
        
        if src == "엑셀 업로드(.xlsx)":
            up = st.file_uploader("판매량(계획_실적).xlsx 형식", type=["xlsx"], key="sales_uploader")
            if up is not None:
                excel_bytes = up.getvalue()
                base_info = f"소스: 업로드 파일 — {up.name}"
        else:
            path = Path(__file__).parent / DEFAULT_SALES_XLSX
            if path.exists():
                excel_bytes = path.read_bytes()
                base_info = f"소스: 레포 파일 — {DEFAULT_SALES_XLSX}"
            else:
                base_info = f"레포 경로에 {DEFAULT_SALES_XLSX} 파일이 없습니다. 엑셀을 업로드해주세요."
        st.caption(base_info)

    if excel_bytes is not None:
        sheets = load_all_sheets(excel_bytes)
        long_dict = build_long_dict(sheets)

        tab_labels: List[str] = []
        if "부피" in long_dict:
            tab_labels.append("부피 기준 (천m³)")
        if "열량" in long_dict:
            tab_labels.append("열량 기준 (GJ)")

        if not tab_labels:
            st.info("유효한 시트를 찾지 못했습니다. 엑셀 파일 내 시트명을 확인해 주세요. ('계획_부피', '실적_부피' 등)")
        else:
            tabs = st.tabs(tab_labels)
            for tab_label, tab in zip(tab_labels, tabs):
                with tab:
                    if tab_label.startswith("부피"):
                        df_long = long_dict.get("부피", pd.DataFrame())
                        unit = "천m³"
                        prefix = "vol_"  # 부피 탭용 고유 키
                    else:
                        df_long = long_dict.get("열량", pd.DataFrame()).copy()
                        unit = "GJ"
                        prefix = "gj_"   # 열량 탭용 고유 키

                    # 기존 그래프 섹션 호출
                    monthly_trend_section(df_long, unit_label=unit, key_prefix=prefix)

if __name__ == "__main__":
    main()
