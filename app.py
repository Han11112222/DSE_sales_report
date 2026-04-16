import io
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import numpy as np
import matplotlib as mpl
import plotly.express as px
import plotly.graph_objects as go
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

# 요청하신 그룹 순서 및 매핑
GROUP_ORDER = ["가정용", "산업용", "업무용", "영업용", "기타"]

USE_COL_TO_GROUP: Dict[str, str] = {
    "취사용": "가정용", "개별난방용": "가정용", "중앙난방용": "가정용", "자가열전용": "가정용",
    "산업용": "산업용",
    "업무난방용": "업무용", "냉방용": "업무용", "주한미군": "업무용",
    "일반용": "영업용",
    "수송용(CNG)": "기타", "수송용(BIO)": "기타", "열병합용": "기타", "열병합용1": "기타", "열병합용2": "기타", "연료전지용": "기타", "열전용설비용": "기타",
}

# ─────────────────────────────────────────────────────────
# 데이터 처리 유틸
# ─────────────────────────────────────────────────────────
def _clean_base(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "Unnamed: 0" in out.columns: out = out.drop(columns=["Unnamed: 0"])
    out["연"] = pd.to_numeric(out["연"], errors="coerce").astype("Int64")
    out["월"] = pd.to_numeric(out["월"], errors="coerce").astype("Int64")
    return out

def make_long(plan_df: pd.DataFrame, actual_df: pd.DataFrame) -> pd.DataFrame:
    plan_df = _clean_base(plan_df)
    actual_df = _clean_base(actual_df)
    records = []
    for label, df in [("계획", plan_df), ("실적", actual_df)]:
        for col in df.columns:
            if col in ["연", "월"]: continue
            group = USE_COL_TO_GROUP.get(col, "기타")
            base = df[["연", "월"]].copy()
            base["그룹"] = group
            base["계획/실적"] = label
            base["값"] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
            records.append(base)
    long_df = pd.concat(records, ignore_index=True).dropna(subset=["연", "월"])
    long_df["연"] = long_df["연"].astype(int)
    long_df["월"] = long_df["월"].astype(int)
    return long_df[long_df["연"].isin(range(2022, 2027))]

def load_data(excel_bytes):
    xls = pd.ExcelFile(io.BytesIO(excel_bytes), engine="openpyxl")
    sheets = {name: xls.parse(name) for name in ["계획_부피", "실적_부피", "계획_열량", "실적_열량"] if name in xls.sheet_names}
    long_dict = {}
    if "계획_부피" in sheets and "실적_부피" in sheets:
        long_dict["부피"] = make_long(sheets["계획_부피"], sheets["실적_부피"])
    if "계획_열량" in sheets and "실적_열량" in sheets:
        long_dict["열량"] = make_long(sheets["계획_열량"], sheets["실적_열량"])
    return long_dict

# ─────────────────────────────────────────────────────────
# 그래프 섹션 1: 월별 추이 (실적 + 26년 미래 계획 점선)
# ─────────────────────────────────────────────────────────
def render_monthly_trend(df, unit, prefix):
    st.markdown("### 📈 월별 판매량 추이 (2022-2026)")
    
    c1, c2, c3 = st.columns([1, 1, 2])
    with c1: sel_year = st.selectbox("기준 연도", options=[2022,2023,2024,2025,2026], index=4, key=f"{prefix}y")
    with c2: sel_month = st.selectbox("실적 마감 월", options=list(range(1, 13)), index=2, key=f"{prefix}m")
    with c3: sel_group = st.selectbox("그룹 선택", options=["총량"] + GROUP_ORDER, index=0, key=f"{prefix}g")

    plot_df = df[df["그룹"] == sel_group] if sel_group != "총량" else df
    
    fig = go.Figure()
    colors = {"2022": "#E2E8F0", "2023": "#CBD5E0", "2024": "#A0AEC0", "2025": "#718096", "2026": "#2B6CB0"}

    for year in range(2022, 2027):
        year_str = str(year)
        # 과거 연도 및 26년 현재까지의 실적
        year_act = plot_df[(plot_df["연"] == year) & (plot_df["계획/실적"] == "실적")]
        if year == 2026: year_act = year_act[year_act["월"] <= sel_month]
        
        year_act = year_act.groupby("월")["값"].sum().reset_index()
        if not year_act.empty:
            fig.add_trace(go.Scatter(x=year_act["월"], y=year_act["값"], mode='lines+markers', 
                                     name=f"{year}년 실적", line=dict(color=colors[year_str], width=2 if year < 2026 else 4)))

        # 2026년 미래 구간 계획 (점선)
        if year == 2026:
            y26_plan = plot_df[(plot_df["연"] == 2026) & (plot_df["계획/실적"] == "계획") & (plot_df["월"] >= sel_month)].groupby("월")["값"].sum().reset_index()
            if not y26_plan.empty:
                fig.add_trace(go.Scatter(x=y26_plan["월"], y=y26_plan["값"], mode='lines+markers', 
                                         name="2026년 계획(예상)", line=dict(color=colors["2026"], width=3, dash='dot')))

    fig.update_layout(xaxis=dict(dtick=1, title="월"), yaxis=dict(title=f"판매량({unit})"), hovermode="x unified", legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────────────────
# 그래프 섹션 2: 기간별 용도 누적 실적 (스택 그래프)
# ─────────────────────────────────────────────────────────
def render_stacked_chart(df, unit, prefix):
    st.markdown("---")
    st.markdown("### 🧱 기간별 용도 누적 실적 (스택)")
    
    c1, c2 = st.columns([2, 2])
    with c1: plot_years = st.multiselect("연도 선택", options=range(2022, 2027), default=list(range(2022, 2027)), key=f"{prefix}sy")
    with c2: period = st.radio("기간 선택", ["연간", "상반기(1~6월)", "하반기(7~12월)"], horizontal=True, key=f"{prefix}p")

    # 데이터 필터링 (실적만 사용)
    stack_df = df[(df["연"].isin(plot_years)) & (df["계획/실적"] == "실적")]
    if period == "상반기(1~6월)": stack_df = stack_df[stack_df["월"] <= 6]
    elif period == "하반기(7~12월)": stack_df = stack_df[stack_df["월"] > 6]

    grp_data = stack_df.groupby(["연", "그룹"])["값"].sum().reset_index()
    
    # 그룹 순서 강제 적용
    grp_data["그룹"] = pd.Categorical(grp_data["그룹"], categories=GROUP_ORDER, ordered=True)
    grp_data = grp_data.sort_values(["연", "그룹"])

    fig = px.bar(grp_data, x="연", y="값", color="그룹", barmode="stack",
                 category_orders={"그룹": GROUP_ORDER},
                 color_discrete_sequence=px.colors.qualitative.Pastel)
    
    # 합계 라인 추가
    total_line = grp_data.groupby("연")["값"].sum().reset_index()
    fig.add_trace(go.Scatter(x=total_line["연"], y=total_line["값"], mode='lines+markers+text', 
                             name="합계", text=total_line["값"].apply(lambda x: f"{x:,.0f}"),
                             textposition="top center", line=dict(color="black", dash="dash")))

    fig.update_layout(xaxis=dict(dtick=1), yaxis=dict(title=f"판매량({unit})"), legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────────────────
# 메인 실행
# ─────────────────────────────────────────────────────────
def main():
    st.sidebar.header("📂 데이터 설정")
    src = st.sidebar.radio("소스", ["레포 파일", "직접 업로드"])
    excel_bytes = None
    if src == "직접 업로드":
        up = st.sidebar.file_uploader("XLSX 업로드", type="xlsx")
        if up: excel_bytes = up.getvalue()
    else:
        p = Path(__file__).parent / DEFAULT_SALES_XLSX
        if p.exists(): excel_bytes = p.read_bytes()

    if excel_bytes:
        data_dict = load_data(excel_bytes)
        tabs = st.tabs([f"{k} 기준" for k in data_dict.keys()])
        for (k, df), tab in zip(data_dict.items(), tabs):
            with tab:
                unit = "천m³" if k == "부피" else "GJ"
                render_monthly_trend(df, unit, k)
                render_stacked_chart(df, unit, k)
    else:
        st.warning("데이터 파일을 업로드하거나 레포지토리에 추가해주세요.")

if __name__ == "__main__":
    main()
