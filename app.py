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

# 요청하신 그룹 순서
GROUP_ORDER = ["가정용", "산업용", "업무용", "영업용", "기타"]

# 2번째 사진과 동일한 스택 그래프 색상 맵핑 (Highcharts 기본 색상 느낌)
COLOR_MAP = {
    "가정용": "#0b5ed7",  # 진한 파랑
    "산업용": "#7cb5ec",  # 연한 하늘색
    "업무용": "#f15c80",  # 핑크/빨강
    "영업용": "#e4d354",  # 노랑
    "기타": "#90ed7d"     # 연두색
}

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
# 그래프 섹션 1: 월별 추이 (1번째 사진 UI 반영 & 2026년 4~12월 계획량 점선)
# ─────────────────────────────────────────────────────────
def render_monthly_trend(df, unit, prefix):
    st.markdown("### 📈 월별 추이 그래프")
    
    # 지시사항 1, 2: 기준연도/월 삭제 & 연도 다중선택 배치 (디폴트 24~26년)
    c1, c2 = st.columns([3, 1])
    with c1: 
        sel_years = st.multiselect("연도 선택(그래프)", options=[2022, 2023, 2024, 2025, 2026], default=[2024, 2025, 2026], key=f"{prefix}my")
    with c2: 
        st.markdown("<div style='padding-top:28px;font-size:14px;color:#666;'>집계 기준: <b>연 누적 (단월)</b></div>", unsafe_allow_html=True)

    # 지시사항: 기존에 있던 용도 버튼 복구 및 유지
    try:
        sel_group = st.segmented_control("그룹 선택", options=["총량"] + GROUP_ORDER, selection_mode="single", default="영업용", key=f"{prefix}sg")
    except:
        sel_group = st.radio("그룹 선택", options=["총량"] + GROUP_ORDER, index=GROUP_ORDER.index("영업용")+1, horizontal=True, key=f"{prefix}rd")

    if not sel_years:
        st.info("연도를 하나 이상 선택해주세요.")
        return

    plot_df = df[df["그룹"] == sel_group] if sel_group != "총량" else df
    
    fig = go.Figure()
    # 1번째 사진의 연도별 선 색상과 동일하게 구성 (24=파랑, 25=빨강, 26=초록)
    line_colors = {2022: "#9467bd", 2023: "#8c564b", 2024: "#1f77b4", 2025: "#d62728", 2026: "#2ca02c"}

    for year in sorted(sel_years):
        year_str = str(year)
        c = line_colors.get(year, "#1f77b4")
        
        # [실적 부분]
        y_act = plot_df[(plot_df["연"] == year) & (plot_df["계획/실적"] == "실적")]
        
        # 지시사항 4: 2026년은 3월까지만 실적 표시
        if year == 2026:
            y_act = y_act[y_act["월"] <= 3]
            
        y_act_grp = y_act.groupby("월")["값"].sum().reset_index()
        if not y_act_grp.empty:
            fig.add_trace(go.Scatter(x=y_act_grp["월"], y=y_act_grp["값"], mode='lines+markers', 
                                     name=f"{year}년 실적", line=dict(color=c, width=2)))

        # [계획 부분] 2026년 4월~12월 계획량 점선으로 표현
        if year == 2026:
            # 3월 실적의 마지막 점과 자연스럽게 이어지도록 월 >= 3 으로 데이터 가져옴
            y26_plan = plot_df[(plot_df["연"] == 2026) & (plot_df["계획/실적"] == "계획") & (plot_df["월"] >= 3)].groupby("월")["값"].sum().reset_index()
            if not y26_plan.empty:
                fig.add_trace(go.Scatter(x=y26_plan["월"], y=y26_plan["값"], mode='lines+markers', 
                                         name="2026년 계획(4~12월)", line=dict(color=c, width=2, dash='dash')))

    fig.update_layout(xaxis=dict(dtick=1, title="월"), yaxis=dict(title=f"판매량({unit})"), hovermode="x unified", legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────────────────
# 그래프 섹션 2: 기간별 용도 누적 실적 (2번째 사진 UI 및 색상 복구)
# ─────────────────────────────────────────────────────────
def render_stacked_chart(df, unit, prefix):
    st.markdown("---")
    st.markdown("### 🧱 연간 용도별 실적 판매량 (당월)")
    
    c1, c2 = st.columns([2, 2])
    with c1: plot_years = st.multiselect("연도 선택(스택 그래프)", options=[2022, 2023, 2024, 2025, 2026], default=[2024, 2025, 2026], key=f"{prefix}stk_y")
    with c2: period = st.radio("기간", ["연간", "상반기(1~6월)", "하반기(7~12월)"], horizontal=True, key=f"{prefix}stk_p")

    # 실적 데이터만 집계
    stack_df = df[(df["연"].isin(plot_years)) & (df["계획/실적"] == "실적")]
    if period == "상반기(1~6월)": stack_df = stack_df[stack_df["월"] <= 6]
    elif period == "하반기(7~12월)": stack_df = stack_df[stack_df["월"] > 6]

    grp_data = stack_df.groupby(["연", "그룹"])["값"].sum().reset_index()
    
    # 지시사항: 가정용, 산업용, 업무용, 영업용, 기타 순서 강제 적용
    grp_data["그룹"] = pd.Categorical(grp_data["그룹"], categories=GROUP_ORDER, ordered=True)
    grp_data = grp_data.sort_values(["연", "그룹"])

    # 지시사항 3: 스택 그래프 색상을 2번째 사진과 동일하게 적용
    fig = px.bar(grp_data, x="연", y="값", color="그룹", barmode="stack",
                 category_orders={"그룹": GROUP_ORDER},
                 color_discrete_map=COLOR_MAP,
                 text_auto=',.0f')
    
    # 2번째 사진과 동일한 "합계" 점선(보라색) 및 "가정용" 점선(회색) 추가
    total_line = grp_data.groupby("연")["값"].sum().reset_index()
    fig.add_trace(go.Scatter(x=total_line["연"], y=total_line["값"], mode='lines+markers+text', 
                             name="합계", text=total_line["값"].apply(lambda x: f"{x:,.0f}"),
                             textposition="top center", line=dict(color="#8085e9", dash="dash", width=2)))
    
    home_line = grp_data[grp_data["그룹"] == "가정용"].groupby("연")["값"].sum().reset_index()
    if not home_line.empty:
        fig.add_trace(go.Scatter(x=home_line["연"], y=home_line["값"], mode='lines+markers', 
                                 name="가정용", line=dict(color="#cccccc", dash="dot", width=2)))

    # 범례 위치를 사진처럼 우측으로 배치
    fig.update_layout(xaxis=dict(dtick=1), yaxis=dict(title=f"판매량({unit})"), legend=dict(title="그룹", orientation="v", x=1.02, y=0.8))
    st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────────────────
# 메인 실행
# ─────────────────────────────────────────────────────────
def main():
    st.sidebar.header("📂 데이터 설정")
    src = st.sidebar.radio("데이터 소스", ["레포 파일 사용", "엑셀 업로드"])
    excel_bytes = None
    if src == "엑셀 업로드":
        up = st.sidebar.file_uploader("판매량 엑셀 파일 업로드", type="xlsx")
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
        st.warning("데이터 파일을 로드할 수 없습니다.")

if __name__ == "__main__":
    main()
