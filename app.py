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
st.set_page_config(page_title="DSE 판매량 분석 보고", layout="wide")

DEFAULT_SALES_XLSX = "판매량(계획_실적).xlsx"

# 요청하신 그룹 순서
GROUP_ORDER = ["가정용", "산업용", "업무용", "영업용", "기타"]

# 막대그래프 및 꺾은선(실적)용 연도별 색상 팔레트 (진한 푸른색, 회색, 파란색)
BAR_PALETTE = ["#1f497d", "#808080", "#4292c6"]

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
def center_style(styler):
    styler = styler.set_properties(**{"text-align": "center"})
    styler = styler.set_table_styles([dict(selector="th", props=[("text-align", "center")])])
    return styler

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
            if col not in USE_COL_TO_GROUP: continue
            group = USE_COL_TO_GROUP[col]
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
# 그래프 섹션: 연간 추이 그래프 (꺾은선 + 막대 + 데이터 박스)
# ─────────────────────────────────────────────────────────
def render_monthly_trend(df, unit, prefix):
    st.markdown("### 📈 연간 추이 그래프")
    
    # 우측 텍스트 삭제, 가로 비율 유지를 위해 빈 컬럼 배치
    c1, c2 = st.columns([3, 1])
    with c1: 
        sel_years = st.multiselect("연도 선택(그래프)", options=[2022, 2023, 2024, 2025, 2026], default=[2024, 2025, 2026], key=f"{prefix}my")

    try:
        sel_group = st.segmented_control("그룹 선택", options=["총량"] + GROUP_ORDER, selection_mode="single", default="총량", key=f"{prefix}sg")
    except:
        sel_group = st.radio("그룹 선택", options=["총량"] + GROUP_ORDER, index=0, horizontal=True, key=f"{prefix}rd")

    if not sel_years:
        st.info("연도를 하나 이상 선택해주세요.")
        return

    plot_df = df[df["그룹"] == sel_group] if sel_group != "총량" else df
    
    fig_line = go.Figure()
    fig_bar = go.Figure()

    table_data_list = []
    line_y_vals = []

    for i, year in enumerate(sorted(sel_years)):
        c = BAR_PALETTE[i % len(BAR_PALETTE)]
        
        if year == 2026:
            # 2026년 1~12월 계획 데이터
            y26_plan = plot_df[(plot_df["연"] == 2026) & (plot_df["계획/실적"] == "계획")].groupby("월")["값"].sum().reset_index()
            # 2026년 1~3월 실적 데이터
            y26_act = plot_df[(plot_df["연"] == 2026) & (plot_df["계획/실적"] == "실적") & (plot_df["월"] <= 3)].groupby("월")["값"].sum().reset_index()
            
            # 1. 2026년 계획 (검정색 실선 및 검정색 막대)
            if not y26_plan.empty:
                fig_line.add_trace(go.Scatter(x=y26_plan["월"], y=y26_plan["값"], mode='lines+markers', 
                                         name="2026년 계획", line=dict(color='black', width=2.5)))
                line_y_vals.extend(y26_plan["값"].tolist())
                
                y26_plan_tb = y26_plan.copy()
                y26_plan_tb["표_컬럼"] = "2026년 계획"
                table_data_list.append(y26_plan_tb)
                
                # 계획 막대는 보조 지표이므로 연한 회색/검정으로 배치할 수 있으나 검정으로 통일
                fig_bar.add_trace(go.Bar(x=y26_plan["월"], y=y26_plan["값"], name="2026년 계획", marker_color='black'))
                
            # 2. 2026년 실적 (파란색 점선 및 파란색 막대)
            if not y26_act.empty:
                fig_line.add_trace(go.Scatter(x=y26_act["월"], y=y26_act["값"], mode='lines+markers', 
                                         name="2026년 실적", line=dict(color=c, width=2.5, dash='dot')))
                line_y_vals.extend(y26_act["값"].tolist())
                
                y26_act_tb = y26_act.copy()
                y26_act_tb["표_컬럼"] = "2026년 실적"
                table_data_list.append(y26_act_tb)
                
                fig_bar.add_trace(go.Bar(x=y26_act["월"], y=y26_act["값"], name="2026년 실적", marker_color=c))

        else:
            # 과거 연도 실적 (지정된 팔레트의 실선 및 막대)
            y_act = plot_df[(plot_df["연"] == year) & (plot_df["계획/실적"] == "실적")]
            y_act_grp = y_act.groupby("월")["값"].sum().reset_index()

            if not y_act_grp.empty:
                fig_line.add_trace(go.Scatter(x=y_act_grp["월"], y=y_act_grp["값"], mode='lines+markers', 
                                         name=f"{year}년 실적", line=dict(color=c, width=2.5)))
                line_y_vals.extend(y_act_grp["값"].tolist())

                y_act_tb = y_act_grp.copy()
                y_act_tb["표_컬럼"] = f"{year}년 실적"
                table_data_list.append(y_act_tb)

                fig_bar.add_trace(go.Bar(x=y_act_grp["월"], y=y_act_grp["값"], name=f"{year}년", marker_color=c))

    # Y축 하단 여백 스케일링 최적화 유지 (위아래 갭 축소)
    if line_y_vals:
        min_y = min(line_y_vals)
        max_y = max(line_y_vals)
        y_min_scaled = min_y * 0.7 if min_y > 0 else min_y * 1.1
        y_max_scaled = max_y * 1.05
        fig_line.update_layout(
            height=550, 
            xaxis=dict(dtick=1, title="월"), 
            yaxis=dict(title=f"판매량({unit})", range=[y_min_scaled, y_max_scaled]), 
            hovermode="x unified", 
            legend=dict(orientation="h", y=1.1)
        )
    else:
        fig_line.update_layout(height=550, xaxis=dict(dtick=1, title="월"), yaxis=dict(title=f"판매량({unit})"), hovermode="x unified", legend=dict(orientation="h", y=1.1))
        
    st.plotly_chart(fig_line, use_container_width=True)

    st.markdown(f"##### 📊 {sel_group} 연도별 동월 비교 (막대그래프)")
    fig_bar.update_layout(barmode='group', xaxis=dict(dtick=1, title="월"), yaxis=dict(title=f"판매량({unit})"), hovermode="x unified", legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig_bar, use_container_width=True)

    # 3. 하단 데이터 박스 (4~12월 증감량 공란 처리)
    st.markdown("##### 🔢 월별 상세 데이터표")
    if table_data_list:
        t_df = pd.concat(table_data_list, ignore_index=True)
        table = t_df.pivot_table(index="월", columns="표_컬럼", values="값", aggfunc="sum").sort_index().fillna(0.0)
        
        if "2026년 계획" in table.columns and "2026년 실적" in table.columns:
            table["증감량(차이)"] = table["2026년 실적"] - table["2026년 계획"]
            
            # 4월~12월 공란(NaN) 처리
            table.loc[table.index > 3, "증감량(차이)"] = np.nan
            
            table["증감률(%)"] = np.nan
            valid_mask = (table.index <= 3) & (table["2026년 계획"] != 0)
            table.loc[valid_mask, "증감률(%)"] = (table.loc[valid_mask, "증감량(차이)"] / table.loc[valid_mask, "2026년 계획"]) * 100

        total_row = table.sum(numeric_only=True)
        table.loc["합계"] = total_row
        
        # 합계 증감률 재계산
        if "2026년 계획" in table.columns and "2026년 실적" in table.columns:
            val_diff = table.loc["합계", "증감량(차이)"]
            val_plan = table.loc["합계", "2026년 계획"]
            table.loc["합계", "증감률(%)"] = (val_diff / val_plan * 100) if val_plan != 0 else np.nan

        table = table.reset_index()

        numeric_cols = [col for col in table.columns if col not in ["월", "증감률(%)"]]
        format_dict = {col: "{:,.0f}" for col in numeric_cols}
        if "증감률(%)" in table.columns:
            format_dict["증감률(%)"] = "{:,.1f}%"

        # na_rep="-"를 사용하여 NaN 값(4~12월)을 하이픈(-)으로 표시
        styled = center_style(table.style.format(format_dict, na_rep="-"))
        st.dataframe(styled, use_container_width=True, hide_index=True)

# ─────────────────────────────────────────────────────────
# 메인 실행
# ─────────────────────────────────────────────────────────
def main():
    st.title("📊 DSE 판매량 분석 보고")

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
    else:
        st.warning("데이터 파일을 로드할 수 없습니다.")

if __name__ == "__main__":
    main()
