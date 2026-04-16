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
# 타이틀 변경 요청 반영
st.set_page_config(page_title="DSE 판매량 분석 보고", layout="wide")

DEFAULT_SALES_XLSX = "판매량(계획_실적).xlsx"
MODIFIED_SALES_XLSX = "판매량(계획_실적)_수정.xlsx"

# 요청하신 그룹 순서
GROUP_ORDER = ["가정용", "산업용", "업무용", "영업용", "기타"]

# 차분하고 안정감 있는 색상으로 스택 컬러 변경 (요청 반영)
COLOR_MAP = {
    "가정용": "#4c72b0",  # 차분한 진파랑
    "산업용": "#9ebcda",  # 차분한 하늘색
    "업무용": "#e07a5f",  # 차분한 핑크/다홍
    "영업용": "#e6c253",  # 차분한 샌드 옐로우
    "기타": "#8cce8b"     # 차분한 연두색
}

# 막대그래프용 연도별 푸른색 계열 색상 팔레트 (최근일수록 더욱 진해지도록 구성)
BAR_PALETTE = ["#6baed6", "#4292c6", "#2171b5", "#08519c", "#08306b"]

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
# 그래프 섹션 1: 연간 추이 그래프 (꺾은선 + 막대 + 데이터 박스)
# ─────────────────────────────────────────────────────────
def render_monthly_trend(df, mod_df, unit, prefix):
    st.markdown("### 📈 연간 추이 그래프")
    
    c1, c2 = st.columns([3, 1])
    with c1: 
        sel_years = st.multiselect("연도 선택(그래프)", options=[2022, 2023, 2024, 2025, 2026], default=[2024, 2025, 2026], key=f"{prefix}my")
    with c2: 
        st.markdown("<div style='padding-top:28px;font-size:14px;color:#666;'>집계 기준: <b>단월 판매량</b></div>", unsafe_allow_html=True)

    try:
        sel_group = st.segmented_control("그룹 선택", options=["총량"] + GROUP_ORDER, selection_mode="single", default="총량", key=f"{prefix}sg")
    except:
        sel_group = st.radio("그룹 선택", options=["총량"] + GROUP_ORDER, index=0, horizontal=True, key=f"{prefix}rd")

    mod_toggle = False
    if mod_df is not None:
        mod_toggle = st.toggle("🚀 산업용 물량 수정 (2026년 4~12월 변경사항 표기)", value=False, key=f"{prefix}mod_toggle")

    if not sel_years:
        st.info("연도를 하나 이상 선택해주세요.")
        return mod_toggle

    plot_df = df[df["그룹"] == sel_group] if sel_group != "총량" else df
    
    fig_line = go.Figure()
    fig_bar = go.Figure()

    table_data_list = []

    for i, year in enumerate(sorted(sel_years)):
        c = BAR_PALETTE[i % len(BAR_PALETTE)]
        
        y_act = plot_df[(plot_df["연"] == year) & (plot_df["계획/실적"] == "실적")]
        
        if year == 2026:
            y_act_sub = y_act[y_act["월"] <= 3]
            y_act_grp = y_act_sub.groupby("월")["값"].sum().reset_index()
        else:
            y_act_grp = y_act.groupby("월")["값"].sum().reset_index()

        if not y_act_grp.empty:
            fig_line.add_trace(go.Scatter(x=y_act_grp["월"], y=y_act_grp["값"], mode='lines+markers', 
                                     name=f"{year}년 실적", line=dict(color=c, width=2.5)))

        combined_year_data = y_act_grp.copy()
        combined_year_data["구분"] = "실적"

        if year == 2026:
            y26_plan_only = plot_df[(plot_df["연"] == 2026) & (plot_df["계획/실적"] == "계획") & (plot_df["월"] >= 4)].groupby("월")["값"].sum().reset_index()
            
            if not y26_plan_only.empty:
                y26_act_m3 = y_act_grp[y_act_grp["월"] == 3].copy()
                
                if not y26_act_m3.empty:
                    y26_plan_line = pd.concat([y26_act_m3, y26_plan_only], ignore_index=True)
                else:
                    y26_plan_line = y26_plan_only
                    
                fig_line.add_trace(go.Scatter(x=y26_plan_line["월"], y=y26_plan_line["값"], mode='lines+markers', 
                                         name="2026년 계획(4~12월)", line=dict(color='black', width=2.5, dash='dot')))
                
                y26_plan_only["구분"] = "계획"
                combined_year_data = pd.concat([combined_year_data, y26_plan_only], ignore_index=True)

        if not combined_year_data.empty:
            fig_bar.add_trace(go.Bar(x=combined_year_data["월"], y=combined_year_data["값"], name=f"{year}년", marker_color=c))
            
            combined_year_data["표_컬럼"] = str(year) + "년 " + combined_year_data["구분"]
            combined_year_data["연"] = year
            table_data_list.append(combined_year_data)

        # 2026년 수정 물량 (토글 ON)
        if year == 2026 and mod_toggle and mod_df is not None:
            c_mod = "#e11d48"
            plot_mod_df = mod_df[mod_df["그룹"] == sel_group] if sel_group != "총량" else mod_df
            
            y_act_mod = plot_mod_df[(plot_mod_df["연"] == 2026) & (plot_mod_df["계획/실적"] == "실적") & (plot_mod_df["월"] <= 3)].groupby("월")["값"].sum().reset_index()
            y_plan_mod = plot_mod_df[(plot_mod_df["연"] == 2026) & (plot_mod_df["계획/실적"] == "계획") & (plot_mod_df["월"] >= 4)].groupby("월")["값"].sum().reset_index()
            
            if not y_plan_mod.empty:
                y26_act_m3_mod = y_act_mod[y_act_mod["월"] == 3].copy()
                if not y26_act_m3_mod.empty:
                    y26_plan_line_mod = pd.concat([y26_act_m3_mod, y_plan_mod], ignore_index=True)
                else:
                    y26_plan_line_mod = y_plan_mod
                    
                fig_line.add_trace(go.Scatter(x=y26_plan_line_mod["월"], y=y26_plan_line_mod["값"], mode='lines+markers', 
                                         name="2026년 변경 계획(4~12월)", line=dict(color=c_mod, width=2.5, dash='dash')))
                
            combined_mod_data = y_act_mod.copy()
            if not combined_mod_data.empty:
                combined_mod_data["구분"] = "실적"
            if not y_plan_mod.empty:
                y_plan_mod_copy = y_plan_mod.copy()
                y_plan_mod_copy["구분"] = "계획(변경)"
                combined_mod_data = pd.concat([combined_mod_data, y_plan_mod_copy], ignore_index=True)
                
            if not combined_mod_data.empty:
                fig_bar.add_trace(go.Bar(x=combined_mod_data["월"], y=combined_mod_data["값"], name="2026년 변경", marker_color=c_mod))
                combined_mod_data["표_컬럼"] = "2026년 변경 " + combined_mod_data["구분"]
                combined_mod_data["연"] = "2026년 변경"
                table_data_list.append(combined_mod_data)

    fig_line.update_layout(height=550, xaxis=dict(dtick=1, title="월"), yaxis=dict(title=f"판매량({unit})"), hovermode="x unified", legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig_line, use_container_width=True)

    st.markdown(f"##### 📊 {sel_group} 연도별 동월 비교 (막대그래프)")
    fig_bar.update_layout(barmode='group', xaxis=dict(dtick=1, title="월"), yaxis=dict(title=f"판매량({unit})"), hovermode="x unified", legend=dict(orientation="h", y=1.1))
    st.plotly_chart(fig_bar, use_container_width=True)

    # 3. 하단 데이터 박스 (차이/증감률 열 추가)
    st.markdown("##### 🔢 월별 상세 데이터표")
    if table_data_list:
        t_df = pd.concat(table_data_list, ignore_index=True)
        table = t_df.pivot_table(index="월", columns="표_컬럼", values="값", aggfunc="sum").sort_index().fillna(0.0)
        
        # 수정 조건 시 차이 및 증감률 컬럼 계산
        if mod_toggle and sel_group == "산업용":
            if "2026년 계획" in table.columns and "2026년 변경 계획(변경)" in table.columns:
                table["증감량(차이)"] = table["2026년 변경 계획(변경)"] - table["2026년 계획"]
                table["증감률(%)"] = np.where(table["2026년 계획"] != 0, (table["증감량(차이)"] / table["2026년 계획"]) * 100, 0.0)

        total_row = table.sum(numeric_only=True)
        table.loc["합계"] = total_row
        
        # 합계 행 증감률 재계산
        if mod_toggle and sel_group == "산업용":
            if "2026년 계획" in table.columns and "2026년 변경 계획(변경)" in table.columns:
                val_diff = table.loc["합계", "증감량(차이)"]
                val_plan = table.loc["합계", "2026년 계획"]
                table.loc["합계", "증감률(%)"] = (val_diff / val_plan * 100) if val_plan != 0 else 0.0

        table = table.reset_index()

        # 데이터 포맷팅 분리 적용
        numeric_cols = [col for col in table.columns if col not in ["월", "증감률(%)"]]
        format_dict = {col: "{:,.0f}" for col in numeric_cols}
        if "증감률(%)" in table.columns:
            format_dict["증감률(%)"] = "{:,.1f}%"

        styled = center_style(table.style.format(format_dict))
        st.dataframe(styled, use_container_width=True, hide_index=True)

    return mod_toggle

# ─────────────────────────────────────────────────────────
# 그래프 섹션 2: 연간 용도별 실적 판매량 누적
# ─────────────────────────────────────────────────────────
def render_stacked_chart(df, mod_df, unit, prefix, mod_toggle):
    st.markdown("---")
    st.markdown("### 🧱 연간 용도별 실적 판매량 누적")
    
    c1, c2 = st.columns([2, 2])
    with c1: plot_years = st.multiselect("연도 선택(스택 그래프)", options=[2022, 2023, 2024, 2025, 2026], default=[2024, 2025, 2026], key=f"{prefix}stk_y")
    with c2: period = st.radio("기간", ["연간", "상반기(1~6월)", "하반기(7~12월)"], horizontal=True, key=f"{prefix}stk_p")

    base_df = df[df["연"].isin(plot_years)]
    past_act = base_df[(base_df["연"] < 2026) & (base_df["계획/실적"] == "실적")]
    y26_act = base_df[(base_df["연"] == 2026) & (base_df["계획/실적"] == "실적") & (base_df["월"] <= 3)]
    y26_plan = base_df[(base_df["연"] == 2026) & (base_df["계획/실적"] == "계획") & (base_df["월"] >= 4)]
    
    stack_list = [past_act, y26_act, y26_plan]
    
    if mod_toggle and 2026 in plot_years and mod_df is not None:
        y26_act_mod = mod_df[(mod_df["연"] == 2026) & (mod_df["계획/실적"] == "실적") & (mod_df["월"] <= 3)].copy()
        y26_plan_mod = mod_df[(mod_df["연"] == 2026) & (mod_df["계획/실적"] == "계획") & (mod_df["월"] >= 4)].copy()
        
        y26_act_mod["연"] = "2026년 변경"
        y26_plan_mod["연"] = "2026년 변경"
        
        stack_list.extend([y26_act_mod, y26_plan_mod])
        
    stack_df = pd.concat(stack_list, ignore_index=True)

    if period == "상반기(1~6월)": stack_df = stack_df[stack_df["월"] <= 6]
    elif period == "하반기(7~12월)": stack_df = stack_df[stack_df["월"] > 6]

    stack_df["연_표시"] = stack_df["연"].apply(lambda x: f"{x}년" if isinstance(x, int) else x)
    
    grp_data = stack_df.groupby(["연_표시", "그룹"])["값"].sum().reset_index()
    
    grp_data["그룹"] = pd.Categorical(grp_data["그룹"], categories=GROUP_ORDER, ordered=True)
    
    year_order = [f"{y}년" for y in sorted(plot_years)]
    if mod_toggle and 2026 in plot_years:
        year_order.append("2026년 변경")
        
    grp_data["연_표시"] = pd.Categorical(grp_data["연_표시"], categories=year_order, ordered=True)
    grp_data = grp_data.sort_values(["연_표시", "그룹"])

    fig = px.bar(grp_data, x="연_표시", y="값", color="그룹", barmode="stack",
                 category_orders={"그룹": GROUP_ORDER, "연_표시": year_order},
                 color_discrete_map=COLOR_MAP,
                 text_auto=',.0f')
    
    total_line = grp_data.groupby("연_표시")["값"].sum().reset_index()
    fig.add_trace(go.Scatter(x=total_line["연_표시"], y=total_line["값"], mode='lines+markers+text', 
                             name="합계", text=total_line["값"].apply(lambda x: f"{x:,.0f}" if x>0 else ""),
                             textposition="top center", line=dict(color="#8085e9", dash="dash", width=2)))
    
    home_line = grp_data[grp_data["그룹"] == "가정용"].groupby("연_표시")["값"].sum().reset_index()
    if not home_line.empty:
        fig.add_trace(go.Scatter(x=home_line["연_표시"], y=home_line["값"], mode='lines+markers', 
                                 name="가정용", line=dict(color="#cccccc", dash="dot", width=2)))

    # 스택 바 가로 굵기 2배로 확대 & 세로 높이 20% 축소 반영
    fig.update_traces(selector=dict(type='bar'), width=0.35)
    fig.update_layout(height=550, xaxis=dict(title="연도"), yaxis=dict(title=f"판매량({unit})"), legend=dict(title="그룹", orientation="v", x=1.02, y=0.8))
    st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────────────────
# 메인 실행
# ─────────────────────────────────────────────────────────
def main():
    # 최상단 타이틀 변경 반영
    st.title("📊 DSE 판매량 분석 보고")

    st.sidebar.header("📂 데이터 설정")
    src = st.sidebar.radio("데이터 소스", ["레포 파일 사용", "엑셀 업로드"])
    excel_bytes = None
    mod_bytes = None
    
    if src == "엑셀 업로드":
        up = st.sidebar.file_uploader("판매량(기본) 엑셀 파일 업로드", type="xlsx")
        if up: excel_bytes = up.getvalue()
        
        up_mod = st.sidebar.file_uploader("판매량(수정) 엑셀 파일 업로드 (선택)", type="xlsx")
        if up_mod: mod_bytes = up_mod.getvalue()
    else:
        p = Path(__file__).parent / DEFAULT_SALES_XLSX
        if p.exists(): excel_bytes = p.read_bytes()
        
        p_mod = Path(__file__).parent / MODIFIED_SALES_XLSX
        if p_mod.exists(): mod_bytes = p_mod.read_bytes()

    if excel_bytes:
        data_dict = load_data(excel_bytes)
        mod_dict = load_data(mod_bytes) if mod_bytes else {}
        
        tabs = st.tabs([f"{k} 기준" for k in data_dict.keys()])
        for (k, df), tab in zip(data_dict.items(), tabs):
            with tab:
                unit = "천m³" if k == "부피" else "GJ"
                mod_df = mod_dict.get(k, None)
                
                mod_toggle = render_monthly_trend(df, mod_df, unit, k)
                render_stacked_chart(df, mod_df, unit, k, mod_toggle)
    else:
        st.warning("데이터 파일을 로드할 수 없습니다.")

if __name__ == "__main__":
    main()
