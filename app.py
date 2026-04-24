import io
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import numpy as np
import matplotlib as mpl
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components  # PDF 자동 인쇄 창을 띄우기 위한 라이브러리 추가

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

# 산업용과 업무용 색상 스왑 완료
COLOR_MAP = {
    "가정용": "#4c72b0",
    "산업용": "#e07a5f",  
    "업무용": "#9ebcda",  
    "영업용": "#e6c253",
    "기타": "#8cce8b"
}

# 고정 색상 맵핑
LINE_COLOR_MAP = {
    "2023년 실적": "#9467bd",  # 보라색 (2022년 회색과 구분)
    "2024년 실적": "#1f77b4",  # 파란색
    "2025년 실적": "#2ca02c",  # 녹색
    "2026년 실적": "#d62728",  # 빨간색
    "2026년 계획": "#ff7f0e"   # 오렌지색
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
    
    if "계획_열량" in sheets and "실적_열량" in sheets:
        long_dict["열량"] = make_long(sheets["계획_열량"], sheets["실적_열량"])
        
    if "계획_부피" in sheets and "실적_부피" in sheets:
        long_dict["부피"] = make_long(sheets["계획_부피"], sheets["실적_부피"])
        
    return long_dict

# ─────────────────────────────────────────────────────────
# 그래프 섹션: 연간 추이 그래프 (꺾은선 + 막대 + 데이터 박스)
# ─────────────────────────────────────────────────────────
def render_monthly_trend(df, unit, prefix):
    st.markdown("### 📈 연간 추이 그래프")
    
    c1, c2 = st.columns([3, 1])
    with c1: 
        sel_years = st.multiselect("연도 선택(그래프)", options=[2022, 2023, 2024, 2025, 2026], default=[2022, 2023, 2024, 2025, 2026], key=f"{prefix}my")

    try:
        sel_group = st.segmented_control("그룹 선택", options=["전체"] + GROUP_ORDER, selection_mode="single", default="전체", key=f"{prefix}sg")
    except:
        sel_group = st.radio("그룹 선택", options=["전체"] + GROUP_ORDER, index=0, horizontal=True, key=f"{prefix}rd")

    if not sel_years:
        st.info("연도를 하나 이상 선택해주세요.")
        return

    plot_df = df[df["그룹"] == sel_group] if sel_group != "전체" else df
    
    fig_line = go.Figure()
    fig_bar = go.Figure()

    table_data_list = []
    line_y_vals = []

    for year in sorted(sel_years):
        if year == 2026:
            y26_plan = plot_df[(plot_df["연"] == 2026) & (plot_df["계획/실적"] == "계획")].groupby("월")["값"].sum().reset_index()
            y26_act = plot_df[(plot_df["연"] == 2026) & (plot_df["계획/실적"] == "실적") & (plot_df["월"] <= 3)].groupby("월")["값"].sum().reset_index()
            
            if not y26_plan.empty:
                c_plan = LINE_COLOR_MAP["2026년 계획"]
                fig_line.add_trace(go.Scatter(x=y26_plan["월"], y=y26_plan["값"], mode='markers+lines', name="2026년 계획", line=dict(color=c_plan, width=2.5, dash='dot')))
                line_y_vals.extend(y26_plan["값"].tolist())
                
                y26_plan_tb = y26_plan.copy()
                y26_plan_tb["표_컬럼"] = "2026년 계획"
                table_data_list.append(y26_plan_tb)
                
                fig_bar.add_trace(go.Bar(x=y26_plan["월"], y=y26_plan["값"], name="2026년 계획", marker_color=c_plan))
                
            if not y26_act.empty:
                c_act26 = LINE_COLOR_MAP["2026년 실적"]
                fig_line.add_trace(go.Scatter(x=y26_act["월"], y=y26_act["값"], mode='markers+lines', name="2026년 실적", line=dict(color=c_act26, width=2.5)))
                line_y_vals.extend(y26_act["값"].tolist())
                
                y26_act_tb = y26_act.copy()
                y26_act_tb["표_컬럼"] = "2026년 실적"
                table_data_list.append(y26_act_tb)
                
                fig_bar.add_trace(go.Bar(x=y26_act["월"], y=y26_act["값"], name="2026년 실적", marker_color=c_act26))

        else:
            y_act = plot_df[(plot_df["연"] == year) & (plot_df["계획/실적"] == "실적")]
            y_act_grp = y_act.groupby("월")["값"].sum().reset_index()

            if not y_act_grp.empty:
                key_name = f"{year}년 실적"
                c = LINE_COLOR_MAP.get(key_name, "#808080")
                
                fig_line.add_trace(go.Scatter(x=y_act_grp["월"], y=y_act_grp["값"], mode='markers+lines', name=key_name, line=dict(color=c, width=2.5)))
                line_y_vals.extend(y_act_grp["값"].tolist())

                y_act_tb = y_act_grp.copy()
                y_act_tb["표_컬럼"] = key_name
                table_data_list.append(y_act_tb)

                fig_bar.add_trace(go.Bar(x=y_act_grp["월"], y=y_act_grp["값"], name=f"{year}년", marker_color=c))

    unit_anno = dict(
        xref="paper", yref="paper", 
        x=1.0, y=1.12, 
        xanchor="right", yanchor="bottom", 
        text=f"(단위: {unit})", 
        font=dict(size=13, color="#555"), 
        showarrow=False
    )

    if line_y_vals:
        min_y = min(line_y_vals)
        max_y = max(line_y_vals)
        y_min_scaled = min_y * 0.95 if min_y > 0 else min_y * 1.05
        y_max_scaled = max_y * 1.05
        fig_line.update_layout(
            height=550, 
            xaxis=dict(dtick=1, title="월"), 
            yaxis=dict(title=f"판매량({unit})", range=[y_min_scaled, y_max_scaled], tickformat=",.0f", hoverformat=",.0f"), 
            hovermode="x unified", 
            legend=dict(orientation="h", y=1.1),
            annotations=[unit_anno]
        )
    else:
        fig_line.update_layout(
            height=550, 
            xaxis=dict(dtick=1, title="월"), 
            yaxis=dict(title=f"판매량({unit})", tickformat=",.0f", hoverformat=",.0f"), 
            hovermode="x unified", 
            legend=dict(orientation="h", y=1.1),
            annotations=[unit_anno]
        )
        
    st.plotly_chart(fig_line, use_container_width=True, key=f"{prefix}_main_fig_line")

    st.markdown(f"##### 📊 {sel_group} 연도별 동월 비교 (막대그래프)")
    fig_bar.update_layout(
        barmode='group',
        bargap=0.36,
        xaxis=dict(dtick=1, title="월"), 
        yaxis=dict(title=f"판매량({unit})", tickformat=",.0f", hoverformat=",.0f"), 
        hovermode="x unified", 
        legend=dict(orientation="h", y=1.1),
        annotations=[unit_anno]
    )
    st.plotly_chart(fig_bar, use_container_width=True, key=f"{prefix}_main_fig_bar")

    st.markdown(f"##### 📈 {sel_group} 연간 구성비 추이 그래프")
    
    total_monthly = df.groupby(["연", "월", "계획/실적"])["값"].sum().reset_index(name="총합")
    ratio_line_df = plot_df.groupby(["연", "월", "계획/실적"])["값"].sum().reset_index()
    ratio_line_df = pd.merge(ratio_line_df, total_monthly, on=["연", "월", "계획/실적"])
    ratio_line_df["비중"] = np.where(ratio_line_df["총합"] > 0, (ratio_line_df["값"] / ratio_line_df["총합"]) * 100, 0)

    fig_ratio_line = go.Figure()

    for year in sorted(sel_years):
        if year == 2026:
            y26_plan_r = ratio_line_df[(ratio_line_df["연"] == 2026) & (ratio_line_df["계획/실적"] == "계획")]
            if not y26_plan_r.empty:
                c_plan = LINE_COLOR_MAP["2026년 계획"]
                fig_ratio_line.add_trace(go.Scatter(x=y26_plan_r["월"], y=y26_plan_r["비중"], mode='markers+lines', name="2026년 계획", line=dict(color=c_plan, width=2.5, dash='dot')))
            
            y26_act_r = ratio_line_df[(ratio_line_df["연"] == 2026) & (ratio_line_df["계획/실적"] == "실적") & (ratio_line_df["월"] <= 3)]
            if not y26_act_r.empty:
                c_act26 = LINE_COLOR_MAP["2026년 실적"]
                fig_ratio_line.add_trace(go.Scatter(x=y26_act_r["월"], y=y26_act_r["비중"], mode='markers+lines', name="2026년 실적", line=dict(color=c_act26, width=2.5)))
        else:
            y_act_r = ratio_line_df[(ratio_line_df["연"] == year) & (ratio_line_df["계획/실적"] == "실적")]
            if not y_act_r.empty:
                key_name = f"{year}년 실적"
                c = LINE_COLOR_MAP.get(key_name, "#808080")
                fig_ratio_line.add_trace(go.Scatter(x=y_act_r["월"], y=y_act_r["비중"], mode='markers+lines', name=key_name, line=dict(color=c, width=2.5)))

    if not ratio_line_df["비중"].empty:
        r_min, r_max = ratio_line_df["비중"].min(), ratio_line_df["비중"].max()
        y_range = [max(0, r_min * 0.9), min(100, r_max * 1.1)]
    else:
        y_range = [0, 105]

    fig_ratio_line.update_layout(
        height=450, 
        xaxis=dict(dtick=1, title="월"), 
        yaxis=dict(title="구성비 (%)", range=y_range, tickformat=".1f", ticksuffix="%"), 
        hovermode="x unified", 
        legend=dict(orientation="h", y=1.1)
    )
    st.plotly_chart(fig_ratio_line, use_container_width=True, key=f"{prefix}_main_fig_ratio_line")


    # ─────────────────────────────────────────────────────────
    # 타임 시리즈 그래프
    # ─────────────────────────────────────────────────────────
    st.markdown("##### 📈 전체 용도별 구성비 추이 (타임시리즈)")
    
    ts_col1, ts_col2 = st.columns([1, 3])
    with ts_col1:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        show_ts_ratio = st.toggle("구성비 표기", value=False, key=f"{prefix}_ts_ratio_toggle")
    with ts_col2:
        ts_years = st.multiselect(
            "타임 시리즈 연도 선택 (별도)", 
            options=[2022, 2023, 2024, 2025, 2026], 
            default=[2023, 2024, 2025, 2026], 
            key=f"{prefix}_ts_years"
        )
        
    if ts_years:
        ts_df = df[(df["연"].isin(ts_years)) & (df["계획/실적"] == "실적")].copy()
        
        if not ts_df.empty:
            fig_ts = go.Figure()
            
            ts_df["년월"] = ts_df["연"].astype(str) + "." + ts_df["월"].astype(str).str.zfill(2)
            ts_grp_df = ts_df.groupby(["년월", "그룹"])["값"].sum().reset_index()
            
            ts_pivot = ts_grp_df.pivot(index="년월", columns="그룹", values="값").fillna(0)
            ts_ratio = ts_pivot.div(ts_pivot.sum(axis=1), axis=0).fillna(0) * 100
            
            x_numeric = np.arange(len(ts_ratio.index))
            all_categories = list(ts_ratio.index)
            
            for grp in GROUP_ORDER:
                if grp in ts_ratio.columns:
                    mode_str = 'lines+text' if show_ts_ratio else 'lines'
                    text_arr = []
                    if show_ts_ratio:
                        for m_str, v in zip(ts_ratio.index, ts_ratio[grp]):
                            month_val = int(m_str.split('.')[1])
                            if month_val in [3, 6, 9, 12] and v >= 1.0:
                                text_arr.append(f"{v:.1f}%")
                            else:
                                text_arr.append("")
                    else:
                        text_arr = None
                    
                    fig_ts.add_trace(go.Scatter(
                        x=x_numeric, y=ts_ratio[grp], mode=mode_str, name=grp,
                        line=dict(color=COLOR_MAP.get(grp, "#000"), width=1.5, shape='spline'),
                        stackgroup='one',
                        fillcolor=COLOR_MAP.get(grp, "#000"),
                        text=text_arr,
                        textposition='bottom center', 
                        textfont=dict(size=18, color="white") 
                    ))
            
            if show_ts_ratio:
                for i, m_str in enumerate(ts_ratio.index):
                    if int(m_str.split('.')[1]) in [3, 6, 9, 12]:
                        fig_ts.add_trace(go.Scatter(
                            x=[i, i], y=[0, 100], mode="lines",
                            line=dict(color="rgba(100, 100, 100, 0.7)", width=1.5, dash="dash"),
                            showlegend=False, hoverinfo="skip"
                        ))
            
            tickvals = list(range(len(all_categories)))
            ticktext = list(all_categories)
            range_end = len(all_categories) - 0.5
            
            if 2026 in ts_years and "2026.04" not in all_categories:
                tickvals.append(len(all_categories))
                ticktext.append("2026.04")
                range_end = len(all_categories) + 0.5
                
            fig_ts.update_layout(
                xaxis=dict(
                    title="년월 (YYYY.MM)", 
                    tickmode='array',
                    tickvals=tickvals,
                    ticktext=ticktext,
                    range=[-0.5, range_end],
                    tickangle=-45
                )
            )
                
            fig_ts.update_layout(
                height=540,
                yaxis=dict(title="구성비 (%)", range=[0, 100], tickformat=".0f", ticksuffix="%"),
                hovermode="x unified",
                legend=dict(orientation="h", y=1.1)
            )
            st.plotly_chart(fig_ts, use_container_width=True, key=f"{prefix}_ts_main_chart")
        else:
            st.info("선택한 연도의 실적 데이터가 없습니다.")
    # ─────────────────────────────────────────────────────────


    # ─────────────────────────────────────────────────────────
    # 연간 용도별 구성비 (스택그래프)
    # ─────────────────────────────────────────────────────────
    st.markdown(f"##### 📊 연간 용도별 구성비 (스택그래프)")
    fig_stack = go.Figure()
    
    x_labels = []
    annual_totals = {}
    for year in sorted(sel_years):
        if year == 2026:
            x_labels.append("2026년 실적")
            annual_totals["2026년 실적"] = df[(df["연"] == 2026) & (df["계획/실적"] == "실적") & (df["월"] <= 3)]["값"].sum()
        else:
            label = f"{year}년 실적"
            x_labels.append(label)
            annual_totals[label] = df[(df["연"] == year) & (df["계획/실적"] == "실적")]["값"].sum()
            
    ratios_dict = {grp: [] for grp in GROUP_ORDER}
    for grp in GROUP_ORDER:
        for label in x_labels:
            if label == "2026년 실적":
                val = df[(df["그룹"] == grp) & (df["연"] == 2026) & (df["계획/실적"] == "실적") & (df["월"] <= 3)]["값"].sum()
            else:
                y_int = int(label[:4])
                val = df[(df["그룹"] == grp) & (df["연"] == y_int) & (df["계획/실적"] == "실적")]["값"].sum()
            
            tot = annual_totals.get(label, 0)
            ratio = (val / tot * 100) if tot > 0 else 0
            ratios_dict[grp].append(ratio)
            
        fig_stack.add_trace(go.Bar(
            x=x_labels, 
            y=ratios_dict[grp], 
            name=grp, 
            marker_color=COLOR_MAP.get(grp, "#808080")
        ))

    stack_annotations = []
    for i, label in enumerate(x_labels):
        cum_y = 0
        for grp in GROUP_ORDER:
            val = ratios_dict[grp][i]
            if val >= 1.0:
                mid_y = cum_y + val / 2
                stack_annotations.append(dict(
                    x=label, y=mid_y, xref='x', yref='y',
                    text=f"{val:.1f}%", xanchor='center', yanchor='middle',
                    showarrow=False, font=dict(size=18, color="white")
                ))
            cum_y += val

    fig_stack.update_layout(
        annotations=stack_annotations,
        barmode='stack',
        bargap=0.4,
        xaxis=dict(title="연도 및 구분"),
        yaxis=dict(title="구성비 (%)", range=[0, 100], tickformat=".0f", ticksuffix="%"),
        hovermode="x unified",
        legend=dict(orientation="h", y=1.1)
    )
    st.plotly_chart(fig_stack, use_container_width=True, key=f"{prefix}_main_fig_stack")
    # ─────────────────────────────────────────────────────────

    c_tbl_1, c_tbl_2 = st.columns([3, 1])
    with c_tbl_1:
        st.markdown("##### 🔢 월별 상세 데이터표")
    with c_tbl_2:
        st.markdown(f"<div style='text-align: right; font-size: 13px; color: #555;'><b>(단위: {unit})</b></div>", unsafe_allow_html=True)
        
    if table_data_list:
        t_df = pd.concat(table_data_list, ignore_index=True)
        table = t_df.pivot_table(index="월", columns="표_컬럼", values="값", aggfunc="sum").sort_index().fillna(0.0)
        
        if "2026년 계획" in table.columns and "2026년 실적" in table.columns:
            table["증감량(차이)"] = table["2026년 실적"] - table["2026년 계획"]
            table.loc[table.index > 3, "증감량(차이)"] = np.nan
            table["증감률(%)"] = np.nan
            valid_mask = (table.index <= 3) & (table["2026년 계획"] != 0)
            table.loc[valid_mask, "증감률(%)"] = (table.loc[valid_mask, "증감량(차이)"] / table.loc[valid_mask, "2026년 계획"]) * 100
            
            ytd_plan_sum = table.loc[table.index <= 3, "2026년 계획"].sum()

        total_row = table.sum(numeric_only=True)
        table.loc["합계"] = total_row
        
        if "2026년 계획" in table.columns and "2026년 실적" in table.columns:
            val_diff = table.loc["합계", "증감량(차이)"]
            table.loc["합계", "증감률(%)"] = (val_diff / ytd_plan_sum * 100) if ytd_plan_sum != 0 else np.nan

        table = table.reset_index()
        numeric_cols = [col for col in table.columns if col not in ["월", "증감률(%)"]]
        format_dict = {col: "{:,.0f}" for col in numeric_cols}
        if "증감률(%)" in table.columns:
            format_dict["증감률(%)"] = "{:,.1f}%"

        styled_df = table.style.format(format_dict, na_rep="-")
        styled_df = styled_df.apply(lambda row: ['background-color: #1f497d; color: white;' if row['월'] == '합계' else '' for _ in row], axis=1)
        
        styled = center_style(styled_df)
        st.dataframe(styled, use_container_width=True, hide_index=True)

    # ─────────────────────────────────────────────────────────
    # 보고 일괄 출력 뷰어
    # ─────────────────────────────────────────────────────────
    st.divider()
    st.markdown("### 🖨️ 보고서 일괄 출력 뷰어")
    st.caption("항목과 그룹을 체크하고 버튼을 누르면 선택한 내용만 인쇄용 미리보기 화면에 나열됩니다.")

    st.markdown("##### 1. 출력 항목 선택")
    chk_col1, chk_col2, chk_col3, chk_col4 = st.columns(4)
    with chk_col1:
        prt_line = st.checkbox("연간추이 그래프", value=True, key=f"{prefix}_prt_line")
    with chk_col2:
        prt_bar = st.checkbox("연도별 동월 비교 그래프", value=True, key=f"{prefix}_prt_bar")
    with chk_col3:
        prt_ratio = st.checkbox("전체 연간 구성비 추이 그래프", value=True, key=f"{prefix}_prt_ratio")
    with chk_col4:
        prt_tbl = st.checkbox("월별 상세 데이터표", value=True, key=f"{prefix}_prt_tbl")

    st.markdown("##### 2. 출력 그룹 선택")
    grp_cols = st.columns(6)
    selected_groups = []
    all_grp_names = ["전체"] + GROUP_ORDER
    
    for i, g_name in enumerate(all_grp_names):
        with grp_cols[i]:
            if st.checkbox(g_name, value=True, key=f"{prefix}_prt_grp_{i}"):
                selected_groups.append(g_name)

    if st.button("미리보기", key=f"{prefix}_preview_btn", type="primary"):
        if not selected_groups:
            st.warning("출력할 그룹을 최소 1개 이상 선택해주세요.")
        else:
            st.markdown("<div id='preview-marker' style='display:none;'></div>", unsafe_allow_html=True)
            st.markdown(
                """
                <style>
                div[data-testid="stTable"] table {
                    table-layout: auto !important;
                    width: 100% !important;
                }
                div[data-testid="stTable"] table th,
                div[data-testid="stTable"] table td {
                    height: 35px !important;
                    padding: 0px 8px !important;
                    line-height: 35px !important;
                    vertical-align: middle !important;
                    white-space: nowrap !important;
                    overflow: hidden !important;
                    font-size: 110% !important;
                }
                div[data-testid="stTable"] table tr {
                    height: 35px !important;
                }
                div[data-testid="stTable"] table th:first-child,
                div[data-testid="stTable"] table td:first-child {
                    width: 60px !important;
                    min-width: 60px !important;
                    max-width: 60px !important;
                }
                </style>
                """, unsafe_allow_html=True
            )
            st.markdown("---")
            
            for print_grp in selected_groups:
                st.markdown(f"<div class='print-page-container' style='width: 100%; display: flex; flex-direction: column; align-items: center; page-break-inside: avoid; break-inside: avoid;'>", unsafe_allow_html=True)
                st.markdown(f"<h2 style='text-align: center; color: #1f497d; margin-top: 10px;'>[{print_grp}] 판매량 분석 보고</h2>", unsafe_allow_html=True)

                p_df = df[df["그룹"] == print_grp] if print_grp != "전체" else df

                p_fig_line = go.Figure()
                p_fig_bar = go.Figure()
                p_table_list = []
                p_line_vals = []

                for year in sorted(sel_years):
                    if year == 2026:
                        y26_plan = p_df[(p_df["연"] == 2026) & (p_df["계획/실적"] == "계획")].groupby("월")["값"].sum().reset_index()
                        y26_act = p_df[(p_df["연"] == 2026) & (p_df["계획/실적"] == "실적") & (p_df["월"] <= 3)].groupby("월")["값"].sum().reset_index()

                        if not y26_plan.empty:
                            c_plan = LINE_COLOR_MAP["2026년 계획"]
                            p_fig_line.add_trace(go.Scatter(x=y26_plan["월"], y=y26_plan["값"], mode='markers+lines', name="2026년 계획", line=dict(color=c_plan, width=2.5, dash='dot')))
                            p_line_vals.extend(y26_plan["값"].tolist())
                            y26_plan_tb = y26_plan.copy()
                            y26_plan_tb["표_컬럼"] = "2026년 계획"
                            p_table_list.append(y26_plan_tb)
                            p_fig_bar.add_trace(go.Bar(x=y26_plan["월"], y=y26_plan["값"], name="2026년 계획", marker_color=c_plan))

                        if not y26_act.empty:
                            c_act26 = LINE_COLOR_MAP["2026년 실적"]
                            p_fig_line.add_trace(go.Scatter(x=y26_act["월"], y=y26_act["값"], mode='markers+lines', name="2026년 실적", line=dict(color=c_act26, width=2.5)))
                            p_line_vals.extend(y26_act["값"].tolist())
                            y26_act_tb = y26_act.copy()
                            y26_act_tb["표_컬럼"] = "2026년 실적"
                            p_table_list.append(y26_act_tb)
                            p_fig_bar.add_trace(go.Bar(x=y26_act["월"], y=y26_act["값"], name="2026년 실적", marker_color=c_act26))

                    else:
                        y_act = p_df[(p_df["연"] == year) & (p_df["계획/실적"] == "실적")]
                        y_act_grp = y_act.groupby("월")["값"].sum().reset_index()

                        if not y_act_grp.empty:
                            key_name = f"{year}년 실적"
                            c = LINE_COLOR_MAP.get(key_name, "#808080")
                            p_fig_line.add_trace(go.Scatter(x=y_act_grp["월"], y=y_act_grp["값"], mode='markers+lines', name=key_name, line=dict(color=c, width=2.5)))
                            p_line_vals.extend(y_act_grp["값"].tolist())
                            y_act_tb = y_act_grp.copy()
                            y_act_tb["표_컬럼"] = key_name
                            p_table_list.append(y_act_tb)
                            p_fig_bar.add_trace(go.Bar(x=y_act_grp["월"], y=y_act_grp["값"], name=f"{year}년", marker_color=c))

                shared_y_args = {}
                if p_line_vals:
                    min_y, max_y = min(p_line_vals), max(p_line_vals)
                    y_min_s = min_y * 0.95 if min_y > 0 else min_y * 1.05
                    y_max_s = max_y * 1.05
                    shared_y_args = {"range": [y_min_s, y_max_s]}

                table_ready = False
                styled = None
                if prt_tbl and p_table_list:
                    t_df = pd.concat(p_table_list, ignore_index=True)
                    p_table = t_df.pivot_table(index="월", columns="표_컬럼", values="값", aggfunc="sum").sort_index().fillna(0.0)

                    if "2026년 계획" in p_table.columns and "2026년 실적" in p_table.columns:
                        p_table["증감량(차이)"] = p_table["2026년 실적"] - p_table["2026년 계획"]
                        p_table.loc[p_table.index > 3, "증감량(차이)"] = np.nan
                        p_table["증감률(%)"] = np.nan
                        valid_mask = (p_table.index <= 3) & (p_table["2026년 계획"] != 0)
                        p_table.loc[valid_mask, "증감률(%)"] = (p_table.loc[valid_mask, "증감량(차이)"] / p_table.loc[valid_mask, "2026년 계획"]) * 100
                        ytd_p_sum = p_table.loc[p_table.index <= 3, "2026년 계획"].sum()

                    total_row = p_table.sum(numeric_only=True)
                    p_table.loc["합계"] = total_row

                    if "2026년 계획" in p_table.columns and "2026년 실적" in p_table.columns:
                        val_diff = p_table.loc["합계", "증감량(차이)"]
                        p_table.loc["합계", "증감률(%)"] = (val_diff / ytd_p_sum * 100) if ytd_p_sum != 0 else np.nan

                    p_table = p_table.reset_index()
                    numeric_cols = [col for col in p_table.columns if col not in ["월", "증감률(%)"]]
                    format_dict = {col: "{:,.0f}" for col in numeric_cols}
                    if "증감률(%)" in p_table.columns:
                        format_dict["증감률(%)"] = "{:,.1f}%"

                    styled_df = p_table.style.format(format_dict, na_rep="")
                    styled_df = styled_df.apply(lambda row: ['background-color: #1f497d; color: white;' if row['월'] == '합계' else '' for _ in row], axis=1)
                    
                    styled = center_style(styled_df)
                    try:
                        styled = styled.hide(axis="index")
                    except:
                        styled = styled.hide_index()
                    table_ready = True

                if prt_line and prt_bar:
                    col_left, col_right = st.columns(2)
                    with col_left:
                        p_fig_line.update_layout(height=450, xaxis=dict(dtick=1, title="월"), yaxis=dict(title=f"판매량({unit})", tickformat=",.0f", **shared_y_args), hovermode="x unified", legend=dict(orientation="h", y=1.1, x=0.5, xanchor='center'), annotations=[unit_anno])
                        st.markdown(f"<div style='text-align: center;'><b>■ [{print_grp}] 연간 추이 그래프</b></div>", unsafe_allow_html=True)
                        st.plotly_chart(p_fig_line, use_container_width=True, key=f"prt_line_chart_{prefix}_{print_grp}")
                    with col_right:
                        p_fig_bar.update_layout(barmode='group', bargap=0.36, height=450, xaxis=dict(dtick=1, title="월"), yaxis=dict(title=f"판매량({unit})", tickformat=",.0f", **shared_y_args), hovermode="x unified", legend=dict(orientation="h", y=1.1, x=0.5, xanchor='center'), annotations=[unit_anno])
                        st.markdown(f"<div style='text-align: center;'><b>■ [{print_grp}] 연도별 동월 비교 그래프</b></div>", unsafe_allow_html=True)
                        st.plotly_chart(p_fig_bar, use_container_width=True, key=f"prt_bar_chart_{prefix}_{print_grp}")

                elif prt_line or prt_bar:
                    col_left, col_right = st.columns([1.8, 1])
                    with col_left:
                        if prt_line:
                            p_fig_line.update_layout(height=550, margin=dict(l=10, r=0, t=40, b=10), xaxis=dict(dtick=1, title="월"), yaxis=dict(title=f"판매량({unit})", tickformat=",.0f", **shared_y_args), hovermode="x unified", legend=dict(orientation="h", y=1.1, x=0.5, xanchor='center'), annotations=[unit_anno])
                            st.markdown(f"<div style='text-align: center;'><b>■ [{print_grp}] 연간 추이 그래프</b></div>", unsafe_allow_html=True)
                            st.plotly_chart(p_fig_line, use_container_width=True, key=f"prt_line_single_side_{prefix}_{print_grp}")
                        elif prt_bar:
                            p_fig_bar.update_layout(barmode='group', bargap=0.36, height=550, margin=dict(l=10, r=0, t=40, b=10), xaxis=dict(dtick=1, title="월"), yaxis=dict(title=f"판매량({unit})", tickformat=",.0f", **shared_y_args), hovermode="x unified", legend=dict(orientation="h", y=1.1, x=0.5, xanchor='center'), annotations=[unit_anno])
                            st.markdown(f"<div style='text-align: center;'><b>■ [{print_grp}] 연도별 동월 비교 그래프</b></div>", unsafe_allow_html=True)
                            st.plotly_chart(p_fig_bar, use_container_width=True, key=f"prt_bar_single_side_{prefix}_{print_grp}")

                if prt_ratio:
                    p_ratio_line_df = p_df.groupby(["연", "월", "계획/실적"])["값"].sum().reset_index()
                    p_ratio_line_df = pd.merge(p_ratio_line_df, total_monthly, on=["연", "월", "계획/실적"])
                    p_ratio_line_df["비중"] = np.where(p_ratio_line_df["총합"] > 0, (p_ratio_line_df["값"] / p_ratio_line_df["총합"]) * 100, 0)
                    
                    p_fig_ratio_line = go.Figure()
                    for year in sorted(sel_years):
                        if year == 2026:
                            y26_plan_r = p_ratio_line_df[(p_ratio_line_df["연"] == 2026) & (p_ratio_line_df["계획/실적"] == "계획")]
                            if not y26_plan_r.empty:
                                c_plan = LINE_COLOR_MAP["2026년 계획"]
                                p_fig_ratio_line.add_trace(go.Scatter(x=y26_plan_r["월"], y=y26_plan_r["비중"], mode='markers+lines', name="2026년 계획", line=dict(color=c_plan, width=2.5, dash='dot')))
                            y26_act_r = p_ratio_line_df[(p_ratio_line_df["연"] == 2026) & (p_ratio_line_df["계획/실적"] == "실적") & (p_ratio_line_df["월"] <= 3)]
                            if not y26_act_r.empty:
                                c_act26 = LINE_COLOR_MAP["2026년 실적"]
                                p_fig_ratio_line.add_trace(go.Scatter(x=y26_act_r["월"], y=y26_act_r["비중"], mode='markers+lines', name="2026년 실적", line=dict(color=c_act26, width=2.5)))
                        else:
                            y_act_r = p_ratio_line_df[(p_ratio_line_df["연"] == year) & (p_ratio_line_df["계획/실적"] == "실적")]
                            if not y_act_r.empty:
                                key_name = f"{year}년 실적"
                                c = LINE_COLOR_MAP.get(key_name, "#808080")
                                p_fig_ratio_line.add_trace(go.Scatter(x=y_act_r["월"], y=y_act_r["비중"], mode='markers+lines', name=key_name, line=dict(color=c, width=2.5)))
                    
                    if not p_ratio_line_df["비중"].empty:
                        r_min, r_max = p_ratio_line_df["비중"].min(), p_ratio_line_df["비중"].max()
                        p_y_range = [max(0, r_min * 0.9), min(100, r_max * 1.1)]
                    else:
                        p_y_range = [0, 105]

                    p_fig_ratio_line.update_layout(
                        height=450, 
                        xaxis=dict(dtick=1, title="월"), 
                        yaxis=dict(title="구성비 (%)", range=p_y_range, tickformat=".1f", ticksuffix="%"), 
                        hovermode="x unified", 
                        legend=dict(orientation="h", y=1.1, x=0.5, xanchor='center')
                    )

                    st.markdown(f"<div style='text-align: center; margin-top: 30px;'><b>■ [{print_grp}] 연간 구성비 추이 그래프</b></div>", unsafe_allow_html=True)
                    st.plotly_chart(p_fig_ratio_line, use_container_width=True, key=f"prt_ratio_line_chart_{prefix}_{print_grp}")

                if table_ready and prt_tbl:
                    st.markdown(f"<div style='text-align: center; width: 100%; margin-top: 20px;'><b>■ [{print_grp}] 월별 상세 데이터표</b></div>", unsafe_allow_html=True)
                    st.table(styled)

                st.markdown("</div>", unsafe_allow_html=True)
                st.markdown("<br><br>", unsafe_allow_html=True)
                
            components.html(
                """
                <style>
                @media print {
                    @page { size: A3 landscape !important; margin: 10mm !important; }
                    .main .block-container, .block-container { padding-top: 0 !important; margin-top: 0 !important; padding-left: 0 !important; padding-right: 0 !important; max-width: 100% !important; width: 100% !important; margin: 0 auto !important; }
                    [data-testid="stAppViewContainer"] > section:nth-child(2) { padding-top: 0 !important; max-width: 100% !important; width: 100% !important; margin: 0 auto !important; }
                    header[data-testid="stHeader"], header { display: none !important; }
                    .stHorizontalBlock { justify-content: center !important; gap: 0rem !important; }
                    [data-testid="column"] { padding: 0 15px !important; }
                    #print-btn-container { display: none !important; }
                }
                </style>
                <div id="print-btn-container" style="display: flex; justify-content: center; margin-top: 20px;">
                    <button onclick="printPreview()" style="background-color: #FF4B4B; color: white; border: none; padding: 12px 24px; font-size: 16px; border-radius: 8px; cursor: pointer; font-weight: bold; font-family: sans-serif; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
                        🖨️ PDF 저장하기
                    </button>
                </div>
                <script>
                function printPreview() {
                    var doc = window.parent.document;
                    var marker = doc.getElementById('preview-marker');
                    var hiddenElements = [];
                    if (marker) {
                        var container = marker.closest('[data-testid="stElementContainer"]') || marker.closest('.element-container') || marker.parentNode;
                        var sibling = container.previousElementSibling;
                        while (sibling) {
                            hiddenElements.push({el: sibling, orig: sibling.style.cssText || ''});
                            sibling.style.setProperty('display', 'none', 'important');
                            sibling = sibling.previousElementSibling;
                        }
                    }
                    var extras = doc.querySelectorAll('[data-testid="stSidebar"], header, [data-baseweb="tab-list"], h1');
                    extras.forEach(el => {
                        hiddenElements.push({el: el, orig: el.style.cssText || ''});
                        el.style.setProperty('display', 'none', 'important');
                    });
                    window.parent.print();
                    setTimeout(() => {
                        hiddenElements.forEach(item => { item.el.style.cssText = item.orig; });
                    }, 1500);
                }
                </script>
                """,
                height=80
            )

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
