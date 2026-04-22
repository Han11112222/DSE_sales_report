import io
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import numpy as np
import matplotlib as mpl
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as components

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
GROUP_ORDER = ["가정용", "산업용", "업무용", "영업용", "기타"]

LINE_COLOR_MAP = {
    "2023년 실적": "#9467bd",
    "2024년 실적": "#1f77b4",
    "2025년 실적": "#2ca02c",
    "2026년 실적": "#d62728",
    "2026년 계획": "#ff7f0e"
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
# 그래프 및 보고서 섹션
# ─────────────────────────────────────────────────────────
def render_monthly_trend(df, unit, prefix):
    st.markdown("### 📈 연간 추이 그래프")
    
    c1, c2 = st.columns([3, 1])
    with c1: 
        sel_years = st.multiselect("연도 선택(그래프)", options=[2022, 2023, 2024, 2025, 2026], default=[2024, 2025, 2026], key=f"{prefix}my")

    try:
        sel_group = st.segmented_control("그룹 선택", options=["전체"] + GROUP_ORDER, selection_mode="single", default="전체", key=f"{prefix}sg")
    except:
        sel_group = st.radio("그룹 선택", options=["전체"] + GROUP_ORDER, index=0, horizontal=True, key=f"{prefix}rd")

    if not sel_years:
        st.info("연도를 하나 이상 선택해주세요.")
        return

    plot_df = df[df["그룹"] == sel_group] if sel_group != "전체" else df
    
    fig_line = go.Figure()
    table_data_list = []
    line_y_vals = []

    for year in sorted(sel_years):
        if year == 2026:
            y26_plan = plot_df[(plot_df["연"] == 2026) & (plot_df["계획/실적"] == "계획")].groupby("월")["값"].sum().reset_index()
            y26_act = plot_df[(plot_df["연"] == 2026) & (plot_df["계획/실적"] == "실적") & (plot_df["월"] <= 3)].groupby("월")["값"].sum().reset_index()
            if not y26_plan.empty:
                fig_line.add_trace(go.Scatter(x=y26_plan["월"], y=y26_plan["값"], mode='markers+lines', name="2026년 계획", line=dict(color=LINE_COLOR_MAP["2026년 계획"], width=3, dash='dot')))
                line_y_vals.extend(y26_plan["값"].tolist())
                y26_plan_tb = y26_plan.copy(); y26_plan_tb["표_컬럼"] = "2026년 계획"; table_data_list.append(y26_plan_tb)
            if not y26_act.empty:
                fig_line.add_trace(go.Scatter(x=y26_act["월"], y=y26_act["값"], mode='markers+lines', name="2026년 실적", line=dict(color=LINE_COLOR_MAP["2026년 실적"], width=3)))
                line_y_vals.extend(y26_act["값"].tolist())
                y26_act_tb = y26_act.copy(); y26_act_tb["표_컬럼"] = "2026년 실적"; table_data_list.append(y26_act_tb)
        else:
            y_act_grp = plot_df[(plot_df["연"] == year) & (plot_df["계획/실적"] == "실적")].groupby("월")["값"].sum().reset_index()
            if not y_act_grp.empty:
                key_name = f"{year}년 실적"
                fig_line.add_trace(go.Scatter(x=y_act_grp["월"], y=y_act_grp["값"], mode='markers+lines', name=key_name, line=dict(color=LINE_COLOR_MAP.get(key_name, "#808080"), width=3)))
                line_y_vals.extend(y_act_grp["값"].tolist())
                y_act_tb = y_act_grp.copy(); y_act_tb["표_컬럼"] = key_name; table_data_list.append(y_act_tb)

    unit_anno = dict(xref="paper", yref="paper", x=1.0, y=1.05, xanchor="right", yanchor="bottom", text=f"(단위: {unit})", font=dict(size=14), showarrow=False)
    fig_line.update_layout(height=600, xaxis=dict(dtick=1, title="월"), yaxis=dict(title=f"판매량({unit})", tickformat=",.0f"), hovermode="x unified", legend=dict(orientation="h", y=1.1), annotations=[unit_anno])
    st.plotly_chart(fig_line, use_container_width=True)

    # 상세 데이터표
    if table_data_list:
        t_df = pd.concat(table_data_list, ignore_index=True)
        table = t_df.pivot_table(index="월", columns="표_컬럼", values="값", aggfunc="sum").sort_index().fillna(0.0)
        if "2026년 계획" in table.columns and "2026년 실적" in table.columns:
            table["증감량"] = table["2026년 실적"] - table["2026년 계획"]
            table.loc[table.index > 3, "증감량"] = np.nan
            table["증감률(%)"] = (table["증감량"] / table["2026년 계획"] * 100).replace([np.inf, -np.inf], np.nan)
        
        table.loc["합계"] = table.sum(numeric_only=True)
        table = table.reset_index()
        styled = center_style(table.style.format("{:,.1f}").format({"월": "{}"}))
        st.dataframe(styled, use_container_width=True, hide_index=True)

    # ─────────────────────────────────────────────────────────
    # [수정 포인트] 인쇄용 미리보기 - 그래프 최대화 & 2단 배치
    # ─────────────────────────────────────────────────────────
    st.divider()
    st.markdown("### 🖨️ 인쇄용 미리보기 (A3 가로 최적화)")
    
    selected_groups = st.multiselect("출력할 그룹 선택", options=["전체"] + GROUP_ORDER, default=["전체", "가정용"], key=f"{prefix}_print_sel")

    if st.button("미리보기 활성화", key=f"{prefix}_preview_btn", type="primary"):
        st.markdown("<div id='preview-marker'></div>", unsafe_allow_html=True)
        for print_grp in selected_groups:
            # 개별 섹션 컨테이너: 페이지 나눔 없이 적절한 여백만 주어 한 페이지에 2개씩 담기게 함
            st.markdown(f"<div style='width: 100%; margin-bottom: 50px; border-bottom: 1px dashed #ccc; padding-bottom: 20px;'>", unsafe_allow_html=True)
            st.markdown(f"<h2 style='text-align: center; color: #1f497d;'>[{print_grp}] 판매량 분석 보고</h2>", unsafe_allow_html=True)

            p_df = df[df["그룹"] == print_grp] if print_grp != "전체" else df
            p_fig = go.Figure()
            p_vals = []
            p_list = []

            for year in sorted(sel_years):
                if year == 2026:
                    y26_p = p_df[(p_df["연"] == 2026) & (p_df["계획/실적"] == "계획")].groupby("월")["값"].sum().reset_index()
                    y26_a = p_df[(p_df["연"] == 2026) & (p_df["계획/실적"] == "실적") & (p_df["월"] <= 3)].groupby("월")["값"].sum().reset_index()
                    if not y26_p.empty:
                        p_fig.add_trace(go.Scatter(x=y26_p["월"], y=y26_p["값"], mode='markers+lines', name="26년 계획", line=dict(color=LINE_COLOR_MAP["2026년 계획"], width=4, dash='dot')))
                        p_vals.extend(y26_p["값"].tolist()); y26_p["표_컬럼"] = "26년 계획"; p_list.append(y26_p)
                    if not y26_a.empty:
                        p_fig.add_trace(go.Scatter(x=y26_a["월"], y=y26_a["값"], mode='markers+lines', name="26년 실적", line=dict(color=LINE_COLOR_MAP["2026년 실적"], width=4)))
                        p_vals.extend(y26_a["값"].tolist()); y26_a["표_컬럼"] = "26년 실적"; p_list.append(y26_a)
                else:
                    y_a = p_df[(p_df["연"] == year) & (p_df["계획/실적"] == "실적")].groupby("월")["값"].sum().reset_index()
                    if not y_a.empty:
                        k = f"{year}년 실적"
                        p_fig.add_trace(go.Scatter(x=y_a["월"], y=y_a["값"], mode='markers+lines', name=k, line=dict(color=LINE_COLOR_MAP.get(k, "#808080"), width=4)))
                        p_vals.extend(y_a["값"].tolist()); y_a["표_컬럼"] = k; p_list.append(y_a)

            # [수정] 그래프 높이를 650px로 확대하여 시인성 극대화
            p_fig.update_layout(height=650, margin=dict(l=10, r=10, t=50, b=10), xaxis=dict(dtick=1, tickfont=dict(size=14)), yaxis=dict(tickformat=",.0f", tickfont=dict(size=14)), legend=dict(orientation="h", y=1.08, x=0.5, xanchor='center', font=dict(size=16)))
            
            # 그래프와 표 가로 배치 (그래프 비중 70%)
            col_g, col_t = st.columns([2.3, 1])
            with col_g:
                st.plotly_chart(p_fig, use_container_width=True, config={'displayModeBar': False})
            with col_t:
                if p_list:
                    pt = pd.concat(p_list).pivot_table(index="월", columns="표_컬럼", values="값", aggfunc="sum").sort_index().fillna(0.0)
                    pt.loc["합계"] = pt.sum()
                    st.markdown("<div style='margin-top: 50px;'></div>", unsafe_allow_html=True)
                    st.table(pt.reset_index().style.format("{:,.0f}").format({"월": "{}"}))
            st.markdown("</div>", unsafe_allow_html=True)

        components.html(
            """
            <script>
            function printReport() {
                var doc = window.parent.document;
                var marker = doc.getElementById('preview-marker');
                if (!marker) return;
                
                var container = marker.closest('[data-testid="stElementContainer"]').parentNode;
                var children = container.children;
                var found = false;
                for (var i = 0; i < children.length; i++) {
                    if (children[i].contains(marker)) { found = true; continue; }
                    if (!found) children[i].style.display = 'none';
                }
                
                var sidebar = doc.querySelector('[data-testid="stSidebar"]');
                if (sidebar) sidebar.style.display = 'none';
                var header = doc.querySelector('header');
                if (header) header.style.display = 'none';

                window.parent.print();
                
                setTimeout(() => { window.parent.location.reload(); }, 1000);
            }
            </script>
            <div style="display: flex; justify-content: center;">
                <button onclick="printReport()" style="background-color: #FF4B4B; color: white; border: none; padding: 15px 30px; font-size: 20px; border-radius: 10px; cursor: pointer; font-weight: bold;">
                    🖨️ A3 가로 인쇄 (영역에 맞추기 권장)
                </button>
            </div>
            """, height=100
        )

# ─────────────────────────────────────────────────────────
# 메인 실행
# ─────────────────────────────────────────────────────────
def main():
    st.title("📊 DSE 판매량 분석 보고")
    excel_bytes = None
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
