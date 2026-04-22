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
# 기본 설정 (원본 유지)
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

COLOR_MAP = {
    "가정용": "#4c72b0", "산업용": "#9ebcda", "업무용": "#e07a5f", "영업용": "#e6c253", "기타": "#8cce8b"
}

LINE_COLOR_MAP = {
    "2023년 실적": "#9467bd", "2024년 실적": "#1f77b4", "2025년 실적": "#2ca02c", "2026년 실적": "#d62728", "2026년 계획": "#ff7f0e"
}

USE_COL_TO_GROUP: Dict[str, str] = {
    "취사용": "가정용", "개별난방용": "가정용", "중앙난방용": "가정용", "자가열전용": "가정용",
    "산업용": "산업용",
    "업무난방용": "업무용", "냉방용": "업무용", "주한미군": "업무용",
    "일반용": "영업용",
    "수송용(CNG)": "기타", "수송용(BIO)": "기타", "열병합용": "기타", "열병합용1": "기타", "열병합용2": "기타", "연료전지용": "기타", "열전용설비용": "기타",
}

# ─────────────────────────────────────────────────────────
# 데이터 처리 유틸 (원본 유지)
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
    plan_df = _clean_base(plan_df); actual_df = _clean_base(actual_df)
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
    long_df["연"] = long_df["연"].astype(int); long_df["월"] = long_df["월"].astype(int)
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
# 원본 그래프 렌더링 함수 (유지 및 출력 로직 보완)
# ─────────────────────────────────────────────────────────
def render_monthly_trend(df, unit, prefix):
    st.markdown("### 📈 연간 추이 그래프")
    c1, c2 = st.columns([3, 1])
    with c1: sel_years = st.multiselect("연도 선택(그래프)", options=[2022, 2023, 2024, 2025, 2026], default=[2024, 2025, 2026], key=f"{prefix}my")
    try: sel_group = st.segmented_control("그룹 선택", options=["전체"] + GROUP_ORDER, selection_mode="single", default="전체", key=f"{prefix}sg")
    except: sel_group = st.radio("그룹 선택", options=["전체"] + GROUP_ORDER, index=0, horizontal=True, key=f"{prefix}rd")

    if not sel_years: st.info("연도를 하나 이상 선택해주세요."); return
    plot_df = df[df["그룹"] == sel_group] if sel_group != "전체" else df
    
    fig_line = go.Figure(); fig_bar = go.Figure()
    table_data_list = []; line_y_vals = []

    for year in sorted(sel_years):
        if year == 2026:
            y26_plan = plot_df[(plot_df["연"] == 2026) & (plot_df["계획/실적"] == "계획")].groupby("월")["값"].sum().reset_index()
            y26_act = plot_df[(plot_df["연"] == 2026) & (plot_df["계획/실적"] == "실적") & (plot_df["월"] <= 3)].groupby("월")["값"].sum().reset_index()
            if not y26_plan.empty:
                c_p = LINE_COLOR_MAP["2026년 계획"]
                fig_line.add_trace(go.Scatter(x=y26_plan["월"], y=y26_plan["값"], mode='markers+lines', name="2026년 계획", line=dict(color=c_p, width=2.5, dash='dot')))
                line_y_vals.extend(y26_plan["값"].tolist()); y26_plan["표_컬럼"] = "2026년 계획"; table_data_list.append(y26_plan)
                fig_bar.add_trace(go.Bar(x=y26_plan["월"], y=y26_plan["값"], name="2026년 계획", marker_color=c_p))
            if not y26_act.empty:
                c_a = LINE_COLOR_MAP["2026년 실적"]
                fig_line.add_trace(go.Scatter(x=y26_act["월"], y=y26_act["값"], mode='markers+lines', name="2026년 실적", line=dict(color=c_a, width=2.5)))
                line_y_vals.extend(y26_act["값"].tolist()); y26_act["표_컬럼"] = "2026년 실적"; table_data_list.append(y26_act)
                fig_bar.add_trace(go.Bar(x=y26_act["월"], y=y26_act["값"], name="2026년 실적", marker_color=c_a))
        else:
            y_act_grp = plot_df[(plot_df["연"] == year) & (plot_df["계획/실적"] == "실적")].groupby("월")["값"].sum().reset_index()
            if not y_act_grp.empty:
                key_name = f"{year}년 실적"; c = LINE_COLOR_MAP.get(key_name, "#808080")
                fig_line.add_trace(go.Scatter(x=y_act_grp["월"], y=y_act_grp["값"], mode='markers+lines', name=key_name, line=dict(color=c, width=2.5)))
                line_y_vals.extend(y_act_grp["값"].tolist()); y_act_grp["표_컬럼"] = key_name; table_data_list.append(y_act_grp)
                fig_bar.add_trace(go.Bar(x=y_act_grp["월"], y=y_act_grp["값"], name=f"{year}년", marker_color=c))

    unit_anno = dict(xref="paper", yref="paper", x=1.0, y=1.12, xanchor="right", yanchor="bottom", text=f"(단위: {unit})", font=dict(size=13, color="#555"), showarrow=False)
    fig_line.update_layout(height=550, xaxis=dict(dtick=1, title="월"), yaxis=dict(title=f"판매량({unit})", tickformat=",.0f"), hovermode="x unified", legend=dict(orientation="h", y=1.1), annotations=[unit_anno])
    st.plotly_chart(fig_line, use_container_width=True)

    st.markdown(f"##### 📊 {sel_group} 연도별 동월 비교 (막대그래프)")
    fig_bar.update_layout(barmode='group', bargap=0.36, xaxis=dict(dtick=1, title="월"), yaxis=dict(title=f"판매량({unit})", tickformat=",.0f"), hovermode="x unified", legend=dict(orientation="h", y=1.1), annotations=[unit_anno])
    st.plotly_chart(fig_bar, use_container_width=True)

    if table_data_list:
        t_df = pd.concat(table_data_list, ignore_index=True)
        table = t_df.pivot_table(index="월", columns="표_컬럼", values="값", aggfunc="sum").sort_index().fillna(0.0)
        if "2026년 계획" in table.columns and "2026년 실적" in table.columns:
            table["증감량(차이)"] = table["2026년 실적"] - table["2026년 계획"]
            table.loc[table.index > 3, "증감량(차이)"] = np.nan
            table["증감률(%)"] = np.nan
            valid_mask = (table.index <= 3) & (table["2026년 계획"] != 0)
            table.loc[valid_mask, "증감률(%)"] = (table.loc[valid_mask, "증감량(차이)"] / table.loc[valid_mask, "2026년 계획"]) * 100
        table.loc["합계"] = table.sum(numeric_only=True)
        table = table.reset_index()
        format_dict = {col: "{:,.0f}" for col in table.columns if col not in ["월", "증감률(%)"]}
        if "증감률(%)" in table.columns: format_dict["증감률(%)"] = "{:,.1f}%"
        styled_df = table.style.format(format_dict, na_rep="-").apply(lambda row: ['background-color: #1f497d; color: white;' if row['월'] == '합계' else '' for _ in row], axis=1)
        st.dataframe(center_style(styled_df), use_container_width=True, hide_index=True)

    # ─────────────────────────────────────────────────────────
    # [수정] 보고서 일괄 출력 뷰어 (A3 최적화 및 그래프 최대화)
    # ─────────────────────────────────────────────────────────
    st.divider()
    st.markdown("### 🖨️ 보고서 일괄 출력 뷰어")
    st.markdown("##### 1. 출력 항목 선택")
    chk_col1, chk_col2, chk_col3 = st.columns(3)
    with chk_col1: prt_line = st.checkbox("연간추이 그래프", value=True, key=f"{prefix}_prt_line")
    with chk_col2: prt_bar = st.checkbox("연도별 동월 비교 그래프", value=False, key=f"{prefix}_prt_bar")
    with chk_col3: prt_tbl = st.checkbox("월별 상세 데이터표", value=True, key=f"{prefix}_prt_tbl")

    st.markdown("##### 2. 출력 그룹 선택")
    grp_cols = st.columns(6)
    selected_groups = []
    for i, g_name in enumerate(["전체"] + GROUP_ORDER):
        with grp_cols[i]:
            if st.checkbox(g_name, value=(g_name in ["전체", "가정용"]), key=f"{prefix}_p_grp_{i}"):
                selected_groups.append(g_name)

    if st.button("미리보기", key=f"{prefix}_preview_btn", type="primary"):
        st.markdown("<div id='preview-marker' style='display:none;'></div>", unsafe_allow_html=True)
        # [핵심] 글자 크기 110% 및 레이아웃 강제 고정 CSS
        st.markdown("""
            <style>
            div[data-testid="stTable"] table { width: 100% !important; table-layout: auto !important; }
            div[data-testid="stTable"] table th, div[data-testid="stTable"] table td {
                height: 35px !important; padding: 0px 8px !important; line-height: 35px !important;
                vertical-align: middle !important; white-space: nowrap !important; font-size: 110% !important;
            }
            div[data-testid="stTable"] table th:first-child, div[data-testid="stTable"] table td:first-child {
                width: 60px !important; min-width: 60px !important;
            }
            .print-page-container { page-break-inside: avoid; break-inside: avoid; width: 100% !important; margin-bottom: 20px; }
            </style>
            """, unsafe_allow_html=True)
        
        for print_grp in selected_groups:
            st.markdown(f"<div class='print-page-container'>", unsafe_allow_html=True)
            st.markdown(f"<h2 style='text-align: center; color: #1f497d; margin-top: 5px;'>[{print_grp}] 판매량 분석 보고</h2>", unsafe_allow_html=True)
            p_df = df[df["그룹"] == print_grp] if print_grp != "전체" else df
            p_fig = go.Figure(); p_list = []; p_line_y = []

            for year in sorted(sel_years):
                if year == 2026:
                    y26p = p_df[(p_df["연"] == 2026) & (p_df["계획/실적"] == "계획")].groupby("월")["값"].sum().reset_index()
                    y26a = p_df[(p_df["연"] == 2026) & (p_df["계획/실적"] == "실적") & (p_df["월"] <= 3)].groupby("월")["값"].sum().reset_index()
                    if not y26p.empty:
                        p_fig.add_trace(go.Scatter(x=y26p["월"], y=y26p["값"], mode='markers+lines', name="26년 계획", line=dict(color=LINE_COLOR_MAP["2026년 계획"], width=4, dash='dot')))
                        p_line_y.extend(y26p["값"].tolist()); y26p["표_컬럼"] = "2026년 계획"; p_list.append(y26p)
                    if not y26a.empty:
                        p_fig.add_trace(go.Scatter(x=y26a["월"], y=y26a["값"], mode='markers+lines', name="26년 실적", line=dict(color=LINE_COLOR_MAP["2026년 실적"], width=4)))
                        p_line_y.extend(y26a["값"].tolist()); y26a["표_컬럼"] = "2026년 실적"; p_list.append(y26a)
                else:
                    y_act = p_df[(p_df["연"] == year) & (p_df["계획/실적"] == "실적")].groupby("월")["값"].sum().reset_index()
                    if not y_act.empty:
                        k = f"{year}년 실적"; c = LINE_COLOR_MAP.get(k, "#808080")
                        p_fig.add_trace(go.Scatter(x=y_act["월"], y=y_act["값"], mode='markers+lines', name=k, line=dict(color=c, width=4)))
                        p_line_y.extend(y_act["값"].tolist()); y_act["표_컬럼"] = k; p_list.append(y_act)

            # [수정] 1그래프 + 1표 레이아웃: 그래프 높이 650px로 최대화
            col_l, col_r = st.columns([1.5, 1])
            with col_l:
                p_fig.update_layout(height=650, xaxis=dict(dtick=1), yaxis=dict(tickformat=",.0f"), legend=dict(orientation="h", y=1.1, x=0.5, xanchor='center'), margin=dict(t=50, b=20))
                st.plotly_chart(p_fig, use_container_width=True, key=f"p_chart_{prefix}_{print_grp}")
            with col_r:
                if prt_tbl and p_list:
                    pt = pd.concat(p_list).pivot_table(index="월", columns="표_컬럼", values="값", aggfunc="sum").sort_index().fillna(0.0)
                    if "2026년 계획" in pt.columns and "2026년 실적" in pt.columns:
                        pt["증감량(차이)"] = pt["2026년 실적"] - pt["2026년 계획"]
                        pt.loc[pt.index > 3, "증감량(차이)"] = np.nan
                        pt["증감률(%)"] = (pt["증감량(차이)"] / pt["2026년 계획"] * 100).replace([np.inf, -np.inf], np.nan)
                    pt.loc["합계"] = pt.sum(numeric_only=True)
                    pt = pt.reset_index()
                    f_d = {col: "{:,.0f}" for col in pt.columns if col not in ["월", "증감률(%)"]}
                    if "증감률(%)" in pt.columns: f_d["증감률(%)"] = "{:,.1f}%"
                    st.table(center_style(pt.style.format(f_d, na_rep="").apply(lambda r: ['background-color: #1f497d; color: white; font-weight: bold;' if r['월'] == '합계' else '' for _ in r], axis=1)))
            st.markdown("</div>", unsafe_allow_html=True)

        components.html("""
            <style>
            @media print {
                @page { size: A3 landscape !important; margin: 5mm !important; }
                header, [data-testid="stSidebar"], [data-testid="stHeader"] { display: none !important; }
                .main .block-container { padding-top: 0 !important; margin-top: 0 !important; max-width: 98% !important; }
                #print-btn-container { display: none !important; }
            }
            </style>
            <div id="print-btn-container" style="display: flex; justify-content: center; margin-top: 20px;">
                <button onclick="printReport()" style="background-color: #FF4B4B; color: white; border: none; padding: 15px 30px; font-size: 20px; border-radius: 10px; cursor: pointer; font-weight: bold;">
                    🖨️ PDF 저장하기 (A3 가로)
                </button>
            </div>
            <script>
            function printReport() {
                var doc = window.parent.document;
                var marker = doc.getElementById('preview-marker');
                var children = marker.closest('[data-testid="stElementContainer"]').parentNode.children;
                var found = false;
                for (var i = 0; i < children.length; i++) {
                    if (children[i].contains(marker)) { found = true; continue; }
                    if (!found) children[i].style.display = 'none';
                }
                window.parent.print();
                setTimeout(() => { window.parent.location.reload(); }, 1000);
            }
            </script>
            """, height=100)

# ─────────────────────────────────────────────────────────
# 메인 실행 (원본 유지)
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
            with tab: render_monthly_trend(df, "천m³" if k == "부피" else "GJ", k)
    else: st.warning("데이터 파일을 로드할 수 없습니다.")

if __name__ == "__main__": main()
