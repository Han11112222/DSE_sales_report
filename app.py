import io
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
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
st.set_page_config(page_title="도시가스 판매량 분석 보고서", layout="wide")

DEFAULT_SALES_XLSX = "판매량(계획_실적).xlsx"

# 엑셀 헤더 → 분석 그룹 매핑 (판매량용)
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
    "수송용(CNG)": "수송용",
    "수송용(BIO)": "수송용",
    "열병합용": "열병합",
    "열병합용1": "열병합",
    "열병합용2": "열병합",
    "연료전지용": "연료전지",
    "열전용설비용": "열전용설비용",
}

GROUP_OPTIONS: List[str] = [
    "총량",
    "가정용",
    "영업용",
    "업무용",
    "산업용",
    "수송용",
    "열병합",
    "연료전지",
    "열전용설비용",
]

# 색상
COLOR_PLAN = "rgba(0, 90, 200, 1)"
COLOR_ACT = "rgba(0, 150, 255, 1)"
COLOR_PREV = "rgba(190, 190, 190, 1)"
COLOR_DIFF = "rgba(0, 80, 160, 1)"

# ─────────────────────────────────────────────────────────
# 공통 유틸
# ─────────────────────────────────────────────────────────
def fmt_num_safe(v) -> str:
    if pd.isna(v):
        return "-"
    try:
        return f"{float(v):,.0f}"
    except Exception:
        return "-"

def fmt_rate(v: float) -> str:
    if pd.isna(v) or np.isnan(v):
        return "-"
    return f"{float(v):,.1f}%"

def center_style(styler):
    """모든 표 숫자 가운데 정렬용 공통 스타일."""
    styler = styler.set_properties(**{"text-align": "center"})
    styler = styler.set_table_styles(
        [dict(selector="th", props=[("text-align", "center")])]
    )
    return styler

def highlight_subtotal(s):
    """표의 '💡 소계', '💡 총계', '💡 합계' 행을 연한 회색으로 하이라이트."""
    is_subtotal = s.astype(str).str.contains('💡 소계|💡 총계|💡 합계')
    return ['background-color: #f8f9fa; font-weight: bold;' if is_subtotal.any() else '' for _ in s]

def _clean_base(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "Unnamed: 0" in out.columns:
        out = out.drop(columns=["Unnamed: 0"])
    out["연"] = pd.to_numeric(out["연"], errors="coerce").astype("Int64")
    out["월"] = pd.to_numeric(out["월"], errors="coerce").astype("Int64")
    return out

def keyword_group(col: str) -> Optional[str]:
    """판매량 컬럼명이 약간 달라도 잡히도록 키워드 기반 보정."""
    c = str(col)
    if "열병합" in c: return "열병합"
    if "연료전지" in c: return "연료전지"
    if "수송용" in c: return "수송용"
    if "열전용" in c: return "열전용설비용"
    if c in ["산업용"]: return "산업용"
    if c in ["일반용"]: return "영업용"
    if any(k in c for k in ["취사용", "난방용", "자가열"]): return "가정용"
    if any(k in c for k in ["업무", "냉방", "주한미군"]): return "업무용"
    return None

def make_long(plan_df: pd.DataFrame, actual_df: pd.DataFrame) -> pd.DataFrame:
    """판매량 wide → long (연·월·그룹·용도·계획/실적·값)."""
    plan_df = _clean_base(plan_df)
    actual_df = _clean_base(actual_df)

    records = []
    for label, df in [("계획", plan_df), ("실적", actual_df)]:
        for col in df.columns:
            if col in ["연", "월"]:
                continue

            group = USE_COL_TO_GROUP.get(col)
            if group is None:
                group = keyword_group(col)
            if group is None:
                continue

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
    return long_df

def load_all_sheets(excel_bytes: bytes) -> Dict[str, pd.DataFrame]:
    """판매량 파일 시트 로드"""
    xls = pd.ExcelFile(io.BytesIO(excel_bytes), engine="openpyxl")
    needed = ["계획_부피", "실적_부피", "계획_열량", "실적_열량"]
    out: Dict[str, pd.DataFrame] = {}
    for name in needed:
        if name in xls.sheet_names:
            out[name] = xls.parse(name)
    return out

def build_long_dict(sheets: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    """판매량 long dict"""
    long_dict: Dict[str, pd.DataFrame] = {}
    if ("계획_부피" in sheets) and ("실적_부피" in sheets):
        long_dict["부피"] = make_long(sheets["계획_부피"], sheets["실적_부피"])
    if ("계획_열량" in sheets) and ("실적_열량" in sheets):
        long_dict["열량"] = make_long(sheets["계획_열량"], sheets["실적_열량"])
    return long_dict

def apply_period_filter(df: pd.DataFrame, sel_year: int, sel_month: int, agg_mode: str) -> pd.DataFrame:
    """기준 연/월 + 당월/연누적 공통 필터."""
    if df.empty:
        return df

    base = df[df["연"] == sel_year].copy()
    if agg_mode == "당월":
        base = base[base["월"] == sel_month]
    else:  # "연 누적"
        base = base[base["월"] <= sel_month]
    return base

def apply_period_filter_multi_years(df: pd.DataFrame, sel_month: int, agg_mode: str, years: List[int]) -> pd.DataFrame:
    """여러 연도 비교 차트용: 기준월/모드로 각 연도 동일 기간만 남김."""
    if df.empty:
        return df

    base = df[df["연"].isin(years)].copy()
    if agg_mode == "당월":
        base = base[base["월"] == sel_month]
    else:
        base = base[base["월"] <= sel_month]
    return base

def render_section_selector(
    long_df: pd.DataFrame,
    title: str,
    key_prefix: str,
    fixed_mode: Optional[str] = None,
    show_mode: bool = True
) -> Tuple[int, int, str, List[int]]:
    """각 섹션별 기준선택 UI."""
    st.markdown(f"#### ✅ {title} 기준 선택")

    if long_df.empty:
        st.info("연도 정보가 없습니다.")
        return 0, 1, "연 누적", []

    years_all = sorted(long_df["연"].unique().tolist())

    df_for_default = long_df.copy()
    if {"계획/실적", "값"}.issubset(df_for_default.columns):
        mask = (
            (df_for_default["계획/실적"] == "실적")
            & df_for_default["값"].notna()
            & (df_for_default["값"] != 0)
        )
        if mask.any():
            df_for_default = df_for_default[mask]

    if not df_for_default.empty:
        default_year = int(df_for_default["연"].max())
    else:
        default_year = years_all[-1] if years_all else 2025

    months_for_default_year = sorted(
        df_for_default[df_for_default["연"] == default_year]["월"].unique().tolist()
    )
    if not months_for_default_year:
        months_for_default_year = sorted(
            long_df[long_df["연"] == default_year]["월"].unique().tolist()
        )
    default_month_global = months_for_default_year[-1] if months_for_default_year else 1

    c1, c2, c3 = st.columns([1.2, 1.2, 1.6])

    with c1:
        sel_year = st.selectbox(
            "기준 연도",
            options=years_all,
            index=years_all.index(default_year) if default_year in years_all else 0,
            key=f"{key_prefix}year",
        )

    months_options = list(range(1, 13))
    df_sel = long_df[long_df["연"] == sel_year].copy()
    months_actual: List[int] = []
    
    if {"계획/실적", "값"}.issubset(df_sel.columns):
        m = (
            (df_sel["계획/실적"] == "실적")
            & df_sel["값"].notna()
            & (df_sel["값"] != 0)
        )
        months_actual = sorted(df_sel[m]["월"].unique().tolist())

    if months_actual:
        default_month_for_sel_year = months_actual[-1]
    else:
        default_month_for_sel_year = default_month_global

    if default_month_for_sel_year < 1: default_month_for_sel_year = 1
    if default_month_for_sel_year > 12: default_month_for_sel_year = 12

    with c2:
        sel_month = st.selectbox(
            "기준 월",
            options=months_options,
            index=months_options.index(default_month_for_sel_year),
            key=f"{key_prefix}month",
        )

    if fixed_mode in ["당월", "연 누적"]:
        agg_mode = fixed_mode
        with c3:
            st.markdown(
                "<div style='padding-top:28px;font-size:14px;color:#666;'>집계 기준: <b>연 누적</b></div>"
                if fixed_mode == "연 누적"
                else "<div style='padding-top:28px;font-size:14px;color:#666;'>집계 기준: <b>당월</b></div>",
                unsafe_allow_html=True,
            )
    else:
        if show_mode:
            with c3:
                agg_mode = st.radio(
                    "집계 기준",
                    ["당월", "연 누적"],
                    index=0,
                    horizontal=True,
                    key=f"{key_prefix}mode",
                )
        else:
            agg_mode = "연 누적"
            with c3:
                st.markdown(
                    "<div style='padding-top:28px;font-size:14px;color:#666;'>집계 기준: <b>연 누적</b></div>",
                    unsafe_allow_html=True,
                )

    st.markdown(
        f"<div style='margin-top:-4px;font-size:13px;color:#666;'>"
        f"선택 기준: <b>{sel_year}년 {sel_month}월</b> · {agg_mode}"
        f"</div>",
        unsafe_allow_html=True,
    )

    return sel_year, sel_month, agg_mode, years_all

def render_metric_card(icon: str, title: str, main: str, sub: str = "", color: str = "#1f77b4"):
    html = f"""
    <div style="
        background-color:#ffffff;
        border-radius:22px;
        padding:24px 26px 20px 26px;
        box-shadow:0 4px 18px rgba(0,0,0,0.06);
        height:100%;
        display:flex;
        flex-direction:column;
        justify-content:flex-start;
    ">
        <div style="font-size:44px; line-height:1; margin-bottom:8px;">{icon}</div>
        <div style="font-size:18px; font-weight:650; color:#444; margin-bottom:6px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{title}</div>
        <div style="font-size:28px; font-weight:750; color:{color}; margin-bottom:8px; white-space: nowrap; letter-spacing:-0.5px;">{main}</div>
        <div style="font-size:14px; color:#444; min-height:20px; font-weight:500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">{sub}</div>
    </div>
    """
    st.markdown(html, unsafe_allow_html=True)

def render_rate_donut(rate: float, color: str, title: str = ""):
    if pd.isna(rate) or np.isnan(rate):
        st.markdown("<div style='font-size:14px;color:#999;text-align:center;'>데이터 없음</div>", unsafe_allow_html=True)
        return

    filled = max(min(float(rate), 200.0), 0.0)
    empty = max(100.0 - filled, 0.0)

    fig = go.Figure(
        data=[go.Pie(
            values=[filled, empty],
            hole=0.7,
            sort=False,
            direction="clockwise",
            marker=dict(colors=[color, "#e5e7eb"]),
            textinfo="none",
        )]
    )
    fig.update_layout(
        showlegend=False,
        width=200,
        height=230,
        margin=dict(l=0, r=0, t=40, b=0),
        title=dict(text=title, font=dict(size=14, color="#666"), x=0.5, xanchor='center', y=0.98) if title else None,
        annotations=[dict(
            text=f"{rate:.1f}%",
            x=0.5, y=0.5,
            showarrow=False,
            font=dict(size=22, color=color, family="NanumGothic"),
        )],
    )
    st.plotly_chart(fig, use_container_width=False)

# ─────────────────────────────────────────────────────────
# 판매량 분석 섹션 컴포넌트
# ─────────────────────────────────────────────────────────
def monthly_core_dashboard(long_df: pd.DataFrame, unit_label: str, key_prefix: str = ""):
    st.markdown("## 📌 월간 핵심 대시보드")

    if long_df.empty:
        st.info("데이터가 없습니다.")
        return

    sel_year, sel_month, agg_mode, years_all = render_section_selector(
        long_df, "월간 핵심 대시보드", key_prefix + "dash_base_"
    )
    mode_tag = "당월" if agg_mode == "당월" else "연도누적(연 누적)"

    base_this = apply_period_filter(long_df, sel_year, sel_month, agg_mode)
    plan_total = base_this[base_this["계획/실적"] == "계획"]["값"].sum()
    act_total = base_this[base_this["계획/실적"] == "실적"]["값"].sum()

    prev_year = sel_year - 1
    has_prev = prev_year in years_all
    if has_prev:
        base_prev = apply_period_filter(long_df, prev_year, sel_month, agg_mode)
        prev_total = base_prev[base_prev["계획/실적"] == "실적"]["값"].sum()
    else:
        base_prev = pd.DataFrame([])
        prev_total = np.nan

    plan_diff = act_total - plan_total if not pd.isna(plan_total) else np.nan
    plan_rate = (act_total / plan_total * 100.0) if (plan_total and plan_total > 0) else np.nan

    prev_diff = act_total - prev_total if not pd.isna(prev_total) else np.nan
    prev_rate = (act_total / prev_total * 100.0) if (prev_total and prev_total > 0) else np.nan

    st.markdown("<br>", unsafe_allow_html=True)
    k1, k2, k3 = st.columns(3)

    with k1:
        render_metric_card("📘", f"계획 합계 ({unit_label})", fmt_num_safe(plan_total), "", color="#2563eb")
    with k2:
        sub2 = f"계획대비 차이 {fmt_num_safe(plan_diff)} · 달성률({mode_tag}) {fmt_rate(plan_rate)}"
        render_metric_card("📗", f"실적 합계 ({unit_label})", fmt_num_safe(act_total), sub2, color="#16a34a")
    with k3:
        if pd.isna(prev_total):
            main_prev = "-"
            sub3 = "전년 데이터 없음"
        else:
            main_prev = fmt_num_safe(prev_total)
            sub3 = f"전년대비 차이 {fmt_num_safe(prev_diff)} · 증감률({mode_tag}) {fmt_rate(prev_rate)}"
        render_metric_card("📙", f"전년 동월{' 누적' if agg_mode=='연 누적' else ''} 실적 ({unit_label})", main_prev, sub3, color="#f97316")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("#### 🎯 달성률 요약")

    d1, d2, d3, d4, d5 = st.columns([1, 2, 1, 2, 1])
    with d2:
        render_rate_donut(plan_rate, "#16a34a", "계획 달성률")
        st.caption(f"({mode_tag})")
    with d4:
        render_rate_donut(prev_rate, "#f97316", "전년대비 증감률")
        st.caption(f"({mode_tag})")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("#### ⚠️ 특이사항 (계획·전년 대비 편차 핵심 이슈)")

    if base_this.empty:
        st.info("선택 기준에 해당하는 데이터가 없습니다.")
        return

    try:
        det = base_this.copy()
        det["그룹/용도"] = det["그룹"] + " / " + det["용도"]

        pivot = (
            det.pivot_table(index="그룹/용도", columns="계획/실적", values="값", aggfunc="sum")
            .fillna(0.0).rename_axis(None, axis=1).reset_index()
        )

        for c in ["계획", "실적"]:
            if c not in pivot.columns: pivot[c] = 0.0

        pivot["계획대비차이"] = pivot["실적"] - pivot["계획"]
        pivot["계획달성률(%)"] = np.where(pivot["계획"] != 0, (pivot["실적"] / pivot["계획"]) * 100.0, np.nan)

        if has_prev:
            prev_only = apply_period_filter(long_df, prev_year, sel_month, agg_mode)
            prev_only = prev_only[prev_only["계획/실적"] == "실적"].copy()
            prev_only["그룹/용도"] = prev_only["그룹"] + " / " + prev_only["용도"]
            prev_grp = prev_only.groupby("그룹/용도", as_index=False)["값"].sum().rename(columns={"값": "전년실적"})
            pivot = pivot.merge(prev_grp, on="그룹/용도", how="left")
        else:
            pivot["전년실적"] = np.nan

        pivot["전년대비차이"] = pivot["실적"] - pivot["전년실적"]
        pivot["전년대비증감률(%)"] = np.where(pivot["전년실적"] != 0, (pivot["실적"] / pivot["전년실적"]) * 100.0, np.nan)

        if pivot.empty:
            st.markdown("<div style='font-size:14px;color:#666;'>표시할 특이사항이 없습니다.</div>", unsafe_allow_html=True)
            return

        plan_rank = pivot.copy()
        plan_rank["_abs_plan"] = plan_rank["계획대비차이"].abs()
        plan_rank = plan_rank.sort_values("_abs_plan", ascending=False)

        prev_rank = pivot.copy()
        prev_rank = prev_rank[~prev_rank["전년실적"].isna()]
        prev_rank["_abs_prev"] = prev_rank["전년대비차이"].abs()
        prev_rank = prev_rank.sort_values("_abs_prev", ascending=False)

        picked_rows = []
        if len(plan_rank) >= 1: picked_rows.append(plan_rank.iloc[0])
        if len(prev_rank) >= 1: picked_rows.append(prev_rank.iloc[0])
        else:
            if len(plan_rank) >= 2: picked_rows.append(plan_rank.iloc[1])

        core_issues = pd.DataFrame(picked_rows).drop_duplicates(subset=["그룹/용도"])
        if len(core_issues) < 2:
            for _, row in plan_rank.iterrows():
                if row["그룹/용도"] not in core_issues["그룹/용도"].values:
                    core_issues = pd.concat([core_issues, row.to_frame().T], ignore_index=True)
                if len(core_issues) >= 2: break
        core_issues = core_issues.head(2)

        show_cols = ["그룹/용도", "계획", "실적", "계획대비차이", "계획달성률(%)", "전년실적", "전년대비차이", "전년대비증감률(%)"]
        disp = core_issues[show_cols].copy()

        for c in ["계획", "실적", "계획대비차이", "전년실적", "전년대비차이"]: disp[c] = disp[c].apply(fmt_num_safe)
        for c in ["계획달성률(%)", "전년대비증감률(%)"]: disp[c] = disp[c].apply(fmt_rate)

        styled = center_style(disp.astype(str).style)
        html_table = styled.to_html(index=False, escape=False)
        st.markdown(f"<div style='border-radius:12px; overflow-x:auto; border:1px solid #eee;'>{html_table}</div>", unsafe_allow_html=True)

    except Exception:
        st.markdown("<div style='font-size:14px;color:#666;'>특이사항 계산 중 오류가 발생해 표시를 생략했습니다.</div>", unsafe_allow_html=True)

def monthly_trend_section(long_df: pd.DataFrame, unit_label: str, key_prefix: str = ""):
    st.markdown("### 📈 월별 추이 그래프")

    if long_df.empty:
        st.info("데이터가 없습니다.")
        return

    sel_year, sel_month, agg_mode, years_all = render_section_selector(
        long_df, "월별 추이 그래프", key_prefix + "trend_base_", fixed_mode="연 누적", show_mode=False
    )

    preferred_years = [y for y in [2021, 2022, 2023, 2024, 2025] if y in years_all]
    if sel_year not in preferred_years: preferred_years = preferred_years + [sel_year]
    default_years = preferred_years if preferred_years else [sel_year]

    sel_years = st.multiselect("연도 선택(그래프)", options=years_all, default=default_years, key=f"{key_prefix}trend_years")
    if not sel_years:
        st.info("표시할 연도를 한 개 이상 선택해 주세요.")
        return

    try:
        sel_group = st.segmented_control("그룹 선택", GROUP_OPTIONS, selection_mode="single", default="총량", key=f"{key_prefix}trend_group")
    except Exception:
        sel_group = st.radio("그룹 선택", GROUP_OPTIONS, index=0, horizontal=True, key=f"{key_prefix}trend_group_radio")

    base = long_df[long_df["연"].isin(sel_years)].copy()
    base = apply_period_filter_multi_years(base, sel_month, agg_mode, sel_years)

    if sel_group != "총량":
        base = base[base["그룹"] == sel_group]

    plot_df = base.groupby(["연", "월", "계획/실적"], as_index=False)["값"].sum().sort_values(["연", "월", "계획/실적"])
    if plot_df.empty:
        st.info("선택 조건에 해당하는 데이터가 없습니다.")
        return

    plot_df["라벨"] = plot_df["연"].astype(str) + "년 · " + ("" if sel_group == "총량" else sel_group + " · ") + plot_df["계획/실적"]

    fig = px.line(
        plot_df, x="월", y="값", color="라벨", line_dash="계획/실적",
        category_orders={"계획/실적": ["실적", "계획"]}, line_dash_map={"실적": "solid", "계획": "dash"}, markers=True,
    )
    fig.update_layout(
        xaxis=dict(dtick=1), yaxis_title=f"판매량 ({unit_label})", legend_title="연도 / 구분",
        margin=dict(l=10, r=10, t=60, b=10), legend=dict(orientation="h", yanchor="bottom", y=1.12, xanchor="right", x=1),
    )
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("##### 🔢 월별 수치표")
    table = plot_df.pivot_table(index="월", columns="라벨", values="값", aggfunc="sum").sort_index().fillna(0.0)
    total_row = table.sum(numeric_only=True)
    table.index = table.index.astype(object)
    table.loc["소계"] = total_row
    table = table.reset_index()
    numeric_cols = [c for c in table.columns if c != "월"]
    styled = center_style(table.style.format({col: "{:,.0f}" for col in numeric_cols}))
    st.dataframe(styled, use_container_width=True, hide_index=True)

def yearly_summary_section(long_df: pd.DataFrame, unit_label: str, key_prefix: str = ""):
    st.markdown("### 📊 계획대비 실적 요약 — 그룹별 분석")

    if long_df.empty:
        st.info("데이터가 없습니다.")
        return

    sel_year, sel_month, agg_mode, years_all = render_section_selector(long_df, "연간 계획대비 실적 요약", key_prefix + "summary_base_")
    mode_tag = "당월" if agg_mode == "당월" else f"1~{sel_month}월 연 누적"

    col2, col3 = st.columns([2, 1.5])
    with col2:
        view_mode = st.radio("표시 기준", ["그룹별 합계", "그룹·용도 세부"], index=0, horizontal=True, key=f"{key_prefix}summary_mode")
    with col3:
        include_prev = st.toggle("(Y-1) 포함", value=False, key=f"{key_prefix}summary_prev")

    base_this = apply_period_filter(long_df, sel_year, sel_month, agg_mode)
    if base_this.empty:
        st.info("선택 기준에 데이터가 없습니다.")
        return

    prev_year = sel_year - 1
    base_prev = apply_period_filter(long_df, prev_year, sel_month, agg_mode) if (include_prev and prev_year in years_all) else pd.DataFrame([])
    if not base_prev.empty: base_prev = base_prev[base_prev["계획/실적"] == "실적"]

    if view_mode == "그룹별 합계":
        grp_this = base_this.groupby(["그룹", "계획/실적"], as_index=False)["값"].sum()
        idx_col = "그룹"
        grp_prev = base_prev.groupby("그룹", as_index=False)["값"].sum().rename(columns={"값": "전년실적"}) if not base_prev.empty else pd.DataFrame([])
    else:
        base_this["그룹/용도"] = base_this["그룹"] + " / " + base_this["용도"]
        grp_this = base_this.groupby(["그룹/용도", "계획/실적"], as_index=False)["값"].sum()
        idx_col = "그룹/용도"
        if not base_prev.empty:
            base_prev["그룹/용도"] = base_prev["그룹"] + " / " + base_prev["용도"]
            grp_prev = base_prev.groupby("그룹/용도", as_index=False)["값"].sum().rename(columns={"값": "전년실적"})
        else: grp_prev = pd.DataFrame([])

    pivot = grp_this.pivot(index=idx_col, columns="계획/실적", values="값").fillna(0.0)
    for c in ["계획", "실적"]:
        if c not in pivot.columns: pivot[c] = 0.0

    pivot["차이(실적-계획)"] = pivot["실적"] - pivot["계획"]
    pivot["달성률(%)"] = np.where(pivot["계획"] != 0, (pivot["실적"] / pivot["계획"]) * 100.0, np.nan)
    pivot = pivot[["계획", "실적", "차이(실적-계획)", "달성률(%)"]]

    plan_series = grp_this[grp_this["계획/실적"] == "계획"].set_index(idx_col)["값"]
    act_series = grp_this[grp_this["계획/실적"] == "실적"].set_index(idx_col)["값"]
    prev_series = grp_prev.set_index(idx_col)["전년실적"] if not grp_prev.empty else pd.Series(dtype=float)

    cats = sorted(set(plan_series.index) | set(act_series.index) | set(prev_series.index))
    y_plan = [plan_series.get(c, 0.0) for c in cats]
    y_act = [act_series.get(c, 0.0) for c in cats]
    y_prev = [prev_series.get(c, 0.0) for c in cats] if not prev_series.empty else None

    st.markdown(f"#### 📊 {sel_year}년 {mode_tag} 그룹별 계획·실적 막대그래프")
    fig_bar = go.Figure()
    fig_bar.add_bar(x=cats, y=y_plan, name=f"{sel_year} 계획", marker_color=COLOR_PLAN)
    fig_bar.add_bar(x=cats, y=y_act, name=f"{sel_year} 실적", marker_color=COLOR_ACT)
    if include_prev and y_prev is not None:
        fig_bar.add_bar(x=cats, y=y_prev, name=f"{prev_year} 실적", marker_color=COLOR_PREV)

    fig_bar.update_traces(width=0.25, selector=dict(type="bar"))
    fig_bar.update_layout(barmode="group", xaxis_title=idx_col, yaxis_title=f"기준기간 합계 ({unit_label})", margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig_bar, use_container_width=True)

    st.markdown("##### 🔢 기준기간 요약 표")
    pivot_reset = pivot.reset_index()
    styled = center_style(pivot_reset.style.format({"계획": "{:,.0f}", "실적": "{:,.0f}", "차이(실적-계획)": "{:,.0f}", "달성률(%)": "{:,.1f}"}))
    st.dataframe(styled, use_container_width=True, hide_index=True)

def plan_vs_actual_usage_section(long_df: pd.DataFrame, unit_label: str, key_prefix: str = ""):
    st.markdown("### 🧮 계획대비 월별 실적 (용도 선택)")

    if long_df.empty:
        st.info("데이터가 없습니다.")
        return

    sel_year, sel_month, agg_mode, years_all = render_section_selector(long_df, "계획대비 월별 실적", key_prefix + "pv_base_", fixed_mode="연 누적", show_mode=False)
    mode_tag = f"1~{sel_month}월 연 누적"

    groups_all = sorted(g for g in long_df["그룹"].unique() if g is not None)
    available_groups = ["총량"] + [g for g in GROUP_OPTIONS if g != "총량" and g in groups_all]
    
    if not available_groups:
        st.info("선택 가능한 그룹이 없습니다.")
        return

    col1, col2 = st.columns([2, 1.5])
    with col1:
        try:
            sel_group = st.segmented_control("용도(그룹) 선택", available_groups, selection_mode="single", default="총량", key=f"{key_prefix}pv_group")
        except Exception:
            sel_group = st.radio("용도(그룹) 선택", available_groups, index=available_groups.index("총량"), horizontal=True, key=f"{key_prefix}pv_group_radio")
    with col2:
        include_prev = st.toggle("(Y-1) 포함", value=False, key=f"{key_prefix}pv_prev")

    period = st.radio("기간", ["연간", "상반기(1~6월)", "하반기(7~12월)"], index=0, horizontal=False, key=f"{key_prefix}pv_period")

    base = long_df.copy() if sel_group == "총량" else long_df[long_df["그룹"] == sel_group].copy()

    if period == "상반기(1~6월)":
        base = base[base["월"].between(1, 6)]
        period_label = "상반기"
    elif period == "하반기(7~12월)":
        base = base[base["월"].between(7, 12)]
        period_label = "하반기"
    else:
        period_label = "연간"

    base_year = apply_period_filter(base, sel_year, sel_month, agg_mode)
    if base_year.empty:
        st.info("선택 기준에 해당하는 데이터가 없습니다.")
        return

    prev_year = sel_year - 1
    base_prev = apply_period_filter(base, prev_year, sel_month, agg_mode) if (include_prev and prev_year in years_all) else pd.DataFrame([])
    if not base_prev.empty: base_prev = base_prev[base_prev["계획/실적"] == "실적"]

    bars = base_year.groupby(["월", "계획/실적"], as_index=False)["값"].sum().sort_values(["월", "계획/실적"])
    plan_series = bars[bars["계획/실적"] == "계획"].set_index("월")["값"].sort_index()
    actual_series = bars[bars["계획/실적"] == "실적"].set_index("월")["값"].sort_index()

    months_all = sorted(set(plan_series.index) | set(actual_series.index))
    plan_aligned = plan_series.reindex(months_all).fillna(0.0)
    actual_aligned = actual_series.reindex(months_all).fillna(0.0)
    diff_series = actual_aligned - plan_aligned

    fig = go.Figure()
    for status, name, color in [("계획", f"{sel_year}년 계획", COLOR_PLAN), ("실적", f"{sel_year}년 실적", COLOR_ACT)]:
        sub = bars[bars["계획/실적"] == status]
        if not sub.empty: fig.add_bar(x=sub["월"], y=sub["값"], name=name, width=0.25, marker_color=color)

    if include_prev and not base_prev.empty:
        prev_group = base_prev.groupby("월", as_index=False)["값"].sum().sort_values("월")
        fig.add_bar(x=prev_group["월"], y=prev_group["값"], name=f"{prev_year}년 실적", width=0.25, marker_color=COLOR_PREV)

    if len(diff_series) > 0:
        fig.add_scatter(
            x=months_all, y=diff_series.values, mode="lines+markers+text", name="증감(실적-계획)", yaxis="y2",
            line=dict(color=COLOR_DIFF, width=2), marker=dict(color=COLOR_DIFF),
            text=[f"{v:,.0f}" for v in diff_series.values], textposition="top center", textfont=dict(size=11),
        )

    fig.update_layout(
        title=f"{sel_year}년 {sel_group} 판매량 및 증감 ({period_label}, {mode_tag})", xaxis_title="월", yaxis_title=f"판매량 ({unit_label})",
        xaxis=dict(dtick=1), margin=dict(l=10, r=10, t=40, b=10), barmode="group",
        yaxis2=dict(title="증감(실적-계획)", overlaying="y", side="right", showgrid=False),
    )
    st.plotly_chart(fig, use_container_width=True)

def half_year_stacked_section(long_df: pd.DataFrame, unit_label: str, key_prefix: str = ""):
    st.markdown("### 🧱 기간별 용도 누적 실적 (스택형 막대 + 라인)")

    if long_df.empty:
        st.info("데이터가 없습니다.")
        return

    sel_year, sel_month, agg_mode, years_all = render_section_selector(long_df, "기간별 용도 누적 실적", key_prefix + "stack_base_")

    preferred_years = [y for y in [2021, 2022, 2023, 2024, 2025] if y in years_all]
    if sel_year not in preferred_years: preferred_years = preferred_years + [sel_year]
    default_years = preferred_years if preferred_years else [sel_year]

    sel_years = st.multiselect("연도 선택(스택 그래프)", options=years_all, default=default_years, key=f"{key_prefix}stack_years")
    if not sel_years:
        st.info("연도를 한 개 이상 선택해 주세요.")
        return

    period = st.radio("기간", ["연간", "상반기(1~6월)", "하반기(7~12월)"], index=0, horizontal=True, key=f"{key_prefix}period")

    base = long_df[(long_df["연"].isin(sel_years)) & (long_df["계획/실적"] == "실적")].copy()

    if period == "상반기(1~6월)":
        base = base[base["월"].between(1, 6)]
        period_label = "상반기(1~6월)"
    elif period == "하반기(7~12월)":
        base = base[base["월"].between(7, 12)]
        period_label = "하반기(7~12월)"
    else:
        period_label = "연간"

    base = apply_period_filter_multi_years(base, sel_month, agg_mode, sel_years)
    if base.empty:
        st.info("선택 기준에 해당하는 데이터가 없습니다.")
        return

    grp = base.groupby(["연", "그룹"], as_index=False)["값"].sum()
    total_per_year = grp.groupby("연")["값"].transform("sum")
    grp["비중(%)"] = (grp["값"] / total_per_year) * 100
    grp["비중텍스트"] = grp.apply(lambda r: f"({r['값']:,.0f}, {r['비중(%)']:.1f}%)" if r["비중(%)"] >= 1.5 else "", axis=1)

    fig = px.bar(
        grp, x="연", y="값", color="그룹", barmode="stack", text="비중텍스트",
        hover_data={"값": ":,.0f", "비중(%)": ":.1f", "비중텍스트": False}
    )
    fig.update_traces(width=0.4, textposition="inside", insidetextanchor="middle", textangle=0, textfont=dict(size=11), selector=dict(type="bar"))

    total = grp.groupby("연", as_index=False)["값"].sum().rename(columns={"값": "합계"})
    home = grp[grp["그룹"] == "가정용"].groupby("연", as_index=False)["값"].sum().rename(columns={"값": "가정용"})

    fig.add_scatter(
        x=total["연"], y=total["합계"], mode="lines+markers+text", name="합계", line=dict(dash="dash"),
        text=total["합계"].apply(lambda v: f"{v:,.0f}"), textposition="top center", textfont=dict(size=11),
    )

    if not home.empty:
        fig.add_scatter(x=home["연"], y=home["가정용"], mode="lines+markers", name="가정용", line=dict(dash="dot"))

    mode_tag = "당월" if agg_mode == "당월" else f"1~{sel_month}월 연 누적"
    fig.update_layout(title=f"{period_label} 용도별 실적 판매량 ({mode_tag})", xaxis_title="연도", yaxis_title=f"판매량 ({unit_label})", margin=dict(l=10, r=10, t=40, b=10))
    st.plotly_chart(fig, use_container_width=True)

# ─────────────────────────────────────────────────────────
# 메인 앱 레이아웃
# ─────────────────────────────────────────────────────────
def main():
    st.title("📊 판매량 분석 보고서 (판매량 분석 단독 탭)")

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
            st.info("유효한 시트를 찾지 못했습니다. 파일 시트명을 확인해 주세요. ('계획_부피', '실적_부피', '계획_열량', '실적_열량' 등)")
        else:
            tabs = st.tabs(tab_labels)
            for tab_label, tab in zip(tab_labels, tabs):
                with tab:
                    if tab_label.startswith("부피"):
                        df_long = long_dict.get("부피", pd.DataFrame())
                        unit = "천m³"
                        prefix = "sales_vol_"
                    else:
                        df_long = long_dict.get("열량", pd.DataFrame()).copy()
                        unit = "GJ"
                        prefix = "sales_gj_"

                    monthly_core_dashboard(df_long, unit_label=unit, key_prefix=prefix + "dash_")
                    st.markdown("---")
                    st.markdown("## 📊 실적 분석")
                    monthly_trend_section(df_long, unit_label=unit, key_prefix=prefix + "trend_")
                    half_year_stacked_section(df_long, unit_label=unit, key_prefix=prefix + "stack_")
                    st.markdown("---")
                    st.markdown("## 📏 계획대비 분석")
                    yearly_summary_section(df_long, unit_label=unit, key_prefix=prefix + "summary_")
                    plan_vs_actual_usage_section(df_long, unit_label=unit, key_prefix=prefix + "pv_")

if __name__ == "__main__":
    main()
