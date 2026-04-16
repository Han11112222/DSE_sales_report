import io
from pathlib import Path
from typing import Dict, List, Optional

import pandas as pd
import matplotlib as mpl
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
    "총량", "가정용", "영업용", "업무용", "산업용", "수송용", "열병합", "연료전지", "열전용설비용"
]

# ─────────────────────────────────────────────────────────
# 공통 데이터 유틸
# ─────────────────────────────────────────────────────────
def _clean_base(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if "Unnamed: 0" in out.columns:
        out = out.drop(columns=["Unnamed: 0"])
    out["연"] = pd.to_numeric(out["연"], errors="coerce").astype("Int64")
    out["월"] = pd.to_numeric(out["월"], errors="coerce").astype("Int64")
    return out

def keyword_group(col: str) -> Optional[str]:
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
# 핵심 UI 컴포넌트 : 최근 3년 동향 분석 (실적 + 계획 연동)
# ─────────────────────────────────────────────────────────
def three_year_trend_section(long_df: pd.DataFrame, unit_label: str, key_prefix: str = ""):
    st.markdown("### 📈 최근 3년간 용도별 실적 및 향후 계획")

    if long_df.empty:
        st.info("데이터가 없습니다.")
        return

    # 기준 선택 UI
    years_all = sorted(long_df["연"].dropna().unique().tolist())
    default_year = years_all[-1] if years_all else 2026

    st.markdown("#### ✅ 분석 기준 선택")
    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        sel_year = st.selectbox("기준 연도", options=years_all, index=years_all.index(default_year) if default_year in years_all else 0, key=f"{key_prefix}year")
    with c2:
        sel_month = st.selectbox("실적 마감 월 (이후는 계획으로 표시)", options=list(range(1, 13)), index=2, key=f"{key_prefix}month") 
    with c3:
        sel_group = st.selectbox("용도(그룹) 선택", options=GROUP_OPTIONS, index=0, key=f"{key_prefix}group")

    # 데이터 필터링 및 집계
    df_g = long_df[long_df["그룹"] == sel_group] if sel_group != "총량" else long_df.copy()
    grp_df = df_g.groupby(["연", "월", "계획/실적"], as_index=False)["값"].sum()

    y3 = sel_year - 2
    y2 = sel_year - 1
    y1 = sel_year

    y3_df = grp_df[(grp_df["연"] == y3) & (grp_df["계획/실적"] == "실적")].sort_values("월")
    y2_df = grp_df[(grp_df["연"] == y2) & (grp_df["계획/실적"] == "실적")].sort_values("월")
    
    # 당해 연도 실적 (1월 ~ 선택 월)
    y1_act = grp_df[(grp_df["연"] == y1) & (grp_df["계획/실적"] == "실적") & (grp_df["월"] <= sel_month)].sort_values("월")
    # 당해 연도 계획 (선택 월 초과 ~ 12월)
    y1_plan_sub = grp_df[(grp_df["연"] == y1) & (grp_df["계획/실적"] == "계획") & (grp_df["월"] > sel_month)].sort_values("월")

    # 차트 그리기
    fig = go.Figure()

    # 과거 2년 실적 (연한 회색 계열)
    if not y3_df.empty:
        fig.add_trace(go.Scatter(x=y3_df["월"], y=y3_df["값"], mode='lines+markers', name=f"{y3}년 실적", line=dict(color='#CBD5E0', width=2)))
    if not y2_df.empty:
        fig.add_trace(go.Scatter(x=y2_df["월"], y=y2_df["값"], mode='lines+markers', name=f"{y2}년 실적", line=dict(color='#A0AEC0', width=2)))

    # 당해 연도 실적 (진한 파란색)
    if not y1_act.empty:
        fig.add_trace(go.Scatter(x=y1_act["월"], y=y1_act["값"], mode='lines+markers', name=f"{y1}년 실적(1~{sel_month}월)", line=dict(color='#2B6CB0', width=3.5)))

    # 당해 연도 계획 (점선, 실적의 마지막 포인트부터 자연스럽게 이어지도록 처리)
    if not y1_plan_sub.empty:
        last_act = y1_act[y1_act["월"] == sel_month]
        if not last_act.empty:
            y1_plan_conn = pd.concat([last_act, y1_plan_sub])
        else:
            y1_plan_conn = y1_plan_sub
            
        fig.add_trace(go.Scatter(x=y1_plan_conn["월"], y=y1_plan_conn["값"], mode='lines+markers', name=f"{y1}년 계획({sel_month+1}~12월)", line=dict(color='#2B6CB0', width=3.5, dash='dot')))

    fig.update_layout(
        xaxis=dict(dtick=1, title="월"),
        yaxis=dict(title=f"판매량 ({unit_label})"),
        margin=dict(l=10, r=10, t=50, b=10),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        hovermode="x unified"
    )

    st.plotly_chart(fig, use_container_width=True)

    # ─────────────────────────────────────────────────────────
    # 하단 데이터 표 구성
    # ─────────────────────────────────────────────────────────
    st.markdown("##### 🔢 월별 상세 데이터 (실적 및 계획)")
    months = list(range(1, 13))
    table_data = {"월": months}

    def get_vals(df, m_list):
        v_dict = dict(zip(df["월"], df["값"]))
        return [v_dict.get(m, 0.0) for m in m_list]

    if not y3_df.empty: table_data[f"{y3}년 실적"] = get_vals(y3_df, months)
    if not y2_df.empty: table_data[f"{y2}년 실적"] = get_vals(y2_df, months)

    # 당해 연도는 실적(1~N월) + 계획(N+1~12월) 데이터를 결합하여 표기
    y1_combined = pd.concat([y1_act, y1_plan_sub])
    if not y1_combined.empty:
        table_data[f"{y1}년 (실적+계획)"] = get_vals(y1_combined, months)

    df_table = pd.DataFrame(table_data)
    
    # 보기 편하게 가로로 변환 (월이 컬럼이 되도록)
    df_t = df_table.set_index("월").T
    df_t.columns = [f"{m}월" for m in df_t.columns]
    df_t["합계"] = df_t.sum(axis=1)

    styled = df_t.style.format("{:,.0f}").set_properties(**{'text-align': 'center'})
    st.dataframe(styled, use_container_width=True)

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

                    # 함수 호출 시 key_prefix 전달
                    three_year_trend_section(df_long, unit_label=unit, key_prefix=prefix)

if __name__ == "__main__":
    main()
