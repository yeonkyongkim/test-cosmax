import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

st.set_page_config(page_title="COSMAX 시제품 안정성 대시보드", layout="wide")
st.title("COSMAX 시제품 안정성 분석 대시보드")

import os

DATA_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cosmax_day3_dummy_data_numeric.xlsx")

if not os.path.exists(DATA_PATH):
    st.error(f"엑셀 파일을 찾을 수 없습니다: {DATA_PATH}")
    st.stop()

# --- 데이터 로드 ---
@st.cache_data
def load_data(path):
    xl = pd.ExcelFile(path)
    product = pd.read_excel(xl, sheet_name="시제품정보")
    test = pd.read_excel(xl, sheet_name="안정성테스트결과")
    codebook = pd.read_excel(xl, sheet_name="코드북")
    return product, test, codebook

product_raw, test_raw, codebook = load_data(DATA_PATH)

# --- 코드북으로 라벨 매핑 생성 ---
def build_mapping(codebook, sheet, column):
    subset = codebook[(codebook["시트명"] == sheet) & (codebook["컬럼명"] == column)]
    mapping = {}
    for _, row in subset.iterrows():
        code = row["코드"]
        desc = row["설명"]
        if "원본 범주" in desc:
            label = desc.split('"')[1] if '"' in desc else desc
        elif "원본 고유값" in desc:
            label = desc.replace("원본 고유값 ", "").split("에")[0]
        else:
            label = desc
        try:
            mapping[int(code)] = label
        except ValueError:
            pass
    return mapping

# 시제품정보 매핑
product = product_raw.copy()
for col in ["시제품코드", "제품유형", "제형", "개발단계", "목표피부타입", "주요컨셉", "담당팀"]:
    m = build_mapping(codebook, "시제품정보", col)
    if m:
        product[col] = product[col].map(m).fillna(product_raw[col].astype(str))

# 작성일 변환
product["작성일"] = pd.to_datetime(product_raw["작성일"].astype(str), format="%Y%m%d")

# 안정성테스트결과 매핑
test = test_raw.copy()
for col in ["시제품코드", "테스트조건", "보관온도", "보관기간_주", "색상변화등급", "향변화여부", "분리현상여부", "판정결과", "비고"]:
    m = build_mapping(codebook, "안정성테스트결과", col)
    if m:
        test[col] = test[col].map(m).fillna(test_raw[col].astype(str))

# 보관온도를 숫자로 표시 (단위 추가)
test["보관온도"] = test["보관온도"].astype(str) + "°C"
# 보관기간도 숫자 + 주
test["보관기간_주"] = test["보관기간_주"].astype(str) + "주"
# 측정일 변환
test["측정일"] = pd.to_datetime(test_raw["측정일"].astype(str), format="%Y%m%d")
# 색상변화등급 원본값 복원 (0,1,2,3)
color_grade_map = build_mapping(codebook, "안정성테스트결과", "색상변화등급")
# 비고 NaN 처리
test["비고"] = test["비고"].fillna("-")

# ============================================================
# 상단 필터
# ============================================================
st.markdown("---")
st.subheader("필터")
fc1, fc2, fc3 = st.columns(3)
with fc1:
    sel_product = st.multiselect("시제품코드", options=sorted(product["시제품코드"].unique()), default=sorted(product["시제품코드"].unique()))
with fc2:
    sel_type = st.multiselect("제품유형", options=sorted(product["제품유형"].unique()), default=sorted(product["제품유형"].unique()))
with fc3:
    sel_condition = st.multiselect("테스트조건", options=sorted(test["테스트조건"].unique()), default=sorted(test["테스트조건"].unique()))

# 필터 적용
product_f = product[product["시제품코드"].isin(sel_product) & product["제품유형"].isin(sel_type)]
test_f = test[test["시제품코드"].isin(sel_product) & test["테스트조건"].isin(sel_condition)]

# ============================================================
# KPI 카드
# ============================================================
st.markdown("---")
k1, k2, k3, k4 = st.columns(4)
k1.metric("시제품 수", f"{len(product_f)}개")
k2.metric("테스트 건수", f"{len(test_f)}건")
pass_rate = (test_f["판정결과"] == "적합").mean() * 100 if len(test_f) > 0 else 0
k3.metric("적합 판정률", f"{pass_rate:.1f}%")
k4.metric("평균 pH", f"{test_raw.loc[test_f.index, 'pH'].mean():.2f}" if len(test_f) > 0 else "-")

# ============================================================
# 탭 구성
# ============================================================
tab1, tab2, tab4, tab3 = st.tabs(["시제품 현황", "안정성 테스트 분석", "교차분석", "원본 데이터"])

# --- 탭 1: 시제품 현황 ---
with tab1:
    # 제품유형 분포
    st.markdown("#### 제품유형 분포")
    col_chart, col_table = st.columns([3, 2])
    with col_chart:
        fig = px.pie(product_f, names="제품유형", title="", hole=0.4)
        st.plotly_chart(fig, use_container_width=True)
    with col_table:
        tbl = product_f["제품유형"].value_counts().reset_index()
        tbl.columns = ["제품유형", "건수"]
        st.dataframe(tbl, use_container_width=True, hide_index=True)

    # 제형 분포
    st.markdown("#### 제형 분포")
    col_chart, col_table = st.columns([3, 2])
    with col_chart:
        fig = px.pie(product_f, names="제형", title="", hole=0.4)
        st.plotly_chart(fig, use_container_width=True)
    with col_table:
        tbl = product_f["제형"].value_counts().reset_index()
        tbl.columns = ["제형", "건수"]
        st.dataframe(tbl, use_container_width=True, hide_index=True)

    # 담당팀별 개발단계
    st.markdown("#### 담당팀별 개발단계 현황")
    col_chart, col_table = st.columns([3, 2])
    with col_chart:
        fig = px.histogram(product_f, x="담당팀", color="개발단계", barmode="group")
        fig.update_layout(xaxis_title="담당팀", yaxis_title="건수")
        st.plotly_chart(fig, use_container_width=True)
    with col_table:
        tbl = product_f.groupby(["담당팀", "개발단계"]).size().reset_index(name="건수")
        st.dataframe(tbl, use_container_width=True, hide_index=True)

    # 주요컨셉별 목표 피부타입
    st.markdown("#### 주요컨셉별 목표 피부타입")
    col_chart, col_table = st.columns([3, 2])
    with col_chart:
        fig = px.histogram(product_f, x="주요컨셉", color="목표피부타입", barmode="group")
        fig.update_layout(xaxis_title="주요컨셉", yaxis_title="건수")
        st.plotly_chart(fig, use_container_width=True)
    with col_table:
        tbl = product_f.groupby(["주요컨셉", "목표피부타입"]).size().reset_index(name="건수")
        st.dataframe(tbl, use_container_width=True, hide_index=True)

# --- 탭 2: 안정성 테스트 분석 ---
with tab2:
    # 판정결과 분포
    st.markdown("#### 판정결과 분포")
    col_chart, col_table = st.columns([3, 2])
    with col_chart:
        fig = px.histogram(test_f, x="판정결과", color="판정결과",
                           color_discrete_map={"적합": "#2ecc71", "경미변화": "#f39c12", "재검토": "#e74c3c"})
        fig.update_layout(showlegend=False, xaxis_title="판정결과", yaxis_title="건수")
        st.plotly_chart(fig, use_container_width=True)
    with col_table:
        tbl = test_f["판정결과"].value_counts().reset_index()
        tbl.columns = ["판정결과", "건수"]
        st.dataframe(tbl, use_container_width=True, hide_index=True)

    # 테스트조건별 판정결과
    st.markdown("#### 테스트조건별 판정결과")
    col_chart, col_table = st.columns([3, 2])
    with col_chart:
        fig = px.histogram(test_f, x="테스트조건", color="판정결과", barmode="group",
                           color_discrete_map={"적합": "#2ecc71", "경미변화": "#f39c12", "재검토": "#e74c3c"})
        fig.update_layout(xaxis_title="테스트조건", yaxis_title="건수")
        st.plotly_chart(fig, use_container_width=True)
    with col_table:
        tbl = test_f.groupby(["테스트조건", "판정결과"]).size().reset_index(name="건수")
        st.dataframe(tbl, use_container_width=True, hide_index=True)

    # 테스트조건별 점도 분포
    st.markdown("#### 테스트조건별 점도(cP) 분포")
    col_chart, col_table = st.columns([3, 2])
    with col_chart:
        fig = px.box(test_f, x="테스트조건", y=test_raw.loc[test_f.index, "점도_cP"], color="테스트조건")
        fig.update_layout(showlegend=False, xaxis_title="테스트조건", yaxis_title="점도 (cP)")
        st.plotly_chart(fig, use_container_width=True)
    with col_table:
        tbl = test_f.copy()
        tbl["점도_cP"] = test_raw.loc[test_f.index, "점도_cP"]
        tbl_stats = tbl.groupby("테스트조건")["점도_cP"].agg(["mean", "min", "max"]).round(0).astype(int).reset_index()
        tbl_stats.columns = ["테스트조건", "평균", "최소", "최대"]
        st.dataframe(tbl_stats, use_container_width=True, hide_index=True)

    # 테스트조건별 pH 분포
    st.markdown("#### 테스트조건별 pH 분포")
    col_chart, col_table = st.columns([3, 2])
    with col_chart:
        fig = px.box(test_f, x="테스트조건", y=test_raw.loc[test_f.index, "pH"], color="테스트조건")
        fig.update_layout(showlegend=False, xaxis_title="테스트조건", yaxis_title="pH")
        st.plotly_chart(fig, use_container_width=True)
    with col_table:
        tbl = test_f.copy()
        tbl["pH"] = test_raw.loc[test_f.index, "pH"]
        tbl_stats = tbl.groupby("테스트조건")["pH"].agg(["mean", "min", "max"]).round(4).reset_index()
        tbl_stats.columns = ["테스트조건", "평균", "최소", "최대"]
        st.dataframe(tbl_stats, use_container_width=True, hide_index=True)

    # 보관기간별 점도 변화
    st.markdown("#### 시제품별 점도 추이 (보관기간별)")
    test_trend = test_f.copy()
    test_trend["점도_cP"] = test_raw.loc[test_f.index, "점도_cP"]
    test_trend["pH_val"] = test_raw.loc[test_f.index, "pH"]
    col_chart, col_table = st.columns([3, 2])
    with col_chart:
        fig = px.line(test_trend.sort_values("보관기간_주"), x="보관기간_주", y="점도_cP",
                      color="시제품코드", markers=True)
        fig.update_layout(xaxis_title="보관기간", yaxis_title="점도 (cP)")
        st.plotly_chart(fig, use_container_width=True)
    with col_table:
        tbl_pivot = test_trend.pivot_table(index="시제품코드", columns="보관기간_주", values="점도_cP", aggfunc="mean").round(0)
        st.dataframe(tbl_pivot, use_container_width=True)

    # 향변화/분리현상 여부
    st.markdown("#### 향변화 / 분리현상 발생 여부")
    col_chart1, col_table1, col_chart2, col_table2 = st.columns([2, 1, 2, 1])
    with col_chart1:
        fig = px.pie(test_f, names="향변화여부", title="향변화", hole=0.4,
                     color_discrete_map={"N": "#2ecc71", "Y": "#e74c3c"})
        st.plotly_chart(fig, use_container_width=True)
    with col_table1:
        tbl = test_f["향변화여부"].value_counts().reset_index()
        tbl.columns = ["향변화여부", "건수"]
        st.dataframe(tbl, use_container_width=True, hide_index=True)
    with col_chart2:
        fig = px.pie(test_f, names="분리현상여부", title="분리현상", hole=0.4,
                     color_discrete_map={"N": "#2ecc71", "Y": "#e74c3c"})
        st.plotly_chart(fig, use_container_width=True)
    with col_table2:
        tbl = test_f["분리현상여부"].value_counts().reset_index()
        tbl.columns = ["분리현상여부", "건수"]
        st.dataframe(tbl, use_container_width=True, hide_index=True)

# --- 탭 4: 교차분석 ---
with tab4:
    # 시제품정보 + 안정성테스트를 병합
    merged = test_f.merge(product_f[["시제품코드", "제품유형", "제형", "개발단계", "목표피부타입", "주요컨셉", "담당팀"]],
                          on="시제품코드", how="left")
    merged["점도_cP"] = test_raw.loc[test_f.index, "점도_cP"]
    merged["pH"] = test_raw.loc[test_f.index, "pH"]

    # 범주형 컬럼 목록
    cat_cols = ["제품유형", "제형", "개발단계", "목표피부타입", "주요컨셉", "담당팀",
                "테스트조건", "보관온도", "보관기간_주", "판정결과", "향변화여부", "분리현상여부", "색상변화등급"]
    num_cols = ["점도_cP", "pH"]

    st.markdown("### 교차분석")
    st.caption("행과 열에 원하는 변수를 선택하면 교차표와 시각화를 자동으로 생성합니다.")

    cr1, cr2, cr3 = st.columns(3)
    with cr1:
        row_var = st.selectbox("행 (Row)", options=cat_cols, index=0)
    with cr2:
        col_var = st.selectbox("열 (Column)", options=cat_cols, index=cat_cols.index("판정결과"))
    with cr3:
        value_opt = st.selectbox("값 (Value)", options=["건수"] + num_cols, index=0)

    st.markdown(f"**전체 표본 수: {len(merged)}건**")

    if len(merged) == 0:
        st.warning("필터 조건에 해당하는 데이터가 없습니다.")
    else:
        # --- 교차표 생성 ---
        if value_opt == "건수":
            cross = pd.crosstab(merged[row_var], merged[col_var], margins=True, margins_name="합계")
            cross_pct = pd.crosstab(merged[row_var], merged[col_var], normalize="index", margins=False).round(4) * 100
        else:
            cross = merged.pivot_table(index=row_var, columns=col_var, values=value_opt, aggfunc="mean").round(2)
            cross_n = merged.pivot_table(index=row_var, columns=col_var, values=value_opt, aggfunc="count").fillna(0).astype(int)
            cross_pct = None

        # --- 시각화 + 표 ---
        col_chart, col_table = st.columns([3, 2])

        with col_chart:
            if value_opt == "건수":
                # 합계 행/열 제외하고 시각화
                plot_data = cross.drop("합계", axis=0, errors="ignore").drop("합계", axis=1, errors="ignore")
                fig = px.bar(
                    plot_data.reset_index().melt(id_vars=row_var, var_name=col_var, value_name="건수"),
                    x=row_var, y="건수", color=col_var, barmode="group",
                    title=f"{row_var} × {col_var} 교차 (건수)"
                )
                st.plotly_chart(fig, use_container_width=True)
            else:
                fig = px.bar(
                    cross.reset_index().melt(id_vars=row_var, var_name=col_var, value_name=value_opt),
                    x=row_var, y=value_opt, color=col_var, barmode="group",
                    title=f"{row_var} × {col_var} 교차 (평균 {value_opt})"
                )
                st.plotly_chart(fig, use_container_width=True)

        with col_table:
            st.markdown(f"**교차표 ({value_opt})**")
            st.dataframe(cross, use_container_width=True)

        # 표본 수 표 (수치형 선택 시)
        if value_opt != "건수":
            st.markdown("**셀별 표본 수 (N)**")
            st.dataframe(cross_n, use_container_width=True)

        # 비율 표 (건수일 때만) — 첫 열에 표본수 포함
        if cross_pct is not None:
            st.markdown("**행 기준 비율 (%)**")
            cross_no_margin = cross.drop("합계", axis=1, errors="ignore").drop("합계", axis=0, errors="ignore")
            row_n = cross_no_margin.sum(axis=1).astype(int)
            pct_display = cross_pct.round(1).copy()
            pct_display.insert(0, "표본수(N)", row_n)
            st.dataframe(pct_display, use_container_width=True)

        # --- 엑셀 다운로드 ---
        def to_excel():
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                cross.to_excel(writer, sheet_name="교차표")
                if value_opt != "건수":
                    cross_n.to_excel(writer, sheet_name="표본수(N)")
                if cross_pct is not None:
                    cross_pct.round(1).to_excel(writer, sheet_name="행기준비율(%)")
            return buf.getvalue()

        st.download_button(
            label="교차분석 결과 엑셀 다운로드",
            data=to_excel(),
            file_name=f"교차분석_{row_var}x{col_var}_{value_opt}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # --- 히트맵 ---
        st.markdown(f"**히트맵: {row_var} × {col_var}**")
        heat_data = cross.drop("합계", axis=0, errors="ignore").drop("합계", axis=1, errors="ignore") if value_opt == "건수" else cross
        fig_heat = px.imshow(
            heat_data,
            text_auto=True,
            color_continuous_scale="Blues",
            labels=dict(x=col_var, y=row_var, color=value_opt),
            aspect="auto"
        )
        fig_heat.update_layout(height=400)
        st.plotly_chart(fig_heat, use_container_width=True)

# --- 탭 3: 원본 데이터 ---
with tab3:
    st.subheader("시제품정보")
    st.dataframe(product_f, use_container_width=True)
    st.subheader("안정성테스트결과")
    st.dataframe(test_f, use_container_width=True)
