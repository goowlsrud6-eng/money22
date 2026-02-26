import streamlit as st
import pandas as pd
from datetime import datetime

# 페이지 설정
st.set_page_config(page_title="통합 자금 관리 시스템", layout="wide")

# 1. 초기 데이터 구조 설정
if 'history_logs' not in st.session_state:
    st.session_state.history_logs = []

# 대분류 데이터 정의
CATEGORIES = ["건기식", "사입", "제작(국내)", "수입(외화)", "물류비", "물품대"]

# 탭 구성
tab1, tab2 = st.tabs(["📝 입금 기록 및 관리", "📊 월별 대시보드"])

# --- Tab 1: 입금 기록 및 관리 ---
with tab1:
    st.header("새로운 입금 내역 기록")
    
    with st.container(border=True):
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            category = st.selectbox("대분류", CATEGORIES)
        with col2:
            vendor = st.text_input("업체명", placeholder="예: 우일코리아")
        with col3:
            product = st.text_input("상품명", placeholder="예: 여름 파자마")
        with col4:
            order_type = st.text_input("발주유형/차수", placeholder="예: 초도, 리오더 1차")

        col5, col6, col7, col8 = st.columns(4)
        with col5:
            pay_date = st.date_input("입금일", datetime.now())
        with col6:
            deposit = st.number_input("실제 입금 금액", value=0, step=1000)
        with col7:
            advance = st.number_input("선급금 처리 (+적립/-차감)", value=0, step=1000)
        with col8:
            note = st.text_input("송금 사유/메모")

        if st.button("🚀 기록 저장하기", use_container_width=True):
            new_log = {
                "입금일": pay_date.strftime("%Y-%m-%d"),
                "월": pay_date.strftime("%Y-%m"),
                "대분류": category,
                "업체명": vendor,
                "상품명": product,
                "발주유형": order_type,
                "실제입금액": deposit,
                "선급금변동": advance,
                "비고": note
            }
            st.session_state.history_logs.append(new_log)
            st.success("데이터가 성공적으로 저장되었습니다!")
            st.rerun()

    st.divider()
    
    # --- 필터링 구역 ---
    st.subheader("🔍 내역 검색 및 필터")
    if st.session_state.history_logs:
        df_full = pd.DataFrame(st.session_state.history_logs)
        
        # 필터 레이아웃
        f_col1, f_col2, f_col3 = st.columns(3)
        
        with f_col1:
            selected_cats = st.multiselect("대분류 필터", options=CATEGORIES, default=[])
        with f_col2:
            all_vendors = sorted(df_full["업체명"].unique())
            selected_vendors = st.multiselect("업체명 필터", options=all_vendors, default=[])
        with f_col3:
            all_products = sorted(df_full["상품명"].unique())
            selected_products = st.multiselect("상품명 필터", options=all_products, default=[])
            
        # 데이터 필터링 로직
        filtered_df = df_full.copy()
        if selected_cats:
            filtered_df = filtered_df[filtered_df["대분류"].isin(selected_cats)]
        if selected_vendors:
            filtered_df = filtered_df[filtered_df["업체명"].isin(selected_vendors)]
        if selected_products:
            filtered_df = filtered_df[filtered_df["상품명"].isin(selected_products)]
            
        st.write(f"✅ 검색 결과: {len(filtered_df)} 건")
        
        # --- 수정 및 관리 표 ---
        edited_df = st.data_editor(
            filtered_df,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "대분류": st.column_config.SelectboxColumn(options=CATEGORIES),
                "입금일": st.column_config.DateColumn(),
                "실제입금액": st.column_config.NumberColumn(format="%d 원"),
                "선급금변동": st.column_config.NumberColumn(format="%d 원"),
            },
            key="history_editor"
        )
        
        # 변경사항 저장 로직 (필터링된 상태에서의 수정을 원본에 반영)
        if st.button("💾 변경사항 최종 저장 (필터 적용 상태 포함)", type="primary"):
            # 필터링되지 않은 나머지 데이터와 수정된 데이터를 합치는 작업
            # (단순화를 위해 현재 에디터의 내용을 원본 인덱스에 맞춰 업데이트하거나 전체 교체)
            # 여기서는 편의상 필터링된 결과만 수정했을 때도 전체 로그에 반영되도록 처리합니다.
            
            # 1. 원본 데이터프레임에서 필터링되지 않은 데이터들 추출
            non_filtered_df = df_full.drop(filtered_df.index)
            # 2. 수정된 데이터와 합치기
            final_df = pd.concat([non_filtered_df, edited_df]).sort_index()
            st.session_state.history_logs = final_df.to_dict('records')
            st.success("변경사항이 성공적으로 반영되었습니다!")
            st.rerun()
            
        # 다운로드 버튼 (필터링된 결과만 다운로드)
        csv = filtered_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 현재 필터링된 내역 다운로드",
            data=csv,
            file_name=f"입금내역_필터결과_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
        )
    else:
        st.info("기록된 내역이 없습니다.")

# --- Tab 2: 월별 대시보드 ---
with tab2:
    st.header("📅 월별 지출 요약")
    if st.session_state.history_logs:
        df_dash = pd.DataFrame(st.session_state.history_logs)
        available_months = sorted(df_dash["월"].unique(), reverse=True)
        selected_month = st.selectbox("조회할 월 선택", available_months)
        
        month_df = df_dash[df_dash["월"] == selected_month]
        
        m_col1, m_col2, m_col3 = st.columns(3)
        m_col1.metric(f"{selected_month} 총 실제입금", f"{month_df['실제입금액'].sum():,} 원")
        m_col2.metric("선급금 변동 합계", f"{month_df['선급금변동'].sum():,} 원")
        m_col3.metric("기록 건수", f"{len(month_df)} 건")
        
        st.divider()
        st.subheader(f"📍 {selected_month} 카테고리별 지출 현황")
        cat_summary = month_df.groupby("대분류")["실제입금액"].sum().reset_index()
        st.table(cat_summary)
    else:
        st.info("데이터가 충분하지 않습니다.")