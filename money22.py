import streamlit as st
import pandas as pd
from datetime import datetime

# 페이지 설정
st.set_page_config(page_title="통합 자금 관리 시스템", layout="wide")

# 1. 초기 데이터 구조 설정 (데이터베이스 대용)
if 'history_logs' not in st.session_state:
    st.session_state.history_logs = []

# 대분류 데이터 정의
CATEGORIES = ["건기식", "사입", "제작(국내)", "수입(외화)", "물류비", "물품대"]

# 탭 구성: 입금 기록창 / 월별 대시보드
tab1, tab2 = st.tabs(["📝 입금 기록창", "📊 월별 대시보드"])

# --- Tab 1: 입금 기록창 ---
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

    st.divider()
    
    # 최근 기록 리스트 및 다운로드
    st.subheader("📋 전체 입금 내역")
    if st.session_state.history_logs:
        df_logs = pd.DataFrame(st.session_state.history_logs)
        st.dataframe(df_logs, use_container_width=True)
        
        # 다운로드 버튼
        csv = df_logs.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 전체 내역 엑셀(CSV) 다운로드",
            data=csv,
            file_name=f"입금내역_추출_{datetime.now().strftime('%Y%m%d')}.csv",
            mime="text/csv",
        )
    else:
        st.info("기록된 내역이 없습니다.")

# --- Tab 2: 월별 대시보드 ---
with tab2:
    st.header("📅 월별 지출 요약")
    
    if st.session_state.history_logs:
        df_dash = pd.DataFrame(st.session_state.history_logs)
        
        # 월별 선택 필터
        available_months = sorted(df_dash["월"].unique(), reverse=True)
        selected_month = st.selectbox("조회할 월 선택", available_months)
        
        month_df = df_dash[df_dash["월"] == selected_month]
        
        # 요약 통계 카드
        m_col1, m_col2, m_col3 = st.columns(3)
        m_col1.metric(f"{selected_month} 총 실제입금", f"{month_df['실제입금액'].sum():,} 원")
        m_col2.metric("선급금 사용/적립 합계", f"{month_df['선급금변동'].sum():,} 원")
        m_col3.metric("기록 건수", f"{len(month_df)} 건")
        
        st.divider()
        st.subheader(f"📍 {selected_month} 카테고리별 지출")
        # 카테고리별 합계 표
        cat_summary = month_df.groupby("대분류")["실제입금액"].sum().reset_index()
        st.table(cat_summary)
    else:
        st.info("데이터가 충분하지 않아 대시보드를 생성할 수 없습니다.")