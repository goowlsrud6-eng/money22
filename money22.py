import streamlit as st
import pandas as pd
from datetime import datetime

# 페이지 설정
st.set_page_config(page_title="자금 관리 시스템 v4", layout="wide")

# 1. 데이터 초기화
if 'history_logs' not in st.session_state:
    st.session_state.history_logs = []

CATEGORIES = ["건기식", "사입", "제작(국내)", "수입(외화)", "물류비", "물품대"]
CURRENCIES = ["KRW", "USD", "CNY"]

# 탭 구성
tab1, tab2, tab3 = st.tabs(["📝 입금/정산 입력", "📊 상세 필터 및 대시보드", "📂 엑셀 일괄 업로드"])

# 데이터프레임 전처리
if st.session_state.history_logs:
    df_all = pd.DataFrame(st.session_state.history_logs)
    # 숫자형 데이터 강제 변환
    for col in ['발주총액', '실제입금액', '선급금변동']:
        if col in df_all.columns:
            df_all[col] = pd.to_numeric(df_all[col], errors='coerce').fillna(0)
else:
    df_all = pd.DataFrame()

# --- Tab 1: 입금/정산 입력 ---
with tab1:
    st.header("새로운 입금 내역 작성")
    with st.container(border=True):
        c1, c2, c3, c4 = st.columns(4)
        category = c1.selectbox("대분류", CATEGORIES)
        vendor = c2.text_input("업체명 (예: 우일코리아)")
        product = c3.text_input("상품명 (예: 에어메쉬)")
        order_type = c4.text_input("발주유형 (예: 초도, 리오더1)")

        # 기존 발주 정보 자동 찾기 (발주총액 및 통화)
        found_total = 0.0
        found_curr = "KRW"
        if not df_all.empty:
            match = df_all[(df_all['업체명'] == vendor) & (df_all['상품명'] == product) & (df_all['발주유형'] == order_type)]
            if not match.empty:
                found_total = float(match.iloc[-1]['발주총액'])
                found_curr = match.iloc[-1]['통화']

        st.divider()
        c5, c6, c7 = st.columns(3)
        currency = c5.selectbox("통화", CURRENCIES, index=CURRENCIES.index(found_curr))
        total_order_amt = c6.number_input(f"해당 건 발주 총액 ({currency})", value=found_total)
        pay_date = c7.date_input("입금일", datetime.now())

        c8, c9, c10 = st.columns(3)
        deposit = c8.number_input(f"이번 실제 입금액 ({currency})", value=0.0)
        advance_change = c9.number_input(f"선급금 변동 (+적립/-차감) ({currency})", value=0.0)
        note = c10.text_input("메모/사유")

        if st.button("🚀 기록 저장", use_container_width=True, type="primary"):
            new_log = {
                "입금일": pay_date.strftime("%Y-%m-%d"),
                "월": pay_date.strftime("%Y-%m"),
                "대분류": category, "업체명": vendor, "상품명": product, "발주유형": order_type,
                "통화": currency, "발주총액": total_order_amt,
                "실제입금액": deposit, "선급금변동": advance_change, "비고": note
            }
            st.session_state.history_logs.append(new_log)
            st.success("데이터가 저장되었습니다.")
            st.rerun()

# --- Tab 2: 상세 필터 및 대시보드 ---
with tab2:
    if not df_all.empty:
        # 1. 4단 필터 구역
        st.subheader("🔍 상세 필터 검색")
        f1, f2, f3, f4 = st.columns(4)
        with f1:
            f_cat = st.multiselect("대분류 필터", options=CATEGORIES)
        with f2:
            f_vendor = st.multiselect("업체명 필터", options=sorted(df_all['업체명'].unique()))
        with f3:
            f_product = st.multiselect("품목명 필터", options=sorted(df_all['상품명'].unique()))
        with f4:
            f_order = st.multiselect("발주구분 필터", options=sorted(df_all['발주유형'].unique()))

        # 필터링 로직
        filtered_df = df_all.copy()
        if f_cat: filtered_df = filtered_df[filtered_df['대분류'].isin(f_cat)]
        if f_vendor: filtered_df = filtered_df[filtered_df['업체명'].isin(f_vendor)]
        if f_product: filtered_df = filtered_df[filtered_df['상품명'].isin(f_product)]
        if f_order: filtered_df = filtered_df[filtered_df['발주유형'].isin(f_order)]

        # 2. 요약 요약 (통화별로 구분)
        st.subheader("🚩 통화별 정산 요약")
        summary = filtered_df.groupby(['대분류', '업체명', '상품명', '발주유형', '통화']).agg({
            '발주총액': 'last',
            '실제입금액': 'sum',
            '선급금변동': 'sum'
        }).reset_index()
        summary['미결제 잔액'] = summary['발주총액'] - summary['실제입금액']
        st.dataframe(summary.style.format(precision=0), use_container_width=True)

        st.divider()
        
        # 3. 상세 기록 편집
        st.subheader("📋 상세 입금 내역 (수정 및 삭제 가능)")
        edited_df = st.data_editor(filtered_df, use_container_width=True, num_rows="dynamic")
        
        if st.button("💾 변경사항 최종 저장"):
            non_filtered = df_all.drop(filtered_df.index)
            st.session_state.history_logs = pd.concat([non_filtered, edited_df]).to_dict('records')
            st.success("수정사항이 반영되었습니다.")
            st.rerun()
            
        csv = filtered_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 필터링된 내역 다운로드(CSV)", csv, f"report_{datetime.now().strftime('%Y%m%d')}.csv")
    else:
        st.info("데이터가 없습니다. [내역 입력] 탭에서 기록을 시작하세요.")

# --- Tab 3: 엑셀 일괄 업로드 ---
with tab3:
    st.header("📂 엑셀 일괄 업로드")
    uploaded_file = st.file_uploader("CSV 또는 XLSX 파일", type=['csv', 'xlsx'])
    
    if uploaded_file:
        try:
            up_df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
            if st.button("✅ 데이터 일괄 추가"):
                up_df['입금일'] = pd.to_datetime(up_df['입금일'], errors='coerce')
                up_df['월'] = up_df['입금일'].dt.strftime('%Y-%m')
                up_df['입금일'] = up_df['입금일'].dt.strftime('%Y-%m-%d')
                
                for col in ['발주총액', '실제입금액', '선급금변동']:
                    if col in up_df.columns:
                        up_df[col] = pd.to_numeric(up_df[col], errors='coerce').fillna(0)
                
                st.session_state.history_logs.extend(up_df.fillna("").to_dict('records'))
                st.success(f"{len(up_df)}건 추가 완료!")
                st.rerun()
        except Exception as e:
            st.error(f"오류 발생: {e}")