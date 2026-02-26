import streamlit as st
import pandas as pd
from datetime import datetime

# 페이지 설정
st.set_page_config(page_title="통합 자금 관리 시스템 v3", layout="wide")

# 1. 초기 데이터 구조 설정 (데이터베이스 연동 전 세션 상태 유지)
if 'history_logs' not in st.session_state:
    st.session_state.history_logs = []

# 상수 정의
CATEGORIES = ["건기식", "사입", "제작(국내)", "수입(외화)", "물류비", "물품대"]
CURRENCIES = ["KRW", "USD", "CNY"]

# 탭 구성
tab1, tab2, tab3 = st.tabs(["📝 입금/정산 입력", "📊 실시간 잔액 대시보드", "📂 엑셀 일괄 업로드"])

# 데이터프레임 미리 준비
df_all = pd.DataFrame(st.session_state.history_logs) if st.session_state.history_logs else pd.DataFrame()

# --- Tab 1: 입금/정산 입력 ---
with tab1:
    st.header("새로운 입금 내역 작성")
    with st.container(border=True):
        c1, c2, c3, c4 = st.columns(4)
        category = c1.selectbox("대분류", CATEGORIES)
        vendor = c2.text_input("업체명 (예: 우일코리아)")
        product = c3.text_input("상품명 (예: 에어메쉬)")
        order_type = c4.text_input("발주유형 (예: 초도, 리오더1)")

        # 기존에 입력된 발주정보가 있다면 발주총액과 통화를 자동으로 찾아옴
        found_total = 0.0
        found_curr = "KRW"
        if not df_all.empty:
            match = df_all[(df_all['업체명'] == vendor) & (df_all['상품명'] == product) & (df_all['발주유형'] == order_type)]
            if not match.empty:
                found_total = float(match.iloc[-1]['발주총액'])
                found_curr = match.iloc[-1]['통화']

        st.divider()
        c5, c6, c7, c8 = st.columns(4)
        currency = c5.selectbox("통화 선택", CURRENCIES, index=CURRENCIES.index(found_curr))
        total_order_amt = c6.number_input("발주 총액 (확인/수정)", value=found_total)
        pay_date = c7.date_input("입금일", datetime.now())
        
        exchange_rate = 1.0
        if currency != "KRW":
            exchange_rate = c8.number_input(f"적용 환율 (1 {currency} 당 KRW)", value=1.0, format="%.2f")
        else:
            c8.info("한화 결제 (환율 1.0)")

        c9, c10, c11 = st.columns(3)
        deposit = c9.number_input(f"이번 실제 입금액 ({currency})", value=0.0)
        advance_change = c10.number_input(f"선급금 변동 (+적립 / -차감) ({currency})", value=0.0)
        note = c11.text_input("메모/사유")

        if st.button("🚀 기록 저장", use_container_width=True, type="primary"):
            new_log = {
                "입금일": pay_date.strftime("%Y-%m-%d"),
                "월": pay_date.strftime("%Y-%m"),
                "대분류": category, 
                "업체명": vendor, 
                "상품명": product, 
                "발주유형": order_type,
                "통화": currency, 
                "발주총액": total_order_amt, 
                "환율": exchange_rate,
                "실제입금액": deposit, 
                "선급금변동": advance_change,
                "원화환산입금액": deposit * exchange_rate, 
                "비고": note
            }
            st.session_state.history_logs.append(new_log)
            st.success(f"{vendor} - {product} 내역이 저장되었습니다.")
            st.rerun()

# --- Tab 2: 잔액 대시보드 및 수정 ---
with tab2:
    if not df_all.empty:
        # 1. 품목별 요약 (가로형 대시보드)
        st.subheader("🚩 품목별 정산 잔액 현황")
        summary = df_all.groupby(['업체명', '상품명', '발주유형', '통화']).agg({
            '발주총액': 'last',      
            '실제입금액': 'sum',     
            '선급금변동': 'sum'      
        }).reset_index()

        summary['현재 미결제 잔액'] = summary['발주총액'] - summary['실제입금액']
        summary.columns = ['업체명', '상품명', '차수', '통화', '발주총액', '누적 실입금액', '선급금 잔액(보유)', '미결제 잔액']
        st.dataframe(summary.style.format(precision=0), use_container_width=True)

        st.divider()

        # 2. 검색 필터
        st.subheader("🔍 전체 내역 필터 및 수정")
        f_col1, f_col2 = st.columns(2)
        with f_col1:
            f_vendor = st.multiselect("업체명 검색", options=df_all['업체명'].unique())
        with f_col2:
            f_cat = st.multiselect("대분류 검색", options=CATEGORIES)

        filtered_df = df_all.copy()
        if f_vendor:
            filtered_df = filtered_df[filtered_df['업체명'].isin(f_vendor)]
        if f_cat:
            filtered_df = filtered_df[filtered_df['대분류'].isin(f_cat)]
        
        # 3. 데이터 편집기 (수정/삭제)
        edited_df = st.data_editor(
            filtered_df, 
            use_container_width=True, 
            num_rows="dynamic",
            column_config={
                "실제입금액": st.column_config.NumberColumn(format="%.0f"),
                "발주총액": st.column_config.NumberColumn(format="%.0f"),
                "선급금변동": st.column_config.NumberColumn(format="%.0f"),
            }
        )
        
        if st.button("💾 데이터 수정사항 저장"):
            # 필터링된 상태에서 수정된 내용을 원본 세션에 반영
            non_filtered = df_all.drop(filtered_df.index)
            st.session_state.history_logs = pd.concat([non_filtered, edited_df]).to_dict('records')
            st.success("변경사항이 저장되었습니다.")
            st.rerun()
            
        # 다운로드 버튼
        csv = edited_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 현재 리스트 다운로드 (CSV)", csv, f"settlement_{datetime.now().strftime('%Y%m%d')}.csv", "text/csv")
    else:
        st.info("기록된 데이터가 없습니다.")

# --- Tab 3: 엑셀 일괄 업로드 ---
with tab3:
    st.header("📂 엑셀 일괄 업로드")
    st.write("기존 데이터를 한 번에 등록하려면 파일을 업로드하세요.")
    
    with st.expander("📌 업로드 전 필수 확인 (양식 가이드)"):
        st.write("- 필수 열: 입금일, 대분류, 업체명, 상품명, 발주유형, 통화, 발주총액, 환율, 실제입금액, 선급금변동, 비고")
        st.write("- 날짜 형식: YYYY-MM-DD (예: 2026-02-26)")
        st.write("- 선급금: 보낼 때는 (+), 차감할 때는 (-) 금액으로 입력")

    uploaded_file = st.file_uploader("CSV 또는 XLSX 파일 선택", type=['csv', 'xlsx'])
    
    if uploaded_file:
        try:
            if uploaded_file.name.endswith('.csv'):
                up_df = pd.read_csv(uploaded_file)
            else:
                up_df = pd.read_excel(uploaded_file)
            
            st.write("🔍 업로드 데이터 미리보기 (상위 5건):")
            st.dataframe(up_df.head())

            if st.button("✅ 시스템에 데이터 최종 추가"):
                # [날짜 형식 오류 해결 코드]
                # 1. 입금일을 강제로 날짜 데이터로 변환 (잘못된 형식은 NaT로 처리)
                up_df['입금일'] = pd.to_datetime(up_df['입금일'], errors='coerce')
                
                # 2. 날짜에서 월(YYYY-MM) 정보를 안전하게 추출
                up_df['월'] = up_df['입금일'].dt.strftime('%Y-%m')
                
                # 3. 입금일 데이터를 다시 문자열로 정규화 (YYYY-MM-DD)
                up_df['입금일'] = up_df['입금일'].dt.strftime('%Y-%m-%d')
                
                # 4. 원화 환산액 계산
                up_df['원화환산입금액'] = up_df['실제입금액'] * up_df['환율']
                
                # 5. 빈 값(NaN) 처리
                up_df = up_df.fillna("")

                # 6. 기존 데이터와 합치기
                st.session_state.history_logs.extend(up_df.to_dict('records'))
                st.success(f"성공적으로 {len(up_df)}건의 데이터를 추가했습니다.")
                st.rerun()
        except Exception as e:
            st.error(f"업로드 중 오류가 발생했습니다: {e}")