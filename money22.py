import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="통합 자금 관리 시스템 v2", layout="wide")

if 'history_logs' not in st.session_state:
    st.session_state.history_logs = []

CATEGORIES = ["건기식", "사입", "제작(국내)", "수입(외화)", "물류비", "물품대"]
CURRENCIES = ["KRW", "USD", "CNY"]

tab1, tab2, tab3 = st.tabs(["📝 내역 입력", "📊 잔액 및 월별 대시보드", "📂 엑셀 업로드"])

# --- Tab 1: 내역 입력 ---
with tab1:
    st.header("입금 및 정산 기록")
    with st.container(border=True):
        c1, c2, c3, c4 = st.columns(4)
        category = c1.selectbox("대분류", CATEGORIES)
        vendor = c2.text_input("업체명")
        product = c3.text_input("상품명")
        order_type = c4.text_input("발주유형/차수")

        c5, c6, c7, c8 = st.columns(4)
        currency = c5.selectbox("통화", CURRENCIES)
        total_order_amt = c6.number_input("해당 건 발주 총액", value=0.0)
        pay_date = c7.date_input("입금일", datetime.now())
        exchange_rate = 1.0
        if currency != "KRW":
            exchange_rate = c8.number_input("적용 환율 (1외화당 KRW)", value=1.0)
        else:
            c8.write("한화 결제 (환율 1.0)")

        c9, c10, c11 = st.columns(3)
        deposit = c9.number_input("실제 입금 금액 (외화면 외화기준)", value=0.0)
        advance = c10.number_input("선급금 처리 (+적립/-차감)", value=0.0)
        note = c11.text_input("메모")

        if st.button("🚀 기록 저장", use_container_width=True):
            new_log = {
                "입금일": pay_date.strftime("%Y-%m-%d"),
                "월": pay_date.strftime("%Y-%m"),
                "대분류": category, "업체명": vendor, "상품명": product, "발주유형": order_type,
                "통화": currency, "발주총액": total_order_amt, "환율": exchange_rate,
                "입금액(원화환산)": deposit * exchange_rate,
                "실제입금액": deposit, "선급금변동": advance, "비고": note
            }
            st.session_state.history_logs.append(new_log)
            st.rerun()

# --- Tab 2: 잔액 대시보드 ---
with tab2:
    if st.session_state.history_logs:
        df = pd.DataFrame(st.session_state.history_logs)
        
        # [핵심] 업체/상품별 잔액 계산
        st.subheader("🚩 품목별 정산 잔액 현황")
        # 같은 품목(업체+상품+차수)끼리 묶어서 계산
        summary = df.groupby(['업체명', '상품명', '발주유형', '통화', '발주총액']).agg({
            '실제입금액': 'sum',
            '선급금변동': 'sum'
        }).reset_index()
        
        summary['잔금(외화/원화)'] = summary['발주총액'] - summary['실제입금액']
        st.dataframe(summary, use_container_width=True)

        st.divider()
        st.subheader("📋 전체 내역 수정 및 관리")
        edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")
        if st.button("💾 변경사항 저장"):
            st.session_state.history_logs = edited_df.to_dict('records')
            st.rerun()
    else:
        st.info("데이터가 없습니다.")

# --- Tab 3: 엑셀 업로드 ---
with tab3:
    st.header("엑셀 일괄 업로드")
    st.write("아래 버튼으로 샘플 양식을 다운받아 작성 후 업로드하세요.")
    
    # 샘플 데이터 생성
    sample_data = pd.DataFrame([{
        "입금일": "2026-02-26", "대분류": "제작(국내)", "업체명": "우일코리아", "상품명": "에어메쉬",
        "발주유형": "초도", "통화": "KRW", "발주총액": 1000000, "환율": 1, "실제입금액": 300000, "선급금변동": 300000, "비고": "선급금지불"
    }])
    st.download_button("📥 엑셀 양식 다운로드", sample_data.to_csv(index=False).encode('utf-8-sig'), "template.csv", "text/csv")
    
    uploaded_file = st.file_uploader("작성한 파일을 선택하세요", type=['csv', 'xlsx'])
    if uploaded_file:
        up_df = pd.read_csv(uploaded_file) if uploaded_file.name.endswith('.csv') else pd.read_excel(uploaded_file)
        if st.button("파일 데이터 추가하기"):
            # 기존 데이터에 합치기
            up_df['월'] = up_df['입금일'].apply(lambda x: x[:7])
            up_df['입금액(원화환산)'] = up_df['실제입금액'] * up_df['환율']
            st.session_state.history_logs.extend(up_df.to_dict('records'))
            st.success("업로드 완료!")
            st.rerun()