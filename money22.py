import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
import plotly.express as px

# 1. 페이지 설정
st.set_page_config(page_title="자금 관리 시스템 v5", layout="wide", page_icon="💰")

# 2. DB 설정 및 초기화 (캐시 처리로 속도 향상)
@st.cache_resource
def get_db_connection():
    # 데이터베이스 파일명을 바꿔서 깨끗한 상태로 시작하고 싶다면 이름을 변경하세요.
    conn = sqlite3.connect('finance_management_v3.db', check_same_thread=False)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS history 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, pay_date TEXT, month TEXT, 
                  category TEXT, vendor TEXT, product TEXT, order_type TEXT, 
                  currency TEXT, total_order_amt REAL, deposit REAL, 
                  advance_change REAL, note TEXT)''')
    conn.commit()
    return conn

conn = get_db_connection()

# 데이터 로드 함수
def load_data():
    return pd.read_sql("SELECT * FROM history", conn)

df_all = load_data()
CATEGORIES = ["건기식", "사입", "제작(국내)", "수입(외화)", "물류비", "물품대"]
CURRENCIES = ["KRW", "USD", "CNY"]

# 메인 화면 구성
st.title("💰 자금 관리 시스템 v5")
tab1, tab2, tab3, tab4 = st.tabs(["📝 입금/정산 입력", "📊 실시간 대시보드", "🔍 상세 내역 관리", "📂 엑셀 일괄 업로드"])

# --- Tab 1: 입금/정산 입력 ---
with tab1:
    st.header("✨ 내역 등록")
    with st.form("my_input_form", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns(4)
        category = c1.selectbox("대분류", CATEGORIES)
        vendor = c2.text_input("업체명")
        product = c3.text_input("상품명")
        order_type = c4.text_input("발주유형")

        st.markdown("---")
        c5, c6, c7 = st.columns(3)
        currency = c5.selectbox("통화", CURRENCIES)
        total_order_amt = c6.number_input("해당 건 발주 총액", min_value=0.0)
        pay_date = c7.date_input("입금일", datetime.now())

        c8, c9, c10 = st.columns(3)
        deposit = c8.number_input("이번 실제 입금액", min_value=0.0)
        advance_change = c9.number_input("선급금 변동", value=0.0)
        note = c10.text_input("메모")

        # 제출 버튼 (폼 안에 위치)
        submitted = st.form_submit_button("🚀 데이터 저장하기", use_container_width=True)
        
        if submitted:
            if vendor and product:
                cur = conn.cursor()
                cur.execute('''INSERT INTO history (pay_date, month, category, vendor, product, 
                               order_type, currency, total_order_amt, deposit, advance_change, note) 
                               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                            (pay_date.strftime("%Y-%m-%d"), pay_date.strftime("%Y-%m"), 
                             category, vendor, product, order_type, currency, 
                             total_order_amt, deposit, advance_change, note))
                conn.commit()
                st.success("데이터가 성공적으로 저장되었습니다!")
                st.rerun()
            else:
                st.error("업체명과 상품명은 꼭 입력해주세요!")

# --- Tab 2: 실시간 대시보드 ---
with tab2:
    if not df_all.empty:
        st.subheader("📊 월별 입금액 현황")
        chart_df = df_all.groupby(['month', 'currency'])['deposit'].sum().reset_index()
        fig = px.bar(chart_df, x='month', y='deposit', color='currency', barmode='group',
                     labels={'deposit':'입금 합계', 'month':'기준월'})
        st.plotly_chart(fig, use_container_width=True)
        
        st.divider()
        st.subheader("🚩 업체별 정산 현황")
        summary = df_all.groupby(['업체명', '상품명', 'currency']).agg({
            'total_order_amt': 'max',
            'deposit': 'sum'
        }).reset_index()
        summary['미결제 잔액'] = summary['total_order_amt'] - summary['deposit']
        st.dataframe(summary, use_container_width=True)
    else:
        st.info("입력된 데이터가 없습니다.")

# --- Tab 3: 상세 내역 관리 ---
with tab3:
    st.header("🔍 내역 수정 및 삭제")
    if not df_all.empty:
        edited = st.data_editor(df_all, use_container_width=True, num_rows="dynamic")
        if st.button("💾 수정사항 저장하기"):
            edited.to_sql('history', conn, if_exists='replace', index=False)
            st.success("데이터가 업데이트되었습니다.")
            st.rerun()
    else:
        st.info("표시할 데이터가 없습니다.")

# --- Tab 4: 📂 엑셀 일괄 업로드 ---
with tab4:
    st.header("📂 엑셀(CSV) 업로드")
    
    # 1. 양식 제공
    st.subheader("1. 양식 다운로드")
    csv_sample = pd.DataFrame(columns=["pay_date", "category", "vendor", "product", "order_type", "currency", "total_order_amt", "deposit", "advance_change", "note"])
    st.download_button(
        label="📥 업로드용 CSV 양식 다운로드",
        data=csv_sample.to_csv(index=False).encode('utf-8-sig'),
        file_name="finance_template.csv",
        mime="text/csv"
    )
    
    st.divider()
    
    # 2. 업로드
    st.subheader("2. 파일 업로드")
    file = st.file_uploader("작성한 CSV 파일을 선택하세요", type=['csv'])
    if file:
        try:
            up_df = pd.read_csv(file)
            st.write("미리보기:", up_df.head())
            if st.button("✅ 데이터 일괄 추가하기"):
                up_df['month'] = pd.to_datetime(up_df['pay_date']).dt.strftime('%Y-%m')
                up_df.to_sql('history', conn, if_exists='append', index=False)
                st.success(f"{len(up_df)}건 추가 완료!")
                st.rerun()
        except Exception as e:
            st.error(f"오류가 발생했습니다: {e}")