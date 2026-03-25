import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
import plotly.express as px

# 1. 페이지 설정 (가장 상단에 위치해야 함)
st.set_page_config(page_title="자금 관리 시스템 v5", layout="wide", page_icon="💰")

# 2. 보안 설정 (비밀번호)
def check_password():
    if "password_correct" not in st.session_state:
        st.session_state["password_correct"] = False

    if st.session_state["password_correct"]:
        return True

    st.header("🔒 시스템 접속 보안")
    password = st.text_input("접속 비밀번호를 입력하세요", type="password")
    if st.button("로그인"):
        if password == "1234": # <--- 비밀번호 수정 가능
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("비밀번호가 틀렸습니다.")
    return False

# 3. DB 설정 및 초기화
@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_management_v2.db', check_same_thread=False)
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS history 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, pay_date TEXT, month TEXT, 
                  category TEXT, vendor TEXT, product TEXT, order_type TEXT, 
                  currency TEXT, total_order_amt REAL, deposit REAL, 
                  advance_change REAL, note TEXT)''')
    conn.commit()
    return conn

# 프로그램 메인 로직
if check_password():
    conn = get_db_connection()
    
    # 데이터 로드
    def load_data():
        return pd.read_sql("SELECT * FROM history", conn)

    df_all = load_data()
    CATEGORIES = ["건기식", "사입", "제작(국내)", "수입(외화)", "물류비", "물품대"]
    CURRENCIES = ["KRW", "USD", "CNY"]

    tab1, tab2, tab3, tab4 = st.tabs(["📝 입금/정산 입력", "📊 실시간 대시보드", "🔍 상세 내역 관리", "📂 엑셀 일괄 업로드"])

    # --- Tab 1: 입금/정산 입력 ---
    with tab1:
        st.header("✨ 내역 등록")
        # 폼 시작
        with st.form("my_input_form"):
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

            # 에러 방지용 제출 버튼 (반드시 form 안에 위치)
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
                    st.success("성공적으로 저장되었습니다!")
                    st.rerun()
                else:
                    st.error("업체명과 상품명을 입력해주세요.")

    # --- Tab 2: 대시보드 ---
    with tab2:
        if not df_all.empty:
            st.subheader("📊 월별 입금 추이")
            chart_df = df_all.groupby(['month', 'currency'])['deposit'].sum().reset_index()
            fig = px.bar(chart_df, x='month', y='deposit', color='currency', barmode='group')
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("데이터가 없습니다.")

    # --- Tab 3: 내역 관리 ---
    with tab3:
        if not df_all.empty:
            edited = st.data_editor(df_all, use_container_width=True, num_rows="dynamic")
            if st.button("💾 변경사항 저장"):
                edited.to_sql('history', conn, if_exists='replace', index=False)
                st.success("수정 완료!")
                st.rerun()

    # --- Tab 4: 엑셀 업로드 ---
    with tab4:
        st.subheader("📂 데이터 업로드")
        csv_sample = pd.DataFrame(columns=["pay_date", "category", "vendor", "product", "order_type", "currency", "total_order_amt", "deposit", "advance_change", "note"])
        st.download_button("📥 양식 다운로드", csv_sample.to_csv(index=False).encode('utf-8-sig'), "template.csv")
        
        file = st.file_uploader("CSV 선택", type=['csv'])
        if file:
            up_df = pd.read_csv(file)
            if st.button("✅ 일괄 추가하기"):
                up_df['month'] = pd.to_datetime(up_df['pay_date']).dt.strftime('%Y-%m')
                up_df.to_sql('history', conn, if_exists='append', index=False)
                st.success("업로드 완료!")
                st.rerun()