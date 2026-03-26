import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
import plotly.express as px

# 1. 페이지 설정
st.set_page_config(page_title="자금 관리 시스템 v8", layout="wide", page_icon="💰")

# 2. DB 설정 및 초기화
@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_v8.db', check_same_thread=False)
    c = conn.cursor()
    # 발주 마스터: 유형(type) 필드 추가
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (order_id TEXT PRIMARY KEY, order_date TEXT, vendor TEXT, 
                  product TEXT, category TEXT, currency TEXT, total_amt REAL, is_closed INTEGER DEFAULT 0)''')
    # 입금 내역
    c.execute('''CREATE TABLE IF NOT EXISTS payments 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, order_id TEXT, pay_date TEXT, 
                  deposit REAL, advance REAL, note TEXT, krw_val REAL)''')
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비"]
CURRENCIES = ["KRW", "USD", "CNY"]

# --- 환율 로직 (입금월 기준) ---
def get_exchange_rate(date_obj, currency):
    if currency in ["KRW", "한화"] or not currency: return 1.0
    # 실제 운영 시 월평균 환율 계산 로직으로 대체 가능
    rates = {"USD": 1350.0, "CNY": 190.0} 
    return rates.get(currency, 1.0)

# --- 데이터 로드 ---
def load_orders():
    return pd.read_sql("SELECT * FROM orders", conn)

def load_payments():
    return pd.read_sql("SELECT * FROM payments", conn)

st.title("💰 자금 관리 시스템 v8 (유형별/상태별 관리)")

tab1, tab2, tab3, tab4 = st.tabs(["📥 발주 등록(ERP)", "💸 입금 기록", "🔍 상세 내역 관리", "📊 현황 대시보드"])

# --- Tab 1: 발주서 등록 (유형 구분 포함) ---
with tab1:
    st.header("📄 이카운트 발주서 등록")
    with st.expander("신규 발주 정보 입력 (또는 엑셀 파싱)", expanded=True):
        c1, c2, c3 = st.columns(3)
        order_id = c1.text_input("발주번호 (ERP 전표번호)")
        order_date = c2.date_input("발주일", datetime.now())
        category = c3.selectbox("업무 유형", CATEGORIES)
        
        c4, c5, c6 = st.columns(3)
        vendor = c4.text_input("거래처명")
        product = c5.text_input("상품명")
        curr = c6.selectbox("통화", CURRENCIES)
        
        total_amt = st.number_input("발주 총액 (외화 또는 한화)", min_value=0.0)
        
        if st.button("🚀 발주 마스터 저장", use_container_width=True):
            if order_id and vendor:
                cur = conn.cursor()
                cur.execute("INSERT OR REPLACE INTO orders VALUES (?, ?, ?, ?, ?, ?, ?, 0)", 
                            (order_id, order_date.strftime("%Y-%m-%d"), vendor, product, category, curr, total_amt))
                conn.commit()
                st.success(f"[{category}] {vendor} 발주 건이 등록되었습니다.")
                st.rerun()

# --- Tab 2: 입금 기록 (진행 중인 건만 표시) ---
with tab2:
    st.header("📝 입금 내역 입력")
    orders = load_orders()
    active_orders = orders[orders['is_closed'] == 0]
    
    if not active_orders.empty:
        with st.form("payment_form", clear_on_submit=True):
            # 선택 창에 유형(category)을 함께 표시하여 구분 용이하게 함
            selected_oid = st.selectbox("진행 중인 발주 선택", 
                                        options=active_orders['order_id'],
                                        format_func=lambda x: f"[{active_orders[active_orders['order_id']==x]['category'].values[0]}] {active_orders[active_orders['order_id']==x]['vendor'].values[0]} - {x}")
            
            c1, c2, c3 = st.columns(3)
            p_date = c1.date_input("입금일")
            dep_amt = c2.number_input("이번 입금액", min_value=0.0)
            adv_amt = c3.number_input("선급금 변동", value=0.0)
            p_note = st.text_input("송금 사유")
            
            if st.form_submit_button("💰 입금 저장"):
                order_info = active_orders[active_orders['order_id'] == selected_oid].iloc[0]
                rate = get_exchange_rate(p_date, order_info['currency'])
                krw_val = dep_amt * rate
                
                cur = conn.cursor()
                cur.execute("INSERT INTO payments (order_id, pay_date, deposit, advance, note, krw_val) VALUES (?, ?, ?, ?, ?, ?)",
                            (selected_oid, p_date.strftime("%Y-%m-%d"), dep_amt, adv_amt, p_note, krw_val))
                conn.commit()
                st.success("입금 내역이 DB에 반영되었습니다.")
    else:
        st.info("현재 진행 중인(미마감) 발주 건이 없습니다.")

# --- Tab 3: 상세 내역 관리 (진행/마감 분리 및 회색 음영) ---
with tab3:
    st.header("🔍 상세 내역 및 마감 처리")
    
    all_orders = load_orders()
    all_payments = load_payments()
    
    col_filter1, col_filter2 = st.columns(2)
    view_status = col_filter1.radio("상태 필터", ["진행 중", "마감 완료"], horizontal=True)
    view_category = col_filter2.multiselect("유형 필터", CATEGORIES, default=CATEGORIES)
    
    is_closed_flag = 1 if view_status == "마감 완료" else 0
    display_orders = all_orders[(all_orders['is_closed'] == is_closed_flag) & (all_orders['category'].isin(view_category))]
    
    if not display_orders.empty:
        for idx, row in display_orders.iterrows():
            # 마감 건은 시각적으로 구분 (Expand 배경색 등은 테마에 따라 다름)
            title_prefix = "✅ [마감]" if is_closed_flag == 1 else "⏳ [진행]"
            with st.expander(f"{title_prefix} {row['category']} | {row['vendor']} | {row['product']} ({row['order_id']})"):
                
                # 해당 발주 건의 모든 입금 내역
                related_pays = all_payments[all_payments['order_id'] == row['order_id']]
                total_deposit = related_pays['deposit'].sum()
                balance = row['total_amt'] - total_deposit
                
                # 요약 지표
                m1, m2, m3 = st.columns(3)
                m1.metric("발주 총액", f"{row['total_amt']:,.2f} {row['currency']}")
                m2.metric("누적 입금", f"{total_deposit:,.2f}")
                m3.metric("미결제 잔액", f"{balance:,.2f}", delta=-total_deposit)
                
                # 상세 내역 테이블
                st.dataframe(related_pays.drop(columns=['id', 'order_id']), use_container_width=True)
                
                # 마감/해제 버튼
                if is_closed_flag == 0:
                    if st.button(f"🚩 마감하기 ({row['order_id']})"):
                        conn.cursor().execute("UPDATE orders SET is_closed = 1 WHERE order_id = ?", (row['order_id'],))
                        conn.commit()
                        st.rerun()
                else:
                    st.caption("이 건은 마감되어 수정이 제한됩니다.")
                    if st.button(f"🔓 마감 취소 ({row['order_id']})"):
                        conn.cursor().execute("UPDATE orders SET is_closed = 0 WHERE order_id = ?", (row['order_id'],))
                        conn.commit()
                        st.rerun()
    else:
        st.write("해당 조건의 데이터가 없습니다.")

# --- Tab 4: 현황 대시보드 (유형별 합계) ---
with tab4:
    st.header("📊 업무 유형별 지출 현황")
    if not all_payments.empty:
        # 발주 정보와 입금 내역 결합
        df_dash = all_payments.merge(all_orders[['order_id', 'category', 'vendor']], on='order_id')
        
        fig = px.pie(df_dash, values='krw_val', names='category', title="유형별 지출 비중 (한화 환산 기준)")
        st.plotly_chart(fig, use_container_width=True)
        
        fig2 = px.bar(df_dash, x='category', y='krw_val', color='vendor', title="유형별/거래처별 지출 금액")
        st.plotly_chart(fig2, use_container_width=True)