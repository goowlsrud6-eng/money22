import streamlit as st
import pandas as pd
import sqlite3
import os
import shutil
from datetime import datetime
import plotly.express as px

# --- 1. 데이터 안전장치 (자동 백업) ---
def run_backup():
    if not os.path.exists('backups'):
        os.makedirs('backups')
    today_str = datetime.now().strftime("%Y%m%d")
    db_file = 'finance_v12.db'
    backup_file = f"backups/backup_{today_str}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)
        return True
    return False

# 2. 페이지 설정
st.set_page_config(page_title="자금 관리 시스템 v12", layout="wide", page_icon="💰")
run_backup()

# 3. DB 연결 및 테이블 설정
@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_v12.db', check_same_thread=False)
    c = conn.cursor()
    # 발주 마스터
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

# --- 환율 도우미 ---
def get_exchange_rate(currency):
    rates = {"USD": 1350.0, "CNY": 190.0, "한화": 1.0, "KRW": 1.0}
    return rates.get(currency, 1.0)

# --- 메인 UI ---
st.title("💰 자금 관리 통합 시스템 v12")
tabs = st.tabs(["📥 발주서 등록", "📝 입금 건별 입력", "📂 입금 엑셀 업로드", "🔍 상세 정산 관리"])

# --- Tab 1: 발주서 등록 ---
with tabs[0]:
    st.header("📄 발주 마스터 등록")
    with st.form("order_reg"):
        c1, c2, c3 = st.columns(3)
        oid = c1.text_input("발주번호 (ERP 전표번호)")
        odate = c2.date_input("발주일", datetime.now())
        ocat = c3.selectbox("유형", ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비"])
        
        c4, c5, c6 = st.columns(3)
        ovendor = c4.text_input("거래처명")
        oprod = c5.text_input("대표 상품명 (품목 외 n건)")
        ocurr = c6.selectbox("통화", ["한화", "USD", "CNY"])
        
        ototal = st.number_input("발주 합계 금액", min_value=0.0)
        
        if st.form_submit_button("🚀 발주 정보 저장"):
            if oid and ovendor:
                conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?, ?, ?, ?, ?, ?, ?, 0)", 
                                     (oid, odate.strftime("%Y-%m-%d"), ovendor, oprod, ocat, ocurr, ototal))
                conn.commit()
                st.success(f"발주번호 {oid} 저장 완료")

# --- Tab 2: 입금 건별 입력 (살려낸 기능!) ---
with tabs[1]:
    st.header("📝 한 건씩 입금 기록")
    o_df = pd.read_sql("SELECT * FROM orders WHERE is_closed = 0", conn)
    
    if not o_df.empty:
        with st.form("single_payment_form", clear_on_submit=True):
            # 진행 중인 발주 선택
            sel_oid = st.selectbox("연동할 발주번호 선택", options=o_df['order_id'], 
                                   format_func=lambda x: f"{x} | {o_df[o_df['order_id']==x]['vendor'].values[0]} ({o_df[o_df['order_id']==x]['product'].values[0]})")
            
            c1, c2, c3 = st.columns(3)
            p_date = c1.date_input("입금일", datetime.now())
            p_dep = c2.number_input("입금액", min_value=0.0)
            p_adv = c3.number_input("선급금 변동", value=0.0)
            
            p_note = st.text_input("송금 사유 및 메모")
            
            if st.form_submit_button("💾 입금 내역 저장"):
                # 해당 발주의 통화 확인 후 환산
                curr = o_df[o_df['order_id'] == sel_oid]['currency'].values[0]
                rate = get_exchange_rate(curr)
                
                cur = conn.cursor()
                cur.execute('''INSERT INTO payments (order_id, pay_date, deposit, advance, note, krw_val) 
                               VALUES (?, ?, ?, ?, ?, ?)''',
                            (sel_oid, p_date.strftime("%Y-%m-%d"), p_dep, p_adv, p_note, p_dep * rate))
                conn.commit()
                st.success("기록되었습니다!")
    else:
        st.info("현재 진행 중인 발주 건이 없습니다. 발주서를 먼저 등록해 주세요.")

# --- Tab 3: 입금 엑셀 업로드 ---
with tabs[2]:
    st.header("📂 입금 내역 엑셀 업로드")
    template = pd.DataFrame(columns=["발주번호", "입금일", "입금액", "통화", "선급금변동", "송금사유"])
    st.download_button("📥 업로드 양식(CSV) 받기", template.to_csv(index=False).encode('utf-8-sig'), "pay_template.csv")
    
    up_file = st.file_uploader("파일 업로드", type=['csv'])
    if up_file:
        df_up = pd.read_csv(up_file)
        if st.button("✅ 엑셀 데이터 일괄 저장"):
            for _, r in df_up.iterrows():
                rate = get_exchange_rate(r['통화'])
                conn.cursor().execute('''INSERT INTO payments (order_id, pay_date, deposit, advance, note, krw_val) 
                                       VALUES (?, ?, ?, ?, ?, ?)''',
                                    (r['발주번호'], r['입금일'], r['입금액'], r['선급금변동'], r['송금사유'], float(r['입금액']) * rate))
            conn.commit()
            st.success(f"{len(df_up)}건 일괄 저장 완료!")

# --- Tab 4: 상세 정산 관리 ---
with tabs[3]:
    st.header("🔍 상세 정산 현황")
    orders_all = pd.read_sql("SELECT * FROM orders", conn)
    pays_all = pd.read_sql("SELECT * FROM payments", conn)
    
    if not orders_all.empty:
        # 데이터 계산
        p_sum = pays_all.groupby('order_id').agg({'deposit': 'sum', 'advance': 'sum'}).reset_index()
        main_df = orders_all.merge(p_sum, on='order_id', how='left').fillna(0)
        main_df['잔금'] = main_df['total_amt'] - main_df['deposit']
        
        status = st.radio("상태 필터", ["진행 중", "마감 완료"], horizontal=True)
        target = 1 if status == "마감 완료" else 0
        
        disp = main_df[main_df['is_closed'] == target]
        st.data_editor(disp, use_container_width=True)
        
        if status == "진행 중":
            to_close = st.selectbox("마감할 발주번호", disp['order_id'].unique())
            if st.button("🚩 마감 처리"):
                conn.cursor().execute("UPDATE orders SET is_closed = 1 WHERE order_id = ?", (to_close,))
                conn.commit()
                st.rerun()