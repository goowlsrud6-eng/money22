import streamlit as st
import pandas as pd
import sqlite3
import os
import shutil
from datetime import datetime

# --- 1. 데이터 안전장치 (자동 백업) ---
def run_backup():
    if not os.path.exists('backups'):
        os.makedirs('backups')
    today_str = datetime.now().strftime("%Y%m%d")
    db_file = 'finance_v15.db'
    backup_file = f"backups/backup_{today_str}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)
        return True
    return False

# 2. 페이지 설정 및 DB 연결
st.set_page_config(page_title="자금 관리 시스템 v15", layout="wide", page_icon="💰")
run_backup()

@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_v15.db', check_same_thread=False)
    c = conn.cursor()
    # 거래처 마스터: 계좌 정보 컬럼 추가
    c.execute('''CREATE TABLE IF NOT EXISTS vendors 
                 (name TEXT PRIMARY KEY, bank TEXT, account TEXT, holder TEXT)''')
    # 발주 마스터
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (order_id TEXT PRIMARY KEY, order_date TEXT, vendor TEXT, 
                  product TEXT, category TEXT, currency TEXT, total_amt REAL, is_closed INTEGER DEFAULT 0)''')
    # 입금 내역: 송금 관련 정보 컬럼 추가
    c.execute('''CREATE TABLE IF NOT EXISTS payments 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, order_id TEXT, pay_date TEXT, 
                  category TEXT, vendor TEXT, product TEXT, currency TEXT,
                  deposit REAL, advance REAL, note TEXT, krw_val REAL,
                  bank TEXT, account TEXT, holder TEXT)''')
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비"]
CURRENCIES = ["한화", "USD", "CNY"]

# --- 데이터 로드 함수 ---
def get_vendor_data():
    return pd.read_sql("SELECT * FROM vendors ORDER BY name ASC", conn)

# --- 메인 UI 구성 ---
tabs = st.tabs(["📝 입금 건별 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세 정산 관리", "⚙️ 거래처 및 계좌 관리"])

# --- Tab 5: 거래처 및 계좌 관리 (확장) ---
with tabs[4]:
    st.header("⚙️ 거래처 및 송금 계좌 관리")
    st.write("업체별 계좌 정보를 등록해두면 입금 입력 시 자동으로 불러옵니다.")
    
    with st.expander("➕ 신규 업체 및 계좌 등록", expanded=True):
        c1, c2, c3, c4 = st.columns(4)
        v_name = c1.text_input("업체명")
        v_bank = c2.text_input("은행명")
        v_acc = c3.text_input("계좌번호")
        v_hold = c4.text_input("예금주")
        
        if st.button("💾 업체 정보 저장"):
            if v_name:
                conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?, ?, ?, ?)", 
                                     (v_name, v_bank, v_acc, v_hold))
                conn.commit()
                st.success(f"'{v_name}' 정보가 저장되었습니다.")
                st.rerun()

    st.divider()
    st.subheader("📋 등록된 업체 리스트")
    v_df = get_vendor_data()
    edited_v = st.data_editor(v_df, use_container_width=True, num_rows="dynamic", key="vendor_editor")
    if st.button("🗑️ 변경사항(삭제/수정) 반영"):
        edited_v.to_sql('vendors', conn, if_exists='replace', index=False)
        st.success("업체 정보가 업데이트되었습니다.")
        st.rerun()

# --- Tab 1: 입금 건별 입력 (계좌 정보 연동) ---
with tabs[0]:
    st.header("📝 입금 내역 입력")
    orders_df = pd.read_sql("SELECT order_id, vendor, product FROM orders WHERE is_closed = 0", conn)
    v_data = get_vendor_data()
    
    with st.form("payment_form_v15", clear_on_submit=True):
        col_oid, col_date = st.columns([2, 1])
        sel_oid = col_oid.selectbox("🔗 발주서 연동", options=["직접 입력(발주서 없음)"] + list(orders_df['order_id']))
        p_date = col_date.date_input("📅 입금일", datetime.now())
        
        st.divider()
        
        c1, c2, c3 = st.columns(3)
        p_cat = c1.selectbox("📁 유형", CATEGORIES)
        p_vendor = c2.selectbox("🏢 업체명 선택", options=["선택하세요"] + list(v_data['name']))
        p_prod = c3.text_input("📦 상품명/항목명")
        
        # 계좌 정보 표시 구역 (선택된 업체에 따라 자동 매칭 안내)
        bank_info, acc_info, hold_info = "", "", ""
        if p_vendor != "선택하세요":
            v_info = v_data[v_data['name'] == p_vendor].iloc[0]
            bank_info, acc_info, hold_info = v_info['bank'], v_info['account'], v_info['holder']
            st.caption(f"🏦 **송금정보 자동연동:** {bank_info} | {acc_info} (예금주: {hold_info})")

        c4, c5, c6 = st.columns(3)
        p_curr = c4.selectbox("💱 통화", CURRENCIES)
        p_dep = c5.number_input("💵 입금액", min_value=0.0)
        p_adv = c6.number_input("🧧 선급금 변동", value=0.0)
        
        p_note = st.text_input("📝 송금 사유 및 메모")
        
        if st.form_submit_button("💾 데이터 저장하기"):
            if p_vendor == "선택하세요":
                st.error("업체명을 선택해주세요.")
            else:
                rate = 1350.0 if p_curr == "USD" else (190.0 if p_curr == "CNY" else 1.0)
                oid_val = sel_oid if sel_oid != "직접 입력(발주서 없음)" else None
                
                cur = conn.cursor()
                cur.execute('''INSERT INTO payments (order_id, pay_date, category, vendor, product, currency, 
                               deposit, advance, note, krw_val, bank, account, holder) 
                               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                            (oid_val, p_date.strftime("%Y-%m-%d"), p_cat, p_vendor, p_prod, p_curr, 
                             p_dep, p_adv, p_note, p_dep * rate, bank_info, acc_info, hold_info))
                conn.commit()
                st.success("입금 내역과 송금 정보가 함께 저장되었습니다!")

# --- Tab 4: 상세 정산 관리 (계좌 정보 포함) ---
with tabs[3]:
    st.header("🔍 상세 내역 조회")
    p_all = pd.read_sql("SELECT * FROM payments", conn)
    if not p_all.empty:
        # 엑셀처럼 옆으로 긴 대장을 위해 컬럼 순서 조정
        cols = ['pay_date', 'vendor', 'holder', 'bank', 'account', 'category', 'product', 'deposit', 'currency', 'note']
        st.data_editor(p_all[cols], use_container_width=True)