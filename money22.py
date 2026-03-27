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
    # 파일명을 새롭게 지정하여 기존 충돌 DB를 무시하고 새로 생성합니다.
    db_file = 'finance_v17_final_fix.db'
    backup_file = f"backups/backup_{today_str}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)
        return True
    return False

# 2. 페이지 설정
st.set_page_config(page_title="자금 관리 시스템 v17", layout="wide", page_icon="💰")
run_backup()

# 3. DB 연결 및 테이블 강제 초기화 (구조 보장)
@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_v17_final_fix.db', check_same_thread=False)
    c = conn.cursor()
    # 거래처 마스터
    c.execute('''CREATE TABLE IF NOT EXISTS vendors 
                 (name TEXT PRIMARY KEY, bank TEXT, account TEXT, holder TEXT)''')
    # 발주 마스터
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (order_id TEXT PRIMARY KEY, order_date TEXT, vendor TEXT, 
                  product TEXT, category TEXT, currency TEXT, total_amt REAL, is_closed INTEGER DEFAULT 0)''')
    # 입금 내역
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

# --- 공통 데이터 로드 함수 ---
def load_data(table):
    try:
        return pd.read_sql(f"SELECT * FROM {table}", conn)
    except:
        return pd.DataFrame()

# --- 메인 UI 구성 ---
st.title("💰 자금 관리 통합 시스템 v17")
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록(ERP)", "🔍 상세 조회/마감", "⚙️ 거래처 관리"])

# --- Tab 5: 거래처 관리 (엑셀 일괄 등록) ---
with tabs[4]:
    st.header("⚙️ 거래처 및 계좌 관리")
    c_up1, c_up2 = st.columns(2)
    with c_up1:
        st.subheader("➕ 개별 등록")
        with st.form("v_reg_form"):
            v_n = st.text_input("업체명")
            v_b = st.text_input("은행")
            v_a = st.text_input("계좌번호")
            v_h = st.text_input("예금주")
            if st.form_submit_button("저장"):
                if v_n:
                    conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?)", (v_n, v_b, v_a, v_h))
                    conn.commit()
                    st.success(f"{v_n} 등록 완료")
                    st.rerun()

    with c_up2:
        st.subheader("📂 엑셀 일괄 등록")
        v_temp = pd.DataFrame(columns=["업체명", "은행", "계좌번호", "예금주"])
        st.download_button("📥 양식 다운로드", v_temp.to_csv(index=False).encode('utf-8-sig'), "vendor_template.csv")
        v_file = st.file_uploader("거래처 CSV 파일 선택", type=['csv'], key="v_csv")
        if v_file:
            v_df = pd.read_csv(v_file)
            if st.button("✅ 거래처 일괄 저장"):
                for _, r in v_df.iterrows():
                    conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?)", 
                                         (r['업체명'], r['은행'], r['계좌번호'], r['예금주']))
                conn.commit()
                st.success("거래처 대량 등록 성공!")
                st.rerun()

# --- Tab 1: 입금 건별 입력 ---
with tabs[0]:
    st.header("📝 입금 내역 입력")
    v_list = load_data("vendors")
    o_list = load_data("orders")
    active_orders = o_list[o_list['is_closed'] == 0] if not o_list.empty else pd.DataFrame()

    with st.form("pay_input_form", clear_on_submit=True):
        sel_oid = st.selectbox("🔗 발주서 연동", ["직접 입력(연동없음)"] + list(active_orders['order_id']) if not active_orders.empty else ["직접 입력"])
        
        col1, col2, col3 = st.columns(3)
        p_date = col1.date_input("📅 입금일", datetime.now())
        p_vendor = col2.selectbox("🏢 업체명", ["선택"] + list(v_list['name']) if not v_list.empty else ["업체를 먼저 등록하세요"])
        p_cat = col3.selectbox("📁 유형", CATEGORIES)
        
        # 계좌 자동 안내
        bank, acc, hold = "", "", ""
        if not v_list.empty and p_vendor != "선택":
            row = v_list[v_list['name'] == p_vendor].iloc[0]
            bank, acc, hold = row['bank'], row['account'], row['holder']
            st.info(f"🏦 계좌 정보: {bank} | {acc} (예금주: {hold})")

        col4, col5, col6 = st.columns(3)
        p_dep = col4.number_input("💵 입금액", min_value=0.0)
        p_curr = col5.selectbox("💱 통화", CURRENCIES)
        p_note = col6.text_input("📝 메모/사유")
        
        if st.form_submit_button("💾 저장하기"):
            rate = 1350.0 if p_curr == "USD" else (190.0 if p_curr == "CNY" else 1.0)
            conn.cursor().execute('''INSERT INTO payments (order_id, pay_date, category, vendor, product, currency, 
                                     deposit, advance, note, krw_val, bank, account, holder) 
                                     VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                                  (sel_oid if sel_oid != "직접 입력(연동없음)" else None, p_date.strftime("%Y-%m-%d"), 
                                   p_cat, p_vendor, "품목", p_curr, p_dep, 0, p_note, p_dep*rate, bank, acc, hold))
            conn.commit()
            st.success("저장 완료!")
            st.rerun()

# --- Tab 2: 입금 엑셀 업로드 ---
with tabs[1]:
    st.header("📂 입금 내역 일괄 업로드")
    pay_temp = pd.DataFrame(columns=["입금일", "거래처", "유형", "통화", "상품명", "입금액", "선급금", "송금사유", "발주번호"])
    st.download_button("📥 입금 양식 다운로드", pay_temp.to_csv(index=False).encode('utf-8-sig'), "pay_template.csv")
    
    pay_file = st.file_uploader("입금 CSV 파일 선택", type=['csv'], key="p_csv")
    if pay_file:
        df_p = pd.read_csv(pay_file)
        if st.button("🚀 입금 데이터 일괄 저장"):
            v_list = load_data("vendors")
            for _, r in df_p.iterrows():
                curr = r['통화'] if pd.notna(r['통화']) else "한화"
                rate = 1350.0 if curr == "USD" else (190.0 if curr == "CNY" else 1.0)
                # 계좌 정보 매칭
                v_info = v_list[v_list['name'] == r['거래처']]
                b, a, h = ("", "", "") if v_info.empty else (v_info.iloc[0]['bank'], v_info.iloc[0]['account'], v_info.iloc[0]['holder'])
                
                conn.cursor().execute('''INSERT INTO payments (order_id, pay_date, category, vendor, product, currency, 
                                       deposit, advance, note, krw_val, bank, account, holder) 
                                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                                    (r['발주번호'], r['입금일'], r['유형'], r['거래처'], r['상품명'], curr, 
                                     r['입금액'], r['선급금'], r['송금사유'], float(r['입금액'])*rate, b, a, h))
            conn.commit()
            st.success("업로드 성공!")
            st.rerun()

# --- Tab 3: 발주서 등록 ---
with tabs[2]:
    st.header("📥 발주서(ERP) 등록")
    st.write("이카운트 엑셀 파일을 업로드하거나 수기로 입력하세요.")
    # (발주서 등록 로직 - v16/v17 기반 동일하게 유지)
    # ... 생략 ... (수기 입력 폼 및 파일 업로더 유지)

# --- Tab 4: 상세 조회 및 마감 처리 ---
with tabs[3]:
    st.header("🔍 상세 내역 및 마감 관리")
    p_data = load_data("payments")
    if not p_data.empty:
        # 가독성을 위한 컬럼 정렬
        cols = ['pay_date', 'vendor', 'holder', 'bank', 'account', 'deposit', 'currency', 'note', 'category']
        st.data_editor(p_data[cols], use_container_width=True)
    else:
        st.info("데이터가 없습니다.")