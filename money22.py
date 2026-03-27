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
    # 파일명을 v15_final로 변경하여 이전 DB와 충돌 방지
    db_file = 'finance_v15_final.db'
    backup_file = f"backups/backup_{today_str}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)
        return True
    return False

# 2. 페이지 설정
st.set_page_config(page_title="자금 관리 시스템 v15", layout="wide", page_icon="💰")
run_backup()

# 3. DB 연결 및 초기화 (빈 화면 방지를 위해 테이블 생성 로직 강화)
@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_v15_final.db', check_same_thread=False)
    c = conn.cursor()
    # 거래처 마스터 (계좌 정보 포함)
    c.execute('''CREATE TABLE IF NOT EXISTS vendors 
                 (name TEXT PRIMARY KEY, bank TEXT, account TEXT, holder TEXT)''')
    # 발주 마스터
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (order_id TEXT PRIMARY KEY, order_date TEXT, vendor TEXT, 
                  product TEXT, category TEXT, currency TEXT, total_amt REAL, is_closed INTEGER DEFAULT 0)''')
    # 입금 내역 (송금 정보 포함)
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
def get_vendor_data():
    return pd.read_sql("SELECT * FROM vendors ORDER BY name ASC", conn)

def load_table(table_name):
    return pd.read_sql(f"SELECT * FROM {table_name}", conn)

# --- 메인 화면 구성 ---
st.title("💰 자금 관리 통합 시스템 v15")

# 탭 메뉴 구성
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📝 입금 건별 입력", 
    "📂 입금 엑셀 업로드", 
    "📥 발주서 등록", 
    "🔍 상세 정산 관리", 
    "⚙️ 거래처 및 계좌 관리"
])

# --- Tab 1: 입금 건별 입력 (계좌 자동 연동) ---
with tab1:
    st.header("📝 입금 내역 기록")
    orders_df = load_table("orders")
    active_orders = orders_df[orders_df['is_closed'] == 0]
    v_data = get_vendor_data()
    
    with st.form("payment_form_v15_final", clear_on_submit=True):
        c_oid, c_date = st.columns([2, 1])
        sel_oid = c_oid.selectbox("🔗 발주서 연동", options=["직접 입력(발주서 없음)"] + list(active_orders['order_id']))
        p_date = c_date.date_input("📅 입금일", datetime.now())
        
        st.divider()
        
        col1, col2, col3 = st.columns(3)
        p_cat = col1.selectbox("📁 유형", CATEGORIES)
        p_vendor = col2.selectbox("🏢 업체명 선택", options=["선택하세요"] + list(v_data['name']))
        p_prod = col3.text_input("📦 상품명/항목명")
        
        # 계좌 정보 실시간 안내
        bank_info, acc_info, hold_info = "", "", ""
        if p_vendor != "선택하세요":
            v_info = v_data[v_data['name'] == p_vendor].iloc[0]
            bank_info, acc_info, hold_info = v_info['bank'], v_info['account'], v_info['holder']
            st.info(f"🏦 **송금 예정 계좌:** {bank_info} | {acc_info} (예금주: {hold_info})")

        col4, col5, col6 = st.columns(3)
        p_curr = col4.selectbox("💱 통화", CURRENCIES)
        p_dep = col5.number_input("💵 입금액", min_value=0.0)
        p_adv = col6.number_input("🧧 선급금 변동", value=0.0)
        
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
                st.success("입금 내역이 저장되었습니다.")
                st.rerun()

# --- Tab 2: 입금 엑셀 업로드 ---
with tab2:
    st.header("📂 엑셀 일괄 업로드")
    template = pd.DataFrame(columns=["입금일", "거래처", "유형", "통화", "상품명", "입금액", "선급금", "송금사유", "발주번호"])
    st.download_button("📥 업로드 양식 다운로드", template.to_csv(index=False).encode('utf-8-sig'), "pay_template.csv")
    
    up_file = st.file_uploader("파일 업로드", type=['csv'])
    if up_file:
        df_up = pd.read_csv(up_file)
        if st.button("✅ 데이터 일괄 저장"):
            for _, r in df_up.iterrows():
                curr = r['통화'] if pd.notna(r['통화']) else "한화"
                rate = 1350.0 if curr == "USD" else (190.0 if curr == "CNY" else 1.0)
                # 계좌 정보는 거래처 마스터에서 찾아오기
                v_name = r['거래처']
                v_master = v_data[v_data['name'] == v_name]
                b, a, h = ("", "", "") if v_master.empty else (v_master.iloc[0]['bank'], v_master.iloc[0]['account'], v_master.iloc[0]['holder'])
                
                conn.cursor().execute('''INSERT INTO payments (order_id, pay_date, category, vendor, product, currency, 
                                       deposit, advance, note, krw_val, bank, account, holder) 
                                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                                    (r['발주번호'], r['입금일'], r['유형'], v_name, r['상품명'], curr, 
                                     r['입금액'], r['선급금'], r['송금사유'], float(r['입금액']) * rate, b, a, h))
            conn.commit()
            st.success("일괄 업로드 완료!")
            st.rerun()

# --- Tab 3: 발주서 등록 ---
with tab3:
    st.header("📥 발주서(Master) 등록")
    with st.form("order_reg_v15"):
        c1, c2, c3 = st.columns(3)
        oid = c1.text_input("발주번호 (ERP 전표번호)")
        odate = c2.date_input("발주일", datetime.now())
        ocat = c3.selectbox("유형", CATEGORIES)
        
        c4, c5, c6 = st.columns(3)
        ovendor = c4.selectbox("거래처 선택", options=["선택하세요"] + list(v_data['name']))
        oprod = c5.text_input("대표 상품명")
        ocurr = c6.selectbox("통화", CURRENCIES)
        
        ototal = st.number_input("발주 총액", min_value=0.0)
        
        if st.form_submit_button("🚀 발주서 저장"):
            if oid and ovendor != "선택하세요":
                conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?, ?, ?, ?, ?, ?, ?, 0)", 
                                     (oid, odate.strftime("%Y-%m-%d"), ovendor, oprod, ocat, ocurr, ototal))
                conn.commit()
                st.success(f"발주번호 {oid} 저장 완료")
                st.rerun()

# --- Tab 4: 상세 정산 관리 ---
with tab4:
    st.header("🔍 상세 내역 조회")
    p_all = load_table("payments")
    if not p_all.empty:
        # 엑셀 대장 스타일 순서
        cols = ['pay_date', 'vendor', 'holder', 'bank', 'account', 'category', 'product', 'deposit', 'currency', 'note']
        st.data_editor(p_all[cols], use_container_width=True)
    else:
        st.info("조회할 데이터가 없습니다.")

# --- Tab 5: 거래처 및 계좌 관리 ---
with tab5:
    st.header("⚙️ 업체 정보 관리")
    with st.expander("➕ 업체 및 계좌 등록", expanded=True):
        col1, col2, col3, col4 = st.columns(4)
        v_name = col1.text_input("업체명")
        v_bank = col2.text_input("은행명")
        v_acc = col3.text_input("계좌번호")
        v_hold = col4.text_input("예금주")
        if st.button("💾 업체 정보 저장"):
            if v_name:
                conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?, ?, ?, ?)", (v_name, v_bank, v_acc, v_hold))
                conn.commit()
                st.success("저장되었습니다.")
                st.rerun()
    
    st.divider()
    v_list = get_vendor_data()
    if not v_list.empty:
        edited_v = st.data_editor(v_list, use_container_width=True, num_rows="dynamic")
        if st.button("🗑️ 변경사항(삭제/수정) 적용"):
            edited_v.to_sql('vendors', conn, if_exists='replace', index=False)
            st.success("업체 정보가 업데이트되었습니다.")
            st.rerun()