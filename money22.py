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
    db_file = 'finance_v17_final.db'
    backup_file = f"backups/backup_{today_str}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)
        return True
    return False

# 2. 페이지 설정 및 DB 연결
st.set_page_config(page_title="자금 관리 시스템 v17", layout="wide", page_icon="💰")
run_backup()

@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_v17_final.db', check_same_thread=False)
    c = conn.cursor()
    # 거래처 마스터
    c.execute('CREATE TABLE IF NOT EXISTS vendors (name TEXT PRIMARY KEY, bank TEXT, account TEXT, holder TEXT)')
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

# --- 이카운트 발주서 파싱 함수 ---
def parse_ecount_order(file):
    try:
        # 이카운트 엑셀 로드 (실제 구조에 따라 skiprows 조정이 필요할 수 있음)
        df = pd.read_excel(file)
        # 전표번호 위치 (샘플 데이터 기준: 2행 1열 등)
        raw_oid = str(df.iloc[0, 0]) if "전표번호" in str(df.iloc[0,0]) else "자동입력필요"
        oid = raw_oid.split(":")[-1].strip() if ":" in raw_oid else raw_oid
        
        # 거래처명 위치 추출 (샘플 데이터 기반)
        vendor = "미확인 업체"
        for i in range(len(df)):
            if "수신" in str(df.iloc[i, 0]):
                vendor = str(df.iloc[i, 0]).split(":")[-1].strip()
                break
        
        # 품목명 및 합계 (품목코드 컬럼 기준 아래 데이터 합산)
        # 이 부분은 엑셀 양식의 헤더 위치에 따라 유동적입니다.
        total_amt = 0
        product_name = "품목 정보 확인필요"
        
        if '합계' in df.columns:
            total_amt = df['합계'].sum()
        elif 'USD 합계' in df.columns:
            total_amt = df['USD 합계'].sum()
            
        return {"order_id": oid, "vendor": vendor, "total_amt": total_amt}
    except:
        return None

# --- 메인 UI ---
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀", "📥 발주서 등록(ERP)", "🔍 상세 조회", "⚙️ 거래처 관리"])

# --- Tab 5: 거래처 관리 (엑셀 일괄 등록 추가) ---
with tabs[4]:
    st.header("⚙️ 거래처 및 계좌 관리")
    
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        st.subheader("➕ 개별 등록")
        with st.form("v_reg"):
            v1, v2, v3, v4 = st.columns(4)
            name = v1.text_input("업체명")
            bank = v2.text_input("은행")
            acc = v3.text_input("계좌번호")
            hold = v4.text_input("예금주")
            if st.form_submit_button("저장"):
                conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?)", (name, bank, acc, hold))
                conn.commit()
                st.rerun()

    with col_up2:
        st.subheader("📂 엑셀 일괄 등록")
        v_template = pd.DataFrame(columns=["업체명", "은행", "계좌번호", "예금주"])
        st.download_button("📥 거래처 양식 다운로드", v_template.to_csv(index=False).encode('utf-8-sig'), "vendor_template.csv")
        v_file = st.file_uploader("거래처 CSV 업로드", type=['csv'])
        if v_file:
            v_df = pd.read_csv(v_file)
            if st.button("✅ 거래처 일괄 저장"):
                for _, r in v_df.iterrows():
                    conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?)", (r['업체명'], r['은행'], r['계좌번호'], r['예금주']))
                conn.commit()
                st.success(f"{len(v_df)}개 업체 등록 완료")
                st.rerun()

# --- Tab 3: 발주서 등록 (이카운트 엑셀 파싱) ---
with tabs[2]:
    st.header("📥 이카운트 발주서 등록")
    ecount_file = st.file_uploader("이카운트 엑셀 파일(.xlsx)", type=['xlsx'])
    if ecount_file:
        res = parse_ecount_order(ecount_file)
        with st.form("order_confirm"):
            c1, c2, c3 = st.columns(3)
            oid = c1.text_input("발주번호", value=res['order_id'] if res else "")
            odate = c2.date_input("발주일", datetime.now())
            ocat = c3.selectbox("유형", CATEGORIES)
            
            c4, c5, c6 = st.columns(3)
            ovendor = c4.text_input("거래처", value=res['vendor'] if res else "")
            oprod = c5.text_input("대표품목")
            ocurr = c6.selectbox("통화", CURRENCIES)
            
            ototal = st.number_input("총액", value=float(res['total_amt']) if res else 0.0)
            
            if st.form_submit_button("🚀 발주서 확정 저장"):
                conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", 
                                     (oid, odate.strftime("%Y-%m-%d"), ovendor, oprod, ocat, ocurr, ototal))
                conn.commit()
                st.success("발주서 등록 완료")

# --- Tab 1: 입금 건별 입력 ---
with tabs[0]:
    st.header("📝 입금 내역 입력")
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    o_data = pd.read_sql("SELECT * FROM orders WHERE is_closed = 0", conn)
    
    with st.form("pay_form"):
        sel_oid = st.selectbox("연동 발주서", ["직접 입력"] + list(o_data['order_id']))
        c1, c2, c3 = st.columns(3)
        p_date = c1.date_input("입금일")
        p_vendor = c2.selectbox("업체명", ["선택"] + list(v_data['name']))
        p_cat = c3.selectbox("유형", CATEGORIES)
        
        # 계좌 정보 자동 안내
        bank, acc, hold = "", "", ""
        if p_vendor != "선택":
            row = v_data[v_data['name'] == p_vendor].iloc[0]
            bank, acc, hold = row['bank'], row['account'], row['holder']
            st.info(f"🏦 계좌: {bank} / {acc} (예금주: {hold})")
            
        p_dep = st.number_input("입금액", min_value=0.0)
        p_curr = st.selectbox("통화 ", CURRENCIES)
        p_note = st.text_input("메모")
        
        if st.form_submit_button("저장"):
            rate = 1350.0 if p_curr == "USD" else (190.0 if p_curr == "CNY" else 1.0)
            conn.cursor().execute('''INSERT INTO payments (order_id, pay_date, category, vendor, product, currency, 
                                     deposit, advance, note, krw_val, bank, account, holder) 
                                     VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                                  (sel_oid if sel_oid != "직접 입력" else None, p_date.strftime("%Y-%m-%d"), 
                                   p_cat, p_vendor, "품목", p_curr, p_dep, 0, p_note, p_dep*rate, bank, acc, hold))
            conn.commit()
            st.success("입금 완료")

# --- Tab 4: 상세 조회 (마감/회색 처리 예고) ---
with tabs[3]:
    st.header("🔍 상세 내역 조회")
    p_all = pd.read_sql("SELECT * FROM payments", conn)
    st.data_editor(p_all, use_container_width=True)