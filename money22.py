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
    db_file = 'finance_v19_fixed.db' 
    backup_file = f"backups/backup_{today_str}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)
        return True
    return False

# 2. 페이지 설정
st.set_page_config(page_title="자금 관리 시스템 v19", layout="wide", page_icon="💰")
run_backup()

# 3. DB 연결 및 테이블 강제 생성 (구조 보장)
@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_v19_fixed.db', check_same_thread=False)
    c = conn.cursor()
    # 거래처 마스터
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

def load_data(table):
    return pd.read_sql(f"SELECT * FROM {table}", conn)

# --- 메인 UI ---
st.title("💰 자금 관리 통합 시스템 v19")
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록(ERP/수기)", "🔍 상세 조회/마감", "⚙️ 거래처 관리"])

# --- Tab 1: 입금 건별 입력 ---
with tabs[0]:
    st.header("📝 입금 내역 입력")
    v_list = load_data("vendors")
    o_list = load_data("orders")
    active_orders = o_list[o_list['is_closed'] == 0] if not o_list.empty else pd.DataFrame()

    with st.form("pay_input_form_v19"):
        sel_oid = st.selectbox("🔗 발주서 연동", ["직접 입력"] + list(active_orders['order_id']) if not active_orders.empty else ["직접 입력"])
        c1, c2, c3 = st.columns(3)
        p_date = c1.date_input("📅 입금일", datetime.now())
        p_vendor = c2.selectbox("🏢 업체명", ["선택"] + list(v_list['name']) if not v_list.empty else ["선택"])
        p_cat = c3.selectbox("📁 유형", CATEGORIES)
        
        bank, acc, hold = "", "", ""
        if p_vendor != "선택" and not v_list.empty:
            row = v_list[v_list['name'] == p_vendor].iloc[0]
            bank, acc, hold = row['bank'], row['account'], row['holder']
            st.info(f"🏦 계좌 정보: {bank} | {acc} (예금주: {hold})")

        c4, c5, c6 = st.columns(3)
        p_dep = c4.number_input("💵 입금액", min_value=0.0)
        p_curr = c5.selectbox("💱 통화", CURRENCIES)
        p_note = c6.text_input("📝 메모")
        
        if st.form_submit_button("💾 저장"):
            rate = 1350.0 if p_curr == "USD" else (190.0 if p_curr == "CNY" else 1.0)
            conn.cursor().execute('''INSERT INTO payments (order_id, pay_date, category, vendor, product, currency, 
                                     deposit, advance, note, krw_val, bank, account, holder) 
                                     VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                                  (sel_oid if sel_oid != "직접 입력" else None, p_date.strftime("%Y-%m-%d"), 
                                   p_cat, p_vendor, "품목", p_curr, p_dep, 0, p_note, p_dep*rate, bank, acc, hold))
            conn.commit()
            st.success("저장되었습니다.")
            st.rerun()

# --- Tab 2: 입금 엑셀 업로드 ---
with tabs[1]:
    st.header("📂 입금 내역 일괄 업로드")
    pay_temp = pd.DataFrame(columns=["입금일", "거래처", "유형", "통화", "상품명", "입금액", "선급금", "송금사유", "발주번호"])
    st.download_button("📥 입금 양식(CSV) 다운로드", pay_temp.to_csv(index=False).encode('utf-8-sig'), "pay_template.csv")
    
    pay_file = st.file_uploader("입금 내역 CSV 파일 선택", type=['csv'])
    if pay_file:
        df_p = pd.read_csv(pay_file)
        if st.button("🚀 데이터 일괄 저장"):
            v_list = load_data("vendors")
            for _, r in df_p.iterrows():
                curr = r['통화'] if pd.notna(r['통화']) else "한화"
                rate = 1350.0 if curr == "USD" else (190.0 if curr == "CNY" else 1.0)
                v_info = v_list[v_list['name'] == r['거래처']]
                b, a, h = ("", "", "") if v_info.empty else (v_info.iloc[0]['bank'], v_info.iloc[0]['account'], v_info.iloc[0]['holder'])
                conn.cursor().execute('''INSERT INTO payments (order_id, pay_date, category, vendor, product, currency, 
                                       deposit, advance, note, krw_val, bank, account, holder) 
                                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                                    (r['발주번호'], r['입금일'], r['유형'], r['거래처'], r['상품명'], curr, 
                                     r['입금액'], r['선급금'], r['송금사유'], float(r['입금액'])*rate, b, a, h))
            conn.commit()
            st.success("업로드 완료!")
            st.rerun()

# --- Tab 3: 발주서 등록 (파일 업로드 + 수기) ---
with tabs[2]:
    st.header("📥 발주서 등록")
    col_o1, col_o2 = st.columns(2)
    with col_o1:
        st.subheader("📄 이카운트 엑셀 업로드")
        ecount_file = st.file_uploader("이카운트 엑셀(.xlsx)", type=['xlsx'])
        if ecount_file:
            st.warning("분석 기능을 불러오는 중입니다. 현재는 오른쪽 수기 폼을 이용해 주세요.")
    with col_o2:
        st.subheader("✍️ 발주서 수기 등록")
        v_list = load_data("vendors")
        with st.form("order_reg_v19"):
            o_id = st.text_input("발주번호 (전표번호)")
            o_date = st.date_input("발주일", datetime.now())
            o_cat = st.selectbox("유형", CATEGORIES)
            o_vendor = st.selectbox("거래처", ["선택"] + list(v_list['name']) if not v_list.empty else ["업체 없음"])
            o_prod = st.text_input("상품명")
            o_total = st.number_input("발주 총액", min_value=0.0)
            o_curr = st.selectbox("통화 ", CURRENCIES)
            if st.form_submit_button("🚀 발주서 저장"):
                if o_id and o_vendor != "선택":
                    conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", 
                                         (o_id, o_date.strftime("%Y-%m-%d"), o_vendor, o_prod, o_cat, o_curr, o_total))
                    conn.commit()
                    st.success("발주서 등록 성공!")
                    st.rerun()

# --- Tab 4: 상세 조회 및 마감 처리 ---
with tabs[3]:
    st.header("🔍 상세 내역 관리")
    p_all = load_data("payments")
    if not p_all.empty:
        # 가독성을 위한 컬럼 순서
        disp_cols = ['pay_date', 'vendor', 'holder', 'bank', 'account', 'deposit', 'currency', 'category', 'product', 'note']
        st.data_editor(p_all[disp_cols], use_container_width=True)
    else:
        st.info("데이터가 없습니다.")

# --- Tab 5: 거래처 관리 ---
with tabs[4]:
    st.header("⚙️ 거래처 및 계좌 관리")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("➕ 개별 등록")
        with st.form("v_reg_v19"):
            vn = st.text_input("업체명")
            vb = st.text_input("은행")
            va = st.text_input("계좌번호")
            vh = st.text_input("예금주")
            if st.form_submit_button("저장"):
                if vn:
                    conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?)", (vn, vb, va, vh))
                    conn.commit()
                    st.rerun()
    with c2:
        st.subheader("📂 엑셀 일괄 등록")
        v_temp = pd.DataFrame(columns=["업체명", "은행", "계좌번호", "예금주"])
        st.download_button("📥 양식 받기", v_temp.to_csv(index=False).encode('utf-8-sig'), "vendor_template.csv")
        v_file = st.file_uploader("거래처 CSV 업로드", type=['csv'], key="vendor_csv")
        if v_file:
            v_df = pd.read_csv(v_file)
            if st.button("🚀 일괄 저장"):
                for _, r in v_df.iterrows():
                    conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?)", (r['업체명'], r['은행'], r['계좌번호'], r['예금주']))
                conn.commit()
                st.success("등록 완료!")
                st.rerun()