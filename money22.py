import streamlit as st
import pandas as pd
import sqlite3
import os
import shutil
from datetime import datetime

# --- 1. 데이터 안전장치 ---
def run_backup():
    if not os.path.exists('backups'): os.makedirs('backups')
    today_str = datetime.now().strftime("%Y%m%d")
    db_file = 'finance_v20_final.db'
    backup_file = f"backups/backup_{today_str}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

# 2. 페이지 설정
st.set_page_config(page_title="자금 관리 v20", layout="wide", page_icon="💰")
run_backup()

# 3. DB 연결 및 한글 컬럼명 매칭
@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_v20_final.db', check_same_thread=False)
    c = conn.cursor()
    c.execute('CREATE TABLE IF NOT EXISTS vendors (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT)')
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (발주번호 TEXT PRIMARY KEY, 발주일 TEXT, 거래처명 TEXT, 
                  상품명 TEXT, 유형 TEXT, 통화 TEXT, 발주총액 REAL, 마감여부 INTEGER DEFAULT 0)''')
    c.execute('''CREATE TABLE IF NOT EXISTS payments 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 발주번호 TEXT, 입금일 TEXT, 
                  유형 TEXT, 거래처명 TEXT, 상품명 TEXT, 통화 TEXT,
                  실입금액 REAL, 선급금사용액 REAL, 메모 TEXT, 한화환산액 REAL,
                  은행 TEXT, 계좌번호 TEXT, 예금주 TEXT)''')
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비"]
CURRENCIES = ["한화", "USD", "CNY"]

def load_table(table):
    return pd.read_sql(f"SELECT * FROM {table}", conn)

# --- 스타일 함수: 마감 건 회색 음영 ---
def style_closed_rows(row):
    if row['마감여부'] == 1:
        return ['background-color: #f0f0f0; color: #999; text-decoration: line-through'] * len(row)
    return [''] * len(row)

# --- 메인 UI ---
st.title("💰 자금 관리 시스템 v20")
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리"])

# --- Tab 5: 거래처 관리 (등록 확인 리스트 추가) ---
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    c1, c2 = st.columns(2)
    with c1:
        with st.form("v_reg"):
            vn, vb, va, vh = st.columns(4)
            n = vn.text_input("업체명")
            b = vb.text_input("은행")
            a = va.text_input("계좌번호")
            h = vh.text_input("예금주")
            if st.form_submit_button("거래처 등록"):
                conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?)", (n,b,a,h))
                conn.commit()
                st.rerun()
    with c2:
        v_temp = pd.DataFrame(columns=["거래처명", "은행", "계좌번호", "예금주"])
        st.download_button("📥 양식 다운로드", v_temp.to_csv(index=False).encode('utf-8-sig'), "vendor_temp.csv")
        v_file = st.file_uploader("거래처 일괄 업로드", type=['csv'])
        if v_file:
            v_df = pd.read_csv(v_file)
            if st.button("🚀 일괄 저장"):
                v_df.to_sql('vendors', conn, if_exists='append', index=False)
                conn.commit()
                st.rerun()
    
    st.subheader("📋 등록된 거래처 목록")
    st.dataframe(load_table("vendors"), use_container_width=True)

# --- Tab 3: 발주서 등록 (리스트 표시 추가) ---
with tabs[2]:
    st.header("📥 발주서 등록")
    v_list = load_table("vendors")
    with st.expander("➕ 신규 발주서 수기 등록", expanded=True):
        with st.form("order_reg"):
            c1, c2, c3, c4 = st.columns(4)
            o_id = c1.text_input("발주번호")
            o_date = c2.date_input("발주일")
            o_cat = c3.selectbox("유형", CATEGORIES)
            o_vendor = c4.selectbox("거래처", ["선택"] + list(v_list['거래처명']) if not v_list.empty else ["업체없음"])
            
            c5, c6, c7 = st.columns([2,1,1])
            o_prod = c5.text_input("상품명")
            o_total = c6.number_input("발주총액", min_value=0.0)
            o_curr = c7.selectbox("통화", CURRENCIES)
            if st.form_submit_button("🚀 발주서 저장"):
                conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", 
                                     (o_id, o_date.strftime("%Y-%m-%d"), o_vendor, o_prod, o_cat, o_curr, o_total))
                conn.commit()
                st.rerun()

    st.subheader("📑 발주 리스트 (진행/마감)")
    o_df = load_table("orders")
    if not o_df.empty:
        st.dataframe(o_df.style.apply(style_closed_rows, axis=1), use_container_width=True)

# --- Tab 1: 입금 입력 (선급금 구분) ---
with tabs[0]:
    st.header("📝 입금 내역 입력")
    orders = load_table("orders")
    active_orders = orders[orders['마감여부'] == 0]
    v_data = load_table("vendors")
    
    with st.form("pay_form"):
        sel_oid = st.selectbox("🔗 발주번호 연동 (없으면 직접입력)", ["직접 입력"] + list(active_orders['발주번호']) if not active_orders.empty else ["직접 입력"])
        c1, c2, c3 = st.columns(3)
        p_date = c1.date_input("입금일")
        p_vendor = c2.selectbox("업체명", ["선택"] + list(v_data['거래처명']))
        p_cat = c3.selectbox("유형", CATEGORIES)
        
        c4, c5 = st.columns(2)
        p_dep = c4.number_input("💵 실제 입금액 (통장 찍힌 금액)", min_value=0.0)
        p_adv_use = c5.number_input("🧧 기존 선급금에서 차감할 금액", min_value=0.0)
        
        p_curr = st.selectbox("통화 ", CURRENCIES)
        p_note = st.text_input("메모")
        
        if st.form_submit_button("💾 저장"):
            rate = 1350.0 if p_curr == "USD" else (190.0 if p_curr == "CNY" else 1.0)
            # 계좌정보 가져오기
            v_info = v_data[v_data['거래처명'] == p_vendor].iloc[0]
            conn.cursor().execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금사용액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                     VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                                  (sel_oid if sel_oid != "직접 입력" else None, p_date.strftime("%Y-%m-%d"), 
                                   p_cat, p_vendor, "품목", p_curr, p_dep, p_adv_use, p_note, (p_dep + p_adv_use)*rate, 
                                   v_info['은행'], v_info['계좌번호'], v_info['예금주']))
            conn.commit()
            st.success("입금 정보가 저장되었습니다.")

# --- Tab 4: 상세내역 관리 (필터/마감/음영) ---
with tabs[3]:
    st.header("🔍 상세내역 조회 및 필터")
    p_all = load_table("payments")
    o_all = load_table("orders")
    
    if not p_all.empty:
        # 필터링 섹션
        f1, f2, f3 = st.columns(3)
        with f1: f_vendor = st.multiselect("업체별 필터", p_all['거래처명'].unique())
        with f2: f_cat = st.multiselect("유형별 필터", CATEGORIES)
        with f3: 
            p_all['월'] = pd.to_datetime(p_all['입금일']).dt.strftime('%Y-%m')
            f_month = st.multiselect("월별 필터", sorted(p_all['월'].unique(), reverse=True))

        df_filtered = p_all.copy()
        if f_vendor: df_filtered = df_filtered[df_filtered['거래처명'].isin(f_vendor)]
        if f_cat: df_filtered = df_filtered[df_filtered['유형'].isin(f_cat)]
        if f_month: df_filtered = df_filtered[df_filtered['월'].isin(f_month)]

        # 발주 마감 여부 결합
        df_final = df_filtered.merge(o_all[['발주번호', '마감여부']], on='발주번호', how='left').fillna(0)
        
        st.subheader("📊 조회 결과")
        st.dataframe(df_final.style.apply(style_closed_rows, axis=1), use_container_width=True)
        
        # 마감 체크 기능
        st.divider()
        c_m1, c_m2 = st.columns([3, 1])
        target_oid = c_m1.selectbox("마감 처리할 발주번호 선택", o_all[o_all['마감여부']==0]['발주번호'].unique())
        if c_m2.button("🚩 해당 발주 마감하기"):
            conn.cursor().execute("UPDATE orders SET 마감여부 = 1 WHERE 발주번호 = ?", (target_oid,))
            conn.commit()
            st.rerun()