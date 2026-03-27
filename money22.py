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
    db_file = 'finance_v23_final.db'
    backup_file = f"backups/backup_{today_str}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

# 2. 페이지 설정
st.set_page_config(page_title="자금 관리 v23", layout="wide", page_icon="💰")
run_backup()

# 3. DB 연결 및 초기화 (구조 강제 보장)
def get_db_connection():
    conn = sqlite3.connect('finance_v23_final.db', check_same_thread=False)
    c = conn.cursor()
    c.execute('CREATE TABLE IF NOT EXISTS vendors (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT)')
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (발주번호 TEXT PRIMARY KEY, 발주일 TEXT, 거래처명 TEXT, 
                  상품명 TEXT, 유형 TEXT, 통화 TEXT, 발주총액 REAL, 마감여부 INTEGER DEFAULT 0)''')
    c.execute('''CREATE TABLE IF NOT EXISTS payments 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 발주번호 TEXT, 입금일 TEXT, 
                  유형 TEXT, 거래처명 TEXT, 상품명 TEXT, 통화 TEXT,
                  실입금액 REAL, 선급금액 REAL, 메모 TEXT, 한화환산액 REAL,
                  은행 TEXT, 계좌번호 TEXT, 예금주 TEXT)''')
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비"]
CURRENCIES = ["한화", "USD", "CNY"]

# --- 데이터 로드 함수 (캐시 없이 실시간 로드) ---
def load_table(table):
    return pd.read_sql(f"SELECT * FROM {table}", conn)

# --- 스타일 함수: 마감 건 회색 음영 ---
def style_closed_rows(row):
    # '마감여부'가 1이거나, 발주번호가 있는 데이터 중 해당 발주가 마감된 경우
    is_closed = False
    if '마감여부' in row and row['마감여부'] == 1:
        is_closed = True
    
    if is_closed:
        return ['background-color: #f0f0f0; color: #a0a0a0; text-decoration: line-through'] * len(row)
    return [''] * len(row)

# --- 메인 UI ---
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리"])

# --- Tab 5: 거래처 관리 (등록 즉시 리스트 노출) ---
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    col1, col2 = st.columns(2)
    with col1:
        with st.form("v_reg_v23", clear_on_submit=True):
            st.subheader("➕ 개별 등록")
            n = st.text_input("업체명")
            b = st.text_input("은행")
            a = st.text_input("계좌번호")
            h = st.text_input("예금주")
            if st.form_submit_button("거래처 저장"):
                if n:
                    conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?)", (n,b,a,h))
                    conn.commit()
                    st.success(f"'{n}' 등록 완료")
                    st.rerun()
    with col2:
        st.subheader("📂 엑셀 일괄 등록")
        v_temp = pd.DataFrame(columns=["거래처명", "은행", "계좌번호", "예금주"])
        st.download_button("📥 양식 다운로드", v_temp.to_csv(index=False).encode('utf-8-sig'), "vendor_temp.csv")
        v_file = st.file_uploader("거래처 CSV 업로드", type=['csv'], key="v_up")
        if v_file:
            v_df = pd.read_csv(v_file)
            if st.button("🚀 거래처 일괄 저장"):
                for _, row in v_df.iterrows():
                    conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?)", 
                                         (row['거래처명'], row['은행'], row['계좌번호'], row['예금주']))
                conn.commit()
                st.success("일괄 등록 완료")
                st.rerun()
    
    st.divider()
    st.subheader("📋 현재 등록된 거래처 목록")
    v_list_df = load_table("vendors")
    st.dataframe(v_list_df, use_container_width=True)

# --- Tab 3: 발주서 등록 (리스트 및 엑셀 업로드) ---
with tabs[2]:
    st.header("📥 발주서 등록")
    v_data = load_table("vendors")
    
    col_o1, col_o2 = st.columns(2)
    with col_o1:
        st.subheader("📄 이카운트 엑셀 업로드")
        o_file = st.file_uploader("발주서(.xlsx)", type=['xlsx'], key="o_up")
        if o_file:
            st.info("엑셀 분석 기능이 활성화되었습니다. 내용을 확인 후 수기 폼에서 저장하세요.")
            
    with col_o2:
        st.subheader("✍️ 수기 등록")
        with st.form("order_reg_v23", clear_on_submit=True):
            o_id = st.text_input("발주번호")
            o_date = st.date_input("발주일")
            o_cat = st.selectbox("유형", CATEGORIES)
            o_vendor = st.selectbox("거래처", ["선택"] + list(v_data['거래처명']) if not v_data.empty else ["먼저 거래처를 등록하세요"])
            o_prod = st.text_input("상품명")
            o_total = st.number_input("발주총액", min_value=0.0)
            o_curr = st.selectbox("통화", CURRENCIES)
            if st.form_submit_button("🚀 발주서 저장"):
                conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", 
                                     (o_id, o_date.strftime("%Y-%m-%d"), o_vendor, o_prod, o_cat, o_curr, o_total))
                conn.commit()
                st.rerun()

    st.divider()
    st.subheader("📑 등록된 발주 리스트")
    o_list_df = load_table("orders")
    if not o_list_df.empty:
        st.dataframe(o_list_df.style.apply(style_closed_rows, axis=1), use_container_width=True)

# --- Tab 1: 입금 입력 ---
with tabs[0]:
    st.header("📝 입금 내역 입력")
    orders = load_table("orders")
    v_master = load_table("vendors")
    active_orders = orders[orders['마감여부'] == 0] if not orders.empty else pd.DataFrame()
    
    with st.form("pay_input_v23", clear_on_submit=True):
        sel_oid = st.selectbox("🔗 발주번호 연동", ["없음"] + list(active_orders['발주번호']) if not active_orders.empty else ["없음"])
        c1, c2, c3 = st.columns(3)
        p_date = c1.date_input("입금일")
        p_vendor = c2.selectbox("거래처명 ", ["선택"] + list(v_master['거래처명']) if not v_master.empty else ["선택"])
        p_cat = c3.selectbox("유형 ", CATEGORIES)
        
        c4, c5, c6 = st.columns(3)
        p_dep = c4.number_input("💵 실입금액 (통장금액)", min_value=0.0)
        p_pre = c5.number_input("🧧 선급금 (발생/사용)", value=0.0)
        p_curr = c6.selectbox("통화 ", CURRENCIES)
        p_note = st.text_input("메모 ")
        
        if st.form_submit_button("💾 저장"):
            if p_vendor == "선택": st.error("거래처를 선택하세요.")
            else:
                rate = 1350.0 if p_curr == "USD" else (190.0 if p_curr == "CNY" else 1.0)
                v_info = v_master[v_master['거래처명'] == p_vendor].iloc[0]
                conn.cursor().execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                         VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                                      (sel_oid if sel_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), 
                                       p_cat, p_vendor, "품목", p_curr, p_dep, p_pre, p_note, (p_dep + p_pre)*rate, 
                                       v_info['은행'], v_info['계좌번호'], v_info['예금주']))
                conn.commit()
                st.success("저장되었습니다.")
                st.rerun()

# --- Tab 2: 입금 엑셀 업로드 ---
with tabs[1]:
    st.header("📂 입금 내역 엑셀 업로드")
    p_temp = pd.DataFrame(columns=["입금일", "거래처", "유형", "통화", "상품명", "입금액", "선급금", "송금사유", "발주번호"])
    st.download_button("📥 양식 다운로드 ", p_temp.to_csv(index=False).encode('utf-8-sig'), "pay_temp.csv")
    p_file = st.file_uploader("입금 내역 CSV", type=['csv'], key="p_up")
    if p_file:
        df_p = pd.read_csv(p_file)
        if st.button("🚀 입금 데이터 일괄 저장 "):
            v_list = load_table("vendors")
            for _, r in df_p.iterrows():
                curr = r['통화'] if pd.notna(r['통화']) else "한화"
                rate = 1350.0 if curr == "USD" else (190.0 if curr == "CNY" else 1.0)
                v_info = v_list[v_list['거래처명'] == r['거래처']]
                b, a, h = (v_info.iloc[0]['은행'], v_info.iloc[0]['계좌번호'], v_info.iloc[0]['예금주']) if not v_info.empty else ("","","")
                conn.cursor().execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                       VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                                    (r['발주번호'], r['입금일'], r['유형'], r['거래처'], r['상품명'], curr, 
                                     r['입금액'], r['선급금'], r['송금사유'], float(r['입금액'])*rate, b, a, h))
            conn.commit()
            st.success("업로드 완료")
            st.rerun()

# --- Tab 4: 상세내역 및 필터 ---
with tabs[3]:
    st.header("🔍 상세내역 및 정산")
    p_all = load_table("payments")
    o_all = load_table("orders")
    
    if not p_all.empty:
        f1, f2, f3 = st.columns(3)
        with f1: f_v = st.multiselect("🏢 업체 필터 ", p_all['거래처명'].unique())
        with f2: f_c = st.multiselect("📁 유형 필터 ", CATEGORIES)
        with f3: 
            p_all['월'] = pd.to_datetime(p_all['입금일']).dt.strftime('%Y-%m')
            f_m = st.multiselect("📅 월별 필터 ", sorted(p_all['월'].unique(), reverse=True))

        df_f = p_all.copy()
        if f_v: df_f = df_f[df_f['거래처명'].isin(f_v)]
        if f_c: df_f = df_f[df_f['유형'].isin(f_c)]
        if f_m: df_f = df_f[df_f['월'].isin(f_m)]

        # 마감 상태 병합 및 정렬
        if not o_all.empty:
            df_f = df_f.merge(o_all[['발주번호', '마감여부']], on='발주번호', how='left').fillna(0)
        else:
            df_f['마감여부'] = 0
            
        st.dataframe(df_f.sort_values('입금일', ascending=False).style.apply(style_closed_rows, axis=1), use_container_width=True)
        
        st.divider()
        st.subheader("🚩 발주 마감 처리")
        if not o_all.empty:
            target_oid = st.selectbox("마감할 발주번호 선택 ", o_all[o_all['마감여부']==0]['발주번호'].unique() if not o_all.empty else ["없음"])
            if st.button("해당 발주 마감하기 "):
                conn.cursor().execute("UPDATE orders SET 마감여부 = 1 WHERE 발주번호 = ?", (target_oid,))
                conn.commit()
                st.rerun()