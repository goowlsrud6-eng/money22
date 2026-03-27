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
    db_file = 'finance_v21_final.db'
    backup_file = f"backups/backup_{today_str}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

# 2. 페이지 설정
st.set_page_config(page_title="자금 관리 v21", layout="wide", page_icon="💰")
run_backup()

# 3. DB 연결 (오류 방지를 위해 INSERT OR REPLACE 로직 적용 준비)
@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_v21_final.db', check_same_thread=False)
    c = conn.cursor()
    # 거래처 마스터
    c.execute('CREATE TABLE IF NOT EXISTS vendors (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT)')
    # 발주 마스터
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (발주번호 TEXT PRIMARY KEY, 발주일 TEXT, 거래처명 TEXT, 
                  상품명 TEXT, 유형 TEXT, 통화 TEXT, 발주총액 REAL, 마감여부 INTEGER DEFAULT 0)''')
    # 입금 내역 (필드명 한글화 및 구조 최적화)
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

def load_table(table):
    try:
        df = pd.read_sql(f"SELECT * FROM {table}", conn)
        return df
    except:
        return pd.DataFrame()

# --- 스타일 함수: 마감 건 회색 음영 ---
def style_closed_rows(row):
    if '마감여부' in row and row['마감여부'] == 1:
        return ['background-color: #f0f0f0; color: #999; border-bottom: 1px solid #ddd'] * len(row)
    return [''] * len(row)

# --- 메인 UI ---
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리"])

# --- Tab 5: 거래처 관리 (중복 오류 해결 버전) ---
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("➕ 개별 등록")
        with st.form("v_reg_v21"):
            n = st.text_input("업체명 (필수)")
            b = st.text_input("은행")
            a = st.text_input("계좌번호")
            h = st.text_input("예금주")
            if st.form_submit_button("거래처 저장"):
                if n:
                    conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?)", (n,b,a,h))
                    conn.commit()
                    st.success(f"'{n}' 등록/갱신 완료")
                    st.rerun()
    with c2:
        st.subheader("📂 엑셀 일괄 등록")
        v_temp = pd.DataFrame(columns=["거래처명", "은행", "계좌번호", "예금주"])
        st.download_button("📥 거래처 양식 다운로드", v_temp.to_csv(index=False).encode('utf-8-sig'), "vendor_temp.csv")
        v_file = st.file_uploader("거래처 CSV 업로드", type=['csv'], key="v_csv_up")
        if v_file:
            v_df = pd.read_csv(v_file)
            if st.button("🚀 데이터 일괄 저장 (중복 시 업데이트)"):
                for _, row in v_df.iterrows():
                    conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?)", 
                                         (row['거래처명'], row['은행'], row['계좌번호'], row['예금주']))
                conn.commit()
                st.success("거래처 정보가 성공적으로 반영되었습니다.")
                st.rerun()
    
    st.subheader("📋 현재 등록된 거래처 목록")
    st.dataframe(load_table("vendors"), use_container_width=True)

# --- Tab 3: 발주서 등록 (리스트 실시간 표시) ---
with tabs[2]:
    st.header("📥 발주서 등록")
    v_data = load_table("vendors")
    with st.form("order_reg_v21"):
        c1, c2, c3, c4 = st.columns(4)
        o_id = c1.text_input("발주번호")
        o_date = c2.date_input("발주일")
        o_cat = c3.selectbox("유형", CATEGORIES)
        o_vendor = c4.selectbox("거래처 선택", ["직접입력"] + list(v_data['거래처명']) if not v_data.empty else ["업체등록 필요"])
        
        c5, c6, c7 = st.columns([2,1,1])
        o_prod = c5.text_input("상품명")
        o_total = c6.number_input("발주총액", min_value=0.0)
        o_curr = c7.selectbox("통화", CURRENCIES)
        if st.form_submit_button("🚀 발주서 확정 저장"):
            conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", 
                                 (o_id, o_date.strftime("%Y-%m-%d"), o_vendor, o_prod, o_cat, o_curr, o_total))
            conn.commit()
            st.success("발주서 등록 완료")
            st.rerun()

    st.subheader("📑 발주 리스트 현황")
    o_df = load_table("orders")
    if not o_df.empty:
        st.dataframe(o_df.style.apply(style_closed_rows, axis=1), use_container_width=True)

# --- Tab 1: 입금 입력 (선급금 구분 강화) ---
with tabs[0]:
    st.header("📝 입금 내역 입력")
    orders = load_table("orders")
    active_orders = orders[orders['마감여부'] == 0] if not orders.empty else pd.DataFrame()
    v_master = load_table("vendors")
    
    with st.form("pay_form_v21"):
        sel_oid = st.selectbox("🔗 발주번호 선택 (해당 시)", ["없음"] + list(active_orders['발주번호']) if not active_orders.empty else ["없음"])
        c1, c2, c3 = st.columns(3)
        p_date = c1.date_input("입금일")
        p_vendor = c2.selectbox("거래처명", ["선택"] + list(v_master['거래처명']) if not v_master.empty else ["선택"])
        p_cat = c3.selectbox("입금유형", CATEGORIES)
        
        st.markdown("---")
        c4, c5, c6 = st.columns(3)
        p_dep = c4.number_input("💰 실입금액 (통장금액)", min_value=0.0)
        p_pre = c5.number_input("🧧 선급금 (발생/사용액)", value=0.0, help="선급금으로 보낸 돈이거나, 이번에 차감할 선급금을 적으세요.")
        p_curr = c6.selectbox("통화 ", CURRENCIES)
        p_note = st.text_input("메모 (송금사유)")
        
        if st.form_submit_button("💾 입금 내역 저장"):
            if p_vendor == "선택": st.error("거래처를 선택하세요.")
            else:
                rate = 1350.0 if p_curr == "USD" else (190.0 if p_curr == "CNY" else 1.0)
                # 계좌 정보 자동 추출
                v_info = v_master[v_master['거래처명'] == p_vendor].iloc[0]
                conn.cursor().execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                         VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                                      (sel_oid if sel_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), 
                                       p_cat, p_vendor, "품목", p_curr, p_dep, p_pre, p_note, (p_dep + p_pre)*rate, 
                                       v_info['은행'], v_info['계좌번호'], v_info['예금주']))
                conn.commit()
                st.success("입금 완료!")
                st.rerun()

# --- Tab 4: 상세내역 관리 (한글화/필터/마감/음영) ---
with tabs[3]:
    st.header("🔍 상세 내역 및 정산 관리")
    p_data = load_table("payments")
    o_data = load_table("orders")
    
    if not p_data.empty:
        # 필터 레이아웃
        f_col1, f_col2, f_col3 = st.columns(3)
        with f_col1: f_v = st.multiselect("🏢 업체 필터", p_data['거래처명'].unique())
        with f_col2: f_c = st.multiselect("📁 유형 필터", CATEGORIES)
        with f_col3: 
            p_data['입금월'] = pd.to_datetime(p_data['입금일']).dt.strftime('%Y-%m')
            f_m = st.multiselect("📅 월별 필터", sorted(p_data['입금월'].unique(), reverse=True))

        filtered_df = p_data.copy()
        if f_v: filtered_df = filtered_df[filtered_df['거래처명'].isin(f_v)]
        if f_c: filtered_df = filtered_df[filtered_df['유형'].isin(f_c)]
        if f_m: filtered_df = filtered_df[filtered_df['입금월'].isin(f_m)]

        # 마감 상태 정보 합치기
        if not o_data.empty:
            final_view = filtered_df.merge(o_data[['발주번호', '마감여부']], on='발주번호', how='left').fillna(0)
        else:
            final_view = filtered_df
            final_view['마감여부'] = 0

        st.subheader("📑 정산 상세 리스트")
        st.dataframe(final_view.style.apply(style_closed_rows, axis=1), use_container_width=True)
        
        # 마감 처리 섹션
        st.markdown("---")
        st.subheader("🚩 발주 마감 설정")
        if not o_data.empty:
            target_oid = st.selectbox("마감할 발주번호 선택", o_data[o_data['마감여부']==0]['발주번호'].unique())
            if st.button("해당 발주 건 마감(회색처리)"):
                conn.cursor().execute("UPDATE orders SET 마감여부 = 1 WHERE 발주번호 = ?", (target_oid,))
                conn.commit()
                st.rerun()