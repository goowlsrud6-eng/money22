import streamlit as st
import pandas as pd
import sqlite3
import os
import shutil
from datetime import datetime

# --- 1. 데이터 안전장치 (자동 백업) ---
def run_backup():
    if not os.path.exists('backups'): os.makedirs('backups')
    today_str = datetime.now().strftime("%Y%m%d")
    db_file = 'finance_v30_final.db'
    backup_file = f"backups/backup_{today_str}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

# 2. 페이지 설정
st.set_page_config(page_title="자금 관리 v30", layout="wide", page_icon="💰")
run_backup()

# 3. DB 연결 및 초기화 (한글 컬럼명 유지)
@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_v30_final.db', check_same_thread=False)
    c = conn.cursor()
    # 거래처 마스터
    c.execute('CREATE TABLE IF NOT EXISTS vendors (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT)')
    # 발주 마스터
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (발주번호 TEXT PRIMARY KEY, 발주일 TEXT, 거래처명 TEXT, 
                  상품명 TEXT, 유형 TEXT, 통화 TEXT, 발주총액 REAL, 마감여부 INTEGER DEFAULT 0)''')
    # 입금 내역
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
    return pd.read_sql(f"SELECT * FROM {table}", conn)

# --- 스타일 함수: 마감 건 회색 음영 ---
def style_closed_rows(row):
    if '마감여부' in row and row['마감여부'] == 1:
        return ['background-color: #f5f5f5; color: #a0a0a0; text-decoration: line-through'] * len(row)
    return [''] * len(row)

# --- 4. 이카운트 엑셀 즉시 분석 및 저장 함수 ---
def process_ecount_and_save(file):
    try:
        df_raw = pd.read_excel(file, header=None)
        # 전표번호 추출 (A1셀 등)
        raw_val = str(df_raw.iloc[0, 0])
        oid = raw_val.split(":")[-1].strip() if ":" in raw_val else raw_val.strip()
        
        # 날짜 추출 (20260304-1 -> 2026-03-04)
        odate = f"{oid[:4]}-{oid[4:6]}-{oid[6:8]}" if len(oid) >= 8 else datetime.now().strftime("%Y-%m-%d")
        
        # 거래처 추출
        vendor = "미지정"
        for i in range(len(df_raw)):
            if "수신" in str(df_raw.iloc[i, 0]):
                vendor = str(df_raw.iloc[i, 0]).split(":")[-1].strip()
                break
        
        # 금액 파싱
        total_amt = 0.0
        currency = "한화"
        f_col = df_raw.iloc[:, 5]
        last_f = f_col.last_valid_index()
        if last_f is not None and isinstance(f_col[last_f], (int, float)) and f_col[last_f] > 100:
            total_amt = float(f_col[last_f]); currency = "외화"
        else:
            a5_val = str(df_raw.iloc[4, 0])
            if "금액" in a5_val:
                total_amt = float(a5_val.split(":")[-1].replace(",", "").strip())
        
        # 즉시 저장
        conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", 
                             (oid, odate, vendor, "이카운트 자동등록", "사입/제작", currency, total_amt))
        conn.commit()
        return True, oid
    except Exception as e:
        return False, str(e)

# --- 5. 메인 UI ---
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록(ERP/수기)", "🔍 상세조회 및 필터", "⚙️ 거래처 관리"])

# --- Tab 1: 입금 입력 (수기) ---
with tabs[0]:
    st.header("📝 입금 내역 수기 입력")
    v_data = load_table("vendors")
    o_data = load_table("orders")
    active_orders = o_data[o_data['마감여부'] == 0] if not o_data.empty else pd.DataFrame()
    
    with st.form("pay_form_v30", clear_on_submit=True):
        sel_oid = st.selectbox("🔗 발주번호 연동", ["없음"] + list(active_orders['발주번호']) if not active_orders.empty else ["없음"])
        c1, c2, c3 = st.columns(3)
        p_date = c1.date_input("입금일")
        p_vendor = c2.selectbox("거래처명", ["선택"] + list(v_data['거래처명']) if not v_data.empty else ["선택"])
        p_cat = c3.selectbox("유형", CATEGORIES)
        
        c4, c5, c6 = st.columns(3)
        p_dep = c4.number_input("💵 실입금액 (통장금액)", min_value=0.0)
        p_pre = c5.number_input("🧧 선급금 (발생/사용)", value=0.0)
        p_curr = c6.selectbox("통화", CURRENCIES)
        p_note = st.text_input("메모(송금사유)")
        
        if st.form_submit_button("입금 저장"):
            if p_vendor == "선택": st.error("거래처를 선택하세요.")
            else:
                rate = 1350.0 if p_curr == "USD" else (190.0 if p_curr == "CNY" else 1.0)
                v_info = v_data[v_data['거래처명'] == p_vendor].iloc[0]
                conn.cursor().execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                         VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                                      (sel_oid if sel_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), 
                                       p_cat, p_vendor, "품목", p_curr, p_dep, p_pre, p_note, (p_dep + p_pre)*rate, 
                                       v_info['은행'], v_info['계좌번호'], v_info['예금주']))
                conn.commit()
                st.success("저장 완료!")
                st.rerun()

# --- Tab 2: 입금 엑셀 업로드 ---
with tabs[1]:
    st.header("📂 입금 내역 엑셀 업로드")
    p_temp = pd.DataFrame(columns=["입금일", "거래처", "유형", "통화", "상품명", "입금액", "선급금", "송금사유", "발주번호"])
    st.download_button("📥 양식 다운로드", p_temp.to_csv(index=False).encode('utf-8-sig'), "pay_temp.csv")
    p_file = st.file_uploader("입금 내역 CSV", type=['csv'])
    if p_file:
        df_p = pd.read_csv(p_file)
        if st.button("🚀 데이터 일괄 저장"):
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
            st.success("업로드 완료!")
            st.rerun()

# --- Tab 3: 발주서 등록 (ERP 자동 + 수기 통합) ---
with tabs[2]:
    st.header("📥 발주서 등록")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("⚡ 1. 이카운트 엑셀 즉시 등록")
        o_file = st.file_uploader("이카운트 발주서(.xlsx)", type=['xlsx'], key="erp_up")
        if o_file:
            success, res = process_ecount_and_save(o_file)
            if success: st.success(f"✅ 발주번호 [{res}] 리스트에 즉시 등록되었습니다!")
            else: st.error(f"❌ 오류: {res}")
    with c2:
        st.subheader("✍️ 2. 발주서 수기 등록")
        v_list = load_table("vendors")
        with st.form("order_manual_v30", clear_on_submit=True):
            m_id = st.text_input("발주번호")
            m_date = st.date_input("발주일", datetime.now())
            m_vendor = st.selectbox("거래처", ["선택"] + list(v_list['거래처명']) if not v_list.empty else ["선택"])
            m_total = st.number_input("발주총액", min_value=0.0)
            if st.form_submit_button("🚀 수기 저장"):
                if m_id and m_vendor != "선택":
                    conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", 
                                         (m_id, m_date.strftime("%Y-%m-%d"), m_vendor, "수기등록", "사입", "한화", m_total))
                    conn.commit()
                    st.rerun()
    st.divider()
    st.subheader("📑 등록된 발주 리스트 (진행/마감)")
    st.dataframe(load_table("orders").style.apply(style_closed_rows, axis=1), use_container_width=True)

# --- Tab 4: 상세조회 및 필터링 (음영/마감 포함) ---
with tabs[3]:
    st.header("🔍 상세 내역 및 정산 관리")
    p_all = load_table("payments")
    o_all = load_table("orders")
    if not p_all.empty:
        f1, f2, f3 = st.columns(3)
        with f1: f_v = st.multiselect("🏢 업체 필터", p_all['거래처명'].unique())
        with f2: f_c = st.multiselect("📁 유형 필터", CATEGORIES)
        with f3: 
            p_all['월'] = pd.to_datetime(p_all['입금일']).dt.strftime('%Y-%m')
            f_m = st.multiselect("📅 월별 필터", sorted(p_all['월'].unique(), reverse=True))
        
        df_f = p_all.copy()
        if f_v: df_f = df_f[df_f['거래처명'].isin(f_v)]
        if f_c: df_f = df_f[df_f['유형'].isin(f_c)]
        if f_m: df_f = df_f[df_f['월'].isin(f_m)]
        
        if not o_all.empty:
            df_f = df_f.merge(o_all[['발주번호', '마감여부']], on='발주번호', how='left').fillna(0)
        st.dataframe(df_f.style.apply(style_closed_rows, axis=1), use_container_width=True)
        
        st.divider()
        target_oid = st.selectbox("마감할 발주번호", o_all[o_all['마감여부']==0]['발주번호'].unique() if not o_all.empty else [])
        if st.button("🚩 해당 발주 마감하기"):
            conn.cursor().execute("UPDATE orders SET 마감여부 = 1 WHERE 발주번호 = ?", (target_oid,))
            conn.commit()
            st.rerun()

# --- Tab 5: 거래처 관리 (리스트 노출) ---
with tabs[4]:
    st.header("⚙️ 거래처 및 계좌 관리")
    cv1, cv2 = st.columns(2)
    with cv1:
        with st.form("v_reg_v30", clear_on_submit=True):
            st.subheader("➕ 개별 등록")
            vn, vb, va, vh = st.text_input("업체명"), st.text_input("은행"), st.text_input("계좌번호"), st.text_input("예금주")
            if st.form_submit_button("거래처 저장"):
                if vn:
                    conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?)", (vn,vb,va,vh))
                    conn.commit()
                    st.rerun()
    with cv2:
        st.subheader("📂 엑셀(CSV) 일괄 등록")
        v_temp = pd.DataFrame(columns=["거래처명", "은행", "계좌번호", "예금주"])
        st.download_button("📥 양식 받기", v_temp.to_csv(index=False).encode('utf-8-sig'), "vendor_temp.csv")
        v_file = st.file_uploader("거래처 CSV 업로드", type=['csv'])
        if v_file:
            v_df = pd.read_csv(v_file)
            if st.button("🚀 거래처 일괄 저장"):
                for _, row in v_df.iterrows():
                    conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?)", (row['거래처명'], row['은행'], row['계좌번호'], row['예금주']))
                conn.commit()
                st.rerun()
    st.dataframe(load_table("vendors"), use_container_width=True)