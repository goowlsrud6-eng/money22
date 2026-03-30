import streamlit as st
import pandas as pd
import sqlite3
import os
import shutil
from datetime import datetime

# --- 1. 데이터 안전장치 ---
def run_backup():
    if not os.path.exists('backups'): os.makedirs('backups')
    db_file = 'finance_v43_final.db'
    backup_file = f"backups/backup_{datetime.now().strftime('%Y%m%d')}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

st.set_page_config(page_title="자금 관리 v43", layout="wide", page_icon="💰")
run_backup()

@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_v43_final.db', check_same_thread=False)
    c = conn.cursor()
    c.execute('CREATE TABLE IF NOT EXISTS vendors (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT, 기본유형 TEXT)')
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
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비", "원단비", "기타"]
CURRENCIES = ["한화", "USD", "CNY"]

def load_table(table):
    return pd.read_sql(f"SELECT * FROM {table}", conn)

def style_closed(row):
    if '마감여부' in row and row['마감여부'] == 1:
        return ['background-color: #f5f5f5; color: #a0a0a0; text-decoration: line-through'] * len(row)
    return [''] * len(row)

# --- 2. 이카운트 정밀 분석 함수 (v43 정식 명칭 고정) ---
def process_ecount_v43(file):
    try:
        df = pd.read_excel(file, header=None)
        raw_oid = str(df.iloc[1, 0]).split(":")[-1].strip() if ":" in str(df.iloc[1,0]) else str(df.iloc[1, 0])
        clean_oid = raw_oid.replace("-", "")
        odate = f"{clean_oid[:4]}-{clean_oid[4:6]}-{clean_oid[6:8]}" if len(clean_oid) >= 8 else datetime.now().strftime("%Y-%m-%d")
        
        vendor = "미지정"
        for i in range(len(df)):
            if "수신" in str(df.iloc[i, 0]):
                vendor = str(df.iloc[i, 0]).split(":")[-1].strip()
                break
        
        f6 = str(df.iloc[5, 5]) if len(df) > 5 else ""
        curr = "USD" if "USD" in f6 else ("CNY" if any(x in f6 for x in ["중국", "CNY"]) else "한화")
        
        prod_col = 1 if curr == "한화" else 2
        prods = df.iloc[6:, prod_col].dropna().astype(str).tolist()
        if prods:
            first_item = prods[0].split("[")[0].strip()
            prod_name = f"{first_item} 외 {len(prods)-1}건" if len(prods) > 1 else first_item
        else: prod_name = "품목미상"
        
        total = 0.0
        if curr != "한화":
            idx = df.iloc[:, 5].last_valid_index()
            total = float(df.iloc[idx, 5]) if idx else 0.0
        else:
            a5 = str(df.iloc[4, 0])
            total = float(a5.split(":")[-1].replace(",", "").strip()) if "금액" in a5 else 0.0

        v_df = load_table("vendors")
        v_type = v_df[v_df['거래처명'] == vendor].iloc[0]['기본유형'] if not v_df[v_df['거래처명'] == vendor].empty else "사입"
        
        conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", (raw_oid, odate, vendor, prod_name, v_type, curr, total))
        conn.commit()
        return True, raw_oid
    except Exception as e:
        return False, str(e)

# --- 3. UI 구성 ---
tabs = st.tabs(["📝 입금 수기 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리"])

# --- Tab 1: 입금 수기 ---
with tabs[0]:
    st.header("📝 입금 내역 수기 입력")
    v_m = load_table("vendors"); o_m = load_table("orders")
    with st.form("p_manual_v43", clear_on_submit=True):
        sel_oid = st.selectbox("🔗 발주번호 연동", ["없음"] + list(o_m['발주번호']) if not o_m.empty else ["없음"])
        c1, c2, c3 = st.columns(3)
        p_d, p_v = c1.date_input("입금일"), c2.selectbox("거래처명", ["선택"] + list(v_m['거래처명']) if not v_m.empty else ["선택"])
        p_cat = c3.selectbox("유형", CATEGORIES)
        c4, c5, c6 = st.columns(3)
        p_dep, p_pre, p_curr = c4.number_input("실입금액"), c5.number_input("선급금액"), c6.selectbox("통화", CURRENCIES)
        p_note = st.text_input("메모")
        if st.form_submit_button("입금 내역 저장"):
            if p_v != "선택":
                rate = 1350.0 if p_curr == "USD" else (190.0 if p_curr == "CNY" else 1.0)
                vi = v_m[v_m['거래처명']==p_v].iloc[0]
                conn.cursor().execute("INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                                     (sel_oid if sel_oid != "없음" else None, p_d.strftime("%Y-%m-%d"), p_cat, p_v, "수기입력", p_curr, p_dep, p_pre, p_note, (p_dep+p_pre)*rate, vi['은행'], vi['계좌번호'], vi['예금주']))
                conn.commit(); st.success("저장 완료!"); st.rerun()

# --- Tab 2: 입금 엑셀 업로드 ---
with tabs[1]:
    st.header("📂 통합 입금 엑셀 업로드")
    tmp = pd.DataFrame(columns=["발주번호", "거래처", "유형", "상품명", "입금일", "실입금액", "선급금액", "송금사유"])
    st.download_button("📥 통합 양식 다운로드", tmp.to_csv(index=False).encode('utf-8-sig'), "payment_v43.csv")
    p_file = st.file_uploader("CSV 업로드", type=['csv'], key="p_csv_v43")
    if p_file:
        if st.button("🚀 데이터 분석 및 저장하기"):
            try:
                df_p = pd.read_csv(p_file)
                o_df = load_table("orders"); v_df = load_table("vendors")
                for _, r in df_p.iterrows():
                    raw_date = pd.to_datetime(r['입금일'], errors='coerce')
                    p_date_str = raw_date.strftime('%Y-%m-%d') if pd.notnull(raw_date) else datetime.now().strftime('%Y-%m-%d')
                    oid = str(r['발주번호']).strip() if pd.notna(r['발주번호']) else None
                    if oid and oid != "nan" and not o_df[o_df['발주번호'] == oid].empty:
                        info = o_df[o_df['발주번호'] == oid].iloc[0]
                        v_n, p_c, p_p, p_cur = info['거래처명'], info['유형'], info['상품명'], info['통화']
                    else:
                        v_n, p_c, p_p, p_cur = r['거래처'], r['유형'], r['상품명'], "한화"
                    v_i = v_df[v_df['거래처명'] == v_n]
                    bank, acc, holder = (v_i.iloc[0]['은행'], v_i.iloc[0]['계좌번호'], v_i.iloc[0]['예금주']) if not v_i.empty else ("","","")
                    rate = 1350.0 if p_cur == "USD" else (190.0 if p_cur == "CNY" else 1.0)
                    total = float(r['실입금액']) + float(r['선급금액'])
                    conn.cursor().execute("INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                                         (oid, p_date_str, p_c, v_n, p_p, p_cur, r['실입금액'], r['선급금액'], r['송금사유'], total*rate, bank, acc, holder))
                conn.commit(); st.success("성공적으로 저장되었습니다!"); st.rerun()
            except Exception as e: st.error(f"에러: {e}")

# --- Tab 3: 발주서 등록 (오타 수정됨) ---
with tabs[2]:
    st.header("📥 발주서 등록")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("⚡ 1. 이카운트 엑셀 즉시 등록")
        o_f = st.file_uploader("이카운트 엑셀(.xlsx)", type=['xlsx'], key="o_f_v43")
        if o_f: 
            # 여기서 process_ecount_v43를 정확히 호출합니다.
            s, r = process_ecount_v43(o_f)
            if s: st.success(f"등록 성공: {r}")
            else: st.error(f"오류: {r}")
    with c2:
        st.subheader("✍️ 2. 발주서 수기 등록")
        v_l = load_table("vendors")
        with st.form("o_manual_v43", clear_on_submit=True):
            m_id = st.text_input("발주번호"); m_date = st.date_input("발주일")
            m_v = st.selectbox("거래처", ["선택"] + list(v_l['거래처명']) if not v_l.empty else ["선택"])
            m_cat = st.selectbox("유형", CATEGORIES); m_total = st.number_input("발주총액")
            if st.form_submit_button("🚀 수기 저장"):
                if m_id and m_v != "선택":
                    conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", (m_id, m_date.strftime("%Y-%m-%d"), m_v, "수기등록", m_cat, "한화", m_total))
                    conn.commit(); st.rerun()
    st.divider(); st.dataframe(load_table("orders").sort_values('발주일', ascending=False).style.apply(style_closed, axis=1), use_container_width=True)

# --- Tab 4: 상세내역 ---
with tabs[3]:
    st.header("🔍 상세 내역 및 정산")
    p_all = load_table("payments"); o_all = load_table("orders")
    if not p_all.empty:
        p_all['입금일_dt'] = pd.to_datetime(p_all['입금일'], errors='coerce')
        p_all['월'] = p_all['입금일_dt'].dt.strftime('%Y-%m').fillna('날짜오류')
        f1, f2, f3 = st.columns(3)
        with f1: f_v = st.multiselect("🏢 업체 필터", p_all['거래처명'].unique())
        with f2: f_c = st.multiselect("📁 유형 필터", CATEGORIES)
        with f3: f_m = st.multiselect("📅 월별 필터", sorted(p_all['월'].unique(), reverse=True))
        df_f = p_all.copy()
        if f_v: df_f = df_f[df_f['거래처명'].isin(f_v)]
        if f_c: df_f = df_f[df_f['유형'].isin(f_c)]
        if f_m: df_f = df_f[df_f['월'].isin(f_m)]
        if not o_all.empty: df_f = df_f.merge(o_all[['발주번호', '마감여부']], on='발주번호', how='left').fillna(0)
        st.dataframe(df_f.drop(columns=['입금일_dt', '월']).style.apply(style_closed, axis=1), use_container_width=True)
        st.divider(); target_oid = st.selectbox("마감할 발주번호 선택", o_all[o_all['마감여부']==0]['발주번호'].unique() if not o_all.empty else [])
        if st.button("🚩 해당 발주 마감"):
            conn.cursor().execute("UPDATE orders SET 마감여부 = 1 WHERE 발주번호 = ?", (target_oid,))
            conn.commit(); st.rerun()

# --- Tab 5: 거래처 관리 ---
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    cv1, cv2 = st.columns(2)
    with cv1:
        with st.form("v_reg_v43", clear_on_submit=True):
            vn, vt = st.text_input("업체명"), st.selectbox("기본유형", CATEGORIES)
            vb, va, vh = st.text_input("은행"), st.text_input("계좌"), st.text_input("예금주")
            if st.form_submit_button("저장"):
                if vn: conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (vn,vb,va,vh,vt)); conn.commit(); st.rerun()
    with cv2:
        v_file = st.file_uploader("거래처 CSV", type=['csv'], key="v_csv_v43")
        if v_file and st.button("🚀 일괄 저장"):
            v_df = pd.read_csv(v_file)
            for _, r in v_df.iterrows():
                conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (r['거래처명'], r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
            conn.commit(); st.rerun()
    st.dataframe(load_table("vendors"), use_container_width=True)