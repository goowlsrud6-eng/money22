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
    db_file = 'finance_v33_final.db'
    backup_file = f"backups/backup_{today_str}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

# 2. 페이지 설정
st.set_page_config(page_title="자금 관리 v33", layout="wide", page_icon="💰")
run_backup()

# 3. DB 연결 및 초기화
@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_v33_final.db', check_same_thread=False)
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
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비"]
CURRENCIES = ["한화", "USD", "CNY"]

def load_table(table):
    return pd.read_sql(f"SELECT * FROM {table}", conn)

# --- 4. 이카운트 정밀 분석 함수 (B2/C7 대응 & 중국/USD 인식) ---
def process_ecount_v33(file):
    try:
        df_raw = pd.read_excel(file, header=None)
        raw_oid = str(df_raw.iloc[0, 0]).split(":")[-1].strip() if ":" in str(df_raw.iloc[0,0]) else str(df_raw.iloc[0, 0])
        oid = raw_oid
        odate = f"{oid[:4]}-{oid[4:6]}-{oid[6:8]}" if len(oid) >= 8 else datetime.now().strftime("%Y-%m-%d")
        
        vendor = "미지정"
        for i in range(len(df_raw)):
            if "수신" in str(df_raw.iloc[i, 0]):
                vendor = str(df_raw.iloc[i, 0]).split(":")[-1].strip()
                break
        
        # 통화 판별: F6 셀에 "중국" 포함 시 CNY
        f6_val = str(df_raw.iloc[5, 5]) if len(df_raw) > 5 else ""
        if "USD" in f6_val: currency = "USD"
        elif "중국" in f6_val or "CNY" in f6_val: currency = "CNY"
        else: currency = "한화"

        # 품목명: 국내(B2), 해외(C7~)
        if currency == "한화":
            raw_prod = str(df_raw.iloc[1, 1]) if len(df_raw) > 1 else "품목미상"
            product_display = raw_prod.split("[")[0].strip()
        else:
            raw_prod_list = df_raw.iloc[6:, 2].dropna().tolist() 
            if raw_prod_list:
                first_prod = str(raw_prod_list[0]).split("[")[0].strip()
                product_display = f"{first_prod} 외 {len(raw_prod_list)-1}건" if len(raw_prod_list) > 1 else first_prod
            else: product_display = "해외품목미상"

        # 금액 파싱
        total_amt = 0.0
        if currency != "한화":
            last_f = df_raw.iloc[:, 5].last_valid_index()
            total_amt = float(df_raw.iloc[last_f, 5]) if last_f is not None else 0.0
        else:
            a5_val = str(df_raw.iloc[4, 0])
            if "금액" in a5_val:
                total_amt = float(a5_val.split(":")[-1].replace(",", "").strip())
        
        # 유형 매칭
        v_df = load_table("vendors")
        matched_v = v_df[v_df['거래처명'] == vendor]
        v_type = matched_v.iloc[0]['기본유형'] if not matched_v.empty else "사입"

        conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", 
                             (oid, odate, vendor, product_display, v_type, currency, total_amt))
        conn.commit()
        return True, oid
    except Exception as e:
        return False, str(e)

# --- 스타일 함수 ---
def style_rows(row):
    if '마감여부' in row and row['마감여부'] == 1:
        return ['background-color: #f5f5f5; color: #a0a0a0; text-decoration: line-through'] * len(row)
    return [''] * len(row)

# --- 5. 메인 UI ---
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리"])

# --- Tab 1: 입금 입력 (수기) ---
with tabs[0]:
    st.header("📝 입금 내역 수기 입력")
    v_master = load_table("vendors")
    o_master = load_table("orders")
    active_orders = o_master[o_master['마감여부'] == 0] if not o_master.empty else pd.DataFrame()
    with st.form("pay_manual", clear_on_submit=True):
        sel_oid = st.selectbox("🔗 발주번호 연동", ["없음"] + list(active_orders['발주번호']) if not active_orders.empty else ["없음"])
        c1, c2, c3 = st.columns(3)
        p_date, p_vendor = c1.date_input("입금일"), c2.selectbox("거래처명", ["선택"] + list(v_master['거래처명']) if not v_master.empty else ["선택"])
        p_cat = c3.selectbox("유형", CATEGORIES)
        c4, c5, c6 = st.columns(3)
        p_dep, p_pre, p_curr = c4.number_input("실입금액"), c5.number_input("선급금액"), c6.selectbox("통화", CURRENCIES)
        p_note = st.text_input("메모")
        if st.form_submit_button("입금 저장"):
            if p_vendor != "선택":
                rate = 1350.0 if p_curr == "USD" else (190.0 if p_curr == "CNY" else 1.0)
                v_info = v_master[v_master['거래처명'] == p_vendor].iloc[0]
                conn.cursor().execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                         VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                                      (sel_oid if sel_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), p_cat, p_vendor, "품목", p_curr, p_dep, p_pre, p_note, (p_dep + p_pre)*rate, v_info['은행'], v_info['계좌번호'], v_info['예금주']))
                conn.commit()
                st.success("저장 완료!"); st.rerun()

# --- Tab 2: 입금 엑셀 업로드 (복구 완료) ---
with tabs[1]:
    st.header("📂 입금 내역 엑셀 업로드")
    st.info("CSV 파일로 일괄 등록합니다.")
    p_temp = pd.DataFrame(columns=["입금일", "거래처", "유형", "통화", "상품명", "입금액", "선급금", "송금사유", "발주번호"])
    st.download_button("📥 양식 다운로드", p_temp.to_csv(index=False).encode('utf-8-sig'), "pay_temp.csv")
    p_file = st.file_uploader("입금 내역 CSV 선택", type=['csv'])
    if p_file and st.button("🚀 입금 내역 일괄 저장"):
        df_p = pd.read_csv(p_file)
        v_list = load_table("vendors")
        for _, r in df_p.iterrows():
            curr = r['통화'] if pd.notna(r['통화']) else "한화"
            rate = 1350.0 if curr == "USD" else (190.0 if curr == "CNY" else 1.0)
            v_info = v_list[v_list['거래처명'] == r['거래처']]
            b, a, h = (v_info.iloc[0]['은행'], v_info.iloc[0]['계좌번호'], v_info.iloc[0]['예금주']) if not v_info.empty else ("","","")
            conn.cursor().execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                   VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                                (r['발주번호'], r['입금일'], r['유형'], r['거래처'], r['상품명'], curr, r['입금액'], r['선급금'], r['송금사유'], float(r['입금액'])*rate, b, a, h))
        conn.commit(); st.success("업로드 성공!"); st.rerun()

# --- Tab 3: 발주서 등록 (자동 + 수기 통합 복구) ---
with tabs[2]:
    st.header("📥 발주서 등록")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("⚡ 1. 이카운트 엑셀 즉시 등록")
        o_file = st.file_uploader("이카운트 발주서(.xlsx)", type=['xlsx'], key="erp_up")
        if o_file:
            success, res = process_ecount_v33(o_file)
            if success: st.success(f"✅ 등록 완료: {res}")
            else: st.error(f"❌ 오류: {res}")
    with col2:
        st.subheader("✍️ 2. 발주서 수기 등록 (복구)")
        v_list = load_table("vendors")
        with st.form("order_manual_v33", clear_on_submit=True):
            m_id = st.text_input("발주번호")
            m_date = st.date_input("발주일")
            m_vendor = st.selectbox("거래처", ["선택"] + list(v_list['거래처명']) if not v_list.empty else ["선택"])
            m_cat = st.selectbox("유형", CATEGORIES)
            m_total = st.number_input("발주총액")
            if st.form_submit_button("🚀 수기 저장"):
                if m_id and m_vendor != "선택":
                    conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", (m_id, m_date.strftime("%Y-%m-%d"), m_vendor, "수기등록", m_cat, "한화", m_total))
                    conn.commit(); st.success("수기 저장 완료!"); st.rerun()
    st.divider()
    st.subheader("📑 발주 리스트")
    st.dataframe(load_table("orders").sort_values('발주일', ascending=False).style.apply(style_rows, axis=1), use_container_width=True)

# --- Tab 4: 상세조회 & Tab 5: 거래처 (동일하게 유지) ---
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    with st.form("v_reg_v33"):
        vn, vt = st.text_input("업체명"), st.selectbox("기본유형", CATEGORIES)
        vb, va, vh = st.text_input("은행"), st.text_input("계좌번호"), st.text_input("예금주")
        if st.form_submit_button("거래처 저장"):
            if vn: conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (vn,vb,va,vh,vt)); conn.commit(); st.rerun()
    st.dataframe(load_table("vendors"), use_container_width=True)

with tabs[3]:
    st.header("🔍 상세 내역")
    p_all = load_table("payments")
    if not p_all.empty:
        st.dataframe(p_all.sort_values('입금일', ascending=False), use_container_width=True)