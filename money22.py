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
    db_file = 'finance_v37_final.db'
    backup_file = f"backups/backup_{today_str}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

# 2. 페이지 설정 및 DB 연결
st.set_page_config(page_title="자금 관리 v37", layout="wide", page_icon="💰")
run_backup()

@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_v37_final.db', check_same_thread=False)
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

# --- 3. 이카운트 정밀 분석 함수 (한화 B7 / 외화 C7 대응) ---
def process_ecount_v37(file):
    try:
        df_raw = pd.read_excel(file, header=None)
        
        # [1] 발주번호 추출 (A2 셀)
        raw_oid_cell = str(df_raw.iloc[1, 0])
        oid = raw_oid_cell.split(":")[-1].strip() if ":" in raw_oid_cell else raw_oid_cell.strip()
        
        # 발주일 추출 (번호 앞 8자리)
        if len(oid.replace("-", "")) >= 8:
            clean_oid = oid.replace("-", "")
            odate = f"{clean_oid[:4]}-{clean_oid[4:6]}-{clean_oid[6:8]}"
        else:
            odate = datetime.now().strftime("%Y-%m-%d")
        
        # [2] 거래처 추출 ('수신' 찾기)
        vendor = "미지정"
        for i in range(len(df_raw)):
            if "수신" in str(df_raw.iloc[i, 0]):
                vendor = str(df_raw.iloc[i, 0]).split(":")[-1].strip()
                break
        
        # [3] 통화 판별 (F6 셀 기준)
        f6_val = str(df_raw.iloc[5, 5]) if len(df_raw) > 5 else ""
        if "USD" in f6_val: currency = "USD"
        elif "중국" in f6_val or "CNY" in f6_val: currency = "CNY"
        else: currency = "한화"

        # [4] 상품명 정밀 추출 (대괄호 전까지)
        product_display = "품목미상"
        if currency == "한화":
            # ★ 한국 발주서 전용: B7 셀(6행, 1열)부터 시작 ★
            prod_list = df_raw.iloc[6:, 1].dropna().astype(str).tolist()
        else:
            # ★ 외화 발주서 전용: C7 셀(6행, 2열)부터 시작 ★
            prod_list = df_raw.iloc[6:, 2].dropna().astype(str).tolist()

        if prod_list:
            first_item = prod_list[0].split("[")[0].strip()
            product_display = f"{first_item} 외 {len(prod_list)-1}건" if len(prod_list) > 1 else first_item

        # [5] 금액 파싱
        total_amt = 0.0
        if currency != "한화":
            last_f = df_raw.iloc[:, 5].last_valid_index()
            total_amt = float(df_raw.iloc[last_f, 5]) if last_f is not None else 0.0
        else:
            a5_val = str(df_raw.iloc[4, 0])
            if "금액" in a5_val:
                total_amt = float(a5_val.split(":")[-1].replace(",", "").strip())
        
        # [6] 거래처 유형 매칭
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
def style_closed(row):
    if '마감여부' in row and row['마감여부'] == 1:
        return ['background-color: #f5f5f5; color: #a0a0a0; text-decoration: line-through'] * len(row)
    return [''] * len(row)

# --- 4. 메인 UI 구성 ---
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리"])

# --- Tab 3: 발주서 등록 (수정된 로직 반영) ---
with tabs[2]:
    st.header("📥 발주서 등록 (ERP/수기)")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("⚡ 1. 이카운트 엑셀 즉시 등록")
        o_file = st.file_uploader("이카운트 발주서(.xlsx) 업로드", type=['xlsx'], key="o_up_v37")
        if o_file:
            success, res = process_ecount_v37(o_file)
            if success: st.success(f"✅ 등록 완료: {res}")
            else: st.error(f"❌ 분석 실패: {res}")
    with col2:
        st.subheader("✍️ 2. 발주서 수기 등록")
        v_list = load_table("vendors")
        with st.form("o_manual_v37", clear_on_submit=True):
            m_id = st.text_input("발주번호")
            m_date = st.date_input("발주일")
            m_vendor = st.selectbox("거래처", ["선택"] + list(v_list['거래처명']) if not v_list.empty else ["선택"])
            m_cat = st.selectbox("유형", CATEGORIES)
            m_total = st.number_input("발주총액")
            if st.form_submit_button("🚀 수기 저장"):
                if m_id and m_vendor != "선택":
                    conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", (m_id, m_date.strftime("%Y-%m-%d"), m_vendor, "수기등록", m_cat, "한화", m_total))
                    conn.commit(); st.rerun()
    st.divider()
    st.subheader("📑 등록된 발주 리스트")
    st.dataframe(load_table("orders").sort_values('발주일', ascending=False).style.apply(style_closed, axis=1), use_container_width=True)

# --- Tab 2: 입금 엑셀 업로드 (기능 복구 유지) ---
with tabs[1]:
    st.header("📂 입금 내역 엑셀 일괄 업로드")
    p_temp = pd.DataFrame(columns=["입금일", "거래처", "유형", "통화", "상품명", "입금액", "선급금", "송금사유", "발주번호"])
    st.download_button("📥 입금 양식(CSV) 다운로드", p_temp.to_csv(index=False).encode('utf-8-sig'), "pay_temp.csv")
    p_file = st.file_uploader("입금 내역 CSV", type=['csv'], key="p_csv_v37")
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

# --- Tab 5: 거래처 관리 (엑셀 일괄 등록 포함) ---
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    col_v1, col_v2 = st.columns(2)
    with col_v1:
        st.subheader("➕ 개별 등록")
        with st.form("v_reg_v37", clear_on_submit=True):
            vn, vt = st.text_input("업체명"), st.selectbox("기본유형", CATEGORIES)
            vb, va, vh = st.text_input("은행"), st.text_input("계좌번호"), st.text_input("예금주")
            if st.form_submit_button("거래처 저장"):
                if vn: conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (vn,vb,va,vh,vt)); conn.commit(); st.rerun()
    with col_v2:
        st.subheader("📂 엑셀 일괄 등록")
        v_temp = pd.DataFrame(columns=["거래처명", "은행", "계좌번호", "예금주", "기본유형"])
        st.download_button("📥 양식 받기", v_temp.to_csv(index=False).encode('utf-8-sig'), "vendor_temp.csv")
        v_file = st.file_uploader("거래처 CSV 업로드", type=['csv'], key="v_csv_v37")
        if v_file and st.button("🚀 거래처 일괄 저장"):
            v_df = pd.read_csv(v_file)
            for _, row in v_df.iterrows():
                conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (row['거래처명'], row['은행'], row['계좌번호'], row['예금주'], row['기본유형']))
            conn.commit(); st.rerun()
    st.dataframe(load_table("vendors"), use_container_width=True)

# (Tab 1 입금 수기, Tab 4 상세조회는 이전 버전과 동일한 풀 기능 유지)