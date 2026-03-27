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
    db_file = 'finance_v24_final.db'
    backup_file = f"backups/backup_{today_str}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

# 2. 페이지 설정
st.set_page_config(page_title="자금 관리 v24", layout="wide", page_icon="💰")
run_backup()

# 3. DB 연결 및 초기화
def get_db_connection():
    conn = sqlite3.connect('finance_v24_final.db', check_same_thread=False)
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

def load_table(table):
    return pd.read_sql(f"SELECT * FROM {table}", conn)

# --- 이카운트 엑셀 분석 함수 ---
def parse_ecount(file):
    try:
        df = pd.read_excel(file)
        # 이카운트 양식 특성상 전표번호, 거래처, 합계 금액 위치 추출 (예시 기반 최적화)
        extracted = {
            "oid": str(df.iloc[0,0]).split(":")[-1].strip() if ":" in str(df.iloc[0,0]) else "",
            "vendor": "",
            "total": 0.0,
            "product": ""
        }
        # 거래처 찾기
        for i in range(len(df)):
            val = str(df.iloc[i,0])
            if "수신" in val:
                extracted["vendor"] = val.split(":")[-1].strip()
                break
        # 금액 합계 찾기
        if '합계' in df.columns: extracted["total"] = df['합계'].sum()
        elif '공급가액' in df.columns: extracted["total"] = df['공급가액'].sum() * 1.1 # 부가세 포함 가정
        
        return extracted
    except:
        return None

# --- 스타일 함수 ---
def style_closed_rows(row):
    if '마감여부' in row and row['마감여부'] == 1:
        return ['background-color: #f0f0f0; color: #a0a0a0; text-decoration: line-through'] * len(row)
    return [''] * len(row)

# --- 메인 UI ---
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리"])

# --- Tab 3: 발주서 등록 (자동 파싱 연결 완성) ---
with tabs[2]:
    st.header("📥 발주서 등록 (ERP 엑셀 자동 분석)")
    v_data = load_table("vendors")
    
    # 엑셀 업로드 구역
    o_file = st.file_uploader("이카운트 발주서(.xlsx)를 여기에 올리세요", type=['xlsx'])
    auto_data = None
    if o_file:
        auto_data = parse_ecount(o_file)
        if auto_data:
            st.success(f"✅ 파일 분석 성공! 아래 [수기 등록] 폼에 데이터가 자동으로 채워졌습니다.")

    st.divider()
    
    # 수기 등록 폼 (자동 파싱 데이터가 있으면 기본값으로 들어감)
    st.subheader("✍️ 발주서 상세 내용 확인 및 저장")
    with st.form("order_reg_v24", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns(4)
        o_id = c1.text_input("발주번호 (전표번호)", value=auto_data["oid"] if auto_data else "")
        o_date = c2.date_input("발주일", datetime.now())
        o_cat = c3.selectbox("유형", CATEGORIES)
        
        # 거래처 자동 매칭 시도
        default_vendor_idx = 0
        if auto_data and not v_data.empty:
            if auto_data["vendor"] in list(v_data['거래처명']):
                default_vendor_idx = list(v_data['거래처명']).index(auto_data["vendor"]) + 1

        o_vendor = c4.selectbox("거래처", ["선택"] + list(v_data['거래처명']), index=default_vendor_idx if not v_data.empty else 0)
        
        c5, c6, c7 = st.columns([2,1,1])
        o_prod = c5.text_input("상품명", value="품목 외 n건" if auto_data else "")
        o_total = c6.number_input("발주총액", min_value=0.0, value=float(auto_data["total"]) if auto_data else 0.0)
        o_curr = c7.selectbox("통화", CURRENCIES)
        
        if st.form_submit_button("🚀 위 내용으로 발주서 최종 저장"):
            if not o_id or o_vendor == "선택":
                st.error("발주번호와 거래처는 필수입니다.")
            else:
                conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", 
                                     (o_id, o_date.strftime("%Y-%m-%d"), o_vendor, o_prod, o_cat, o_curr, o_total))
                conn.commit()
                st.success(f"발주번호 {o_id} 저장 완료!")
                st.rerun()

    st.subheader("📑 등록된 발주 리스트")
    st.dataframe(load_table("orders").style.apply(style_closed_rows, axis=1), use_container_width=True)

# --- Tab 5: 거래처 관리 (리스트 노출 고정) ---
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    col1, col2 = st.columns(2)
    with col1:
        with st.form("v_reg_v24", clear_on_submit=True):
            st.subheader("➕ 개별 등록")
            vn, vb, va, vh = st.text_input("업체명"), st.text_input("은행"), st.text_input("계좌번호"), st.text_input("예금주")
            if st.form_submit_button("거래처 저장"):
                if vn:
                    conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?)", (vn,vb,va,vh))
                    conn.commit()
                    st.rerun()
    with col2:
        st.subheader("📂 엑셀 일괄 등록")
        v_temp = pd.DataFrame(columns=["거래처명", "은행", "계좌번호", "예금주"])
        st.download_button("📥 양식 받기", v_temp.to_csv(index=False).encode('utf-8-sig'), "vendor_temp.csv")
        v_file = st.file_uploader("거래처 CSV 업로드", type=['csv'])
        if v_file and st.button("🚀 일괄 저장"):
            v_df = pd.read_csv(v_file)
            for _, row in v_df.iterrows():
                conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?)", (row['거래처명'], row['은행'], row['계좌번호'], row['예금주']))
            conn.commit()
            st.rerun()
    st.dataframe(load_table("vendors"), use_container_width=True)

# (나머지 입금 입력, 상세내역 탭 코드는 이전과 동일하게 유지...)
# --- Tab 1: 입금 입력 ---
with tabs[0]:
    st.header("📝 입금 내역 입력")
    orders_df = load_table("orders")
    v_master = load_table("vendors")
    active_orders = orders_df[orders_df['마감여부'] == 0] if not orders_df.empty else pd.DataFrame()
    with st.form("pay_input_v24"):
        sel_oid = st.selectbox("🔗 발주번호 연동", ["없음"] + list(active_orders['발주번호']) if not active_orders.empty else ["없음"])
        c1, c2, c3 = st.columns(3)
        p_date = c1.date_input("입금일")
        p_vendor = c2.selectbox("거래처명 ", ["선택"] + list(v_master['거래처명']) if not v_master.empty else ["선택"])
        p_cat = c3.selectbox("유형 ", CATEGORIES)
        c4, c5, c6 = st.columns(3)
        p_dep, p_pre, p_curr = c4.number_input("💰 실입금액"), c5.number_input("🧧 선급금"), c6.selectbox("통화 ", CURRENCIES)
        p_note = st.text_input("메모 ")
        if st.form_submit_button("💾 입금 저장"):
            v_info = v_master[v_master['거래처명'] == p_vendor].iloc[0]
            rate = 1350.0 if p_curr == "USD" else (190.0 if p_curr == "CNY" else 1.0)
            conn.cursor().execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                     VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                                  (sel_oid if sel_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), p_cat, p_vendor, "품목", p_curr, p_dep, p_pre, p_note, (p_dep + p_pre)*rate, v_info['은행'], v_info['계좌번호'], v_info['예금주']))
            conn.commit()
            st.rerun()

# --- Tab 4: 상세내역 ---
with tabs[3]:
    st.header("🔍 상세내역 및 정산")
    p_all = load_table("payments")
    if not p_all.empty:
        st.dataframe(p_all.style.apply(style_closed_rows, axis=1), use_container_width=True)