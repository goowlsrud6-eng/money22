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
    db_file = 'finance_v32_stable.db' # 새 DB 파일로 충돌 방지
    backup_file = f"backups/backup_{today_str}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

# 2. 페이지 설정
st.set_page_config(page_title="자금 관리 v32", layout="wide", page_icon="💰")
run_backup()

# 3. DB 연결 및 테이블 초기화
def get_db_connection():
    conn = sqlite3.connect('finance_v32_stable.db', check_same_thread=False)
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

# --- 4. 이카운트 정밀 분석 함수 (B2 품목, F6 "중국" 인식) ---
def process_ecount_v32(file):
    try:
        df_raw = pd.read_excel(file, header=None)
        
        # [1] 전표번호 및 발주일 추출 (A1셀)
        raw_oid = str(df_raw.iloc[0, 0]).split(":")[-1].strip() if ":" in str(df_raw.iloc[0,0]) else str(df_raw.iloc[0, 0])
        oid = raw_oid
        odate = f"{oid[:4]}-{oid[4:6]}-{oid[6:8]}" if len(oid) >= 8 else datetime.now().strftime("%Y-%m-%d")
        
        # [2] 거래처 추출
        vendor = "미지정"
        for i in range(len(df_raw)):
            if "수신" in str(df_raw.iloc[i, 0]):
                vendor = str(df_raw.iloc[i, 0]).split(":")[-1].strip()
                break
        
        # [3] 통화 판별 (F6 셀 기준: 5행 5열) - "중국" 단어 포함 시 CNY
        f6_val = str(df_raw.iloc[5, 5]) if len(df_raw) > 5 else ""
        if "USD" in f6_val: currency = "USD"
        elif "중국" in f6_val or "CNY" in f6_val: currency = "CNY"
        else: currency = "한화"

        # [4] 품목명 추출 (국내는 B2(1,1), 외화는 C7(6,2) 기반)
        if currency == "한화":
            # 국내용: B2 셀 위치 (인덱스 1, 1) 확인
            raw_prod = str(df_raw.iloc[1, 1]) if len(df_raw) > 1 else "품목미상"
            product_display = raw_prod.split("[")[0].strip()
        else:
            # 외화용: C7 셀부터 시작
            raw_prod_list = df_raw.iloc[6:, 2].dropna().tolist() 
            if raw_prod_list:
                first_prod = str(raw_prod_list[0]).split("[")[0].strip()
                product_display = f"{first_prod} 외 {len(raw_prod_list)-1}건" if len(raw_prod_list) > 1 else first_prod
            else:
                product_display = "외화 품목미상"

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
        v_df = pd.read_sql("SELECT * FROM vendors", conn)
        matched_v = v_df[v_df['거래처명'] == vendor]
        v_type = matched_v.iloc[0]['기본유형'] if not matched_v.empty else "사입"

        # [7] 안전 저장: 리스트에서 사라지지 않게 REPLACE 구문 사용
        conn.cursor().execute('''INSERT OR REPLACE INTO orders 
                                 (발주번호, 발주일, 거래처명, 상품명, 유형, 통화, 발주총액, 마감여부) 
                                 VALUES (?, ?, ?, ?, ?, ?, ?, 0)''', 
                             (oid, odate, vendor, product_display, v_type, currency, total_amt))
        conn.commit()
        return True, oid
    except Exception as e:
        return False, str(e)

# --- 5. 스타일 및 데이터 로드 ---
def load_table(table):
    return pd.read_sql(f"SELECT * FROM {table}", conn)

def style_rows(row):
    if '마감여부' in row and row['마감여부'] == 1:
        return ['background-color: #f5f5f5; color: #a0a0a0; text-decoration: line-through'] * len(row)
    return [''] * len(row)

# --- 6. 메인 UI 구성 ---
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리"])

# --- Tab 3: 발주서 등록 (데이터 유실 방지 강화) ---
with tabs[2]:
    st.header("📥 발주서 자동 등록 (이카운트)")
    st.info("파일을 올리면 리스트에 누적 저장됩니다. (전표번호가 같으면 업데이트됩니다.)")
    o_file = st.file_uploader("이카운트 발주서(.xlsx) 업로드", type=['xlsx'], key="erp_up_v32")
    if o_file:
        success, res = process_ecount_v32(o_file)
        if success: st.success(f"✅ 등록 완료: {res}")
        else: st.error(f"❌ 오류: {res}")
    
    st.divider()
    st.subheader("📑 등록된 발주 리스트 (전체 목록)")
    # 리스트 로드 시 발주일 역순 정렬
    orders_display = load_table("orders")
    if not orders_display.empty:
        st.dataframe(orders_display.sort_values('발주일', ascending=False).style.apply(style_rows, axis=1), use_container_width=True)
    else:
        st.write("등록된 발주서가 없습니다.")

# --- Tab 5: 거래처 관리 ---
with tabs[4]:
    st.header("⚙️ 거래처 및 유형 관리")
    cv1, cv2 = st.columns(2)
    with cv1:
        with st.form("v_reg_v32", clear_on_submit=True):
            vn = st.text_input("업체명 (필수)")
            vt = st.selectbox("기본유형 (발주서 등록 시 자동매칭)", CATEGORIES)
            vb, va, vh = st.text_input("은행"), st.text_input("계좌번호"), st.text_input("예금주")
            if st.form_submit_button("거래처 저장"):
                if vn:
                    conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (vn,vb,va,vh,vt))
                    conn.commit()
                    st.rerun()
    with cv2:
        st.subheader("📂 거래처 일괄 등록")
        v_temp = pd.DataFrame(columns=["거래처명", "은행", "계좌번호", "예금주", "기본유형"])
        st.download_button("📥 양식 받기", v_temp.to_csv(index=False).encode('utf-8-sig'), "vendor_temp.csv")
        v_file = st.file_uploader("거래처 파일 업로드", type=['csv'])
        if v_file and st.button("🚀 일괄 저장"):
            v_df = pd.read_csv(v_file)
            for _, row in v_df.iterrows():
                conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", 
                                     (row['거래처명'], row['은행'], row['계좌번호'], row['예금주'], row['기본유형']))
            conn.commit()
            st.rerun()
    st.dataframe(load_table("vendors"), use_container_width=True)

# --- Tab 1: 입금 입력 ---
with tabs[0]:
    st.header("📝 입금 내역 수기 입력")
    orders_df = load_table("orders")
    v_master = load_table("vendors")
    active_orders = orders_df[orders_df['마감여부'] == 0] if not orders_df.empty else pd.DataFrame()
    with st.form("pay_input_v32", clear_on_submit=True):
        sel_oid = st.selectbox("🔗 발주번호 연동", ["없음"] + list(active_orders['발주번호']) if not active_orders.empty else ["없음"])
        c1, c2, c3 = st.columns(3)
        p_date = c1.date_input("입금일")
        p_vendor = c2.selectbox("거래처명 ", ["선택"] + list(v_master['거래처명']) if not v_master.empty else ["선택"])
        p_cat = c3.selectbox("유형 ", CATEGORIES)
        c4, c5, c6 = st.columns(3)
        p_dep, p_pre, p_curr = c4.number_input("💵 실입금액"), c5.number_input("🧧 선급금"), c6.selectbox("통화 ", ["한화", "USD", "CNY"])
        p_note = st.text_input("메모 ")
        if st.form_submit_button("입금 내역 저장"):
            if p_vendor != "선택":
                rate = 1350.0 if p_curr == "USD" else (190.0 if p_curr == "CNY" else 1.0)
                v_info = v_master[v_master['거래처명'] == p_vendor].iloc[0]
                conn.cursor().execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                         VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                                      (sel_oid if sel_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), p_cat, p_vendor, "품목", p_curr, p_dep, p_pre, p_note, (p_dep + p_pre)*rate, v_info['은행'], v_info['계좌번호'], v_info['예금주']))
                conn.commit()
                st.success("저장 완료!")
                st.rerun()

# --- Tab 4: 상세조회 ---
with tabs[3]:
    st.header("🔍 상세 내역 및 정산")
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
        st.dataframe(df_f.sort_values('입금일', ascending=False).style.apply(style_rows, axis=1), use_container_width=True)