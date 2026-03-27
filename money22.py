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
    db_file = 'finance_v28_final.db'
    backup_file = f"backups/backup_{today_str}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

# 2. 페이지 설정
st.set_page_config(page_title="자금 관리 v28", layout="wide", page_icon="💰")
run_backup()

# 3. DB 연결
def get_db_connection():
    conn = sqlite3.connect('finance_v28_final.db', check_same_thread=False)
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

def style_closed_rows(row):
    if '마감여부' in row and row['마감여부'] == 1:
        return ['background-color: #f0f0f0; color: #a0a0a0; text-decoration: line-through'] * len(row)
    return [''] * len(row)

# --- 4. 이카운트 엑셀 분석 함수 (데이터만 반환, 저장은 버튼으로) ---
def parse_ecount_data(file):
    try:
        df_raw = pd.read_excel(file, header=None)
        raw_oid = str(df_raw.iloc[0, 0]).split(":")[-1].strip() if ":" in str(df_raw.iloc[0,0]) else ""
        odate_val = f"{raw_oid[:4]}-{raw_oid[4:6]}-{raw_oid[6:8]}" if len(raw_oid) >= 8 else datetime.now().strftime("%Y-%m-%d")
        
        vendor_val = ""
        for i in range(len(df_raw)):
            if "수신" in str(df_raw.iloc[i, 0]):
                vendor_val = str(df_raw.iloc[i, 0]).split(":")[-1].strip()
                break
        
        amt = 0.0
        # 한화(A5) 또는 외화(F열 마지막) 체크
        last_f = df_raw.iloc[:, 5].last_valid_index()
        val_f = df_raw.iloc[last_f, 5] if last_f is not None else None
        if isinstance(val_f, (int, float)) and val_f > 100:
            amt = float(val_f)
        else:
            a5_val = str(df_raw.iloc[4, 0])
            if "금액" in a5_val:
                amt = float(a5_val.split(":")[-1].replace(",", "").strip())
        
        return {"oid": raw_oid, "odate": odate_val, "vendor": vendor_val, "total": amt}
    except:
        return None

# --- 5. 메인 UI ---
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록(ERP/수기)", "🔍 상세조회 및 필터", "⚙️ 거래처 관리"])

# --- Tab 3: 발주서 등록 (ERP 업로드와 수기 등록을 동시에 배치) ---
with tabs[2]:
    st.header("📥 발주서 등록")
    v_data = load_table("vendors")
    
    # [A] 이카운트 엑셀 업로드 구역
    st.subheader("📄 1. 이카운트 엑셀 자동 분석")
    o_file = st.file_uploader("이카운트 발주서(.xlsx)를 올리면 아래 수기 폼이 자동으로 채워집니다.", type=['xlsx'])
    parsed = parse_ecount_data(o_file) if o_file else None
    if parsed:
        st.success(f"✅ 분석 완료: 전표번호 {parsed['oid']} 외 정보를 가져왔습니다. 아래 폼에서 확인 후 저장하세요.")

    st.divider()

    # [B] 발주서 수기 등록 폼 (분석된 데이터가 있으면 기본값으로 세팅)
    st.subheader("✍️ 2. 발주서 상세 입력 및 저장")
    with st.form("order_reg_final", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns(4)
        f_oid = c1.text_input("발주번호(전표번호)", value=parsed['oid'] if parsed else "")
        # 날짜 처리
        default_date = datetime.strptime(parsed['odate'], "%Y-%m-%d") if parsed else datetime.now()
        f_date = c2.date_input("발주일", value=default_date)
        f_cat = c3.selectbox("유형", CATEGORIES)
        
        # 거래처 매칭
        v_list = ["선택"] + list(v_data['거래처명']) if not v_data.empty else ["선택"]
        v_idx = v_list.index(parsed['vendor']) if parsed and parsed['vendor'] in v_list else 0
        f_vendor = c4.selectbox("거래처", options=v_list, index=v_idx)
        
        c5, c6, c7 = st.columns([2, 1, 1])
        f_prod = c5.text_input("상품명", value="품목 외 n건" if parsed else "")
        f_total = c6.number_input("발주총액", min_value=0.0, value=float(parsed['total']) if parsed else 0.0)
        f_curr = c7.selectbox("통화", CURRENCIES)
        
        if st.form_submit_button("🚀 발주서 최종 저장"):
            if not f_oid or f_vendor == "선택":
                st.error("발주번호와 거래처는 필수 입력 사항입니다.")
            else:
                conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", 
                                     (f_oid, f_date.strftime("%Y-%m-%d"), f_vendor, f_prod, f_cat, f_curr, f_total))
                conn.commit()
                st.success(f"발주번호 {f_oid} 등록 완료!")
                st.rerun()

    st.subheader("📑 현재 등록된 발주 리스트")
    st.dataframe(load_table("orders").sort_values('발주일', ascending=False).style.apply(style_closed_rows, axis=1), use_container_width=True)

# --- Tab 1: 입금 입력 ---
with tabs[0]:
    st.header("📝 입금 내역 수기 입력")
    orders = load_table("orders")
    v_master = load_table("vendors")
    active_orders = orders[orders['마감여부'] == 0] if not orders.empty else pd.DataFrame()
    
    with st.form("pay_input_v28", clear_on_submit=True):
        sel_oid = st.selectbox("🔗 발주번호 연동", ["없음"] + list(active_orders['발주번호']) if not active_orders.empty else ["없음"])
        c1, c2, c3 = st.columns(3)
        p_date = c1.date_input("입금일")
        p_vendor = c2.selectbox("거래처명 ", ["선택"] + list(v_master['거래처명']) if not v_master.empty else ["선택"])
        p_cat = c3.selectbox("유형 ", CATEGORIES)
        c4, c5, c6 = st.columns(3)
        p_dep = c4.number_input("💰 실입금액 (통장금액)", min_value=0.0)
        p_pre = c5.number_input("🧧 선급금 (발생/사용)", value=0.0)
        p_curr = c6.selectbox("통화 ", CURRENCIES)
        p_note = st.text_input("메모(송금사유) ")
        
        if st.form_submit_button("입금 내역 저장"):
            if p_vendor == "선택": st.error("거래처를 선택하세요.")
            else:
                rate = 1350.0 if p_curr == "USD" else (190.0 if p_curr == "CNY" else 1.0)
                v_info = v_master[v_master['거래처명'] == p_vendor].iloc[0]
                conn.cursor().execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                         VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                                      (sel_oid if sel_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), p_cat, p_vendor, "품목", p_curr, p_dep, p_pre, p_note, (p_dep + p_pre)*rate, v_info['은행'], v_info['계좌번호'], v_info['예금주']))
                conn.commit()
                st.success("저장되었습니다.")
                st.rerun()

# --- Tab 5: 거래처 관리 ---
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    col_v1, col_v2 = st.columns(2)
    with col_v1:
        with st.form("v_reg_v28", clear_on_submit=True):
            st.subheader("➕ 개별 등록")
            vn, vb, va, vh = st.text_input("업체명"), st.text_input("은행"), st.text_input("계좌번호"), st.text_input("예금주")
            if st.form_submit_button("거래처 저장"):
                if vn:
                    conn.cursor().execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?)", (vn,vb,va,vh))
                    conn.commit()
                    st.rerun()
    with col_v2:
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

# (Tab 2: 입금 엑셀, Tab 4: 상세조회는 v27과 동일하게 유지)