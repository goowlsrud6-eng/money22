import streamlit as st
import pandas as pd
import sqlite3
import os
import shutil
from datetime import datetime

# --- 1. 데이터 안전장치 (자동 백업) ---
def run_backup():
    if not os.path.exists('backups'): os.makedirs('backups')
    db_file = 'finance_v45_final.db'
    backup_file = f"backups/backup_{datetime.now().strftime('%Y%m%d')}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

# 2. 페이지 설정 및 DB 연결
st.set_page_config(page_title="자금 관리 v45", layout="wide", page_icon="💰")
run_backup()

@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_v45_final.db', check_same_thread=False)
    c = conn.cursor()
    # 거래처 마스터
    c.execute('CREATE TABLE IF NOT EXISTS vendors (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT, 기본유형 TEXT)')
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
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비", "원단비", "기타"]
CURRENCIES = ["한화", "USD", "CNY"]

# --- 3. 유틸리티 함수 (서식 및 날짜) ---
def format_num(val):
    """천단위 쉼표 + 소수점 2자리 반올림 서식"""
    try:
        return f"{float(val):,.2f}"
    except:
        return "0.00"

def smart_date(date_str):
    """연도 없는 날짜(03월 11일)를 2026년으로 보정"""
    try:
        date_str = str(date_str).strip()
        if "월" in date_str and "일" in date_str:
            return datetime.strptime(f"2026 {date_str}", "%Y %m월 %d일").strftime("%Y-%m-%d")
        # 일반적인 날짜 형식 처리
        return pd.to_datetime(date_str).strftime("%Y-%m-%d")
    except:
        return datetime.now().strftime("%Y-%m-%d")

def style_row(row):
    """마감 여부에 따른 회색 취소선 스타일"""
    if row.get('마감여부') == 1:
        return ['background-color: #f5f5f5; color: #a0a0a0; text-decoration: line-through'] * len(row)
    return [''] * len(row)

# --- 4. 이카운트 분석 함수 ---
def process_ecount_v45(file):
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

        v_df = pd.read_sql("SELECT * FROM vendors", conn)
        v_type = v_df[v_df['거래처명'] == vendor].iloc[0]['기본유형'] if not v_df[v_df['거래처명'] == vendor].empty else "사입"
        
        conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", (raw_oid, odate, vendor, prod_name, v_type, curr, total))
        conn.commit(); return True, raw_oid
    except Exception as e: return False, str(e)

# --- 5. 메인 UI ---
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리"])

# --- Tab 1: 입금 수기 입력 ---
with tabs[0]:
    st.header("📝 입금 내역 수기 입력")
    v_m = pd.read_sql("SELECT * FROM vendors", conn); o_m = pd.read_sql("SELECT * FROM orders", conn)
    active_o = o_m[o_m['마감여부'] == 0] if not o_m.empty else pd.DataFrame()
    with st.form("p_manual_v45", clear_on_submit=True):
        sel_oid = st.selectbox("🔗 발주번호 연동", ["없음"] + list(active_o['발주번호']) if not active_o.empty else ["없음"])
        c1, c2, c3 = st.columns(3)
        p_d, p_v = c1.date_input("입금일"), c2.selectbox("거래처명", ["선택"] + list(v_m['거래처명']) if not v_m.empty else ["선택"])
        p_cat = c3.selectbox("유형", CATEGORIES)
        c4, c5, c6 = st.columns(3)
        p_dep, p_pre, p_cur = c4.number_input("실입금액"), c5.number_input("선급금액"), c6.selectbox("통화", CURRENCIES)
        p_note = st.text_input("메모")
        if st.form_submit_button("저장"):
            if p_v != "선택":
                rate = 1350.0 if p_cur == "USD" else (190.0 if p_cur == "CNY" else 1.0)
                vi = v_m[v_m['거래처명']==p_v].iloc[0]
                conn.cursor().execute("INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                                     (sel_oid if sel_oid != "없음" else None, p_d.strftime("%Y-%m-%d"), p_cat, p_v, "수기입력", p_cur, p_dep, p_pre, p_note, (p_dep+p_pre)*rate, vi['은행'], vi['계좌번호'], vi['예금주']))
                conn.commit(); st.success("저장 완료!"); st.rerun()

# --- Tab 2: 입금 엑셀 업로드 (날짜 보정 및 하이브리드) ---
with tabs[1]:
    st.header("📂 통합 입금 엑셀 업로드")
    tmp = pd.DataFrame(columns=["발주번호", "거래처", "유형", "상품명", "입금일", "실입금액", "선급금액", "송금사유"])
    st.download_button("📥 통합 양식 다운로드", tmp.to_csv(index=False).encode('utf-8-sig'), "payment_v45.csv")
    f = st.file_uploader("CSV 파일 선택", type=['csv'], key="p_csv_v45")
    if f and st.button("🚀 데이터 분석 및 일괄 저장"):
        try:
            df_p = pd.read_csv(f); o_df = pd.read_sql("SELECT * FROM orders", conn); v_df = pd.read_sql("SELECT * FROM vendors", conn)
            for _, r in df_p.iterrows():
                p_date_str = smart_date(r['입금일'])
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
        except Exception as e: st.error(f"에러 발생: {e}")

# --- Tab 3: 발주서 등록 ---
with tabs[2]:
    st.header("📥 발주서 등록")
    c1, c2 = st.columns(2)
    with c1:
        o_f = st.file_uploader("이카운트 엑셀(.xlsx)", type=['xlsx'], key="o_f_v45")
        if o_f: s, r = process_ecount_v45(o_f); st.success(f"등록 성공: {r}") if s else st.error(f"오류: {r}")
    with c2:
        with st.form("o_manual_v45", clear_on_submit=True):
            m_id = st.text_input("발주번호"); m_date = st.date_input("발주일")
            m_v = st.selectbox("거래처", ["선택"] + list(pd.read_sql("SELECT 거래처명 FROM vendors", conn)['거래처명']))
            if st.form_submit_button("수기 저장"):
                conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", (m_id, m_date.strftime("%Y-%m-%d"), m_v, "수기등록", "사입", "한화", 0))
                conn.commit(); st.rerun()
    st.divider()
    o_list = pd.read_sql("SELECT * FROM orders", conn)
    if not o_list.empty:
        o_disp = o_list.copy()
        o_disp['발주총액'] = o_disp['발주총액'].apply(format_num)
        st.dataframe(o_disp.sort_values('발주일', ascending=False).style.apply(style_row, axis=1), use_container_width=True)

# --- Tab 4: 상세내역 및 정산 (삭제/마감/서식) ---
with tabs[3]:
    st.header("🔍 상세 내역 및 정산")
    p_all = pd.read_sql("SELECT * FROM payments", conn); o_all = pd.read_sql("SELECT * FROM orders", conn)
    
    if not p_all.empty:
        # 요약 테이블 (정산 확인용)
        st.subheader("📊 발주번호별 정산 요약")
        sum_df = p_all.groupby('발주번호').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
        if not o_all.empty:
            sum_df = sum_df.merge(o_all[['발주번호', '발주총액', '거래처명']], on='발주번호', how='left')
            sum_df['미입금잔액'] = sum_df['발주총액'] - sum_df['실입금액']
            # 서식 적용
            sum_view = sum_df[['발주번호', '거래처명', '발주총액', '실입금액', '선급금액', '미입금잔액']].copy()
            for col in ['발주총액', '실입금액', '선급금액', '미입금잔액']:
                sum_view[col] = sum_view[col].apply(format_num)
            st.table(sum_view)

        st.subheader("📑 상세 내역")
        # 필터링
        f1, f2 = st.columns(2)
        v_sel = f1.multiselect("업체 필터", p_all['거래처명'].unique())
        df_f = p_all.copy()
        if v_sel: df_f = df_f[df_f['거래처명'].isin(v_sel)]
        if not o_all.empty:
            df_f = df_f.merge(o_all[['발주번호', '마감여부']], on='발주번호', how='left').fillna(0)
        
        # 숫자 서식 적용
        for col in ['실입금액', '선급금액', '한화환산액']:
            df_f[col] = df_f[col].apply(format_num)
        
        st.dataframe(df_f.sort_values('입금일', ascending=False).style.apply(style_row, axis=1), use_container_width=True)
        
        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            del_id = st.number_input("삭제할 ID (id열)", min_value=0)
            if st.button("🗑️ 선택 내역 삭제"):
                conn.execute(f"DELETE FROM payments WHERE id={del_id}"); conn.commit(); st.rerun()
        with c2:
            target_oid = st.selectbox("마감할 발주번호", o_all[o_all['마감여부']==0]['발주번호'].unique() if not o_all.empty else [])
            if st.button("🚩 해당 발주 최종 마감"):
                conn.execute(f"UPDATE orders SET 마감여부=1 WHERE 발주번호='{target_oid}'"); conn.commit(); st.rerun()
    else: st.info("내역이 없습니다.")

# --- Tab 5: 거래처 관리 ---
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    cv1, cv2 = st.columns(2)
    with cv1:
        with st.form("v_reg_v45", clear_on_submit=True):
            vn, vt = st.text_input("업체명"), st.selectbox("기본유형", CATEGORIES)
            vb, va, vh = st.text_input("은행"), st.text_input("계좌"), st.text_input("예금주")
            if st.form_submit_button("저장"):
                if vn: conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (vn,vb,va,vh,vt)); conn.commit(); st.rerun()
    with cv2:
        v_f = st.file_uploader("거래처 CSV", type=['csv'], key="v_csv_v45")
        if v_f and st.button("🚀 일괄 저장"):
            v_df = pd.read_csv(v_f)
            for _, r in v_df.iterrows():
                conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (r[0],r[1],r[2],r[3],r[4]))
            conn.commit(); st.rerun()
    st.dataframe(pd.read_sql("SELECT * FROM vendors", conn), use_container_width=True)