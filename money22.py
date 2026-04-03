import streamlit as st
import pandas as pd
import sqlite3
import os
import shutil
import re
from datetime import datetime

# --- 1. 백업 및 데이터베이스 설정 ---
def run_backup():
    if not os.path.exists('backups'): os.makedirs('backups')
    db_file = 'finance_final_v79.db'
    backup_file = f"backups/backup_{datetime.now().strftime('%Y%m%d')}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

st.set_page_config(page_title="자금 관리 v79", layout="wide", page_icon="💰")
run_backup()

@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_final_v79.db', check_same_thread=False)
    c = conn.cursor()
    c.execute('CREATE TABLE IF NOT EXISTS vendors (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT, 기본유형 TEXT)')
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (발주번호 TEXT PRIMARY KEY, 발주일 TEXT, 발주차수 TEXT, 거래처명 TEXT, 
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

# --- 2. 세션 상태 관리 (상태 유지 및 파일 비우기) ---
if 'alert_msg' not in st.session_state: st.session_state.alert_msg = []
if 'order_up_key' not in st.session_state: st.session_state.order_up_key = 0
if 'pay_up_key' not in st.session_state: st.session_state.pay_up_key = 1000

def add_alert(text): 
    if text not in st.session_state.alert_msg:
        st.session_state.alert_msg.append(text)

def clear_alerts(): 
    st.session_state.alert_msg = []

# --- 3. 유틸리티 함수 ---
def clean_name(name):
    if not name or pd.isna(name): return ""
    return re.sub(r'\s+', '', str(name))

def smart_date(date_str):
    try:
        ds = str(date_str).strip()
        if "월" in ds and "일" in ds: return datetime.strptime(f"2026 {ds}", "%Y %m월 %d일").strftime("%Y-%m-%d")
        return pd.to_datetime(ds).strftime("%Y-%m-%d")
    except: return datetime.now().strftime("%Y-%m-%d")

# --- 4. ERP 발주서 분석 로직 ---
def process_ecount_v79(file):
    try:
        df = pd.read_excel(file, header=None)
        raw_oid = str(df.iloc[1, 0]).split(":")[-1].strip() if ":" in str(df.iloc[1,0]) else str(df.iloc[1, 0])
        clean_oid = raw_oid.replace("-", "")
        odate = f"{clean_oid[:4]}-{clean_oid[4:6]}-{clean_oid[6:8]}" if len(clean_oid) >= 8 else datetime.now().strftime("%Y-%m-%d")
        
        vendor_raw = "미지정"
        for i in range(len(df)):
            if "수신" in str(df.iloc[i, 0]):
                vendor_raw = str(df.iloc[i, 0]).split(":")[-1].strip()
                break
        
        v_master = pd.read_sql("SELECT 거래처명, 기본유형 FROM vendors", conn)
        v_master['clean_key'] = v_master['거래처명'].apply(clean_name)
        match = v_master[v_master['clean_key'] == clean_name(vendor_raw)]
        
        if match.empty:
            add_alert(f"❌ '{vendor_raw}'은(는) 미등록 업체입니다. [거래처 관리]에서 등록 후 다시 시도하세요.")
            return False
        
        v_type, vendor_fixed = match.iloc[0]['기본유형'], match.iloc[0]['거래처명']
        f6 = str(df.iloc[5, 5]) if len(df) > 5 else ""
        curr = "USD" if "USD" in f6 else ("CNY" if any(x in f6 for x in ["중국", "CNY"]) else "한화")
        p_col = 1 if curr == "한화" else 2
        prods = df.iloc[6:, p_col].dropna().astype(str).tolist()
        prod_n = (prods[0].split("[")[0].strip() + (f" 외 {len(prods)-1}건" if len(prods)>1 else "")) if prods else "품목미상"
        
        total = 0.0
        if curr != "한화":
            l_idx = df.iloc[:, 5].last_valid_index()
            total = float(df.iloc[l_idx, 5]) if l_idx else 0.0
        else:
            a5 = str(df.iloc[4, 0]); total = float(a5.split(":")[-1].replace(",", "").strip()) if "금액" in a5 else 0.0

        conn.execute("INSERT OR REPLACE INTO orders (발주번호, 발주일, 발주차수, 거래처명, 상품명, 유형, 통화, 발주총액, 마감여부) VALUES (?,?,?,?,?,?,?,?,0)", 
                     (raw_oid, odate, "", vendor_fixed, prod_n, v_type, curr, total))
        conn.commit()
        return True
    except Exception:
        return False

# --- 5. UI 구성 ---
# 상단 알람 (사용자가 닫기 전까지 유지)
if st.session_state.alert_msg:
    for m in st.session_state.alert_msg:
        st.error(m)
    if st.button("알람 닫기"): 
        clear_alerts()
        st.rerun()

tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리"])

# [Tab 0] 입금 수기
with tabs[0]:
    st.header("📝 입금 내역 수기 입력")
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    o_active = pd.read_sql("SELECT * FROM orders WHERE 마감여부=0", conn)
    with st.form("p_manual_v79", clear_on_submit=True):
        c1, c2 = st.columns(2)
        p_oid = c1.selectbox("🔗 발주번호 연동", ["없음"] + list(o_active['발주번호']) if not o_active.empty else ["없음"])
        p_date = c2.date_input("입금일")
        c3, c4, c5 = st.columns(3)
        p_vn = c3.selectbox("거래처명", ["선택"] + list(v_data['거래처명']) if not v_data.empty else ["선택"])
        p_ct, p_pr = c4.selectbox("유형", CATEGORIES), c5.text_input("상품명")
        c6, c7, c8 = st.columns(3)
        p_dep, p_pre, p_cur = c6.number_input("실입금액", format="%.2f"), c7.number_input("선급금액", format="%.2f"), c8.selectbox("통화", ["한화", "USD", "CNY"])
        p_memo = st.text_input("메모")
        if st.form_submit_button("✅ 저장"):
            if p_vn != "선택":
                rate = 1350.0 if p_cur == "USD" else (190.0 if p_cur == "CNY" else 1.0)
                vi = v_data[v_data['거래처명']==p_vn].iloc[0]
                conn.execute("INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                             (p_oid if p_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), p_ct, p_vn, p_pr, p_cur, p_dep, p_pre, p_memo, (p_dep+p_pre)*rate, vi['은행'], vi['계좌번호'], vi['예금주']))
                conn.commit(); st.success("저장 완료!"); st.rerun()

# [Tab 1] 입금 엑셀 업로드
with tabs[1]:
    st.header("📂 통합 입금 엑셀 업로드")
    f_p = st.file_uploader("입금 CSV 선택", type=['csv'], key=f"pay_up_{st.session_state.pay_up_key}")
    if f_p and st.button("🚀 입금 데이터 일괄 저장"):
        try:
            df_p = pd.read_csv(f_p).dropna(subset=['실입금액', '거래처'], how='all')
            v_l = pd.read_sql("SELECT * FROM vendors", conn); o_l = pd.read_sql("SELECT * FROM orders", conn)
            for _, r in df_p.iterrows():
                vn, oid, pd_s = str(r['거래처']).strip(), str(r['발주번호']).strip(), smart_date(r['입금일'])
                if oid != "nan" and not o_l[o_l['발주번호'] == oid].empty:
                    info = o_l[o_l['발주번호'] == oid].iloc[0]
                    vn, pc, pp, cur = info['거래처명'], info['유형'], info['상품명'], info['통화']
                else: vn, pc, pp, cur = vn, str(r['유형']), str(r['상품명']), "한화"
                vi = v_l[v_l['거래처명'] == vn]
                bk, ac, hd = (vi.iloc[0]['은행'], vi.iloc[0]['계좌번호'], vi.iloc[0]['예금주']) if not vi.empty else ("","","")
                rt = 1350.0 if cur == "USD" else (190.0 if cur == "CNY" else 1.0)
                conn.execute("INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                             (oid if oid != "nan" else None, pd_s, pc, vn, pp, cur, float(r['실입금액']), float(r['선급금액']), r['송금사유'], (float(r['실입금액'])+float(r['선급금액']))*rt, bk, ac, hd))
            conn.commit(); st.success(f"✅ {len(df_p)}건 저장 완료!"); st.session_state.pay_up_key += 1; st.rerun()
        except Exception as e: st.error(f"오류: {e}")

# [Tab 2] 발주서 등록 및 관리
with tabs[2]:
    st.header("📥 발주서 등록 및 관리")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("⚡ 엑셀 일괄 등록")
        of_list = st.file_uploader("발주서(xlsx) 선택", type=['xlsx'], accept_multiple_files=True, key=f"ord_up_{st.session_state.order_up_key}")
        if of_list and st.button("🚀 모든 발주서 일괄 등록 실행"):
            clear_alerts()
            all_ok = True
            scnt = 0
            for of in of_list:
                if process_ecount_v79(of): scnt += 1
                else: all_ok = False
            
            if scnt > 0: st.success(f"✅ {scnt}건 등록 성공!")
            if all_ok: 
                st.session_state.order_up_key += 1
                st.rerun()
    with c2:
        st.subheader("✍️ 수기 등록")
        v_list_m = pd.read_sql("SELECT 거래처명 FROM vendors", conn)
        with st.form("o_manual_v79", clear_on_submit=True):
            c_m1, c_m2 = st.columns(2)
            mi, mt_order = c_m1.text_input("발주번호"), c_m2.text_input("발주차수")
            md, mv = st.date_input("발주일"), st.selectbox("거래처 선택", ["선택"] + list(v_list_m['거래처명']))
            mp = st.text_input("상품명")
            c_m3, c_m4 = st.columns(2)
            mt, m_cur = c_m3.number_input("발주금액", format="%.2f"), c_m4.selectbox("통화", ["한화", "USD", "CNY"])
            if st.form_submit_button("✅ 저장"):
                if mi and mv != "선택":
                    vt = pd.read_sql(f"SELECT 기본유형 FROM vendors WHERE 거래처명='{mv}'", conn).iloc[0]['기본유형']
                    conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,0)", (mi, md.strftime("%Y-%m-%d"), mt_order, mv, mp, vt, m_cur, mt))
                    conn.commit(); st.success("수기 발주 저장 완료!"); st.rerun()
    st.divider()
    o_data = pd.read_sql("SELECT * FROM orders", conn)
    if not o_data.empty:
        st.subheader("📄 등록된 발주서 현황")
        o_disp = o_data[['발주번호', '발주차수', '거래처명', '상품명', '유형', '통화', '발주총액', '마감여부', '발주일']]
        ev_o = st.data_editor(o_disp.sort_values('발주일', ascending=False), hide_index=True, use_container_width=True, disabled=["발주번호"])
        if st.button("💾 발주 정보 소급 저장"):
            for idx, r in ev_o.iterrows():
                conn.execute("UPDATE orders SET 발주일=?, 발주차수=?, 거래처명=?, 상품명=?, 유형=?, 통화=?, 발주총액=?, 마감여부=? WHERE 발주번호=?", (r['발주일'], r['발주차수'], r['거래처명'], r['상품명'], r['유형'], r['통화'], r['발주총액'], r['마감여부'], r['발주번호']))
                conn.execute("UPDATE payments SET 거래처명=?, 유형=? WHERE 발주번호=?", (r['거래처명'], r['유형'], r['발주번호']))
            conn.commit(); st.success("✅ 동기화 완료!"); st.rerun()

# [Tab 3] 상세내역 및 정산
with tabs[3]:
    st.header("🔍 상세 내역 및 통합 정산")
    p_all = pd.read_sql("SELECT * FROM payments", conn); o_all = pd.read_sql("SELECT * FROM orders", conn)
    if not p_all.empty:
        p_all['입금일_dt'] = pd.to_datetime(p_all['입금일'])
        sy = st.selectbox("연도 선택", sorted(p_all['입금일_dt'].dt.year.unique(), reverse=True))
        cat_sum = p_all[p_all['입금일_dt'].dt.year == sy].groupby('유형').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
        st.table(cat_sum.style.format('{:,.2f}'))
        
        st.divider(); ed_p = st.data_editor(p_all.sort_values('입금일', ascending=False).drop(columns=['입금일_dt']), hide_index=True, use_container_width=True, disabled=["id"])
        if st.button("💾 개별 수정 저장"):
            for idx, r in ed_p.iterrows():
                conn.execute("UPDATE payments SET 발주번호=?, 입금일=?, 유형=?, 거래처명=?, 상품명=?, 실입금액=?, 선급금액=?, 메모=? WHERE id=?", (r['발주번호'], r['입금일'], r['유형'], r['거래처명'], r['상품명'], r['실입금액'], r['선급금액'], r['메모'], r['id']))
            conn.commit(); st.success("저장 완료"); st.rerun()
        
        did = st.number_input("삭제 ID", min_value=0, step=1)
        if st.button("🗑️ 행 삭제"): conn.execute(f"DELETE FROM payments WHERE id={did}"); conn.commit(); st.rerun()

# [Tab 4] 거래처 관리
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    cl, cr = st.columns([1.2, 0.8])
    with cl:
        with st.form("v_reg_v79", clear_on_submit=True):
            vn, vt = st.text_input("거래처명"), st.selectbox("유형", CATEGORIES)
            vc1, vc2, vc3 = st.columns(3)
            vb, va, vh = vc1.text_input("은행"), vc2.text_input("계좌"), vc3.text_input("예금주")
            if st.form_submit_button("✅ 저장"):
                if vn: conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (vn,vb,va,vh,vt)); conn.commit(); st.rerun()
    with cr:
        vf = st.file_uploader("거래처 CSV 업로드", type=['csv'])
        if vf and st.button("🚀 업로드"):
            v_up = pd.read_csv(vf)
            for _, r in v_up.iterrows(): conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (r['거래처명'], r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
            conn.commit(); st.success("등록 완료"); st.rerun()
    st.divider()
    v_data_m = pd.read_sql("SELECT * FROM vendors", conn)
    if not v_data_m.empty:
        orig_names = v_data_m['거래처명'].tolist()
        ev_v = st.data_editor(v_data_m, hide_index=True, use_container_width=True)
        if st.button("💾 전체 데이터 동기화 저장"):
            for idx, r in ev_v.iterrows():
                old, new = orig_names[idx], r['거래처명']
                if old != new:
                    conn.execute(f"DELETE FROM vendors WHERE 거래처명='{old}'")
                    conn.execute("INSERT INTO vendors VALUES (?,?,?,?,?)", (new, r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
                    conn.execute("UPDATE orders SET 거래처명=?, 유형=? WHERE 거래처명=?", (new, r['기본유형'], old))
                    conn.execute("UPDATE payments SET 거래처명=?, 유형=? WHERE 거래처명=?", (new, r['기본유형'], old))
                else:
                    conn.execute("UPDATE vendors SET 은행=?, 계좌번호=?, 예금주=?, 기본유형=? WHERE 거래처명=?", (r['은행'], r['계좌번호'], r['예금주'], r['기본유형'], new))
                    conn.execute("UPDATE orders SET 유형=? WHERE 거래처명=?", (r['기본유형'], new))
                    conn.execute("UPDATE payments SET 유형=? WHERE 거래처명=?", (r['기본유형'], new))
            conn.commit(); st.success("✅ 동기화 완료!"); st.rerun()