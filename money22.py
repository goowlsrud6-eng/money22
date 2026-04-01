import streamlit as st
import pandas as pd
import sqlite3
import os
import shutil
import re
from datetime import datetime

# --- 1. 백업 및 DB 설정 ---
def run_backup():
    if not os.path.exists('backups'): os.makedirs('backups')
    db_file = 'finance_final_v70.db'
    backup_file = f"backups/backup_{datetime.now().strftime('%Y%m%d')}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

st.set_page_config(page_title="자금 관리 v70", layout="wide", page_icon="💰")
run_backup()

@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_final_v70.db', check_same_thread=False)
    c = conn.cursor()
    c.execute('CREATE TABLE IF NOT EXISTS vendors (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT, 기본유형 TEXT)')
    c.execute('CREATE TABLE IF NOT EXISTS orders (발주번호 TEXT PRIMARY KEY, 발주일 TEXT, 거래처명 TEXT, 상품명 TEXT, 유형 TEXT, 통화 TEXT, 발주총액 REAL, 마감여부 INTEGER DEFAULT 0)')
    c.execute('''CREATE TABLE IF NOT EXISTS payments (id INTEGER PRIMARY KEY AUTOINCREMENT, 발주번호 TEXT, 입금일 TEXT, 유형 TEXT, 거래처명 TEXT, 상품명 TEXT, 통화 TEXT,
                  실입금액 REAL, 선급금액 REAL, 메모 TEXT, 한화환산액 REAL, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT)''')
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비", "원단비", "기타"]

# --- 2. 세션 상태 관리 (알람 유지 및 업로더 리셋) ---
if 'alert_msg' not in st.session_state: st.session_state.alert_msg = []
if 'up_key' not in st.session_state: st.session_state.up_key = 0

def add_alert(text): st.session_state.alert_msg.append(text)
def clear_alerts(): st.session_state.alert_msg = []

# --- 3. 유틸리티 ---
def clean_name(name):
    if not name or pd.isna(name): return ""
    return re.sub(r'\s+', '', str(name))

def smart_date(date_str):
    try:
        ds = str(date_str).strip()
        if "월" in ds and "일" in ds: return datetime.strptime(f"2026 {ds}", "%Y %m월 %d일").strftime("%Y-%m-%d")
        return pd.to_datetime(ds).strftime("%Y-%m-%d")
    except: return datetime.now().strftime("%Y-%m-%d")

# --- 4. ERP 발주서 분석 로직 (미등록 차단 엔진) ---
def process_ecount_v70(file):
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
        
        # 이름 매칭 확인 (공백 제거 비교)
        v_master = pd.read_sql("SELECT 거래처명, 기본유형 FROM vendors", conn)
        v_master['clean'] = v_master['거래처명'].apply(clean_name)
        match = v_master[v_master['clean'] == clean_name(vendor_raw)]
        
        if match.empty:
            # 일치하지 않으면 등록 실패 처리 및 알람
            add_alert(f"❌ 등록 차단: '{vendor_raw}'는 거래처 관리 리스트에 없습니다. 먼저 등록해 주세요.")
            return False
        
        # 일치하면 정확한 이름과 유형 가져오기
        v_type = match.iloc[0]['기본유형']
        vendor_fixed = match.iloc[0]['거래처명']
        
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

        conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", (raw_oid, odate, vendor_fixed, prod_n, v_type, curr, total))
        conn.commit()
        return True
    except Exception as e:
        add_alert(f"❗ 시스템 오류({file.name}): {str(e)}")
        return False

# --- 5. UI 구성 ---
if st.session_state.alert_msg:
    for m in st.session_state.alert_msg:
        st.error(m)
    if st.button("🚨 알람 확인 및 모두 닫기"): 
        clear_alerts()
        st.rerun()

tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리"])

# [Tab 0] 입금 수기 입력
with tabs[0]:
    st.header("📝 입금 내역 수기 입력")
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    o_active = pd.read_sql("SELECT * FROM orders WHERE 마감여부=0", conn)
    with st.form("p_man_v70", clear_on_submit=True):
        c1, c2 = st.columns(2)
        p_oid = c1.selectbox("🔗 발주번호 연동", ["없음"] + list(o_active['발주번호']) if not o_active.empty else ["없음"])
        p_date = c2.date_input("입금일")
        c3, c4, c5 = st.columns(3)
        p_vn = c3.selectbox("거래처명", ["선택"] + list(v_data['거래처명']) if not v_data.empty else ["선택"])
        p_ct, p_pr = c4.selectbox("유형", CATEGORIES), c5.text_input("상품명")
        c6, c7, c8 = st.columns(3)
        p_dep, p_pre, p_cur = c6.number_input("실입금액"), c7.number_input("선급금액"), c8.selectbox("통화", ["한화", "USD", "CNY"])
        p_memo = st.text_input("메모")
        if st.form_submit_button("✅ 저장"):
            if p_vn != "선택":
                rate = 1350.0 if p_cur == "USD" else (190.0 if p_cur == "CNY" else 1.0)
                vi = v_data[v_data['거래처명']==p_vn].iloc[0]
                conn.execute("INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                             (p_oid if p_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), p_ct, p_vn, p_pr, p_cur, p_dep, p_pre, p_memo, (p_dep+p_pre)*rate, vi['은행'], vi['계좌번호'], vi['예금주']))
                conn.commit(); st.success("성공적으로 저장되었습니다."); st.rerun()

# [Tab 1] 입금 엑셀 업로드
with tabs[1]:
    st.header("📂 통합 입금 엑셀 업로드")
    p_up_key = f"pay_up_{st.session_state.up_key}"
    f_p = st.file_uploader("CSV 선택", type=['csv'], key=p_up_key)
    if f_p and st.button("🚀 입금 데이터 일괄 저장"):
        df_p = pd.read_csv(f_p).dropna(subset=['실입금액', '거래처'], how='all')
        o_df = pd.read_sql("SELECT * FROM orders", conn); v_df = pd.read_sql("SELECT * FROM vendors", conn)
        for _, r in df_p.iterrows():
            vn, oid = str(r['거래처']).strip(), str(r['발주번호']).strip()
            pd_s = smart_date(r['입금일'])
            if oid != "nan" and not o_df[o_df['발주번호'] == oid].empty:
                info = o_df[o_df['발주번호'] == oid].iloc[0]
                vn, pc, pp, cur = info['거래처명'], info['유형'], info['상품명'], info['통화']
            else: vn, pc, pp, cur = vn, str(r['유형']), str(r['상품명']), "한화"
            vi = v_df[v_df['거래처명'] == vn]
            bk, ac, hd = (vi.iloc[0]['은행'], vi.iloc[0]['계좌번호'], vi.iloc[0]['예금주']) if not vi.empty else ("","","")
            rt = 1350.0 if cur == "USD" else (190.0 if cur == "CNY" else 1.0)
            conn.execute("INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                         (oid if oid != "nan" else None, pd_s, pc, vn, pp, cur, float(r['실입금액']), float(r['선급금액']), r['송금사유'], (float(r['실입금액'])+float(r['선급금액']))*rt, bk, ac, hd))
        conn.commit(); st.success("✅ 입금 내역 일괄 저장 완료!"); st.session_state.up_key += 1; st.rerun()

# [Tab 2] 발주서 등록 (리셋 로직 포함)
with tabs[2]:
    st.header("📥 발주서 등록")
    o_up_key = f"order_up_{st.session_state.up_key}"
    of_list = st.file_uploader("xlsx 다중 선택", type=['xlsx'], accept_multiple_files=True, key=o_up_key)
    if of_list and st.button("🚀 발주서 일괄 등록 실행"):
        clear_alerts()
        scnt = 0
        for of in of_list:
            if process_ecount_v70(of): scnt += 1
        if scnt > 0: st.success(f"✅ {scnt}건의 발주서 등록 완료!")
        st.session_state.up_key += 1 # 파일 비우기
        st.rerun()
    st.divider()
    o_data = pd.read_sql("SELECT * FROM orders", conn)
    if not o_data.empty:
        st.data_editor(o_data.sort_values('발주일', ascending=False), hide_index=True, use_container_width=True)

# [Tab 3] 상세내역 및 정산
with tabs[3]:
    st.header("🔍 상세 내역 및 통합 정산")
    p_all = pd.read_sql("SELECT * FROM payments", conn); o_all = pd.read_sql("SELECT * FROM orders", conn)
    if not p_all.empty:
        st.subheader("📋 유형별 지출 요약")
        p_all['입금일_dt'] = pd.to_datetime(p_all['입금일'])
        sy = st.selectbox("연도 선택", sorted(p_all['입금일_dt'].dt.year.unique(), reverse=True))
        fil_p = p_all[p_all['입금일_dt'].dt.year == sy]
        cat_sum = fil_p.groupby('유형').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
        cat_sum['총합계'] = cat_sum['실입금액'] + cat_sum['선급금액']
        st.table(cat_sum.style.format('{:,.2f}'))
        st.divider()
        st.subheader("📑 상세 내역 수정")
        edited_p = st.data_editor(p_all.sort_values('입금일', ascending=False).drop(columns=['입금일_dt']), hide_index=True, use_container_width=True, disabled=["id"])
        if st.button("💾 입금내역 수정 저장"):
            for idx, r in edited_p.iterrows():
                conn.execute("UPDATE payments SET 발주번호=?, 입금일=?, 유형=?, 거래처명=?, 상품명=?, 실입금액=?, 선급금액=?, 메모=? WHERE id=?", (r['발주번호'], r['입금일'], r['유형'], r['거래처명'], r['상품명'], r['실입금액'], r['선급금액'], r['메모'], r['id']))
            conn.commit(); st.success("수정 완료"); st.rerun()

# [Tab 4] 거래처 관리 (이름 변경 시 강제 소급 동기화)
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    c_l, c_r = st.columns([1.2, 0.8])
    with c_l:
        with st.form("v_reg_v70", clear_on_submit=True):
            st.subheader("➕ 신규 거래처 등록")
            vn, vt = st.text_input("업체명"), st.selectbox("유형", CATEGORIES)
            vc1, vc2, vc3 = st.columns(3)
            vb, va, vh = vc1.text_input("은행"), vc2.text_input("계좌"), vc3.text_input("예금주")
            if st.form_submit_button("✅ 저장"):
                if vn: conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (vn,vb,va,vh,vt)); conn.commit(); st.rerun()
    with c_r:
        st.subheader("📂 엑셀 일괄 업로드")
        vf = st.file_uploader("거래처 CSV 업로드", type=['csv'])
        if vf and st.button("🚀 일괄 저장"):
            v_up = pd.read_csv(vf)
            for _, r in v_up.iterrows(): conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (r['거래처명'], r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
            conn.commit(); st.success("등록 완료"); st.rerun()
    
    st.divider()
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    if not v_data.empty:
        st.subheader("🏢 거래처 목록 (이름 수정 시 과거 발주/입금 내역 이름+유형 동시 소급 업데이트)")
        orig_v_data = v_data.copy()
        edited_v = st.data_editor(v_data, hide_index=True, use_container_width=True, key="v_edit_v70")
        
        if st.button("💾 정보 업데이트 및 전체 소급 동기화"):
            try:
                for i in range(len(orig_v_data)):
                    old_name = orig_v_data.iloc[i]['거래처명']
                    new_row = edited_v.iloc[i]
                    new_name = new_row['거래처명']
                    
                    if old_name != new_name:
                        conn.execute(f"DELETE FROM vendors WHERE 거래처명 = '{old_name}'")
                        conn.execute("INSERT INTO vendors VALUES (?,?,?,?,?)", (new_name, new_row['은행'], new_row['계좌번호'], new_row['예금주'], new_row['기본유형']))
                        # ★ 과거 모든 발주/입금 기록의 이름과 유형을 강제로 새 정보로 덮어씌움 ★
                        conn.execute("UPDATE orders SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_name, new_row['기본유형'], old_name))
                        conn.execute("UPDATE payments SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_name, new_row['기본유형'], old_name))
                    else:
                        conn.execute("UPDATE vendors SET 은행=?, 계좌번호=?, 예금주=?, 기본유형=? WHERE 거래처명=?", (new_row['은행'], new_row['계좌번호'], new_row['예금주'], new_row['기본유형'], new_name))
                        conn.execute("UPDATE orders SET 유형=? WHERE 거래처명=?", (new_row['기본유형'], new_name))
                        conn.execute("UPDATE payments SET 유형=? WHERE 거래처명=?", (new_row['기본유형'], new_name))
                
                conn.commit()
                st.success("✅ 모든 데이터가 최신 거래처 정보로 동기화되었습니다!")
                st.rerun()
            except Exception as e: st.error(f"동기화 실패: {e}")