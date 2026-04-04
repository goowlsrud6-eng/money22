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
    db_file = 'finance_final_v90.db'
    backup_file = f"backups/backup_{datetime.now().strftime('%Y%m%d')}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

st.set_page_config(page_title="자금 관리 v90", layout="wide", page_icon="💰")
run_backup()

@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_final_v90.db', check_same_thread=False)
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

# --- 2. 세션 상태 관리 ---
if 'order_up_key' not in st.session_state: st.session_state.order_up_key = 0
if 'pay_up_key' not in st.session_state: st.session_state.pay_up_key = 1000

# --- 3. 유틸리티 함수 ---
def to_float(val):
    try:
        if val is None or pd.isna(val) or str(val).strip() == "": return 0.0
        return float(str(val).replace(',', ''))
    except: return 0.0

def to_str(val):
    if val is None or pd.isna(val): return ""
    s = str(val).strip()
    if s.lower() in ["nan", "none", ""]: return ""
    return s

def smart_date(date_str):
    try:
        ds = to_str(date_str)
        if not ds: return datetime.now().strftime("%Y-%m-%d")
        if "월" in ds and "일" in ds:
            return datetime.strptime(f"2026 {ds}", "%Y %m월 %d일").strftime("%Y-%m-%d")
        return pd.to_datetime(ds).strftime("%Y-%m-%d")
    except: return datetime.now().strftime("%Y-%m-%d")

# --- 4. ERP 발주서 분석 로직 ---
def process_ecount_v90(file):
    try:
        df = pd.read_excel(file, header=None)
        raw_oid = str(df.iloc[1, 0]).split(":")[-1].strip() if ":" in str(df.iloc[1,0]) else str(df.iloc[1, 0])
        clean_oid = raw_oid.replace("-", "")
        odate = f"{clean_oid[:4]}-{clean_oid[4:6]}-{clean_oid[6:8]}" if len(clean_oid) >= 8 else datetime.now().strftime("%Y-%m-%d")
        
        vendor_raw = ""
        for i in range(len(df)):
            if "수신" in str(df.iloc[i, 0]):
                vendor_raw = str(df.iloc[i, 0]).split(":")[-1].strip()
                break
        
        v_master = pd.read_sql("SELECT 거래처명, 기본유형 FROM vendors", conn)
        v_master['clean_key'] = v_master['거래처명'].apply(lambda x: re.sub(r'\s+', '', str(x)))
        target_key = re.sub(r'\s+', '', vendor_raw)
        
        match = v_master[v_master['clean_key'] == target_key]
        if match.empty: return False, f"⚠️ '{vendor_raw}'은(는) 미등록 업체입니다."
        
        v_type, vendor_fixed = match.iloc[0]['기본유형'], match.iloc[0]['거래처명']
        f6 = str(df.iloc[5, 5]) if len(df) > 5 else ""
        curr = "USD" if "USD" in f6 else ("CNY" if any(x in f6 for x in ["중국", "CNY"]) else "한화")
        p_col = 1 if curr == "한화" else 2
        prods = df.iloc[6:, p_col].dropna().astype(str).tolist()
        prod_n = (prods[0].split("[")[0].strip() + (f" 외 {len(prods)-1}건" if len(prods)>1 else "")) if prods else "품목미상"
        
        last_val_idx = df.iloc[:, 5].last_valid_index()
        total = to_float(df.iloc[last_val_idx, 5]) if curr != "한화" and last_val_idx is not None else to_float(str(df.iloc[4, 0]).split(":")[-1])

        conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,0)", (raw_oid, odate, "", vendor_fixed, prod_n, v_type, curr, total))
        conn.commit(); return True, None
    except: return False, "❗ 발주서 분석 오류"

# --- 5. 메인 UI ---
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리"])

# [Tab 0] 수기 입력
with tabs[0]:
    st.header("📝 입금 내역 수기 입력")
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    o_active = pd.read_sql("SELECT * FROM orders WHERE 마감여부=0", conn)
    with st.form("p_man_v90", clear_on_submit=True):
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
            rate = 1350.0 if p_cur == "USD" else (190.0 if p_cur == "CNY" else 1.0)
            vi = v_data[v_data['거래처명']==p_vn].iloc[0] if p_vn != "선택" else {"은행":"","계좌번호":"","예금주":""}
            conn.execute("INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                         (to_str(p_oid) if p_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), p_ct, p_vn if p_vn != "선택" else "", p_pr, p_cur, p_dep, p_pre, p_memo, (p_dep+p_pre)*rate, vi['은행'], vi['계좌번호'], vi['예금주']))
            conn.commit(); st.success("저장 완료!"); st.rerun()

# [Tab 1] 입금 엑셀 업로드
with tabs[1]:
    st.header("📂 통합 입금 엑셀 업로드")
    template = pd.DataFrame(columns=["발주번호", "거래처", "유형", "상품명", "입금일", "실입금액", "선급금액", "송금사유"])
    st.download_button("📥 샘플 양식 다운로드", template.to_csv(index=False).encode('utf-8-sig'), "payment_template.csv")
    
    f_p = st.file_uploader("CSV 선택", type=['csv'], key=f"pay_up_{st.session_state.pay_up_key}")
    if f_p and st.button("🚀 데이터 일괄 저장"):
        try:
            df_p = pd.read_csv(f_p)
            v_l = pd.read_sql("SELECT * FROM vendors", conn); o_l = pd.read_sql("SELECT * FROM orders", conn)
            for _, r in df_p.iterrows():
                oid, vn_raw = to_str(r.get('발주번호')), to_str(r.get('거래처'))
                if not vn_raw and not oid: continue
                pd_s = smart_date(r.get('입금일'))
                if oid and not o_l[o_l['발주번호'] == oid].empty:
                    info = o_l[o_l['발주번호'] == oid].iloc[0]
                    vn, pc, pp, cur = info['거래처명'], info['유형'], info['상품명'], info['통화']
                else: vn, pc, pp, cur = vn_raw, to_str(r.get('유형')) or "사입", to_str(r.get('상품명')), "한화"
                vi = v_l[v_l['거래처명'] == vn] if vn else pd.DataFrame()
                bk, ac, hd = (vi.iloc[0]['은행'], vi.iloc[0]['계좌번호'], vi.iloc[0]['예금주']) if not vi.empty else ("","","")
                dep, pre = to_float(r.get('실입금액')), to_float(r.get('선급금액'))
                rt = 1350.0 if cur == "USD" else (190.0 if cur == "CNY" else 1.0)
                conn.execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                             (oid if oid else None, pd_s, pc, vn, pp, cur, dep, pre, to_str(r.get('송금사유')), (dep+pre)*rt, bk, ac, hd))
            conn.commit(); st.success("✅ 저장 완료!"); st.session_state.pay_up_key += 1; st.rerun()
        except Exception as e: st.error(f"오류: {e}")

# [Tab 2] 발주서 등록 및 연동 관리
with tabs[2]:
    st.header("📥 발주서 등록 및 관리")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("⚡ 엑셀 일괄 등록")
        of_list = st.file_uploader("발주서(xlsx) 선택", type=['xlsx'], accept_multiple_files=True, key=f"ord_{st.session_state.order_up_key}")
        if of_list and st.button("🚀 모든 발주서 일괄 등록"):
            scnt, errs = 0, []
            for of in of_list:
                res, msg = process_ecount_v90(of)
                if res: scnt += 1
                else: errs.append(msg)
            for em in errs: st.warning(em)
            if scnt > 0: st.success(f"✅ {scnt}건 성공!"); st.session_state.order_up_key += 1; st.rerun()
    with c2:
        st.subheader("✍️ 수기 등록")
        with st.form("o_manual"):
            mi, mt_order = st.text_input("발주번호"), st.text_input("발주차수")
            md, mv = st.date_input("발주일"), st.selectbox("거래처", ["선택"] + list(pd.read_sql("SELECT 거래처명 FROM vendors", conn)['거래처명']))
            mp, mt, m_cur = st.text_input("상품명"), st.number_input("금액"), st.selectbox("통화", ["한화", "USD", "CNY"])
            if st.form_submit_button("✅ 저장"):
                if mi and mv != "선택":
                    vt = pd.read_sql(f"SELECT 기본유형 FROM vendors WHERE 거래처명='{mv}'", conn).iloc[0]['기본유형']
                    conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,0)", (mi, md.strftime("%Y-%m-%d"), mt_order, mv, mp, vt, m_cur, mt))
                    conn.commit(); st.rerun()
    st.divider()
    o_data = pd.read_sql("SELECT * FROM orders", conn)
    if not o_data.empty:
        st.subheader("📄 발주 리스트 (수정 및 동기화)")
        ev_o = st.data_editor(o_data[['발주번호', '발주차수', '거래처명', '상품명', '유형', '통화', '발주총액', '마감여부', '발주일']].sort_values('발주일', ascending=False), hide_index=True, use_container_width=True, disabled=["발주번호"])
        if st.button("💾 정보 업데이트 및 모든 상세내역 소급 적용"):
            for _, r in ev_o.iterrows():
                conn.execute("UPDATE orders SET 발주일=?, 발주차수=?, 거래처명=?, 상품명=?, 유형=?, 통화=?, 발주총액=?, 마감여부=? WHERE 발주번호=?", (r['발주일'], r['발주차수'], r['거래처명'], r['상품명'], r['유형'], r['통화'], r['발주총액'], r['마감여부'], r['발주번호']))
                # 발주번호 기준 입금 상세내역(payments) 강제 연동 (거래처 공란 채우기 핵심 로직)
                conn.execute("UPDATE payments SET 거래처명=?, 유형=?, 상품명=?, 통화=? WHERE 발주번호=?", (r['거래처명'], r['유형'], r['상품명'], r['통화'], r['발주번호']))
            conn.commit(); st.success("✅ 동기화 완료!"); st.rerun()

# [Tab 3] 상세내역 및 통합 정산
with tabs[3]:
    st.header("🔍 상세 내역 및 통합 정산")
    p_all = pd.read_sql("SELECT * FROM payments", conn)
    o_all = pd.read_sql("SELECT * FROM orders", conn)
    
    if not p_all.empty:
        st.subheader("📋 유형별 지출 요약")
        p_all['입금일_dt'] = pd.to_datetime(p_all['입금일'])
        years = sorted(p_all['입금일_dt'].dt.year.unique(), reverse=True)
        sel_y = st.selectbox("연도 선택", years)
        fil_p = p_all[p_all['입금일_dt'].dt.year == sel_y]
        
        if not fil_p.empty:
            cat_sum = fil_p.groupby('유형').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
            cat_sum['총합계'] = cat_sum['실입금액'] + cat_sum['선급금액']
            st.table(cat_sum.style.format({'실입금액': '{:,.2f}', '선급금액': '{:,.2f}', '총합계': '{:,.2f}'}))
        
        st.divider()
        st.subheader("📊 발주번호별 정산 현황")
        # 발주번호별 입금 합계 계산
        p_agg = p_all.groupby('발주번호').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
        if not o_all.empty:
            # 발주 마스터와 조인하여 잔액 계산
            sum_df = pd.merge(o_all[['발주번호', '발주차수', '거래처명', '상품명', '발주총액', '통화']], p_agg, on='발주번호', how='left').fillna(0)
            sum_df['잔액'] = sum_df['발주총액'] - sum_df['실입금액']
            st.table(sum_df.style.format({'발주총액':'{:,.2f}', '실입금액':'{:,.2f}', '선급금액':'{:,.2f}', '잔액':'{:,.2f}'}))
        
        st.divider()
        st.subheader("📑 상세 입금 내역 편집")
        ed_p = st.data_editor(p_all.sort_values('입금일', ascending=False).drop(columns=['입금일_dt']), hide_index=True, use_container_width=True, disabled=["id"])
        if st.button("💾 상세 개별 수정 저장"):
            for _, r in ed_p.iterrows():
                conn.execute("UPDATE payments SET 발주번호=?, 입금일=?, 유형=?, 거래처명=?, 상품명=?, 실입금액=?, 선급금액=?, 메모=? WHERE id=?", (r['발주번호'], r['입금일'], r['유형'], r['거래처명'], r['상품명'], r['실입금액'], r['선급금액'], r['메모'], r['id']))
            conn.commit(); st.success("수정 완료"); st.rerun()
        
        c_del1, c_del2 = st.columns([1, 4])
        did = c_del1.number_input("삭제 ID 입력", min_value=0, step=1)
        if c_del1.button("🗑️ 행 삭제"):
            conn.execute(f"DELETE FROM payments WHERE id={did}")
            conn.commit(); st.rerun()

# [Tab 4] 거래처 관리
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    cl, cr = st.columns([1.2, 0.8])
    with cl:
        with st.form("v_reg"):
            st.subheader("➕ 거래처 수기 등록")
            vn, vt = st.text_input("거래처명"), st.selectbox("유형", CATEGORIES)
            vc1, vc2, vc3 = st.columns(3)
            vb, va, vh = vc1.text_input("은행"), vc2.text_input("계좌"), vc3.text_input("예금주")
            if st.form_submit_button("✅ 저장"):
                if vn: conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (vn,vb,va,vh,vt)); conn.commit(); st.rerun()
    with cr:
        st.subheader("📂 거래처 엑셀 업로드")
        vf = st.file_uploader("거래처 CSV 업로드", type=['csv'])
        if vf and st.button("🚀 업로드 실행"):
            v_up = pd.read_csv(vf)
            for _, r in v_up.iterrows(): 
                conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (r['거래처명'], r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
            conn.commit(); st.success("등록 완료"); st.rerun()
    st.divider()
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    if not v_data.empty:
        st.subheader("🏢 거래처 리스트 (수정 시 전체 소급)")
        orig = v_data['거래처명'].tolist()
        ev_v = st.data_editor(v_data, hide_index=True, use_container_width=True)
        if st.button("💾 거래처명 변경 및 전체 동기화"):
            for idx, r in ev_v.iterrows():
                old, new = orig[idx], r['거래처명']
                if old != new:
                    conn.execute(f"DELETE FROM vendors WHERE 거래처명='{old}'")
                    conn.execute("INSERT INTO vendors VALUES (?,?,?,?,?)", (new, r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
                    conn.execute("UPDATE orders SET 거래처명=?, 유형=? WHERE 거래처명=?", (new, r['기본유형'], old))
                    conn.execute("UPDATE payments SET 거래처명=?, 유형=? WHERE 거래처명=?", (new, r['기본유형'], old))
                else:
                    conn.execute("UPDATE vendors SET 은행=?, 계좌번호=?, 예금주=?, 기본유형=? WHERE 거래처명=?", (r['은행'], r['계좌번호'], r['예금주'], r['기본유형'], new))
            conn.commit(); st.success("✅ 동기화 완료!"); st.rerun()