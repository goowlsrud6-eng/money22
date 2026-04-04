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
    db_file = 'finance_final_v93.db'
    backup_file = f"backups/backup_{datetime.now().strftime('%Y%m%d')}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

st.set_page_config(page_title="자금 관리 v93", layout="wide", page_icon="💰")
run_backup()

@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_final_v93.db', check_same_thread=False)
    c = conn.cursor()
    # 거래처 마스터
    c.execute('CREATE TABLE IF NOT EXISTS vendors (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT, 기본유형 TEXT)')
    # 발주 마스터 (마감여부 포함)
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (발주번호 TEXT PRIMARY KEY, 발주일 TEXT, 발주차수 TEXT, 거래처명 TEXT, 
                  상품명 TEXT, 유형 TEXT, 통화 TEXT, 발주총액 REAL, 마감여부 INTEGER DEFAULT 0)''')
    # 입금 상세 내역 (13개 컬럼)
    c.execute('''CREATE TABLE IF NOT EXISTS payments 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 발주번호 TEXT, 입금일 TEXT, 
                  유형 TEXT, 거래처명 TEXT, 상품명 TEXT, 통화 TEXT,
                  실입금액 REAL, 선급금액 REAL, 메모 TEXT, 한화환산액 REAL,
                  은행 TEXT, 계좌번호 TEXT, 예금주 TEXT)''')
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비", "원단비", "기타"]

# --- 2. 세션 상태 관리 (파일 업로더 초기화용) ---
if 'order_up_key' not in st.session_state: st.session_state.order_up_key = 0
if 'pay_up_key' not in st.session_state: st.session_state.pay_up_key = 1000

# --- 3. 유틸리티 함수 (빈칸 대응 및 데이터 정제) ---
def to_float(val):
    try:
        if val is None or pd.isna(val) or str(val).strip() == "": return 0.0
        return float(str(val).replace(',', ''))
    except: return 0.0

def to_str(val):
    if val is None or pd.isna(val): return ""
    s = str(val).strip()
    return "" if s.lower() in ["nan", "none", ""] else s

def smart_date(date_str):
    try:
        ds = to_str(date_str)
        if not ds: return datetime.now().strftime("%Y-%m-%d")
        # "01월 07일" 형태 대응
        if "월" in ds and "일" in ds:
            return datetime.strptime(f"2026 {ds}", "%Y %m월 %d일").strftime("%Y-%m-%d")
        return pd.to_datetime(ds).strftime("%Y-%m-%d")
    except: return datetime.now().strftime("%Y-%m-%d")

# --- 4. ERP 발주서 분석 엔진 ---
def process_ecount_v93(file):
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
        v_master['clean'] = v_master['거래처명'].apply(lambda x: re.sub(r'\s+', '', str(x)))
        target_key = re.sub(r'\s+', '', vendor_raw)
        
        match = v_master[v_master['clean'] == target_key]
        if match.empty: return False, f"⚠️ '{vendor_raw}' 업체가 거래처 관리에 없습니다."
        
        v_type, vendor_fixed = match.iloc[0]['기본유형'], match.iloc[0]['거래처명']
        f6 = str(df.iloc[5, 5]) if len(df) > 5 else ""
        curr = "USD" if "USD" in f6 else ("CNY" if any(x in f6 for x in ["중국", "CNY"]) else "한화")
        p_col = 1 if curr == "한화" else 2
        prods = df.iloc[6:, p_col].dropna().astype(str).tolist()
        prod_n = (prods[0].split("[")[0].strip() + (f" 외 {len(prods)-1}건" if len(prods)>1 else "")) if prods else "품목미상"
        
        l_idx = df.iloc[:, 5].last_valid_index()
        total = to_float(df.iloc[l_idx, 5]) if curr != "한화" and l_idx else to_float(str(df.iloc[4, 0]).split(":")[-1])

        conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,0)", (raw_oid, odate, "", vendor_fixed, prod_n, v_type, curr, total))
        conn.commit(); return True, None
    except: return False, "❗ 발주서 형식 오류"

# --- 5. UI 섹션 ---
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리"])

# [Tab 0] 입금 수기 입력
with tabs[0]:
    st.header("📝 입금 내역 수기 입력")
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    o_active = pd.read_sql("SELECT 발주번호 FROM orders WHERE 마감여부=0", conn)
    with st.form("manual_pay_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        p_oid = c1.selectbox("🔗 진행중 발주번호 (선택)", ["없음"] + list(o_active['발주번호']))
        p_date = c2.date_input("입금일")
        c3, c4, c5 = st.columns(3)
        p_vn = c3.selectbox("거래처명", ["선택"] + list(v_data['거래처명']))
        p_ct, p_pr = c4.selectbox("유형", CATEGORIES), c5.text_input("상품명")
        c6, c7, c8 = st.columns(3)
        p_dep, p_pre, p_cur = c6.number_input("실입금액", format="%.2f"), c7.number_input("선급금액", format="%.2f"), c8.selectbox("통화", ["한화", "USD", "CNY"])
        p_memo = st.text_input("메모(송금사유)")
        if st.form_submit_button("✅ 입금 내역 저장"):
            rate = 1350.0 if p_cur == "USD" else (190.0 if p_cur == "CNY" else 1.0)
            vi = v_data[v_data['거래처명']==p_vn].iloc[0] if p_vn != "선택" else {}
            conn.execute("INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                         (to_str(p_oid) if p_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), p_ct, p_vn if p_vn != "선택" else "", p_pr, p_cur, p_dep, p_pre, p_memo, (p_dep+p_pre)*rate, vi.get('은행',''), vi.get('계좌번호',''), vi.get('예금주','')))
            conn.commit(); st.success("성공적으로 저장되었습니다."); st.rerun()

# [Tab 1] 입금 엑셀 업로드
with tabs[1]:
    st.header("📂 통합 입금 엑셀 업로드")
    # 샘플 양식 다운로드 로직 (v83에서 복구된 부분)
    sample_df = pd.DataFrame(columns=["발주번호", "거래처", "유형", "상품명", "입금일", "실입금액", "선급금액", "송금사유"])
    st.download_button("📥 입금 업로드 샘플 양식 다운로드", sample_df.to_csv(index=False).encode('utf-8-sig'), "pay_sample.csv", "text/csv")
    
    f_p = st.file_uploader("입금 CSV 파일 선택", type=['csv'], key=f"pay_up_{st.session_state.pay_up_key}")
    if f_p and st.button("🚀 데이터 일괄 저장 실행"):
        try:
            df_p = pd.read_csv(f_p)
            v_l = pd.read_sql("SELECT * FROM vendors", conn); o_l = pd.read_sql("SELECT * FROM orders", conn)
            for _, r in df_p.iterrows():
                oid, vn_raw = to_str(r.get('발주번호')), to_str(r.get('거래처'))
                if not vn_raw and not oid: continue # 둘 다 없으면 패스
                
                pd_s = smart_date(r.get('입금일'))
                # 발주번호가 있으면 발주서 마스터에서 정보를 우선적으로 가져옴 (연동의 핵심)
                if oid and not o_l[o_l['발주번호'] == oid].empty:
                    info = o_l[o_l['발주번호'] == oid].iloc[0]
                    vn, pc, pp, cur = info['거래처명'], info['유형'], info['상품명'], info['통화']
                else: vn, pc, pp, cur = vn_raw, to_str(r.get('유형')) or "사입", to_str(r.get('상품명')), "한화"
                
                vi = v_l[v_l['거래처명'] == vn] if vn else pd.DataFrame()
                dep, pre = to_float(r.get('실입금액')), to_float(r.get('선급금액'))
                rt = 1350.0 if cur == "USD" else (190.0 if cur == "CNY" else 1.0)
                
                conn.execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                             (oid if oid else None, pd_s, pc, vn, pp, cur, dep, pre, to_str(r.get('송금사유')), (dep+pre)*rt, 
                              vi.iloc[0]['은행'] if not vi.empty else "", vi.iloc[0]['계좌번호'] if not vi.empty else "", vi.iloc[0]['예금주'] if not vi.empty else ""))
            conn.commit(); st.success("✅ 일괄 저장 완료!"); st.session_state.pay_up_key += 1; st.rerun()
        except Exception as e: st.error(f"오류 발생: {e}")

# [Tab 2] 발주서 등록 및 마감 관리
with tabs[2]:
    st.header("📥 발주서 등록 및 마감 관리")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("⚡ 엑셀 일괄 등록")
        of_list = st.file_uploader("발주서(xlsx) 선택", type=['xlsx'], accept_multiple_files=True, key=f"ord_{st.session_state.order_up_key}")
        if of_list and st.button("🚀 모든 발주서 등록"):
            scnt, errs = 0, []
            for of in of_list:
                res, msg = process_ecount_v93(of)
                if res: scnt += 1
                else: errs.append(msg)
            for em in errs: st.warning(em) # 탭 내부 경고창
            if scnt > 0: 
                st.success(f"✅ {scnt}건 등록 성공!"); 
                if not errs: st.session_state.order_up_key += 1; st.rerun()
    with c2:
        st.subheader("✍️ 수기 발주 등록")
        v_list = pd.read_sql("SELECT 거래처명 FROM vendors", conn)
        with st.form("o_manual_form"):
            mi, mt_order = st.text_input("발주번호"), st.text_input("발주차수")
            md, mv = st.date_input("발주일"), st.selectbox("거래처 선택", ["선택"] + list(v_list['거래처명']))
            mp, mt, m_cur = st.text_input("상품명"), st.number_input("발주총액", format="%.2f"), st.selectbox("통화", ["한화", "USD", "CNY"])
            if st.form_submit_button("✅ 수기 저장"):
                if mi and mv != "선택":
                    vt = pd.read_sql(f"SELECT 기본유형 FROM vendors WHERE 거래처명='{mv}'", conn).iloc[0]['기본유형']
                    conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,0)", (mi, md.strftime("%Y-%m-%d"), mt_order, mv, mp, vt, m_cur, mt))
                    conn.commit(); st.rerun()
    st.divider()
    o_all_data = pd.read_sql("SELECT * FROM orders", conn)
    if not o_all_data.empty:
        st.subheader("📄 발주 리스트 (수정 및 동기화)")
        show_closed = st.checkbox("마감된 발주 포함해서 보기", value=False)
        disp_o = o_all_data if show_closed else o_all_data[o_all_data['마감여부'] == 0]
        ev_o = st.data_editor(disp_o[['발주번호', '발주차수', '거래처명', '상품명', '유형', '통화', '발주총액', '마감여부', '발주일']].sort_values('발주일', ascending=False), 
                              hide_index=True, use_container_width=True, disabled=["발주번호"],
                              column_config={"마감여부": st.column_config.CheckboxColumn("마감", help="체크 시 정산완료 처리")})
        if st.button("💾 정보 업데이트 및 모든 상세내역 동기화"):
            for _, r in ev_o.iterrows():
                conn.execute("UPDATE orders SET 발주일=?, 발주차수=?, 거래처명=?, 상품명=?, 유형=?, 통화=?, 발주총액=?, 마감여부=? WHERE 발주번호=?", (r['발주일'], r['발주차수'], r['거래처명'], r['상품명'], r['유형'], r['통화'], r['발주총액'], int(r['마감여부']), r['발주번호']))
                # 발주번호 기준 입금 상세내역(payments) 소급 업데이트 (거래처명 연동 핵심)
                conn.execute("UPDATE payments SET 거래처명=?, 유형=?, 상품명=?, 통화=? WHERE 발주번호=?", (r['거래처명'], r['유형'], r['상품명'], r['통화'], r['발주번호']))
            conn.commit(); st.success("✅ 모든 데이터 연동 및 동기화가 완료되었습니다."); st.rerun()

# [Tab 3] 상세내역 및 통합 정산
with tabs[3]:
    st.header("🔍 상세 내역 및 정산 현황")
    p_all = pd.read_sql("SELECT * FROM payments", conn); o_all = pd.read_sql("SELECT * FROM orders", conn)
    if not p_all.empty:
        p_all['dt'] = pd.to_datetime(p_all['입금일'])
        st.subheader("📊 월별 필터 및 검색")
        c1, c2, c3 = st.columns([1,1,2])
        y = c1.selectbox("연도", sorted(p_all['dt'].dt.year.unique(), reverse=True))
        m = c2.selectbox("월", ["전체"] + sorted(list(p_all[p_all['dt'].dt.year==y]['dt'].dt.month.unique())))
        search = c3.text_input("업체명 / 상품명 통합 검색")
        
        fil_p = p_all[p_all['dt'].dt.year == y]
        if m != "전체": fil_p = fil_p[fil_p['dt'].dt.month == m]
        if search: fil_p = fil_p[fil_p['거래처명'].str.contains(search, na=False) | fil_p['상품명'].str.contains(search, na=False)]
        
        if not fil_p.empty:
            cat_sum = fil_p.groupby('유형').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
            cat_sum['총합계'] = cat_sum['실입금액'] + cat_sum['선급금액']
            st.table(cat_sum.style.format({'실입금액':'{:,.2f}', '선급금액':'{:,.2f}', '총합계':'{:,.2f}'}))
        
        st.divider()
        st.subheader("📊 발주별 정산 및 잔액 현황")
        p_agg = p_all.groupby('발주번호').agg({'실입금액':'sum'}).reset_index()
        sum_df = pd.merge(o_all, p_agg, on='발주번호', how='left').fillna(0)
        sum_df['잔액'] = sum_df['발주총액'] - sum_df['실입금액']
        sum_df['상태'] = sum_df['마감여부'].apply(lambda x: "✅ 마감" if x == 1 else "⏳ 진행중")
        st.dataframe(sum_df[['발주번호', '상태', '거래처명', '상품명', '발주총액', '실입금액', '잔액', '통화']].sort_values('상태'), use_container_width=True)

        st.divider()
        st.subheader("📑 상세 내역 리스트 (필터링 결과)")
        ed_p = st.data_editor(fil_p.drop(columns=['dt']).sort_values('입금일', ascending=False), hide_index=True, use_container_width=True, disabled=["id"])
        if st.button("💾 개별 상세 내역 수정 저장"):
            for _, r in ed_p.iterrows():
                conn.execute("UPDATE payments SET 발주번호=?, 입금일=?, 유형=?, 거래처명=?, 상품명=?, 실입금액=?, 선급금액=?, 메모=? WHERE id=?", (r['발주번호'], r['입금일'], r['유형'], r['거래처명'], r['상품명'], r['실입금액'], r['선급금액'], r['메모'], r['id']))
            conn.commit(); st.success("수정 사항이 저장되었습니다."); st.rerun()
        
        did = st.number_input("삭제할 데이터 ID", min_value=0, step=1)
        if st.button("🗑️ 해당 내역 삭제"):
            conn.execute(f"DELETE FROM payments WHERE id={did}")
            conn.commit(); st.rerun()

# [Tab 4] 거래처 관리
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    cl, cr = st.columns([1.2, 0.8])
    with cl:
        with st.form("vendor_reg_form"):
            st.subheader("➕ 거래처 수기 등록")
            vn, vt = st.text_input("거래처명"), st.selectbox("기본 유형", CATEGORIES)
            vc1, vc2, vc3 = st.columns(3)
            vb, va, vh = vc1.text_input("은행명"), vc2.text_input("계좌번호"), vc3.text_input("예금주")
            if st.form_submit_button("✅ 거래처 저장"):
                if vn: conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (vn,vb,va,vh,vt)); conn.commit(); st.rerun()
    with cr:
        st.subheader("📂 엑셀 일괄 업로드")
        vf = st.file_uploader("거래처 CSV 업로드", type=['csv'])
        if vf and st.button("🚀 업로드 실행"):
            pd.read_csv(vf).to_sql('vendors', conn, if_exists='append', index=False)
            st.success("업로드 완료!"); st.rerun()
    st.divider()
    v_master_list = pd.read_sql("SELECT * FROM vendors", conn)
    if not v_master_list.empty:
        st.subheader("🏢 등록된 거래처 목록 (수정 시 전체 소급)")
        orig_v = v_master_list['거래처명'].tolist()
        ev_v = st.data_editor(v_master_list, hide_index=True, use_container_width=True)
        if st.button("💾 거래처 정보 변경 및 모든 데이터 동기화"):
            for idx, r in ev_v.iterrows():
                old, new = orig_v[idx], r['거래처명']
                if old != new:
                    conn.execute(f"DELETE FROM vendors WHERE 거래처명='{old}'")
                    conn.execute("INSERT INTO vendors VALUES (?,?,?,?,?)", (new, r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
                    conn.execute("UPDATE orders SET 거래처명=?, 유형=? WHERE 거래처명=?", (new, r['기본유형'], old))
                    conn.execute("UPDATE payments SET 거래처명=?, 유형=? WHERE 거래처명=?", (new, r['기본유형'], old))
                else:
                    conn.execute("UPDATE vendors SET 은행=?, 계좌번호=?, 예금주=?, 기본유형=? WHERE 거래처명=?", (r['은행'], r['계좌번호'], r['예금주'], r['기본유형'], new))
            conn.commit(); st.success("✅ 거래처 및 관련 모든 데이터가 업데이트되었습니다."); st.rerun()