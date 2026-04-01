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
    db_file = 'finance_final_v73.db'
    backup_file = f"backups/backup_{datetime.now().strftime('%Y%m%d')}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

# 페이지 레이아웃 설정
st.set_page_config(page_title="자금 관리 v73", layout="wide", page_icon="💰")
run_backup()

@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_final_v73.db', check_same_thread=False)
    c = conn.cursor()
    # 거래처 마스터 (기본키: 거래처명)
    c.execute('CREATE TABLE IF NOT EXISTS vendors (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT, 기본유형 TEXT)')
    # 발주 마스터
    c.execute('CREATE TABLE IF NOT EXISTS orders (발주번호 TEXT PRIMARY KEY, 발주일 TEXT, 거래처명 TEXT, 상품명 TEXT, 유형 TEXT, 통화 TEXT, 발주총액 REAL, 마감여부 INTEGER DEFAULT 0)')
    # 입금 및 지출 내역
    c.execute('''CREATE TABLE IF NOT EXISTS payments (id INTEGER PRIMARY KEY AUTOINCREMENT, 발주번호 TEXT, 입금일 TEXT, 유형 TEXT, 거래처명 TEXT, 상품명 TEXT, 통화 TEXT,
                  실입금액 REAL, 선급금액 REAL, 메모 TEXT, 한화환산액 REAL, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT)''')
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비", "원단비", "기타"]

# --- 2. 알람 및 업로더 세션 관리 ---
if 'alert_msg' not in st.session_state: st.session_state.alert_msg = []
if 'up_key' not in st.session_state: st.session_state.up_key = 0

def add_alert(text): st.session_state.alert_msg.append(text)
def clear_alerts(): st.session_state.alert_msg = []

# --- 3. 유틸리티 함수 ---
def clean_name(name):
    """공백 및 특수문자 제거하여 이름 매칭 정확도 향상"""
    if not name or pd.isna(name): return ""
    return re.sub(r'\s+', '', str(name))

def smart_date(date_str):
    """날짜 형식 자동 보정"""
    try:
        ds = str(date_str).strip()
        if "월" in ds and "일" in ds:
            return datetime.strptime(f"2026 {ds}", "%Y %m월 %d일").strftime("%Y-%m-%d")
        return pd.to_datetime(ds).strftime("%Y-%m-%d")
    except:
        return datetime.now().strftime("%Y-%m-%d")

# --- 4. ERP 발주서 분석 로직 (거래처명 엄격 매칭) ---
def process_ecount_v73(file):
    try:
        df = pd.read_excel(file, header=None)
        # 발주번호 추출
        raw_oid = str(df.iloc[1, 0]).split(":")[-1].strip() if ":" in str(df.iloc[1,0]) else str(df.iloc[1, 0])
        clean_oid = raw_oid.replace("-", "")
        odate = f"{clean_oid[:4]}-{clean_oid[4:6]}-{clean_oid[6:8]}" if len(clean_oid) >= 8 else datetime.now().strftime("%Y-%m-%d")
        
        # 거래처명 추출
        vendor_from_excel = "미지정"
        for i in range(len(df)):
            if "수신" in str(df.iloc[i, 0]):
                vendor_from_excel = str(df.iloc[i, 0]).split(":")[-1].strip()
                break
        
        # ★ 오직 '거래처명'을 기준으로 마스터와 대조
        v_master = pd.read_sql("SELECT 거래처명, 기본유형 FROM vendors", conn)
        v_master['clean_key'] = v_master['거래처명'].apply(clean_name)
        target_key = clean_name(vendor_from_excel)
        
        match = v_master[v_master['clean_key'] == target_key]
        
        if match.empty:
            # 매칭 실패 시 등록 차단
            add_alert(f"❌ 등록 실패: 발주서의 '{vendor_from_excel}'이(가) 거래처 리스트에 없습니다. 이름을 확인하거나 먼저 등록하세요.")
            return False
        
        # 정확한 이름과 유형 확정
        v_type = match.iloc[0]['기본유형']
        vendor_fixed = match.iloc[0]['거래처명']
        
        # 통화 및 품목 추출
        f6 = str(df.iloc[5, 5]) if len(df) > 5 else ""
        curr = "USD" if "USD" in f6 else ("CNY" if any(x in f6 for x in ["중국", "CNY"]) else "한화")
        p_col = 1 if curr == "한화" else 2
        prods = df.iloc[6:, p_col].dropna().astype(str).tolist()
        prod_n = (prods[0].split("[")[0].strip() + (f" 외 {len(prods)-1}건" if len(prods)>1 else "")) if prods else "품목미상"
        
        # 금액 추출
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
        add_alert(f"❗ 오류 발생({file.name}): {str(e)}")
        return False

# --- 5. 상단 고정 알람 UI ---
if st.session_state.alert_msg:
    for m in st.session_state.alert_msg:
        st.error(m)
    if st.button("알람 모두 확인 후 닫기"):
        clear_alerts()
        st.rerun()

# --- 6. 메인 UI 구성 ---
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리"])

# [Tab 0] 입금 수기 입력
with tabs[0]:
    st.header("📝 입금 내역 수기 입력")
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    o_active = pd.read_sql("SELECT * FROM orders WHERE 마감여부=0", conn)
    with st.form("p_manual_v73", clear_on_submit=True):
        c1, c2 = st.columns(2)
        p_oid = c1.selectbox("🔗 발주번호 연동", ["없음"] + list(o_active['발주번호']) if not o_active.empty else ["없음"])
        p_date = c2.date_input("입금일")
        c3, c4, c5 = st.columns(3)
        p_vn = c3.selectbox("거래처명", ["선택"] + list(v_data['거래처명']) if not v_data.empty else ["선택"])
        p_ct, p_pr = c4.selectbox("유형", CATEGORIES), c5.text_input("상품명")
        c6, c7, c8 = st.columns(3)
        p_dep, p_pre, p_cur = c6.number_input("실입금액", format="%.2f"), c7.number_input("선급금액", format="%.2f"), c8.selectbox("통화", ["한화", "USD", "CNY"])
        p_memo = st.text_input("메모")
        if st.form_submit_button("✅ 입금 저장"):
            if p_vn != "선택":
                rate = 1350.0 if p_cur == "USD" else (190.0 if p_cur == "CNY" else 1.0)
                vi = v_data[v_data['거래처명']==p_vn].iloc[0]
                conn.execute("INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                             (p_oid if p_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), p_ct, p_vn, p_pr, p_cur, p_dep, p_pre, p_memo, (p_dep+p_pre)*rate, vi['은행'], vi['계좌번호'], vi['예금주']))
                conn.commit(); st.success("성공적으로 저장되었습니다."); st.rerun()

# [Tab 1] 입금 엑셀 업로드
with tabs[1]:
    st.header("📂 통합 입금 엑셀 업로드")
    tmp_df = pd.DataFrame(columns=["발주번호", "거래처", "유형", "상품명", "입금일", "실입금액", "선급금액", "송금사유"])
    st.download_button("📥 양식 다운로드", tmp_df.to_csv(index=False).encode('utf-8-sig'), "payment_template.csv")
    f_p = st.file_uploader("입금 CSV 선택", type=['csv'], key=f"pay_up_{st.session_state.up_key}")
    if f_p and st.button("🚀 입금 데이터 일괄 저장"):
        try:
            df_p = pd.read_csv(f_p).dropna(subset=['실입금액', '거래처'], how='all')
            o_l = pd.read_sql("SELECT * FROM orders", conn); v_l = pd.read_sql("SELECT * FROM vendors", conn)
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
            conn.commit(); st.success(f"✅ 총 {len(df_p)}건의 데이터가 저장되었습니다."); st.session_state.up_key += 1; st.rerun()
        except Exception as e: st.error(f"오류: {e}")

# [Tab 2] 발주서 등록 및 관리
with tabs[2]:
    st.header("📥 발주서 등록 및 관리")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("⚡ 엑셀 일괄 등록")
        of_list = st.file_uploader("발주서(xlsx) 선택", type=['xlsx'], accept_multiple_files=True, key=f"ord_{st.session_state.up_key}")
        if of_list and st.button("🚀 모든 발주서 일괄 업로드 실행"):
            clear_alerts()
            scnt = sum(1 for of in of_list if process_ecount_v73(of))
            if scnt > 0: st.success(f"✅ {scnt}건 등록 완료!"); st.session_state.up_key += 1; st.rerun()
    with c2:
        st.subheader("✍️ 수기 등록")
        v_list_manual = pd.read_sql("SELECT 거래처명 FROM vendors", conn)
        with st.form("o_manual_v73", clear_on_submit=True):
            mi, md = st.text_input("발주번호"), st.date_input("발주일")
            mv = st.selectbox("거래처 선택", ["선택"] + list(v_list_manual['거래처명']))
            mp, mt = st.text_input("상품명"), st.number_input("발주금액", format="%.2f")
            if st.form_submit_button("✅ 수기 저장"):
                if mi and mv != "선택":
                    vt = pd.read_sql(f"SELECT 기본유형 FROM vendors WHERE 거래처명='{mv}'", conn).iloc[0]['기본유형']
                    conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", (mi, md.strftime("%Y-%m-%d"), mv, mp, vt, "한화", mt))
                    conn.commit(); st.success("수기 발주가 저장되었습니다."); st.rerun()
    st.divider()
    o_data = pd.read_sql("SELECT * FROM orders", conn)
    if not o_data.empty:
        st.subheader("📄 등록된 발주서 현황 (더블클릭 수정)")
        ev_o = st.data_editor(o_data.sort_values('발주일', ascending=False), hide_index=True, use_container_width=True, disabled=["발주번호"])
        if st.button("💾 발주 리스트 수정사항 저장"):
            for idx, r in ev_o.iterrows():
                conn.execute("UPDATE orders SET 발주일=?, 거래처명=?, 상품명=?, 유형=?, 통화=?, 발주총액=?, 마감여부=? WHERE 발주번호=?", (r['발주일'], r['거래처명'], r['상품명'], r['유형'], r['통화'], r['발주총액'], r['마감여부'], r['발주번호']))
            conn.commit(); st.success("수정 완료!"); st.rerun()

# [Tab 3] 상세내역 및 통합 정산
with tabs[3]:
    st.header("🔍 상세 내역 및 통합 정산")
    p_all = pd.read_sql("SELECT * FROM payments", conn); o_all = pd.read_sql("SELECT * FROM orders", conn)
    if not p_all.empty:
        # 1. 유형별 요약
        st.subheader("📋 유형별 지출 요약 (월별 필터)")
        p_all['입금일_dt'] = pd.to_datetime(p_all['입금일'])
        f_y, f_m = st.columns(2)
        sy = f_y.selectbox("연도 선택", sorted(p_all['입금일_dt'].dt.year.unique(), reverse=True))
        sm = f_m.selectbox("월 선택", ["전체"] + list(range(1, 13)))
        fil_p = p_all[p_all['입금일_dt'].dt.year == sy]
        if sm != "전체": fil_p = fil_p[fil_p['입금일_dt'].dt.month == int(sm)]
        if not fil_p.empty:
            cat_sum = fil_p.groupby('유형').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
            cat_sum['총합계'] = cat_sum['실입금액'] + cat_sum['선급금액']
            st.table(cat_sum.style.format({'실입금액': '{:,.2f}', '선급금액': '{:,.2f}', '총합계': '{:,.2f}'}))
        
        st.divider()
        # 2. 상세내역 리스트
        st.subheader("📑 상세 입금 내역 리스트 (최신순)")
        ed_p = st.data_editor(p_all.sort_values('입금일', ascending=False).drop(columns=['입금일_dt']), hide_index=True, use_container_width=True, disabled=["id"])
        if st.button("💾 입금 내역 수정 저장"):
            for idx, r in ed_p.iterrows():
                conn.execute("UPDATE payments SET 발주번호=?, 입금일=?, 유형=?, 거래처명=?, 상품명=?, 실입금액=?, 선급금액=?, 메모=? WHERE id=?", (r['발주번호'], r['입금일'], r['유형'], r['거래처명'], r['상품명'], r['실입금액'], r['선급금액'], r['메모'], r['id']))
            conn.commit(); st.success("상세 내역이 수정되었습니다."); st.rerun()
        
        did = st.number_input("삭제 ID 입력", min_value=0, step=1)
        if st.button("🗑️ 해당 ID 행 삭제"):
            conn.execute(f"DELETE FROM payments WHERE id={did}"); conn.commit(); st.rerun()

        st.divider()
        # 3. 발주번호별 정산 상황
        st.subheader("📊 발주번호별 정산 상황 (최신 발주순)")
        sum_df = p_all.groupby('발주번호').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
        if not o_all.empty:
            sum_df = sum_df.merge(o_all[['발주번호', '발주일', '거래처명', '발주총액']], on='발주번호', how='left')
            sum_df = sum_df.sort_values('발주일', ascending=False)
            sum_df['잔액'] = sum_df['발주총액'].fillna(0) - sum_df['실입금액']
            st.table(sum_df[['발주번호', '발주일', '거래처명', '발주총액', '실입금액', '선급금액', '잔액']].style.format({'발주총액': '{:,.2f}', '실입금액': '{:,.2f}', '선급금액': '{:,.2f}', '잔액': '{:,.2f}'}))

# [Tab 4] 거래처 관리 (★ 거래처명 중심 소급 동기화 엔진)
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    cl, cr = st.columns([1.2, 0.8])
    with cl:
        with st.form("v_reg_v73", clear_on_submit=True):
            st.subheader("➕ 거래처 수기 등록")
            vn, vt = st.text_input("거래처명"), st.selectbox("유형", CATEGORIES)
            vc1, vc2, vc3 = st.columns(3)
            vb, va, vh = vc1.text_input("은행명"), vc2.text_input("계좌번호"), vc3.text_input("예금주")
            if st.form_submit_button("✅ 저장"):
                if vn: conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (vn,vb,va,vh,vt)); conn.commit(); st.rerun()
    with cr:
        st.subheader("📂 거래처 엑셀 업로드")
        vf = st.file_uploader("거래처 CSV 업로드", type=['csv'])
        if vf and st.button("🚀 일괄 저장 실행"):
            v_up = pd.read_csv(vf)
            for _, r in v_up.iterrows(): conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (r['거래처명'], r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
            conn.commit(); st.success("거래처 정보 등록 완료"); st.rerun()

    st.divider()
    v_data_master = pd.read_sql("SELECT * FROM vendors", conn)
    if not v_data_master.empty:
        st.subheader("🏢 거래처 리스트 (이름 수정 시 발주/입금 내역 이름+유형 전체 소급 업데이트)")
        orig_names_list = v_data_master['거래처명'].tolist()
        ev_v = st.data_editor(v_data_master, hide_index=True, use_container_width=True, key="v_ed_v73")
        if st.button("💾 거래처명 기준 모든 데이터 동기화 저장"):
            try:
                for idx, r in ev_v.iterrows():
                    old_name = orig_names_list[idx]
                    new_name = r['거래처명']
                    
                    if old_name != new_name:
                        # 1. 마스터 이름 교체 (기존 데이터 삭제 후 재삽입)
                        conn.execute(f"DELETE FROM vendors WHERE 거래처명 = '{old_name}'")
                        conn.execute("INSERT INTO vendors VALUES (?,?,?,?,?)", (new_name, r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
                        # 2. ★ 오직 [거래처명]을 기준으로 과거 발주/입금 내역 강제 소급 적용 ★
                        conn.execute("UPDATE orders SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_name, r['기본유형'], old_name))
                        conn.execute("UPDATE payments SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_name, r['기본유형'], old_name))
                    else:
                        # 이름은 같고 나머지 정보나 유형만 바뀐 경우
                        conn.execute("UPDATE vendors SET 은행=?, 계좌번호=?, 예금주=?, 기본유형=? WHERE 거래처명=?", (r['은행'], r['계좌번호'], r['예금주'], r['기본유형'], new_name))
                        conn.execute("UPDATE orders SET 유형=? WHERE 거래처명=?", (r['기본유형'], new_name))
                        conn.execute("UPDATE payments SET 유형=? WHERE 거래처명=?", (r['기본유형'], new_name))
                
                conn.commit()
                st.success("✅ [거래처명] 기준 모든 데이터 동기화 및 업데이트가 완료되었습니다!"); st.rerun()
            except Exception as e: st.error(f"동기화 오류: {e}")