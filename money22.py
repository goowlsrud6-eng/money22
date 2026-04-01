import streamlit as st
import pandas as pd
import sqlite3
import os
import shutil
from datetime import datetime

# --- 1. 백업 및 DB 설정 ---
def run_backup():
    if not os.path.exists('backups'): os.makedirs('backups')
    db_file = 'finance_v59_final.db'
    backup_file = f"backups/backup_{datetime.now().strftime('%Y%m%d')}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

st.set_page_config(page_title="자금 관리 v59", layout="wide", page_icon="💰")
run_backup()

@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_v59_final.db', check_same_thread=False)
    c = conn.cursor()
    c.execute('CREATE TABLE IF NOT EXISTS vendors (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT, 기본유형 TEXT)')
    c.execute('CREATE TABLE IF NOT EXISTS orders (발주번호 TEXT PRIMARY KEY, 발주일 TEXT, 거래처명 TEXT, 상품명 TEXT, 유형 TEXT, 통화 TEXT, 발주총액 REAL, 마감여부 INTEGER DEFAULT 0)')
    c.execute('''CREATE TABLE IF NOT EXISTS payments (id INTEGER PRIMARY KEY AUTOINCREMENT, 발주번호 TEXT, 입금일 TEXT, 유형 TEXT, 거래처명 TEXT, 상품명 TEXT, 통화 TEXT,
                  실입금액 REAL, 선급금액 REAL, 메모 TEXT, 한화환산액 REAL, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT)''')
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비", "원단비", "기타"]

# --- 2. 유틸리티 함수 (서식/날짜/음영) ---
def format_num_display(val):
    """천단위 쉼표 + 소수점 2자리 서식 (표시용)"""
    try:
        return "{:,.2f}".format(float(val))
    except:
        return "0.00"

def smart_date(date_str):
    """연도 없는 날짜 보정"""
    try:
        ds = str(date_str).strip()
        if "월" in ds and "일" in ds:
            return datetime.strptime(f"2026 {ds}", "%Y %m월 %d일").strftime("%Y-%m-%d")
        return pd.to_datetime(ds).strftime("%Y-%m-%d")
    except: return datetime.now().strftime("%Y-%m-%d")

def style_row(row):
    """마감 시 회색 음영 및 취소선"""
    if row.get('마감여부') == 1:
        return ['background-color: #f0f0f0; color: #a0a0a0; text-decoration: line-through'] * len(row)
    return [''] * len(row)

# --- 3. ERP 분석 로직 ---
def process_ecount_v59(file):
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
        
        # 거래처 유형 연동
        v_check = pd.read_sql(f"SELECT 기본유형 FROM vendors WHERE 거래처명 = '{vendor}'", conn)
        v_type = v_check.iloc[0]['기본유형'] if not v_check.empty else "사입"
        
        f6 = str(df.iloc[5, 5]) if len(df) > 5 else ""
        curr = "USD" if "USD" in f6 else ("CNY" if any(x in f6 for x in ["중국", "CNY"]) else "한화")
        p_col = 1 if curr == "한화" else 2
        prods = df.iloc[6:, p_col].dropna().astype(str).tolist()
        prod_n = (prods[0].split("[")[0].strip() + (f" 외 {len(prods)-1}건" if len(prods)>1 else "")) if prods else "품목미상"
        
        total = 0.0
        if curr != "한화":
            last_idx = df.iloc[:, 5].last_valid_index()
            total = float(df.iloc[last_idx, 5]) if last_idx is not None else 0.0
        else:
            a5 = str(df.iloc[4, 0])
            total = float(a5.split(":")[-1].replace(",", "").strip()) if "금액" in a5 else 0.0

        conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", (raw_oid, odate, vendor, prod_n, v_type, curr, total))
        conn.commit(); return True, raw_oid
    except Exception as e: return False, str(e)

# --- 4. 메인 UI ---
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리"])

# [Tab 1] 입금 수기 입력
with tabs[0]:
    st.header("📝 입금 내역 수기 입력")
    v_m = pd.read_sql("SELECT * FROM vendors", conn); o_m = pd.read_sql("SELECT * FROM orders", conn)
    active_o = o_m[o_m['마감여부'] == 0] if not o_m.empty else pd.DataFrame()
    with st.form("p_manual_v59", clear_on_submit=True):
        sel_oid = st.selectbox("🔗 발주번호 연동", ["없음"] + list(active_o['발주번호']) if not active_o.empty else ["없음"])
        c1, c2, c3 = st.columns(3)
        p_d, p_v = c1.date_input("입금일"), c2.selectbox("거래처명", ["선택"] + list(v_m['거래처명']) if not v_m.empty else ["선택"])
        p_cat = c3.selectbox("유형", CATEGORIES)
        p_prod = st.text_input("상품명 (품목)")
        c4, c5, c6 = st.columns(3)
        p_dep, p_pre, p_cur = c4.number_input("실입금액"), c5.number_input("선급금액"), c6.selectbox("통화", ["한화", "USD", "CNY"])
        p_note = st.text_input("메모(송금사유)")
        if st.form_submit_button("저장"):
            if p_v != "선택":
                rate = 1350.0 if p_cur == "USD" else (190.0 if p_cur == "CNY" else 1.0)
                vi = v_m[v_m['거래처명']==p_v].iloc[0]
                conn.cursor().execute("INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                                     (sel_oid if sel_oid != "없음" else None, p_d.strftime("%Y-%m-%d"), p_cat, p_v, p_prod, p_cur, p_dep, p_pre, p_note, (p_dep+p_pre)*rate, vi['은행'], vi['계좌번호'], vi['예금주']))
                conn.commit(); st.success("저장 완료!"); st.rerun()

# [Tab 2] 입금 엑셀 업로드
with tabs[1]:
    st.header("📂 통합 입금 엑셀 업로드")
    tmp = pd.DataFrame(columns=["발주번호", "거래처", "유형", "상품명", "입금일", "실입금액", "선급금액", "송금사유"])
    st.download_button("📥 입금 양식 다운로드", tmp.to_csv(index=False).encode('utf-8-sig'), "payment_v59.csv")
    f_p = st.file_uploader("CSV 파일 선택", type=['csv'], key="p_csv_v59")
    if f_p and st.button("🚀 입금 데이터 일괄 저장"):
        try:
            df_p = pd.read_csv(f_p).dropna(subset=['실입금액', '거래처'], how='all')
            o_df = pd.read_sql("SELECT * FROM orders", conn); v_df = pd.read_sql("SELECT * FROM vendors", conn)
            for _, r in df_p.iterrows():
                pd_s = smart_date(r['입금일']); oid = str(r['발주번호']).strip() if pd.notna(r['발주번호']) else "nan"
                if oid != "nan" and not o_df[o_df['발주번호'] == oid].empty:
                    info = o_df[o_df['발주번호'] == oid].iloc[0]
                    vn, pc, pp, cur = info['거래처명'], info['유형'], info['상품명'], info['통화']
                else: vn, pc, pp, cur = str(r['거래처']), str(r['유형']), str(r['상품명']), "한화"
                if vn == "nan": continue
                vi = v_df[v_df['거래처명'] == vn]
                bk, ac, hd = (vi.iloc[0]['은행'], vi.iloc[0]['계좌번호'], vi.iloc[0]['예금주']) if not vi.empty else ("","","")
                rt = 1350.0 if cur == "USD" else (190.0 if cur == "CNY" else 1.0)
                s_amt, p_amt = float(r['실입금액']) if pd.notna(r['실입금액']) else 0.0, float(r['선급금액']) if pd.notna(r['선급금액']) else 0.0
                conn.cursor().execute("INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                                     (oid if oid != "nan" else None, pd_s, pc, vn, pp, cur, s_amt, p_amt, r['송금사유'], (s_amt+p_amt)*rt, bk, ac, hd))
            conn.commit(); st.success("저장 성공!"); st.rerun()
        except Exception as e: st.error(f"에러: {e}")

# [Tab 3] 발주서 등록
with tabs[2]:
    st.header("📥 발주서 등록")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("⚡ 1. 이카운트 엑셀 일괄 등록")
        of_list = st.file_uploader("발주서(.xlsx)들을 한 번에 선택", type=['xlsx'], accept_multiple_files=True, key="of_v59")
        if of_list and st.button("🚀 모든 파일 등록"):
            for of in of_list: process_ecount_v59(of)
            st.rerun()
    with c2:
        st.subheader("✍️ 2. 발주서 수기 등록")
        with st.form("o_man_v59", clear_on_submit=True):
            mi, md = st.text_input("발주번호"), st.date_input("발주일")
            mv = st.selectbox("거래처", ["선택"] + list(pd.read_sql("SELECT 거래처명 FROM vendors", conn)['거래처명']))
            mp, mt = st.text_input("상품명"), st.number_input("발주총액")
            if st.form_submit_button("수기 저장"):
                if mi and mv != "선택":
                    conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,0)", (mi, md.strftime("%Y-%m-%d"), mv, mp, "사입", "한화", mt))
                    conn.commit(); st.rerun()
    st.divider()
    o_list = pd.read_sql("SELECT * FROM orders", conn)
    if not o_list.empty:
        st.subheader("📄 발주서 리스트 (마감 건 음영)")
        st.dataframe(o_list.sort_values('발주일', ascending=False).style.apply(style_row, axis=1).format({'발주총액': '{:,.2f}'}), use_container_width=True, hide_index=True)

# [Tab 4] 상세내역 및 통합 정산
with tabs[3]:
    st.header("🔍 상세 내역 및 통합 정산 (수정 가능)")
    p_df = pd.read_sql("SELECT * FROM payments", conn); o_df = pd.read_sql("SELECT * FROM orders", conn)
    if not p_df.empty:
        # 1. 유형별 요약
        st.subheader("📋 유형별 지출 요약 (물류비/원단비 등)")
        cat_sum = p_df.groupby('유형').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
        cat_sum['총합계'] = cat_sum['실입금액'] + cat_sum['선급금액']
        st.table(cat_sum.style.format({'실입금액': '{:,.2f}', '선급금액': '{:,.2f}', '총합계': '{:,.2f}'}))

        # 2. 발주번호별 요약
        st.subheader("📊 발주번호별 정산 요약")
        sum_df = p_df[p_df['발주번호'].notna()].groupby('발주번호').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
        if not o_df.empty:
            sum_df = sum_df.merge(o_df[['발주번호', '발주총액', '거래처명']], on='발주번호', how='left')
            sum_df['미입금잔액'] = sum_df['발주총액'].fillna(0) - sum_df['실입금액']
            st.table(sum_df.style.format({'발주총액': '{:,.2f}', '실입금액': '{:,.2f}', '선급금액': '{:,.2f}', '미입금잔액': '{:,.2f}'}))

        # 3. 상세 편집 테이블 (천단위 서식 적용)
        st.subheader("📑 입금 상세 내역 편집")
        edited_p = st.data_editor(p_df, hide_index=True, use_container_width=True, disabled=["id"], key="p_editor_v59")
        if st.button("💾 모든 수정사항 저장"):
            for idx, r in edited_p.iterrows():
                conn.execute("UPDATE payments SET 발주번호=?, 입금일=?, 유형=?, 거래처명=?, 상품명=?, 실입금액=?, 선급금액=?, 메모=? WHERE id=?",
                             (r['발주번호'], r['입금일'], r['유형'], r['거래처명'], r['상품명'], r['실입금액'], r['선급금액'], r['메모'], r['id']))
            conn.commit(); st.success("DB 업데이트 완료!"); st.rerun()

        st.divider()
        c_del, c_cl = st.columns(2)
        with c_del:
            did = st.number_input("지울 행의 ID 입력", min_value=0, step=1)
            if st.button("🗑️ 해당 행 완전 삭제"):
                conn.execute(f"DELETE FROM payments WHERE id={did}"); conn.commit(); st.success("삭제됨"); st.rerun()
        with c_cl:
            toid = st.selectbox("마감할 발주번호", o_df[o_df['마감여부']==0]['발주번호'].unique() if not o_df.empty else [])
            if st.button("🚩 해당 발주 마감"):
                conn.execute(f"UPDATE orders SET 마감여부=1 WHERE 발주번호='{toid}'"); conn.commit(); st.rerun()
    else: st.info("데이터가 없습니다.")

# [Tab 5] 거래처 관리
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    # 엑셀 등록
    c_u1, c_u2 = st.columns(2)
    with c_u1:
        st.subheader("📂 엑셀 일괄 등록")
        v_tmp = pd.DataFrame(columns=["거래처명", "은행", "계좌번호", "예금주", "기본유형"])
        st.download_button("📥 거래처 양식 다운로드", v_tmp.to_csv(index=False).encode('utf-8-sig'), "vendor_temp.csv")
    with c_u2:
        vf = st.file_uploader("거래처 CSV 업로드", type=['csv'], key="v_csv_v59")
        if vf and st.button("🚀 일괄 업로드 실행"):
            v_df_up = pd.read_csv(vf)
            for _, r in v_df_up.iterrows():
                conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (r['거래처명'], r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
            conn.commit(); st.rerun()

    st.divider()
    v_df = pd.read_sql("SELECT * FROM vendors", conn)
    if not v_df.empty:
        st.subheader("🏢 등록 거래처 목록 (더블클릭 수정 가능)")
        edited_v = st.data_editor(v_df, hide_index=True, use_container_width=True, key="v_editor_v59")
        if st.button("💾 거래처 수정사항 저장"):
            for idx, r in edited_v.iterrows():
                conn.execute("UPDATE vendors SET 은행=?, 계좌번호=?, 예금주=?, 기본유형=? WHERE 거래처명=?",
                             (r['은행'], r['계좌번호'], r['예금주'], r['기본유형'], r['거래처명']))
            conn.commit(); st.rerun()