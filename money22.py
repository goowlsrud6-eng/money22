import streamlit as st
import pandas as pd
import sqlite3
import os
import shutil
import re
from datetime import datetime, timedelta
import plotly.graph_objects as go
import urllib.request

# ==========================================
# 1. 백업 및 데이터베이스 초기화
# ==========================================
def run_backup():
    """매일 첫 접속 시 데이터베이스 백업 생성"""
    if not os.path.exists('backups'):
        os.makedirs('backups')
    db_file = 'finance_final_v136.db'
    today_str = datetime.now().strftime('%Y%m%d')
    backup_file = f"backups/backup_{today_str}.db"
    
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

st.set_page_config(page_title="자금 관리 시스템", layout="wide")
run_backup()

@st.cache_resource
def get_db_connection():
    """테이블 스키마 생성 및 유지"""
    conn = sqlite3.connect('finance_final_v136.db', check_same_thread=False)
    c = conn.cursor()
    # 거래처 정보
    c.execute('CREATE TABLE IF NOT EXISTS vendors (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT, 기본유형 TEXT)')
    # 발주 정보
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (발주번호 TEXT PRIMARY KEY, 발주일 TEXT, 발주차수 TEXT, 거래처명 TEXT, 
                 상품명 TEXT, 유형 TEXT, 통화 TEXT, 발주총액 REAL, 마감여부 INTEGER DEFAULT 0)''')
    # 입금 내역
    c.execute('''CREATE TABLE IF NOT EXISTS payments 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 발주번호 TEXT, 입금일 TEXT, 
                 유형 TEXT, 거래처명 TEXT, 상품명 TEXT, 통화 TEXT,
                 실입금액 REAL, 선급금액 REAL, 메모 TEXT, 한화환산액 REAL,
                 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT)''')
    # 환율 정보
    c.execute('CREATE TABLE IF NOT EXISTS exchange_rates (날짜 TEXT PRIMARY KEY, usd REAL, cny REAL)')
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비", "원단비", "기타"]

if 'order_up_key' not in st.session_state: 
    st.session_state.order_up_key = 0
if 'pay_up_key' not in st.session_state: 
    st.session_state.pay_up_key = 1000

# ==========================================
# 2. 유틸리티 함수 (날짜 버그 완전 해결본)
# ==========================================
def to_float(val):
    try:
        if val is None or pd.isna(val) or str(val).strip() == "": 
            return 0.0
        return float(str(val).replace(',', '').strip())
    except: 
        return 0.0

def to_str(val):
    if val is None or pd.isna(val): 
        return ""
    s = str(val).strip()
    return "" if s.lower() in ["nan", "none", ""] else s

def smart_date(date_val):
    """
    엑셀의 날짜 형식(Timestamp)과 문자열(2025. 1. 7.)을 모두 판별하여 정확한 날짜로 변환
    """
    try:
        if pd.isna(date_val) or str(date_val).strip() == "":
            return datetime.now().strftime("%Y-%m-%d")
        
        # 1. 엑셀 날짜 객체인 경우
        if isinstance(date_val, (datetime, pd.Timestamp)):
            return date_val.strftime("%Y-%m-%d")
            
        # 2. 문자열인 경우 (끝의 점 제거 및 기호 통일)
        ds = str(date_val).strip().rstrip('.')
        ds = ds.replace(" ", "").replace(".", "-").replace("/", "-")
        
        # 3. 파이썬 datetime 변환 시도
        return pd.to_datetime(ds).strftime("%Y-%m-%d")
    except:
        # 끝까지 변환이 안 되면 오늘 날짜
        return datetime.now().strftime("%Y-%m-%d")

@st.cache_data(ttl=3600)
def get_realtime_rate(currency):
    """네이버 환율 크롤러 (디테일: 크롤링 실패 대비 로직 유지)"""
    try:
        url = f"https://finance.naver.com/marketindex/exchangeDetail.naver?marketindexCd=FX_{currency}KRW"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        html = urllib.request.urlopen(req).read().decode('euc-kr', errors='ignore')
        match = re.search(r'<meta property="og:description" content=".*?\s+([\d,.]+)\s+전일대비', html)
        if match: 
            return float(match.group(1).replace(',', ''))
    except: 
        pass
    return 1350.0 if currency == 'USD' else (190.0 if currency == 'CNY' else 1.0)

# ==========================================
# 3. 데이터 처리 엔진
# ==========================================
def process_exchange_csv(file, currency_type):
    try:
        df = pd.read_csv(file)
        df.columns = [c.strip().replace('\ufeff', '') for c in df.columns]
        for _, row in df.iterrows():
            date_val = smart_date(row['날짜'])
            price_val = to_float(row['종가'])
            existing = pd.read_sql(f"SELECT * FROM exchange_rates WHERE 날짜 = '{date_val}'", conn)
            if existing.empty:
                usd = price_val if currency_type == "USD" else 0.0
                cny = price_val if currency_type == "CNY" else 0.0
                conn.execute("INSERT INTO exchange_rates VALUES (?,?,?)", (date_val, usd, cny))
            else:
                col = "usd" if currency_type == "USD" else "cny"
                conn.execute(f"UPDATE exchange_rates SET {col} = ? WHERE 날짜 = ?", (price_val, date_val))
        conn.commit()
        return True
    except:
        return False

def process_ecount_v136(file):
    try:
        df = pd.read_excel(file, header=None)
        raw_oid = str(df.iloc[1, 0]).split(":")[-1].strip() if ":" in str(df.iloc[1,0]) else str(df.iloc[1, 0])
        clean_oid = raw_oid.replace("-", "")
        odate = smart_date(clean_oid[:8])
        
        vendor_raw = ""
        for i in range(len(df)):
            if "수신" in str(df.iloc[i, 0]): 
                vendor_raw = str(df.iloc[i, 0]).split(":")[-1].strip()
                break
                
        v_master = pd.read_sql("SELECT 거래처명, 기본유형 FROM vendors", conn)
        v_master['clean'] = v_master['거래처명'].apply(lambda x: re.sub(r'\s+', '', str(x)).lower())
        match = v_master[v_master['clean'] == re.sub(r'\s+', '', vendor_raw).lower()]
        
        if match.empty: 
            return False, "미등록 업체: [" + vendor_raw + "]"
            
        v_fixed = match.iloc[0]['거래처명']
        v_type_row = pd.read_sql(f"SELECT 기본유형 FROM vendors WHERE 거래처명='{v_fixed}'", conn)
        v_type = v_type_row.iloc[0]['기본유형'] if not v_type_row.empty else "기타"
        
        f6 = str(df.iloc[5, 5]) if len(df) > 5 else ""
        curr = "USD" if "USD" in f6 else ("CNY" if any(x in f6 for x in ["중국", "CNY"]) else "한화")
        p_col = 1 if curr == "한화" else 2
        prods = df.iloc[6:, p_col].dropna().astype(str).tolist()
        prod_n = (prods[0].split("[")[0].strip() + (f" 외 {len(prods)-1}건" if len(prods)>1 else "")) if prods else "품목미상"
        
        l_idx = df.iloc[:, 5].last_valid_index()
        total = to_float(df.iloc[l_idx, 5]) if curr != "한화" and l_idx else to_float(str(df.iloc[4, 0]).split(":")[-1])
        
        conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,0)", 
                     (raw_oid, odate, "", v_fixed, prod_n, v_type, curr, total))
        conn.commit()
        return True, None
    except Exception as e: 
        return False, "분석 오류: " + str(e)

# ==========================================
# 4. 메인 UI (디테일: 합치기 및 필터 레이아웃)
# ==========================================
# [디테일 1] 입금 입력과 업로드 합침
tabs = st.tabs(["입금 등록", "발주서 등록", "상세내역 및 정산", "거래처 관리", "환율 관리"])

# ------------------------------------------
# [Tab 0] 입금 등록 (수기 + 엑셀 통합)
# ------------------------------------------
with tabs[0]:
    st.header("입금 내역 등록")
    v_data_t0 = pd.read_sql("SELECT * FROM vendors", conn)
    o_active_t0 = pd.read_sql("SELECT 발주번호 FROM orders WHERE 마감여부=0", conn)
    
    col_man, col_csv = st.columns([1.5, 1])
    
    with col_man:
        st.subheader("수기 입력")
        with st.form("pay_manual_v136", clear_on_submit=True):
            c1, c2 = st.columns(2)
            p_oid = c1.selectbox("발주번호 연동", ["없음"] + list(o_active_t0['발주번호']))
            p_date = c2.date_input("입금일", value=datetime.now())
            
            c3, c4, c5 = st.columns(3)
            p_vn = c3.selectbox("거래처명", ["선택"] + list(v_data_t0['거래처명']))
            p_ct = c4.selectbox("유형", CATEGORIES)
            p_pr = c5.text_input("상품명")
            
            c6, c7, c8 = st.columns(3)
            p_dep = c6.number_input("실입금액", format="%.2f")
            p_pre = c7.number_input("선급금액", format="%.2f")
            p_cur = c8.selectbox("통화", ["한화", "USD", "CNY"])
            
            p_memo = st.text_input("메모(송금사유)")
            
            if st.form_submit_button("입금 내역 저장"):
                if p_vn == "선택": st.error("거래처를 선택하세요.")
                else:
                    vi = v_data_t0[v_data_t0['거래처명']==p_vn].iloc[0]
                    conn.execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''', 
                                 (to_str(p_oid) if p_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), p_ct, p_vn, p_pr, p_cur, p_dep, p_pre, p_memo, 0.0, vi['은행'], vi['계좌번호'], vi['예금주']))
                    conn.commit(); st.success("저장 완료!"); st.rerun()

    with col_csv:
        st.subheader("엑셀 일괄 업로드")
        pay_tmp = pd.DataFrame(columns=["발주번호", "거래처", "유형", "상품명", "입금일", "실입금액", "선급금액", "송금사유"])
        st.download_button("양식 다운로드", pay_tmp.to_csv(index=False).encode('utf-8-sig'), "pay_template.csv")
        f_p = st.file_uploader("CSV 선택", type=['csv'], key=f"p_up_{st.session_state.pay_up_key}")
        if f_p and st.button("데이터 일괄 저장 실행"):
            try:
                df_p = pd.read_csv(f_p); df_p.columns = [str(c).strip().replace('\ufeff', '') for c in df_p.columns]
                for _, r in df_p.iterrows():
                    oid, vn_raw = to_str(r.get('발주번호')), to_str(r.get('거래처'))
                    if not vn_raw and not oid: continue
                    pd_s = smart_date(r.get('입금일'))
                    if oid and not o_active_t0[o_active_t0['발주번호'] == oid].empty:
                        info = pd.read_sql(f"SELECT * FROM orders WHERE 발주번호='{oid}'", conn).iloc[0]
                        vn, pc, pp, cur = info['거래처명'], info['유형'], info['상품명'], info['통화']
                    else: vn, pc, pp, cur = vn_raw, to_str(r.get('유형')) or "사입", to_str(r.get('상품명')), "한화"
                    vi = v_data_t0[v_data_t0['거래처명'].str.lower() == vn.lower()]
                    dep, pre = to_float(r.get('실입금액')), to_float(r.get('선급금액'))
                    b_b = vi.iloc[0]['은행'] if not vi.empty else ""; b_a = vi.iloc[0]['계좌번호'] if not vi.empty else ""; b_h = vi.iloc[0]['예금주'] if not vi.empty else ""
                    conn.execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''', (oid if oid else None, pd_s, pc, vn, pp, cur, dep, pre, to_str(r.get('송금사유')), 0.0, b_b, b_a, b_h))
                conn.commit(); st.success("성공!"); st.session_state.pay_up_key += 1; st.rerun()
            except Exception as e: st.error(f"오류: {e}")

# ------------------------------------------
# [Tab 1] 발주서 등록 및 마감
# ------------------------------------------
with tabs[1]:
    st.header("발주서 등록 및 마감")
    c1, c2 = st.columns([1, 1.5])
    with c1:
        st.subheader("발주 등록")
        of_list = st.file_uploader("xlsx 파일 선택", type=['xlsx'], accept_multiple_files=True, key=f"ord_{st.session_state.order_up_key}")
        if of_list and st.button("파일 일괄 등록 실행"):
            sc, errs = 0, []
            for of in of_list:
                ok, msg = process_ecount_v136(of)
                if ok: sc += 1
                else: errs.append(f"[{of.name}] {msg}")
            for e in errs: st.warning(e)
            if sc > 0: st.success(f"총 {sc}건 등록 완료!"); st.session_state.order_up_key += 1; st.rerun()
        
        st.divider()
        with st.form("ord_manual_v136"):
            st.subheader("수기 발주 등록")
            mi, m_step = st.text_input("발주번호"), st.text_input("발주차수")
            md, mv = st.date_input("발주일"), st.selectbox("거래처", ["선택"] + list(v_data_t0['거래처명']))
            mp, mt, m_cur = st.text_input("상품명"), st.number_input("금액", format="%.2f"), st.selectbox("통화", ["한화", "USD", "CNY"])
            if st.form_submit_button("저장"):
                existing = pd.read_sql(f"SELECT 발주번호 FROM orders WHERE 발주번호='{mi}'", conn)
                if not mi: st.error("발주번호 필수")
                elif not existing.empty: st.error("이미 존재하는 발주번호")
                elif mv == "선택": st.error("거래처 선택 필수")
                else:
                    vt = pd.read_sql(f"SELECT 기본유형 FROM vendors WHERE 거래처명='{mv}'", conn).iloc[0]['기본유형']
                    conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,0)", (mi, md.strftime("%Y-%m-%d"), m_step, mv, mp, vt, m_cur, mt))
                    conn.commit(); st.rerun()
    with c2:
        o_data = pd.read_sql("SELECT * FROM orders", conn)
        if not o_data.empty:
            st.subheader("발주 리스트 및 마감관리")
            # [디테일 2] 발주총액 천단위 쉼표 소수점 2자리
            ev_o = st.data_editor(o_data.sort_values('발주일', ascending=False), hide_index=True, use_container_width=True, disabled=["발주번호"], 
                                 column_config={"발주총액": st.column_config.NumberColumn("발주총액", format="#,##0.00"), "마감여부": st.column_config.CheckboxColumn("마감")})
            if st.button("정보 소급 적용"):
                for _, r in ev_o.iterrows():
                    conn.execute("UPDATE orders SET 발주일=?, 발주차수=?, 거래처명=?, 상품명=?, 유형=?, 통화=?, 발주총액=?, 마감여부=? WHERE 발주번호=?", 
                                 (r['발주일'], r['발주차수'], r['거래처명'], r['상품명'], r['유형'], r['통화'], r['발주총액'], int(r['마감여부']), r['발주번호']))
                    conn.execute("UPDATE payments SET 거래처명=?, 유형=?, 상품명=?, 통화=? WHERE 발주번호=?", (r['거래처명'], r['유형'], r['상품명'], r['통화'], r['발주번호']))
                conn.commit(); st.rerun()

# ------------------------------------------
# [Tab 2] 상세내역 및 통합 정산
# ------------------------------------------
with tabs[2]:
    st.header("상세 내역 및 통합 정산")
    p_all = pd.read_sql("SELECT * FROM payments", conn); o_all = pd.read_sql("SELECT * FROM orders", conn)
    ex_db = pd.read_sql("SELECT * FROM exchange_rates ORDER BY 날짜 ASC", conn)
    
    # [디테일 3] 필터 왼쪽, 요약 오른쪽 배치
    col_fil, col_sum_top = st.columns([1, 1.2])
    
    with col_fil:
        st.subheader("필터 및 검색")
        p_all['dt'] = pd.to_datetime(p_all['입금일'])
        f_c1, f_c2 = st.columns(2)
        y_val = f_c1.selectbox("기준 연도", sorted(p_all['dt'].dt.year.unique(), reverse=True))
        m_val = f_c2.selectbox("기준 월", ["전체"] + sorted(list(p_all[p_all['dt'].dt.year==y_val]['dt'].dt.month.unique())))
        
        f_c3, f_c4 = st.columns(2)
        cat_filter = f_c3.selectbox("유형 선택", ["전체 유형"] + CATEGORIES)
        # [디테일 3] 업체명 대소문자 무관 검색
        search_txt = f_c4.text_input("업체/상품 검색 (대소문자 무관)")
        
        fil_p = p_all[p_all['dt'].dt.year == y_val].copy()
        if m_val != "전체": fil_p = fil_p[fil_p['dt'].dt.month == m_val]
        if cat_filter != "전체 유형": fil_p = fil_p[fil_p['유형'] == cat_filter]
        if search_txt: fil_p = fil_p[fil_p['거래처명'].str.contains(search_txt, case=False, na=False) | fil_p['상품명'].str.contains(search_txt, case=False, na=False)]
        fil_p = pd.merge(fil_p, o_all[['발주번호', '발주차수']], on='발주번호', how='left')

    with col_sum_top:
        st.subheader("유형별 요약") # [디테일 3] 제목 변경
        if not fil_p.empty:
            cat_sum = fil_p.groupby('유형').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
            st.table(cat_sum.style.format({'실입금액': '{:,.2f}', '선급금액': '{:,.2f}'}))

    st.divider()
    st.subheader("발주번호별 정산 및 미수금 현황")
    # [디테일 3] 발주번호가 있는 건만 표시
    p_agg = p_all[p_all['발주번호'].notnull() & (p_all['발주번호'] != "")].groupby('발주번호').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
    sum_df = pd.merge(o_all, p_agg, on='발주번호', how='left').fillna(0)
    sum_df['잔액'] = sum_df['발주총액'] - sum_df['실입금액']
    sum_df['상태'] = sum_df['마감여부'].apply(lambda x: "✅ 마감" if x == 1 else "⏳ 진행")
    sum_df = sum_df.sort_values(['마감여부', '발주번호'], ascending=[True, False])
    
    disp_sum = sum_df[['발주번호', '발주차수', '상태', '거래처명', '상품명', '발주총액', '실입금액', '선급금액', '잔액', '통화']]
    
    # [디테일 3] 선급금액 > 0 빨간색, 잔액 > 0 파란색
    def style_settlement(row):
        styles = [''] * len(row)
        if row['선급금액'] > 0: styles[disp_sum.columns.get_loc('선급금액')] = 'color: red; font-weight: bold'
        if row['잔액'] > 0: styles[disp_sum.columns.get_loc('잔액')] = 'color: blue; font-weight: bold'
        if row['상태'] == '✅ 마감': styles = ['background-color: #f9f9f9; color: #ccc'] * len(row)
        return styles

    st.dataframe(disp_sum.style.apply(style_settlement, axis=1).format({'발주총액':'{:,.2f}','실입금액':'{:,.2f}','선급금액':'{:,.2f}','잔액':'{:,.2f}'}), 
                 use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("상세 리스트 편집 및 상세")
    ex_db['ym'] = pd.to_datetime(ex_db['날짜']).dt.strftime('%Y-%m')
    m_rates_df = ex_db.groupby('ym').agg({'usd': lambda x: x[x>0].mean(), 'cny': lambda x: x[x>0].mean()}).fillna(0)
    
    # [디테일 3] 직전 월 환율 추적 로직 (에러 방지용 보강)
    def calc_krw_final(row):
        if row['통화'] == '한화': return row['실입금액']
        ym, curr = str(row['입금일'])[:7], row['통화'].lower()
        if ym in m_rates_df.index and m_rates_df.loc[ym, curr] > 0: 
            rate = m_rates_df.loc[ym, curr]
        else:
            past_rates = m_rates_df[m_rates_df.index < ym]
            if not past_rates.empty and past_rates[curr].sum() > 0:
                rate = past_rates[past_rates[curr] > 0].iloc[-1][curr]
            else:
                rate = 1350.0 if row['통화'] == 'USD' else 190.0
        return row['실입금액'] * rate

    fil_p_m = pd.merge(fil_p, o_all[['발주번호', '발주총액']], on='발주번호', how='left').fillna(0)
    if not fil_p_m.empty:
        fil_p_m['예상환산액'] = fil_p_m.apply(calc_krw_final, axis=1)
    else:
        fil_p_m['예상환산액'] = pd.Series(dtype='float64')

    # [디테일 3] 입금일을 발주총액과 실입금액 사이로 & 쉼표/소수점 포맷
    f_cols = ['id', '유형', '발주번호', '거래처명', '상품명', '통화', '발주총액', '입금일', '실입금액', '선급금액', '예상환산액', '메모']
    ed_p = st.data_editor(fil_p_m[f_cols].sort_values('입금일', ascending=False), hide_index=True, use_container_width=True,
                         column_config={c: st.column_config.NumberColumn(c, format="#,##0.00") for c in ['발주총액','실입금액','선급금액','예상환산액']})
    
    b_c1, b_c2 = st.columns([1, 4])
    with b_c1:
        if st.button("상세 수정 저장"):
            for _, r in ed_p.iterrows():
                conn.execute("UPDATE payments SET 유형=?, 거래처명=?, 상품명=?, 실입금액=?, 선급금액=?, 메모=? WHERE id=?", 
                             (r['유형'], r['거래처명'], r['상품명'], r['실입금액'], r['선급금액'], r['메모'], r['id']))
            conn.commit(); st.rerun()
    with b_c2:
        with st.form("del_v136"):
            d_id = st.number_input("삭제 ID", min_value=0, step=1)
            if st.form_submit_button("입금 내역 삭제"):
                conn.execute(f"DELETE FROM payments WHERE id={d_id}"); conn.commit(); st.rerun()

    st.divider()
    st.markdown(f"### 현재 검색 목록 합계 ({len(fil_p_m)}건)")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("총 실입금 환산액", f"{fil_p_m['예상환산액'].sum():,.2f}")
    m2.metric("KRW 실입금 합계", f"{fil_p_m[fil_p_m['통화']=='한화']['실입금액'].sum():,.2f}")
    m3.metric("USD 실입금 합계", f"{fil_p_m[fil_p_m['통화']=='USD']['실입금액'].sum():,.2f}")
    m4.metric("CNY 실입금 합계", f"{fil_p_m[fil_p_m['통화']=='CNY']['실입금액'].sum():,.2f}")

# ------------------------------------------
# [Tab 3] 거래처 관리
# ------------------------------------------
with tabs[3]:
    st.header("거래처 정보 관리")
    # [디테일 4] 거래처 검색 추가
    v_search = st.text_input("거래처명 검색 (계좌정보 확인용)")
    v_data_full = pd.read_sql("SELECT * FROM vendors", conn)
    if v_search:
        v_data_full = v_data_full[v_data_full['거래처명'].str.contains(v_search, case=False, na=False)]
    
    st.data_editor(v_data_full, hide_index=True, use_container_width=True)
    
    with st.expander("신규 거래처 등록"):
        with st.form("new_v_form"):
            n1, n2, n3 = st.columns(3); v_n = n1.text_input("거래처명"); v_t = n2.selectbox("기본유형", CATEGORIES); v_b = n3.text_input("은행")
            n4, n5 = st.columns(2); v_acc = n4.text_input("계좌번호"); v_h = n5.text_input("예금주")
            if st.form_submit_button("신규 거래처 저장"):
                if v_n: conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (v_n, v_b, v_acc, v_h, v_t)); conn.commit(); st.rerun()

# ------------------------------------------
# [Tab 4] 환율 관리 (차트 서식 및 YoY/MoM 리포트 복구)
# ------------------------------------------
with tabs[4]:
    st.header("환율 관리 및 분석")
    cu1, cu2 = st.columns(2)
    with cu1:
        f_u_f = st.file_uploader("USD 환율 CSV", type=['csv'], key="u_csv_tab_v136")
        if f_u_f and st.button("USD 환율 업데이트"): process_exchange_csv(f_u_f, "USD"); st.rerun()
    with cu2:
        f_c_f = st.file_uploader("CNY 환율 CSV", type=['csv'], key="c_csv_tab_v136")
        if f_c_f and st.button("CNY 환율 업데이트"): process_exchange_csv(f_c_f, "CNY"); st.rerun()
    
    ex_db_p = pd.read_sql("SELECT * FROM exchange_rates ORDER BY 날짜 ASC", conn)
    if not ex_db_p.empty:
        # [디테일 5] 날짜 형식 '25.1월' 변경
        ex_db_p['ym_label'] = pd.to_datetime(ex_db_p['날짜']).dt.strftime('%y.%-m월')
        m_mean = ex_db_p.groupby('ym_label', sort=False).agg({'usd': lambda x: x[x>0].mean(), 'cny': lambda x: x[x>0].mean()}).reset_index()
        
        c1, c2 = st.columns(2)
        with c1:
            fig_u = go.Figure(go.Scatter(x=m_mean['ym_label'], y=m_mean['usd'], mode='lines+markers', name='USD'))
            # [디테일 5] USD 세로축 20단위 고정
            fig_u.update_layout(title="USD 평균 추이", yaxis=dict(dtick=20, tickformat=",.2f"), template="plotly_white", height=400)
            st.plotly_chart(fig_u, use_container_width=True)
        with c2:
            fig_c = go.Figure(go.Scatter(x=m_mean['ym_label'], y=m_mean['cny'], mode='lines+markers', name='CNY', line=dict(color='orange')))
            # [디테일 5] CNY 세로축 2단위 고정
            fig_c.update_layout(title="CNY 평균 추이", yaxis=dict(dtick=2, tickformat=",.2f"), template="plotly_white", height=400)
            st.plotly_chart(fig_c, use_container_width=True)
        
        st.subheader("연도별 평균환율") # [디테일 5] 제목 변경
        
        # --- YoY, MoM 리포트 로직 완전 복구 ---
        ex_raw = pd.read_sql("SELECT * FROM exchange_rates ORDER BY 날짜 ASC", conn)
        ex_raw['ym'] = pd.to_datetime(ex_raw['날짜']).dt.strftime('%Y-%m')
        m_stats = ex_raw.groupby('ym').agg({'usd': lambda x: x[x>0].mean(), 'cny': lambda x: x[x>0].mean()}).reset_index().fillna(0)
        m_stats['year'] = m_stats['ym'].str[:4].astype(int); m_stats['month'] = m_stats['ym'].str[5:].astype(int)

        def get_full_report(df, col):
            ys = sorted(df['year'].unique(), reverse=True)
            if not ys: return pd.DataFrame()
            cy, py = ys[0], ys[1] if len(ys) > 1 else None
            res = pd.DataFrame({'월': range(1, 13)})
            res[f'{cy}년'] = res['월'].map(df[df['year'] == cy].set_index('month')[col])
            if py: res[f'{py}년'] = res['월'].map(df[df['year'] == py].set_index('month')[col])
            if py:
                def yoy_fn(r):
                    v1, v2 = r[f'{cy}년'], r[f'{py}년']
                    if pd.notnull(v1) and pd.notnull(v2) and v2 > 0: return f"{(v1-v2):+.2f}({((v1-v2)/v2*100):+.1f}%)"
                    return "-"
                res['전년비(YoY)'] = res.apply(yoy_fn, axis=1)
            df_s = df.sort_values('ym').copy(); df_s['d'], df_s['p'] = df_s[col].diff(), df_s[col].shift(1)
            def mom_fn(r):
                m = r['월']; m_map = df_s[df_s['year'] == cy].set_index('month')
                if m in m_map.index:
                    d, v = m_map.loc[m, 'd'], m_map.loc[m, 'p']
                    if pd.notnull(d) and pd.notnull(v) and v > 0: return f"{d:+.2f}({(d/v*100):+.1f}%)"
                return "-"
            res['전월비(MoM)'] = res.apply(mom_fn, axis=1)
            return res[res[f'{cy}년'].notnull() | (res[f'{py}년'].notnull() if py else False)].reset_index(drop=True)

        rc1, rc2 = st.columns(2)
        with rc1: 
            st.write("**USD 평균 환율**")
            st.table(get_full_report(m_stats, 'usd').style.format(precision=2, thousands=","))
        with rc2: 
            st.write("**CNY 평균 환율**")
            st.table(get_full_report(m_stats, 'cny').style.format(precision=2, thousands=","))
    else:
        st.warning("환율 데이터가 없습니다.")