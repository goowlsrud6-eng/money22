import streamlit as st
import pandas as pd
import sqlite3
import os
import shutil
import re
from datetime import datetime, timedelta
import plotly.graph_objects as go
import urllib.request

# ==============================================================================
# 1. 백업 및 데이터베이스 초기화 (v136 원본 로직 100% 보존)
# ==============================================================================
def run_backup():
    """매일 첫 접속 시 데이터베이스 백업 생성 기능을 수행합니다."""
    if not os.path.exists('backups'):
        os.makedirs('backups')
    db_file = 'finance_final_v136.db'
    today_str = datetime.now().strftime('%Y%m%d')
    backup_file = f"backups/backup_{today_str}.db"
    
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

# 전체 레이아웃 설정
st.set_page_config(page_title="자금 관리 시스템 v136_Final", layout="wide")
run_backup()

@st.cache_resource
def get_db_connection():
    """데이터베이스 연결 및 테이블 스키마 생성"""
    conn = sqlite3.connect('finance_final_v136.db', check_same_thread=False)
    c = conn.cursor()
    # 거래처 마스터 정보
    c.execute('CREATE TABLE IF NOT EXISTS vendors (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT, 기본유형 TEXT)')
    # 발주 정보
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (발주번호 TEXT PRIMARY KEY, 발주일 TEXT, 발주차수 TEXT, 거래처명 TEXT, 
                 상품명 TEXT, 유형 TEXT, 통화 TEXT, 발주총액 REAL, 마감여부 INTEGER DEFAULT 0)''')
    # 입금 내역 정보 (id에 AUTOINCREMENT를 제거하여 수동 관리 가능하게 함 - 수정 5 반영용)
    c.execute('''CREATE TABLE IF NOT EXISTS payments 
                 (id INTEGER PRIMARY KEY, 발주번호 TEXT, 입금일 TEXT, 
                 유형 TEXT, 거래처명 TEXT, 상품명 TEXT, 통화 TEXT,
                 실입금액 REAL, 선급금액 REAL, 메모 TEXT, 한화환산액 REAL,
                 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT)''')
    # 환율 관리 정보
    c.execute('CREATE TABLE IF NOT EXISTS exchange_rates (날짜 TEXT PRIMARY KEY, usd REAL, cny REAL)')
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비", "원단비", "기타"]

# 세션 상태 관리
if 'order_up_key' not in st.session_state: 
    st.session_state.order_up_key = 0
if 'pay_up_key' not in st.session_state: 
    st.session_state.pay_up_key = 1000

# ==============================================================================
# 2. 유틸리티 함수 (숫자, 날짜 처리 보강 및 ID 재사용)
# ==============================================================================
def get_next_available_id():
    """[수정 5] 삭제되어 비어있는 ID 중 가장 작은 번호를 찾음"""
    ids = pd.read_sql("SELECT id FROM payments", conn)['id'].tolist()
    if not ids: return 1
    ids.sort()
    for i in range(1, (max(ids) if ids else 0) + 2):
        if i not in ids: return i
    return 1

def to_float(val):
    """문자열 금액을 실수로 변환하며 쉼표 등을 제거합니다."""
    try:
        if val is None or pd.isna(val) or str(val).strip() == "": 
            return 0.0
        return float(str(val).replace(',', '').strip())
    except: 
        return 0.0

def to_str(val):
    """결측치를 빈 문자열로 처리하여 반환합니다."""
    if val is None or pd.isna(val): 
        return ""
    s = str(val).strip()
    if s.lower() in ["nan", "none", ""]:
        return ""
    return s

def smart_date(date_val):
    """[수정 1] 3/14 같은 형식을 YYYY-03-14로 변환하는 로직 추가"""
    try:
        if pd.isna(date_val) or str(date_val).strip() == "":
            return datetime.now().strftime("%Y-%m-%d")
        if isinstance(date_val, (datetime, pd.Timestamp)):
            return date_val.strftime("%Y-%m-%d")
        
        ds = str(date_val).strip()
        # "3/14" 또는 "03/14" 형식 감지
        if re.match(r'^\d{1,2}/\d{1,2}$', ds) or re.match(r'^\d{1,2}-\d{1,2}$', ds):
            curr_year = datetime.now().year
            ds = f"{curr_year}-{ds.replace('/', '-')}"
            
        ds = ds.rstrip('.').replace(" ", "").replace(".", "-").replace("/", "-")
        return pd.to_datetime(ds).strftime("%Y-%m-%d")
    except:
        return datetime.now().strftime("%Y-%m-%d")

# ==============================================================================
# 3. 데이터 엔진 (CSV 및 엑셀 처리 로직)
# ==============================================================================
def process_exchange_csv(file, currency_type):
    """환율 CSV 분석 및 DB 반영"""
    try:
        df = pd.read_csv(file)
        df.columns = [c.strip().replace('\ufeff', '') for c in df.columns]
        for index, row in df.iterrows():
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
    """이카운트 발주서 엑셀 파일을 정밀 분석하여 DB에 자동 저장합니다."""
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
            return False, f"미등록 업체: [{vendor_raw}]"
            
        v_fixed = match.iloc[0]['거래처명']
        v_type = match.iloc[0]['기본유형']
        
        f6_val = str(df.iloc[5, 5]) if len(df) > 5 else ""
        if "USD" in f6_val: curr = "USD"
        elif any(x in f6_val for x in ["중국", "CNY"]): curr = "CNY"
        else: curr = "한화"
        
        p_col = 1 if curr == "한화" else 2
        prods = df.iloc[6:, p_col].dropna().astype(str).tolist()
        if prods:
            prod_n = prods[0].split("[")[0].strip()
            if len(prods) > 1: prod_n = prod_n + f" 외 {len(prods)-1}건"
        else: prod_n = "품목미상"
        
        l_idx = df.iloc[:, 5].last_valid_index()
        total = to_float(df.iloc[l_idx, 5]) if curr != "한화" and l_idx else to_float(str(df.iloc[4, 0]).split(":")[-1])
        
        conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,0)", (raw_oid, odate, "", v_fixed, prod_n, v_type, curr, total))
        conn.commit()
        return True, None
    except Exception as e:
        return False, str(e)

# ==============================================================================
# 4. 메인 UI 및 탭별 로직
# ==============================================================================
tabs = st.tabs(["입금 등록", "발주서 등록", "상세내역 및 정산", "거래처 관리", "환율 분석"])

# --- [Tab 0] 입금 내역 등록 (상품명 자동 연동 수정) ---
with tabs[0]:
    st.header("입금 내역 등록 및 관리")
    v_data_t0 = pd.read_sql("SELECT * FROM vendors", conn)
    o_active_t0 = pd.read_sql("SELECT 발주번호, 상품명 FROM orders WHERE 마감여부=0", conn)
    
    col_input, col_excel = st.columns([1.5, 1])
    
    with col_input:
        st.subheader("1. 수기 직접 입력")
        with st.form("manual_pay_form_v136", clear_on_submit=True):
            r1c1, r1c2 = st.columns(2)
            p_oid = r1c1.selectbox("발주번호 연동", ["없음"] + list(o_active_t0['발주번호']))
            p_date = r1c2.date_input("입금일자", value=datetime.now())
            
            # [수정 3] 발주번호 선택 시 상품명 연동
            auto_prod_val = ""
            if p_oid != "없음":
                auto_prod_val = o_active_t0[o_active_t0['발주번호'] == p_oid]['상품명'].values[0]
            
            r2c1, r2c2, r2c3 = st.columns(3)
            p_vn = r2c1.selectbox("거래처 선택", ["선택"] + list(v_data_t0['거래처명']))
            p_ct = r2c2.selectbox("유형 분류", CATEGORIES)
            p_pr = r2c3.text_input("상품명(품목)", value=auto_prod_val)
            
            r3c1, r3c2, r3c3 = st.columns(3)
            p_dep = r3c1.number_input("실입금액", format="%.2f")
            p_pre = r3c2.number_input("선급금액", format="%.2f")
            p_cur = r3c3.selectbox("거래통화", ["한화", "USD", "CNY"])
            
            p_memo = st.text_input("비고 (송금 사유 등)")
            
            if st.form_submit_button("입금 내역 저장"):
                if p_vn == "선택":
                    st.error("거래처를 선택해야 합니다.")
                else:
                    new_id = get_next_available_id() # [수정 5]
                    vi = v_data_t0[v_data_t0['거래처명']==p_vn].iloc[0]
                    conn.execute('''
                        INSERT INTO payments (id, 발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    ''', (new_id, to_str(p_oid) if p_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), p_ct, p_vn, p_pr, p_cur, p_dep, p_pre, p_memo, 0.0, vi['은행'], vi['계좌번호'], vi['예금주']))
                    conn.commit()
                    st.success(f"저장되었습니다. (ID: {new_id})")
                    st.rerun()

    with col_excel:
        st.subheader("2. 엑셀 일괄 업로드")
        csv_template = pd.DataFrame(columns=["발주번호", "거래처", "유형", "상품명", "입금일", "실입금액", "선급금액", "송금사유"])
        st.download_button("입금 업로드 양식(CSV) 다운로드", csv_template.to_csv(index=False).encode('utf-8-sig'), "payment_template.csv")
        
        uploaded_pay = st.file_uploader("CSV 파일 선택", type=['csv'], key=f"pay_csv_{st.session_state.pay_up_key}")
        if uploaded_pay and st.button("파일 일괄 저장 실행"):
            try:
                df_up = pd.read_csv(uploaded_pay)
                df_up.columns = [str(c).strip().replace('\ufeff', '') for c in df_up.columns]
                # [수정 3] 연동을 위한 발주 데이터 사전 로드
                o_master_data = pd.read_sql("SELECT 발주번호, 상품명, 거래처명, 유형, 통화 FROM orders", conn)
                
                for _, r in df_up.iterrows():
                    oid_v = to_str(r.get('발주번호'))
                    vn_v = to_str(r.get('거래처'))
                    if not vn_v and not oid_v: continue
                    pd_date = smart_date(r.get('입금일')) # [수정 1]
                    
                    if oid_v and not o_master_data[o_master_data['발주번호'] == oid_v].empty:
                        ord_i = o_master_data[o_master_data['발주번호'] == oid_v].iloc[0]
                        vn_f, pc_f, pp_f, cur_f = ord_i['거래처명'], ord_i['유형'], ord_i['상품명'], ord_i['통화']
                    else:
                        vn_f, pc_f, pp_f, cur_f = vn_v, to_str(r.get('유형')) or "사입", to_str(r.get('상품명')), "한화"
                    
                    vi_match = v_data_t0[v_data_t0['거래처명'].str.lower() == vn_f.lower()]
                    bank_n = vi_match.iloc[0]['은행'] if not vi_match.empty else ""
                    bank_a = vi_match.iloc[0]['계좌번호'] if not vi_match.empty else ""
                    bank_h = vi_match.iloc[0]['예금주'] if not vi_match.empty else ""
                    
                    new_id = get_next_available_id() # [수정 5]
                    conn.execute('''INSERT INTO payments (id, 발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', (new_id, oid_v if oid_v else None, pd_date, pc_f, vn_f, pp_f, cur_f, to_float(r.get('실입금액')), to_float(r.get('선급금액')), to_str(r.get('송금사유')), 0.0, bank_n, bank_a, bank_h))
                    conn.commit() # ID 채우기를 위해 매 행 커밋
                st.success("일괄 저장이 완료되었습니다.")
                st.session_state.pay_up_key += 1
                st.rerun()
            except Exception as e: st.error(f"오류: {e}")

# --- [Tab 1] 발주서 등록 및 관리 ---
with tabs[1]:
    st.header("발주서 등록 및 마감")
    c1, c2 = st.columns([1, 1.5])
    with c1:
        st.subheader("1. 발주 등록")
        ord_template = pd.DataFrame(columns=["발주번호", "발주일", "발주차수", "거래처명", "상품명", "금액", "통화"])
        st.download_button("수기 발주용 양식 받기", ord_template.to_csv(index=False).encode('utf-8-sig'), "order_template.csv")
        
        ord_files = st.file_uploader("발주서(xlsx) 선택 (다중 가능)", type=['xlsx'], accept_multiple_files=True, key=f"ord_f_{st.session_state.order_up_key}")
        if ord_files and st.button("발주서 일괄 분석 실행"):
            sc_cnt = 0
            for f in ord_files:
                ok, msg = process_ecount_v136(f)
                if ok: sc_cnt += 1
                else: st.warning(f"[{f.name}] {msg}")
            if sc_cnt > 0:
                st.success(f"{sc_cnt}건 등록 완료")
                st.session_state.order_up_key += 1
                st.rerun()
        st.divider()
        with st.form("manual_ord_v136"):
            st.write("**직접 발주 입력**")
            m_oid = st.text_input("발주번호")
            m_step = st.text_input("발주차수")
            m_date = st.date_input("발주일")
            m_vn = st.selectbox("거래처", ["선택"] + list(v_data_t0['거래처명']))
            m_prod = st.text_input("상품명")
            m_amt = st.number_input("금액", format="%.2f")
            m_curr = st.selectbox("통화", ["한화", "USD", "CNY"])
            if st.form_submit_button("저장"):
                chk = pd.read_sql(f"SELECT 발주번호 FROM orders WHERE 발주번호='{m_oid}'", conn)
                if not m_oid or not chk.empty or m_vn == "선택":
                    st.error("발주번호 중복 혹은 필수값 누락")
                else:
                    v_t = v_data_t0[v_data_t0['거래처명']==m_vn].iloc[0]['기본유형']
                    conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,0)", (m_oid, m_date.strftime("%Y-%m-%d"), m_step, m_vn, m_prod, v_t, m_curr, m_amt))
                    conn.commit()
                    st.rerun()

    with c2:
        st.subheader("2. 발주 리스트 및 마감 관리")
        show_all = st.checkbox("이미 마감된 발주서까지 모두 보기", value=True)
        o_all_data = pd.read_sql("SELECT * FROM orders", conn)
        if not o_all_data.empty:
            disp_data = o_all_data if show_all else o_all_data[o_all_data['마감여부'] == 0]
            ev_orders = st.data_editor(disp_data.sort_values('발주일', ascending=False), hide_index=True, use_container_width=True, disabled=["발주번호"], 
                                     column_config={"발주총액": st.column_config.NumberColumn("발주총액", format="%,.2f"), "마감여부": st.column_config.CheckboxColumn("마감")})
            if st.button("수정 정보 소급 적용 및 저장"):
                for _, r in ev_orders.iterrows():
                    conn.execute("UPDATE orders SET 발주일=?, 발주차수=?, 거래처명=?, 상품명=?, 유형=?, 통화=?, 발주총액=?, 마감여부=? WHERE 발주번호=?", 
                                 (r['발주일'], r['발주차수'], r['거래처명'], r['상품명'], r['유형'], r['통화'], r['발주총액'], int(r['마감여부']), r['발주번호']))
                    conn.execute("UPDATE payments SET 거래처명=?, 유형=?, 상품명=?, 통화=? WHERE 발주번호=?", (r['거래처명'], r['유형'], r['상품명'], r['통화'], r['발주번호']))
                conn.commit()
                st.rerun()
# ------------------------------------------------------------------------------
# [Tab 2] 상세내역 및 통합 정산 (수정 2: 입금일 수정 / 수정 4: 연도 범위 조회)
# ------------------------------------------------------------------------------
with tabs[2]:
    st.header("상세 내역 및 통합 정산")
    payments_all = pd.read_sql("SELECT * FROM payments", conn)
    orders_all = pd.read_sql("SELECT * FROM orders", conn)
    rates_db = pd.read_sql("SELECT * FROM exchange_rates ORDER BY 날짜 ASC", conn)
    
    col_fil, col_summary = st.columns([1, 1.2])
    
    with col_fil:
        st.subheader("조회 조건 설정")
        if not payments_all.empty:
            payments_all['dt'] = pd.to_datetime(payments_all['입금일'])
            all_years = sorted(payments_all['dt'].dt.year.unique())
            if not all_years: all_years = [datetime.now().year]
            
            # [수정 4] 조회 연도 범위를 선택할 수 있게 변경 (2025~2026 등)
            f_r1c1, f_r1c2 = st.columns(2)
            start_y = f_r1c1.selectbox("시작 연도", all_years, index=0)
            end_y = f_r1c2.selectbox("종료 연도", all_years, index=len(all_years)-1)
            
            target_month = st.selectbox("조회 월 (전체 선택 시 연간 조회)", ["전체"] + sorted(list(payments_all['dt'].dt.month.unique())))
            
            f_r2c1, f_r2c2 = st.columns(2)
            filter_cat = f_r2c1.selectbox("유형 필터", ["전체 유형"] + CATEGORIES)
            search_key = f_r2c2.text_input("업체/상품 검색 (대소문자 무관)")
            
            # 필터링 적용 (연도 범위 적용)
            filtered_df = payments_all[(payments_all['dt'].dt.year >= start_y) & (payments_all['dt'].dt.year <= end_y)].copy()
            
            if target_month != "전체":
                filtered_df = filtered_df[filtered_df['dt'].dt.month == target_month]
            if filter_cat != "전체 유형":
                filtered_df = filtered_df[filtered_df['유형'] == filter_cat]
            if search_key:
                filtered_df = filtered_df[filtered_df['거래처명'].str.contains(search_key, case=False, na=False) | 
                                          filtered_df['상품명'].str.contains(search_key, case=False, na=False)]
            
            if not filtered_df.empty:
                filtered_df = pd.merge(filtered_df, orders_all[['발주번호', '발주차수']], on='발주번호', how='left')
        else:
            st.info("데이터가 없습니다.")
            filtered_df = pd.DataFrame()

    with col_summary:
        st.subheader("유형별 요약")
        if not filtered_df.empty:
            summary_table = filtered_df.groupby('유형').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
            st.table(summary_table.style.format({'실입금액': '{:,.2f}', '선급금액': '{:,.2f}'}))

    st.divider()
    st.subheader("발주번호별 정산 및 미수금 현황")
    pay_agg = payments_all[payments_all['발주번호'].notnull() & (payments_all['발주번호'] != "")].groupby('발주번호').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
    settle_df = pd.merge(orders_all, pay_agg, on='발주번호', how='left').fillna(0)
    settle_df['잔액'] = settle_df['발주총액'] - settle_df['실입금액']
    settle_df['진행상태'] = settle_df['마감여부'].apply(lambda x: "✅ 마감" if x == 1 else "⏳ 진행")
    settle_df = settle_df.sort_values(['마감여부', '발주번호'], ascending=[True, False])
    
    view_settle = settle_df[['발주번호', '발주차수', '진행상태', '거래처명', '상품명', '발주총액', '실입금액', '선급금액', '잔액', '통화']]
    
    def color_settle_styles(row):
        styles = [''] * len(row)
        if row['선급금액'] > 0: styles[view_settle.columns.get_loc('선급금액')] = 'color: red; font-weight: bold'
        if row['잔액'] > 0: styles[view_settle.columns.get_loc('잔액')] = 'color: blue; font-weight: bold'
        if row['진행상태'] == '✅ 마감': styles = ['background-color: #f9f9f9; color: #bbbbbb'] * len(row)
        return styles

    st.dataframe(view_settle.style.apply(color_settle_styles, axis=1).format({'발주총액':'{:,.2f}','실입금액':'{:,.2f}','선급금액':'{:,.2f}','잔액':'{:,.2f}'}), use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("상세 리스트 편집 및 상세")
    
    rates_db['ym'] = pd.to_datetime(rates_db['날짜']).dt.strftime('%Y-%m')
    m_rates_stats = rates_db.groupby('ym').agg({'usd': lambda x: x[x>0].mean(), 'cny': lambda x: x[x>0].mean()}).fillna(0)
    
    def get_conversion_krw_v136(row):
        if row['통화'] == '한화': return row['실입금액']
        ym_key, curr_key = str(row['입금일'])[:7], row['통화'].lower()
        if ym_key in m_rates_stats.index and m_rates_stats.loc[ym_key, curr_key] > 0: 
            rate = m_rates_stats.loc[ym_key, curr_key]
        else:
            past_data = m_rates_stats[m_rates_stats.index < ym_key]
            rate = past_data.iloc[-1][curr_key] if not past_data.empty and past_data[curr_key].sum() > 0 else (1350.0 if row['통화'] == 'USD' else 190.0)
        return row['실입금액'] * rate

    if not filtered_df.empty:
        final_detail_df = pd.merge(filtered_df, orders_all[['발주번호', '발주총액']], on='발주번호', how='left').fillna(0)
        final_detail_df['예상환산액'] = final_detail_df.apply(get_conversion_krw_v136, axis=1)
        
        # [수정 2] 입금일 수기 수정 활성화 (TextColumn으로 설정)
        c_order = ['id', '유형', '발주번호', '거래처명', '상품명', '통화', '발주총액', '입금일', '실입금액', '선급금액', '예상환산액', '메모']
        edited_detail = st.data_editor(
            final_detail_df[c_order].sort_values('입금일', ascending=False), 
            hide_index=True, 
            use_container_width=True,
            column_config={
                "발주총액": st.column_config.NumberColumn("발주총액", format="%,.2f"),
                "실입금액": st.column_config.NumberColumn("실입금액", format="%,.2f"),
                "선급금액": st.column_config.NumberColumn("선급금액", format="%,.2f"),
                "예상환산액": st.column_config.NumberColumn("예상환산액", format="%,.2f"),
                "입금일": st.column_config.TextColumn("입금일") 
            }
        )
        
        bc1, bc2 = st.columns([1, 4])
        with bc1:
            if st.button("상세내역 수정 저장"):
                for _, r in edited_detail.iterrows():
                    fixed_date = smart_date(r['입금일'])
                    conn.execute('''UPDATE payments SET 입금일=?, 유형=?, 거래처명=?, 상품명=?, 실입금액=?, 선급금액=?, 메모=? WHERE id=?''', 
                                 (fixed_date, r['유형'], r['거래처명'], r['상품명'], r['실입금액'], r['선급금액'], r['메모'], r['id']))
                conn.commit(); st.success("수정사항이 반영되었습니다."); st.rerun()
        with bc2:
            with st.form("del_form_v136"):
                d_id = st.number_input("삭제 ID", min_value=0, step=1)
                if st.form_submit_button("입금 내역 삭제"):
                    conn.execute(f"DELETE FROM payments WHERE id={d_id}"); conn.commit(); st.rerun()

        st.divider()
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("조회범위 환산액", f"{final_detail_df['예상환산액'].sum():,.2f}")
        m2.metric("KRW 합계", f"{final_detail_df[final_detail_df['통화']=='한화']['실입금액'].sum():,.2f}")
        m3.metric("USD 합계", f"{final_detail_df[final_detail_df['통화']=='USD']['실입금액'].sum():,.2f}")
        m4.metric("CNY 합계", f"{final_detail_df[final_detail_df['통화']=='CNY']['실입금액'].sum():,.2f}")
    else:
        st.write("데이터가 없습니다.")

# ------------------------------------------------------------------------------
# [Tab 3] 거래처 관리
# ------------------------------------------------------------------------------
with tabs[3]:
    st.header("거래처 정보 관리")
    vc1, vc2 = st.columns([1.2, 0.8])
    with vc1:
        st.subheader("1. 검색 및 수정")
        v_sr = st.text_input("거래처명 검색 (대소문자 무관)")
        v_data = pd.read_sql("SELECT * FROM vendors", conn)
        if v_sr: v_data = v_data[v_data['거래처명'].str.contains(v_sr, case=False, na=False)]
        orig_v = v_data['거래처명'].tolist(); ev_v = st.data_editor(v_data, hide_index=True, use_container_width=True)
        if st.button("정보 동기화 저장"):
            for idx, r in ev_v.iterrows():
                old_n, new_n = orig_v[idx], r['거래처명']
                if old_n != new_n:
                    conn.execute(f"DELETE FROM vendors WHERE 거래처명='{old_n}'")
                    conn.execute("INSERT INTO vendors VALUES (?,?,?,?,?)", (new_n, r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
                    conn.execute("UPDATE orders SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_n, r['기본유형'], old_n))
                    conn.execute("UPDATE payments SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_n, r['기본유형'], old_n))
                else:
                    conn.execute("UPDATE vendors SET 은행=?, 계좌번호=?, 예금주=?, 기본유형=? WHERE 거래처명=?", (r['은행'], r['계좌번호'], r['예금주'], r['기본유형'], r['거래처명']))
            conn.commit(); st.rerun()
    with vc2:
        st.subheader("2. 등록 및 업로드")
        with st.form("nv_form_v136"):
            vn, vt, vb = st.text_input("거래처명"), st.selectbox("유형", CATEGORIES), st.text_input("은행")
            vac, vh = st.text_input("계좌번호"), st.text_input("예금주")
            if st.form_submit_button("거래처 저장"):
                if vn: conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (vn, vb, vac, vh, vt)); conn.commit(); st.rerun()
        st.divider()
        v_tmp = pd.DataFrame(columns=["거래처명", "은행", "계좌번호", "예금주", "기본유형"])
        st.download_button("양식 다운로드", v_tmp.to_csv(index=False).encode('utf-8-sig'), "vendor_template.csv")
        vf = st.file_uploader("거래처 CSV 선택", type=['csv'], key="v_csv_up")
        if vf and st.button("일괄 저장 실행"):
            v_up = pd.read_csv(vf)
            for _, r in v_up.iterrows():
                conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (r['거래처명'], r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
            conn.commit(); st.rerun()

# ------------------------------------------------------------------------------
# [Tab 4] 환율 관리
# ------------------------------------------------------------------------------
with tabs[4]:
    st.header("환율 관리 및 정밀 분석")
    cu1, cu2 = st.columns(2)
    with cu1:
        fu = st.file_uploader("USD CSV", type=['csv'], key="tab4_u")
        if fu and st.button("USD 업데이트"): process_exchange_csv(fu, "USD"); st.rerun()
    with cu2:
        fc = st.file_uploader("CNY CSV", type=['csv'], key="tab4_c")
        if fc and st.button("CNY 업데이트"): process_exchange_csv(fc, "CNY"); st.rerun()
    ex_db_p = pd.read_sql("SELECT * FROM exchange_rates ORDER BY 날짜 ASC", conn)
    if not ex_db_p.empty:
        ex_db_p['ym_l'] = pd.to_datetime(ex_db_p['날짜']).dt.strftime('%y.%-m월')
        m_mean = ex_db_p.groupby('ym_l', sort=False).agg({'usd': lambda x: x[x>0].mean(), 'cny': lambda x: x[x>0].mean()}).reset_index()
        cc1, cc2 = st.columns(2)
        with cc1:
            fig_u = go.Figure(go.Scatter(x=m_mean['ym_l'], y=m_mean['usd'], mode='lines+markers', name='USD'))
            fig_u.update_layout(title="USD 평균 추이", yaxis=dict(dtick=20, tickformat=",.2f"), template="plotly_white", height=400); st.plotly_chart(fig_u, use_container_width=True)
        with cc2:
            fig_c = go.Figure(go.Scatter(x=m_mean['ym_l'], y=m_mean['cny'], mode='lines+markers', name='CNY', line=dict(color='orange')))
            fig_c.update_layout(title="CNY 평균 추이", yaxis=dict(dtick=2, tickformat=",.2f"), template="plotly_white", height=400); st.plotly_chart(fig_c, use_container_width=True)
        st.subheader("연도별 평균환율 분석 (2025 -> 2026)")
        report_raw = pd.read_sql("SELECT * FROM exchange_rates ORDER BY 날짜 ASC", conn); report_raw['ym'] = pd.to_datetime(report_raw['날짜']).dt.strftime('%Y-%m')
        stats_df = report_raw.groupby('ym').agg({'usd': lambda x: x[x>0].mean(), 'cny': lambda x: x[x>0].mean()}).reset_index().fillna(0)
        stats_df['yr'], stats_df['mo'] = stats_df['ym'].str[:4].astype(int), stats_df['ym'].str[5:].astype(int)

        def get_yr_analysis_v136(df, col_nm):
            yrs = sorted(df['yr'].unique())
            if not yrs: return pd.DataFrame()
            res_df = pd.DataFrame({'월': [f"{i}월" for i in range(1, 13)]})
            for y in yrs:
                m_map = df[df['yr'] == y].set_index('mo')[col_nm]
                res_df[f'{y}년'] = res_df['월'].apply(lambda x: m_map.get(int(x.replace('월','')), 0))
            if len(yrs) >= 2:
                cy, py = yrs[-1], yrs[-2]
                res_df['전년비(YoY)'] = res_df.apply(lambda r: f"{(r[f'{cy}년']-r[f'{py}년']):+.2f}({((r[f'{cy}년']-r[f'{py}년'])/r[f'{py}년']*100):+.1f}%)" if r[f'{py}년']>0 and r[f'{cy}년']>0 else "-", axis=1)
            df_srt = df.sort_values('ym').copy(); df_srt['diff'], df_srt['prev'] = df_srt[col_nm].diff(), df_srt[col_nm].shift(1)
            res_df['전월비(MoM)'] = res_df['월'].apply(lambda x: f"{df_srt[(df_srt['yr']==yrs[-1]) & (df_srt['mo']==int(x.replace('월','')))]['diff'].values[0]:+.2f}" if not df_srt[(df_srt['yr']==yrs[-1]) & (df_srt['mo']==int(x.replace('월','')))].empty else "-")
            mask = (res_df[f'{yrs[0]}년'] > 0)
            for y in yrs[1:]: mask = mask | (res_df[f'{y}년'] > 0)
            return res_df[mask].reset_index(drop=True)

        rc1, rc2 = st.columns(2)
        with rc1: st.write("**USD 평균**"); st.table(get_yr_analysis_v136(stats_df, 'usd').style.format(precision=2, thousands=","))
        with rc2: st.write("**CNY 평균**"); st.table(get_yr_analysis_v136(stats_df, 'cny').style.format(precision=2, thousands=","))
    else: st.warning("데이터가 없습니다.")