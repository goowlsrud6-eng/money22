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
# 1. 백업 및 데이터베이스 초기화 (원본 구조 100% 유지)
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

# 페이지 설정
st.set_page_config(page_title="자금 관리 시스템 v136_Final", layout="wide")
run_backup()

@st.cache_resource
def get_db_connection():
    """테이블 스키마 생성 및 유지 로직"""
    conn = sqlite3.connect('finance_final_v136.db', check_same_thread=False)
    c = conn.cursor()
    # 거래처 정보 테이블
    c.execute('''CREATE TABLE IF NOT EXISTS vendors 
                 (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT, 기본유형 TEXT)''')
    # 발주 정보 테이블
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (발주번호 TEXT PRIMARY KEY, 발주일 TEXT, 발주차수 TEXT, 거래처명 TEXT, 
                 상품명 TEXT, 유형 TEXT, 통화 TEXT, 발주총액 REAL, 마감여부 INTEGER DEFAULT 0)''')
    # 입금 내역 테이블
    c.execute('''CREATE TABLE IF NOT EXISTS payments 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 발주번호 TEXT, 입금일 TEXT, 
                 유형 TEXT, 거래처명 TEXT, 상품명 TEXT, 통화 TEXT,
                 실입금액 REAL, 선급금액 REAL, 메모 TEXT, 한화환산액 REAL,
                 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT)''')
    # 환율 정보 테이블
    c.execute('CREATE TABLE IF NOT EXISTS exchange_rates (날짜 TEXT PRIMARY KEY, usd REAL, cny REAL)')
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비", "원단비", "기타"]

# 업로드 상태 유지를 위한 키 관리
if 'order_up_key' not in st.session_state: 
    st.session_state.order_up_key = 0
if 'pay_up_key' not in st.session_state: 
    st.session_state.pay_up_key = 1000

# ==============================================================================
# 2. 유틸리티 함수
# ==============================================================================
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
    if s.lower() in ["nan", "none", ""]:
        return ""
    return s

def smart_date(date_val):
    """다양한 날짜 형식을 시스템 표준(YYYY-MM-DD)으로 변환"""
    try:
        if pd.isna(date_val) or str(date_val).strip() == "":
            return datetime.now().strftime("%Y-%m-%d")
        if isinstance(date_val, (datetime, pd.Timestamp)):
            return date_val.strftime("%Y-%m-%d")
        ds = str(date_val).strip().rstrip('.')
        ds = ds.replace(" ", "").replace(".", "-").replace("/", "-")
        return pd.to_datetime(ds).strftime("%Y-%m-%d")
    except:
        return datetime.now().strftime("%Y-%m-%d")

# ==============================================================================
# 3. 데이터 처리 엔진
# ==============================================================================
def process_exchange_csv(file, currency_type):
    """환율 CSV 파일 처리"""
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
    """이카운트 발주서 엑셀 분석 및 저장"""
    try:
        df = pd.read_excel(file, header=None)
        # 발주번호 추출
        raw_oid = str(df.iloc[1, 0]).split(":")[-1].strip() if ":" in str(df.iloc[1,0]) else str(df.iloc[1, 0])
        clean_oid = raw_oid.replace("-", "")
        odate = smart_date(clean_oid[:8])
        
        # 거래처명 추출 (수신 항목 탐색)
        vendor_raw = ""
        for i in range(len(df)):
            if "수신" in str(df.iloc[i, 0]): 
                vendor_raw = str(df.iloc[i, 0]).split(":")[-1].strip()
                break
        
        # 거래처 마스터와 매칭 (대소문자 및 공백 제거)
        v_master = pd.read_sql("SELECT 거래처명, 기본유형 FROM vendors", conn)
        v_master['clean'] = v_master['거래처명'].apply(lambda x: re.sub(r'\s+', '', str(x)).lower())
        match = v_master[v_master['clean'] == re.sub(r'\s+', '', vendor_raw).lower()]
        
        if match.empty: 
            return False, f"미등록 업체입니다: [{vendor_raw}]"
            
        v_fixed = match.iloc[0]['거래처명']
        v_type = match.iloc[0]['기본유형']
        
        # 통화 및 품목 정보 추출
        f6_val = str(df.iloc[5, 5]) if len(df) > 5 else ""
        if "USD" in f6_val:
            curr = "USD"
        elif any(x in f6_val for x in ["중국", "CNY"]):
            curr = "CNY"
        else:
            curr = "한화"
            
        p_col = 1 if curr == "한화" else 2
        prods = df.iloc[6:, p_col].dropna().astype(str).tolist()
        if prods:
            prod_name = prods[0].split("[")[0].strip()
            if len(prods) > 1:
                prod_name = prod_name + f" 외 {len(prods)-1}건"
        else:
            prod_name = "품목미상"
            
        # 총액 추출
        l_idx = df.iloc[:, 5].last_valid_index()
        if curr != "한화" and l_idx:
            total = to_float(df.iloc[l_idx, 5])
        else:
            total = to_float(str(df.iloc[4, 0]).split(":")[-1])
            
        conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,0)", 
                     (raw_oid, odate, "", v_fixed, prod_name, v_type, curr, total))
        conn.commit()
        return True, None
    except Exception as e: 
        return False, f"분석 중 오류 발생: {str(e)}"

# ==============================================================================
# 4. 메인 UI (탭 구성 및 상세 로직)
# ==============================================================================
tabs = st.tabs(["입금 내역 등록", "발주서 관리", "상세내역 및 정산", "거래처 정보 관리", "환율 분석 및 관리"])

# ------------------------------------------------------------------------------
# [Tab 0] 입금 내역 등록 (수기 입력 + 엑셀 업로드 통합)
# ------------------------------------------------------------------------------
with tabs[0]:
    st.header("입금 내역 등록")
    v_data_master = pd.read_sql("SELECT * FROM vendors", conn)
    o_active_list = pd.read_sql("SELECT 발주번호 FROM orders WHERE 마감여부=0", conn)
    
    col_input, col_excel = st.columns([1.5, 1])
    
    with col_input:
        st.subheader("1. 수기 입력 양식")
        with st.form("pay_manual_entry", clear_on_submit=True):
            r1c1, r1c2 = st.columns(2)
            p_oid = r1c1.selectbox("연동할 발주번호", ["없음"] + list(o_active_list['발주번호']))
            p_date = r1c2.date_input("입금 일자", value=datetime.now())
            
            r2c1, r2c2, r2c3 = st.columns(3)
            p_vn = r2c1.selectbox("입금 거래처", ["선택하세요"] + list(v_data_master['거래처명']))
            p_ct = r2c2.selectbox("유형 분류", CATEGORIES)
            p_pr = r2c3.text_input("상품명(적요)")
            
            r3c1, r3c2, r3c3 = st.columns(3)
            p_dep = r3c1.number_input("실입금액 (실송금액)", format="%.2f")
            p_pre = r3c2.number_input("선급금액", format="%.2f")
            p_cur = r3c3.selectbox("거래 통화", ["한화", "USD", "CNY"])
            
            p_memo = st.text_input("비고/메모 (송금 사유 등)")
            
            if st.form_submit_button("입금 내역 저장하기"):
                if p_vn == "선택하세요":
                    st.error("거래처를 선택해야 저장할 수 있습니다.")
                else:
                    vendor_info = v_data_master[v_data_master['거래처명']==p_vn].iloc[0]
                    target_oid = to_str(p_oid) if p_oid != "없음" else None
                    conn.execute('''
                        INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                    ''', (target_oid, p_date.strftime("%Y-%m-%d"), p_ct, p_vn, p_pr, p_cur, p_dep, p_pre, p_memo, 0.0, 
                          vendor_info['은행'], vendor_info['계좌번호'], vendor_info['예금주']))
                    conn.commit()
                    st.success("내역이 저장되었습니다.")
                    st.rerun()

    with col_excel:
        st.subheader("2. 엑셀(CSV) 일괄 업로드")
        # 양식 다운로드
        csv_template = pd.DataFrame(columns=["발주번호", "거래처", "유형", "상품명", "입금일", "실입금액", "선급금액", "송금사유"])
        st.download_button("입금 업로드 양식 받기", csv_template.to_csv(index=False).encode('utf-8-sig'), "payment_template.csv")
        
        uploaded_csv = st.file_uploader("파일 선택 (.csv)", type=['csv'], key=f"pay_csv_{st.session_state.pay_up_key}")
        if uploaded_csv and st.button("파일 데이터 일괄 처리"):
            try:
                df_upload = pd.read_csv(uploaded_csv)
                df_upload.columns = [str(c).strip().replace('\ufeff', '') for c in df_upload.columns]
                
                success_count = 0
                for _, row_data in df_upload.iterrows():
                    oid_val = to_str(row_data.get('발주번호'))
                    vn_raw_val = to_str(row_data.get('거래처'))
                    
                    if not vn_raw_val and not oid_val:
                        continue
                        
                    date_str = smart_date(row_data.get('입금일'))
                    
                    # 발주번호가 있으면 발주 정보 우선 참조
                    if oid_val and not pd.read_sql(f"SELECT 발주번호 FROM orders WHERE 발주번호='{oid_val}'", conn).empty:
                        ord_info = pd.read_sql(f"SELECT * FROM orders WHERE 발주번호='{oid_val}'", conn).iloc[0]
                        v_name, v_cat, v_prod, v_curr = ord_info['거래처명'], ord_info['유형'], ord_info['상품명'], ord_info['통화']
                    else:
                        v_name, v_cat, v_prod, v_curr = vn_raw_val, to_str(row_data.get('유형')) or "사입", to_str(row_data.get('상품명')), "한화"
                    
                    # 거래처 계좌 정보 매칭
                    v_master_match = v_data_master[v_data_master['거래처명'].str.lower() == v_name.lower()]
                    b_name = v_master_match.iloc[0]['은행'] if not v_master_match.empty else ""
                    b_acct = v_master_match.iloc[0]['계좌번호'] if not v_master_match.empty else ""
                    b_owner = v_master_match.iloc[0]['예금주'] if not v_master_match.empty else ""
                    
                    dep_amt = to_float(row_data.get('실입금액'))
                    pre_amt = to_float(row_data.get('선급금액'))
                    
                    conn.execute('''
                        INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                    ''', (oid_val if oid_val else None, date_str, v_cat, v_name, v_prod, v_curr, dep_amt, pre_amt, to_str(row_data.get('송금사유')), 0.0, b_name, b_acct, b_owner))
                    success_count += 1
                
                conn.commit()
                st.success(f"총 {success_count}건의 내역이 업로드되었습니다.")
                st.session_state.pay_up_key += 1
                st.rerun()
            except Exception as e:
                st.error(f"파일 처리 중 오류가 발생했습니다: {e}")

# ------------------------------------------------------------------------------
# [Tab 1] 발주서 등록 및 마감 (마감 필터 복구)
# ------------------------------------------------------------------------------
with tabs[1]:
    st.header("발주서 관리 시스템")
    left_c, right_c = st.columns([1, 1.5])
    
    with left_c:
        st.subheader("발주 데이터 등록")
        order_files = st.file_uploader("xlsx 파일 선택 (다중 가능)", type=['xlsx'], accept_multiple_files=True, key=f"ord_f_{st.session_state.order_up_key}")
        if order_files and st.button("파일 일괄 분석 실행"):
            reg_cnt = 0
            for f in order_files:
                success, msg = process_ecount_v136(f)
                if success: reg_cnt += 1
                else: st.warning(f"[{f.name}] {msg}")
            if reg_cnt > 0:
                st.success(f"{reg_cnt}건의 발주서 등록 완료")
                st.session_state.order_up_key += 1
                st.rerun()
        
        st.divider()
        with st.form("manual_order_form"):
            st.subheader("수기 직접 등록")
            m_oid = st.text_input("발주번호")
            m_step = st.text_input("발주차수")
            m_date = st.date_input("발주일자")
            m_vn = st.selectbox("거래처", ["선택"] + list(v_data_master['거래처명']))
            m_prod = st.text_input("상품명")
            m_amt = st.number_input("발주 금액", format="%.2f")
            m_curr = st.selectbox("통화 단위", ["한화", "USD", "CNY"])
            
            if st.form_submit_button("발주서 수기 저장"):
                dup_chk = pd.read_sql(f"SELECT 발주번호 FROM orders WHERE 발주번호='{m_oid}'", conn)
                if not m_oid: st.error("발주번호를 입력하세요.")
                elif not dup_chk.empty: st.error("이미 존재하는 발주번호입니다.")
                elif m_vn == "선택": st.error("거래처를 선택하세요.")
                else:
                    v_type_info = v_data_master[v_data_master['거래처명']==m_vn].iloc[0]['기본유형']
                    conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,0)", 
                                 (m_oid, m_date.strftime("%Y-%m-%d"), m_step, m_vn, m_prod, v_type_info, m_curr, m_amt))
                    conn.commit()
                    st.rerun()

    with right_c:
        st.subheader("발주 리스트 및 마감 관리")
        # [복구] 마감 데이터 포함 여부 체크박스
        inc_closed = st.checkbox("이미 마감된 발주서도 리스트에 표시", value=True)
        all_orders = pd.read_sql("SELECT * FROM orders", conn)
        
        if not all_orders.empty:
            display_orders = all_orders if inc_closed else all_orders[all_orders['마감여부'] == 0]
            # [디테일] 금액 포맷팅 적용
            edited_orders = st.data_editor(
                display_orders.sort_values('발주일', ascending=False), 
                hide_index=True, 
                use_container_width=True, 
                disabled=["발주번호"], 
                column_config={
                    "발주총액": st.column_config.NumberColumn("발주총액", format="%,.2f"), 
                    "마감여부": st.column_config.CheckboxColumn("마감")
                }
            )
            
            if st.button("수정 내용 소급 적용 및 저장"):
                for index, row in edited_orders.iterrows():
                    # 발주서 테이블 업데이트
                    conn.execute('''UPDATE orders SET 발주일=?, 발주차수=?, 거래처명=?, 상품명=?, 유형=?, 통화=?, 발주총액=?, 마감여부=? 
                                    WHERE 발주번호=?''', 
                                 (row['발주일'], row['발주차수'], row['거래처명'], row['상품명'], row['유형'], row['통화'], row['발주총액'], int(row['마감여부']), row['발주번호']))
                    # [복구] 입금 내역 데이터도 함께 업데이트 (거래처, 유형, 상품명 동기화)
                    conn.execute("UPDATE payments SET 거래처명=?, 유형=?, 상품명=?, 통화=? WHERE 발주번호=?", 
                                 (row['거래처명'], row['유형'], row['상품명'], row['통화'], row['발주번호']))
                conn.commit()
                st.success("정보가 업데이트되었습니다.")
                st.rerun()

# ------------------------------------------------------------------------------
# [Tab 2] 상세내역 및 통합 정산 (색상 강조 및 정밀 포맷팅)
# ------------------------------------------------------------------------------
with tabs[2]:
    st.header("상세 내역 및 통합 정산")
    payments_all = pd.read_sql("SELECT * FROM payments", conn)
    orders_all = pd.read_sql("SELECT * FROM orders", conn)
    rates_db = pd.read_sql("SELECT * FROM exchange_rates ORDER BY 날짜 ASC", conn)
    
    # 상단 필터/요약 레이아웃
    f_col1, f_col2 = st.columns([1, 1.2])
    
    with f_col1:
        st.subheader("조회 조건 설정")
        payments_all['dt'] = pd.to_datetime(payments_all['입금일'])
        flt_r1c1, flt_r1c2 = st.columns(2)
        target_year = flt_r1c1.selectbox("조회 연도", sorted(payments_all['dt'].dt.year.unique(), reverse=True))
        target_month = flt_r1c2.selectbox("조회 월", ["전체"] + sorted(list(payments_all[payments_all['dt'].dt.year==target_year]['dt'].dt.month.unique())))
        
        flt_r2c1, flt_r2c2 = st.columns(2)
        filter_cat = flt_r2c1.selectbox("유형 필터", ["전체 유형"] + CATEGORIES)
        search_key = flt_r2c2.text_input("업체/상품 검색 (대소문자 구분 없음)")
        
        filtered_df = payments_all[payments_all['dt'].dt.year == target_year].copy()
        if target_month != "전체":
            filtered_df = filtered_df[filtered_df['dt'].dt.month == target_month]
        if filter_cat != "전체 유형":
            filtered_df = filtered_df[filtered_df['유형'] == filter_cat]
        if search_key:
            # [디테일] 대소문자 무관 검색 로직
            filtered_df = filtered_df[filtered_df['거래처명'].str.contains(search_key, case=False, na=False) | 
                                      filtered_df['상품명'].str.contains(search_key, case=False, na=False)]
        
        filtered_df = pd.merge(filtered_df, orders_all[['발주번호', '발주차수']], on='발주번호', how='left')

    with f_col2:
        st.subheader("유형별 요약 현황")
        if not filtered_df.empty:
            summary_table = filtered_df.groupby('유형').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
            st.table(summary_table.style.format({'실입금액': '{:,.2f}', '선급금액': '{:,.2f}'}))

    st.divider()
    st.subheader("발주번호별 정산 및 미수금 현황")
    # 발주번호가 있는 데이터만 정산 테이블에 포함
    pay_agg = payments_all[payments_all['발주번호'].notnull() & (payments_all['발주번호'] != "")].groupby('발주번호').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
    settle_df = pd.merge(orders_all, pay_agg, on='발주번호', how='left').fillna(0)
    settle_df['잔액'] = settle_df['발주총액'] - settle_df['실입금액']
    settle_df['진행상태'] = settle_df['마감여부'].apply(lambda x: "✅ 마감" if x == 1 else "⏳ 진행")
    settle_df = settle_df.sort_values(['마감여부', '발주번호'], ascending=[True, False])
    
    view_settle = settle_df[['발주번호', '발주차수', '진행상태', '거래처명', '상품명', '발주총액', '실입금액', '선급금액', '잔액', '통화']]
    
    # [디테일] 선급금(빨강), 잔액(파랑) 강조 스타일
    def color_settle_styles(row):
        styles = [''] * len(row)
        if row['선급금액'] > 0:
            styles[view_settle.columns.get_loc('선급금액')] = 'color: red; font-weight: bold'
        if row['잔액'] > 0:
            styles[view_settle.columns.get_loc('잔액')] = 'color: blue; font-weight: bold'
        if row['진행상태'] == '✅ 마감':
            styles = ['background-color: #f9f9f9; color: #bbbbbb'] * len(row)
        return styles

    st.dataframe(view_settle.style.apply(color_settle_styles, axis=1).format({'발주총액':'{:,.2f}','실입금액':'{:,.2f}','선급금액':'{:,.2f}','잔액':'{:,.2f}'}), 
                 use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("입금 상세 리스트 편집")
    
    # 환율 로직: 직전 월 소급 적용
    rates_db['ym'] = pd.to_datetime(rates_db['날짜']).dt.strftime('%Y-%m')
    monthly_rates = rates_db.groupby('ym').agg({'usd': lambda x: x[x>0].mean(), 'cny': lambda x: x[x>0].mean()}).fillna(0)
    
    def get_conversion_krw(row):
        if row['통화'] == '한화': return row['실입금액']
        ym_key, curr_key = str(row['입금일'])[:7], row['통화'].lower()
        if ym_key in monthly_rates.index and monthly_rates.loc[ym_key, curr_key] > 0: 
            rate = monthly_rates.loc[ym_key, curr_key]
        else:
            past_data = monthly_rates[monthly_rates.index < ym_key]
            if not past_data.empty and past_data[curr_key].sum() > 0:
                rate = past_data[past_data[curr_key] > 0].iloc[-1][curr_key]
            else:
                rate = 1350.0 if row['통화'] == 'USD' else 190.0
        return row['실입금액'] * rate

    final_detail_df = pd.merge(filtered_df, orders_all[['발주번호', '발주총액']], on='발주번호', how='left').fillna(0)
    if not final_detail_df.empty:
        final_detail_df['예상환산액'] = final_detail_df.apply(get_conversion_krw, axis=1)
    else:
        final_detail_df['예상환산액'] = pd.Series(dtype='float64')

    # [디테일] 입금일 위치 조정 및 회계 포맷팅
    col_order = ['id', '유형', '발주번호', '거래처명', '상품명', '통화', '발주총액', '입금일', '실입금액', '선급금액', '예상환산액', '메모']
    
    edited_detail = st.data_editor(
        final_detail_df[col_order].sort_values('입금일', ascending=False), 
        hide_index=True, 
        use_container_width=True,
        column_config={c: st.column_config.NumberColumn(c, format="%,.2f") for c in ['발주총액','실입금액','선급금액','예상환산액']}
    )
    
    btn_col1, btn_col2 = st.columns([1, 4])
    with btn_col1:
        if st.button("상세 수정 내용 저장"):
            for _, r in edited_detail.iterrows():
                conn.execute("UPDATE payments SET 유형=?, 거래처명=?, 상품명=?, 실입금액=?, 선급금액=?, 메모=? WHERE id=?", 
                             (r['유형'], r['거래처명'], r['상품명'], r['실입금액'], r['선급금액'], r['메모'], r['id']))
            conn.commit()
            st.rerun()
    with btn_col2:
        with st.form("delete_payment_form"):
            del_target_id = st.number_input("삭제할 입금 내역의 ID 번호 입력", min_value=0, step=1)
            if st.form_submit_button("해당 ID 데이터 영구 삭제"):
                conn.execute(f"DELETE FROM payments WHERE id={del_target_id}")
                conn.commit()
                st.rerun()

    st.divider()
    st.markdown(f"### 📊 현재 검색 조건 합계 리포트 ({len(final_detail_df)}건)")
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("총 한화 환산액", f"{final_detail_df['예상환산액'].sum():,.2f}")
    m2.metric("KRW 실입금 합계", f"{final_detail_df[final_detail_df['통화']=='한화']['실입금액'].sum():,.2f}")
    m3.metric("USD 실입금 합계", f"{final_detail_df[final_detail_df['통화']=='USD']['실입금액'].sum():,.2f}")
    m4.metric("CNY 실입금 합계", f"{final_detail_df[final_detail_df['통화']=='CNY']['실입금액'].sum():,.2f}")

# ------------------------------------------------------------------------------
# [Tab 3] 거래처 정보 관리 (수정 동기화 및 엑셀 업로드 복구)
# ------------------------------------------------------------------------------
with tabs[3]:
    st.header("거래처 마스터 정보 관리")
    vc_left, vc_right = st.columns([1.2, 0.8])
    
    with vc_left:
        st.subheader("1. 거래처 조회 및 실시간 수정")
        search_v_key = st.text_input("거래처명으로 검색")
        vendors_master = pd.read_sql("SELECT * FROM vendors", conn)
        if search_v_key:
            vendors_master = vendors_master[vendors_master['거래처명'].str.contains(search_v_key, case=False, na=False)]
        
        # [복구] 수정 전 원본 리스트 저장 (이름 변경 감지용)
        original_names = vendors_master['거래처명'].tolist()
        edited_vendors = st.data_editor(vendors_master, hide_index=True, use_container_width=True)
        
        if st.button("거래처 수정 정보 동기화 및 저장"):
            for idx, r in edited_vendors.iterrows():
                old_name = original_names[idx]
                new_name = r['거래처명']
                if old_name != new_name:
                    # 거래처명 변경 시 연결된 모든 테이블 동기화 업데이트
                    conn.execute(f"DELETE FROM vendors WHERE 거래처명='{old_name}'")
                    conn.execute("INSERT INTO vendors VALUES (?,?,?,?,?)", (new_name, r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
                    conn.execute("UPDATE orders SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_name, r['기본유형'], old_name))
                    conn.execute("UPDATE payments SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_name, r['기본유형'], old_name))
                else:
                    conn.execute("UPDATE vendors SET 은행=?, 계좌번호=?, 예금주=?, 기본유형=? WHERE 거래처명=?", 
                                 (r['은행'], r['계좌번호'], r['예금주'], r['기본유형'], r['거래처명']))
            conn.commit()
            st.success("거래처 정보 및 관련 데이터가 모두 동기화되었습니다.")
            st.rerun()
            
    with vc_right:
        st.subheader("2. 등록 및 일괄 업로드")
        with st.expander("신규 거래처 개별 등록"):
            with st.form("new_vendor_manual"):
                nv_n = st.text_input("거래처명")
                nv_t = st.selectbox("기본 유형", CATEGORIES)
                nv_b = st.text_input("은행명")
                nv_a = st.text_input("계좌번호")
                nv_o = st.text_input("예금주명")
                if st.form_submit_button("거래처 추가"):
                    if nv_n:
                        conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (nv_n, nv_b, nv_a, nv_o, nv_t))
                        conn.commit()
                        st.rerun()
        
        with st.expander("거래처 엑셀 일괄 업로드"):
            vendor_tmp = pd.DataFrame(columns=["거래처명", "은행", "계좌번호", "예금주", "기본유형"])
            st.download_button("거래처 양식 다운로드", vendor_tmp.to_csv(index=False).encode('utf-8-sig'), "vendor_template.csv")
            v_file = st.file_uploader("거래처 CSV 선택", type=['csv'], key="v_csv_upload")
            if v_file and st.button("거래처 데이터 일괄 저장"):
                try:
                    v_up_df = pd.read_csv(v_file)
                    v_up_df.columns = [str(c).strip().replace('\ufeff', '') for c in v_up_df.columns]
                    for _, row in v_up_df.iterrows():
                        conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", 
                                     (row['거래처명'], row['은행'], row['계좌번호'], row['예금주'], row['기본유형']))
                    conn.commit()
                    st.success("거래처 일괄 등록 성공!")
                    st.rerun()
                except Exception as e:
                    st.error(f"오류: {e}")

# ------------------------------------------------------------------------------
# [Tab 4] 환율 관리 (디테일: 연도순 2025->2026 고정 및 정밀 분석 리포트)
# ------------------------------------------------------------------------------
with tabs[4]:
    st.header("환율 관리 및 정밀 분석")
    rate_col1, rate_col2 = st.columns(2)
    with rate_col1:
        fu_usd = st.file_uploader("USD 환율 CSV 파일", type=['csv'], key="tab4_usd_csv")
        if fu_usd and st.button("USD 환율 데이터 업데이트"):
            if process_exchange_csv(fu_usd, "USD"):
                st.success("USD 환율이 업데이트되었습니다.")
                st.rerun()
    with rate_col2:
        fu_cny = st.file_uploader("CNY 환율 CSV 파일", type=['csv'], key="tab4_cny_csv")
        if fu_cny and st.button("CNY 환율 데이터 업데이트"):
            if process_exchange_csv(fu_cny, "CNY"):
                st.success("CNY 환율이 업데이트되었습니다.")
                st.rerun()
        
    rates_full = pd.read_sql("SELECT * FROM exchange_rates ORDER BY 날짜 ASC", conn)
    if not rates_full.empty:
        # 날짜 포맷 '25.1월' 형태로 가공
        rates_full['ym_display'] = pd.to_datetime(rates_full['날짜']).dt.strftime('%y.%-m월')
        mean_stats = rates_full.groupby('ym_display', sort=False).agg({'usd': lambda x: x[x>0].mean(), 'cny': lambda x: x[x>0].mean()}).reset_index()
        
        plot1, plot2 = st.columns(2)
        with plot1:
            fig_u = go.Figure(go.Scatter(x=mean_stats['ym_display'], y=mean_stats['usd'], mode='lines+markers', name='USD/KRW'))
            # [디테일] USD 세로축 20단위 고정
            fig_u.update_layout(title="USD 월별 평균 환율 추이", yaxis=dict(dtick=20, tickformat=",.2f"), template="plotly_white", height=400)
            st.plotly_chart(fig_u, use_container_width=True)
        with plot2:
            fig_c = go.Figure(go.Scatter(x=mean_stats['ym_display'], y=mean_stats['cny'], mode='lines+markers', name='CNY/KRW', line=dict(color='orange')))
            # [디테일] CNY 세로축 2단위 고정
            fig_c.update_layout(title="CNY 월별 평균 환율 추이", yaxis=dict(dtick=2, tickformat=",.2f"), template="plotly_white", height=400)
            st.plotly_chart(fig_c, use_container_width=True)
            
        st.subheader("연도별 평균환율 분석 (2025 -> 2026)")
        
        # 리포트용 데이터 가공
        report_raw = pd.read_sql("SELECT * FROM exchange_rates ORDER BY 날짜 ASC", conn)
        report_raw['ym'] = pd.to_datetime(report_raw['날짜']).dt.strftime('%Y-%m')
        report_stats = report_raw.groupby('ym').agg({'usd': lambda x: x[x>0].mean(), 'cny': lambda x: x[x>0].mean()}).reset_index().fillna(0)
        report_stats['year'] = report_stats['ym'].str[:4].astype(int)
        report_stats['month'] = report_stats['ym'].str[5:].astype(int)

        def generate_full_analysis(df, col_name):
            years_list = sorted(df['year'].unique()) # 과거 -> 현재 (2025, 2026)
            if not years_list:
                return pd.DataFrame()
                
            report_df = pd.DataFrame({'월': [f"{i}월" for i in range(1, 13)]})
            
            # [디테일] 연도 컬럼을 2025년, 2026년 순서로 생성
            for yr in years_list:
                month_map = df[df['year'] == yr].set_index('month')[col_name]
                report_df[f'{yr}년'] = report_df['월'].apply(lambda x: month_map.get(int(x.replace('월','')), 0))
            
            # [복구] YoY(전년비) 계산 로직
            if len(years_list) >= 2:
                curr_y, prev_y = years_list[-1], years_list[-2]
                def calc_yoy(row):
                    v_curr, v_prev = row[f'{curr_y}년'], row[f'{prev_y}년']
                    if pd.notnull(v_curr) and pd.notnull(v_prev) and v_prev > 0 and v_curr > 0:
                        diff = v_curr - v_prev
                        pct = (diff / v_prev) * 100
                        return f"{diff:+.2f}({pct:+.1f}%)"
                    return "-"
                report_df['전년비(YoY)'] = report_df.apply(calc_yoy, axis=1)
            
            # [복구] MoM(전월비) 계산 로직
            df_sorted = df.sort_values('ym').copy()
            df_sorted['diff'] = df_sorted[col_name].diff()
            df_sorted['prev_val'] = df_sorted[col_name].shift(1)
            
            def calc_mom(row):
                m_num = int(row['월'].replace('월',''))
                target_row = df_sorted[(df_sorted['year'] == years_list[-1]) & (df_sorted['month'] == m_num)]
                if not target_row.empty:
                    d_val = target_row.iloc[0]['diff']
                    p_val = target_row.iloc[0]['prev_val']
                    if pd.notnull(d_val) and pd.notnull(p_val) and p_val > 0:
                        p_pct = (d_val / p_val) * 100
                        return f"{d_val:+.2f}({p_pct:+.1f}%)"
                return "-"
            report_df['전월비(MoM)'] = report_df.apply(calc_mom, axis=1)
            
            # 데이터가 존재하는 행만 필터링
            final_mask = report_df[f'{years_list[0]}년'] > 0
            for yr in years_list[1:]:
                final_mask = final_mask | (report_df[f'{yr}년'] > 0)
            return report_df[final_mask].reset_index(drop=True)

        rep_col1, rep_col2 = st.columns(2)
        with rep_col1:
            st.write("**[USD/KRW] 전년비/전월비 정밀 분석**")
            usd_rep = generate_full_analysis(report_stats, 'usd')
            st.table(usd_rep.style.format(precision=2, thousands=","))
        with rep_col2:
            st.write("**[CNY/KRW] 전년비/전월비 정밀 분석**")
            cny_rep = generate_full_analysis(report_stats, 'cny')
            st.table(cny_rep.style.format(precision=2, thousands=","))
    else:
        st.warning("분석할 환율 데이터가 없습니다. CSV 파일을 먼저 업데이트 해주세요.")