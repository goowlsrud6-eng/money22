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
    if not os.path.exists('backups'):
        os.makedirs('backups')
    db_file = 'finance_final_v129.db'
    today_str = datetime.now().strftime('%Y%m%d')
    backup_file = f"backups/backup_{today_str}.db"
    
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

st.set_page_config(page_title="자금 관리 시스템", layout="wide")
run_backup()

@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_final_v129.db', check_same_thread=False)
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
# 2. 유틸리티 함수
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

def smart_date(date_str, default_now=True):
    try:
        if pd.isna(date_str) or str(date_str).strip() == "": 
            return datetime.now().strftime("%Y-%m-%d") if default_now else None
        ds = str(date_str).strip().replace(" ", "").replace(".", "-")
        return pd.to_datetime(ds).strftime("%Y-%m-%d")
    except: 
        return datetime.now().strftime("%Y-%m-%d") if default_now else None

@st.cache_data(ttl=3600)
def get_realtime_rate(currency):
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
    except Exception as e: 
        st.error(f"환율 분석 오류: {e}")
        return False

def process_ecount_v129(file):
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
        v_master['clean'] = v_master['거래처명'].apply(lambda x: re.sub(r'\s+', '', str(x)))
        match = v_master[v_master['clean'] == re.sub(r'\s+', '', vendor_raw)]
        
        if match.empty: 
            return False, "미등록 업체: [" + vendor_raw + "]"
            
        v_type = match.iloc[0]['기본유형']
        v_fixed = match.iloc[0]['거래처명']
        
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
# 4. 메인 UI (탭 구성)
# ==========================================
tabs = st.tabs(["입금 입력", "입금 엑셀 업로드", "발주서 등록", "상세내역 및 정산", "거래처 관리", "환율 관리"])

# ------------------------------------------
# [Tab 0] 입금 수기 입력
# ------------------------------------------
with tabs[0]:
    st.header("입금 내역 수기 입력")
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    o_active = pd.read_sql("SELECT 발주번호 FROM orders WHERE 마감여부=0", conn)
    
    with st.form("pay_manual_v129", clear_on_submit=True):
        c1, c2 = st.columns(2)
        p_oid = c1.selectbox("발주번호 연동", ["없음"] + list(o_active['발주번호']))
        p_date = c2.date_input("입금일", value=datetime.now())
        
        c3, c4, c5 = st.columns(3)
        p_vn = c3.selectbox("거래처명", ["선택"] + list(v_data['거래처명']))
        p_ct = c4.selectbox("유형", CATEGORIES)
        p_pr = c5.text_input("상품명")
        
        c6, c7, c8 = st.columns(3)
        p_dep = c6.number_input("실입금액", format="%.2f")
        p_pre = c7.number_input("선급금액", format="%.2f")
        p_cur = c8.selectbox("통화", ["한화", "USD", "CNY"])
        
        p_memo = st.text_input("메모(송금사유)")
        
        if st.form_submit_button("입금 내역 저장"):
            if p_vn == "선택":
                st.error("거래처를 선택하세요.")
            else:
                rate = 1350.0 if p_cur == "USD" else (190.0 if p_cur == "CNY" else 1.0)
                vi = v_data[v_data['거래처명']==p_vn].iloc[0]
                
                conn.execute('''
                    INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                ''', (to_str(p_oid) if p_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), p_ct, p_vn, p_pr, p_cur, p_dep, p_pre, p_memo, (p_dep+p_pre)*rate, vi['은행'], vi['계좌번호'], vi['예금주']))
                
                conn.commit()
                st.success("저장 완료!")
                st.rerun()

# ------------------------------------------
# [Tab 1] 입금 엑셀 업로드
# ------------------------------------------
with tabs[1]:
    st.header("통합 입금 엑셀 업로드")
    pay_tmp = pd.DataFrame(columns=["발주번호", "거래처", "유형", "상품명", "입금일", "실입금액", "선급금액", "송금사유"])
    st.download_button(label="입금 업로드 양식 다운로드", data=pay_tmp.to_csv(index=False).encode('utf-8-sig'), file_name='payment_template.csv')
    
    f_p = st.file_uploader("입금 CSV 선택", type=['csv'], key=f"pay_up_{st.session_state.pay_up_key}")
    if f_p and st.button("데이터 일괄 저장"):
        try:
            df_p = pd.read_csv(f_p)
            df_p.columns = [str(c).strip().replace('\ufeff', '') for c in df_p.columns]
            
            v_l = pd.read_sql("SELECT * FROM vendors", conn)
            o_l = pd.read_sql("SELECT * FROM orders", conn)
            
            for _, r in df_p.iterrows():
                oid = to_str(r.get('발주번호'))
                vn_raw = to_str(r.get('거래처'))
                
                if not vn_raw and not oid: 
                    continue
                
                pd_s = smart_date(r.get('입금일'))
                
                if oid and not o_l[o_l['발주번호'] == oid].empty:
                    info = o_l[o_l['발주번호'] == oid].iloc[0]
                    vn = info['거래처명']
                    pc = info['유형']
                    pp = info['상품명']
                    cur = info['통화']
                else: 
                    vn = vn_raw
                    pc = to_str(r.get('유형')) or "사입"
                    pp = to_str(r.get('상품명'))
                    cur = "한화"
                    
                vi = v_l[v_l['거래처명'] == vn] if vn else pd.DataFrame()
                dep = to_float(r.get('실입금액'))
                pre = to_float(r.get('선급금액'))
                rt = 1350.0 if cur == "USD" else (190.0 if cur == "CNY" else 1.0)
                
                b_bank = vi.iloc[0]['은행'] if not vi.empty else ""
                b_acc = vi.iloc[0]['계좌번호'] if not vi.empty else ""
                b_hold = vi.iloc[0]['예금주'] if not vi.empty else ""
                
                conn.execute('''
                    INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                ''', (oid if oid else None, pd_s, pc, vn, pp, cur, dep, pre, to_str(r.get('송금사유')), (dep+pre)*rt, b_bank, b_acc, b_hold))
                
            conn.commit()
            st.success("일괄 저장 완료!")
            st.session_state.pay_up_key += 1
            st.rerun()
        except Exception as e: 
            st.error(f"오류: {e}")

# ------------------------------------------
# [Tab 2] 발주서 등록 및 마감
# ------------------------------------------
with tabs[2]:
    st.header("발주서 등록 및 마감")
    c_o1, c_o2 = st.columns(2)
    
    with c_o1:
        st.subheader("엑셀 일괄 등록")
        ord_tmp = pd.DataFrame(columns=["발주번호", "발주일", "발주차수", "거래처명", "상품명", "금액", "통화"])
        st.download_button(label="수기용 발주 양식 다운로드", data=ord_tmp.to_csv(index=False).encode('utf-8-sig'), file_name='order_template.csv')
        
        of_list = st.file_uploader("발주서(xlsx) 선택", type=['xlsx'], accept_multiple_files=True, key=f"ord_{st.session_state.order_up_key}")
        if of_list and st.button("일괄 등록 실행"):
            success_cnt = 0
            error_messages = []
            
            for of in of_list:
                ok, msg = process_ecount_v129(of)
                if ok: 
                    success_cnt += 1
                else: 
                    error_messages.append(f"[{of.name}] {msg}")
                    
            for e in error_messages: 
                st.warning(e)
                
            if success_cnt > 0: 
                st.success(f"총 {success_cnt}건 성공!")
                if not error_messages:
                    st.session_state.order_up_key += 1
                    st.rerun()
                    
    with c_o2:
        st.subheader("수기 발주 등록")
        v_list = pd.read_sql("SELECT 거래처명 FROM vendors", conn)
        with st.form("ord_manual_v129"):
            mi = st.text_input("발주번호")
            m_step = st.text_input("발주차수")
            md = st.date_input("발주일")
            mv = st.selectbox("거래처 선택", ["선택"] + list(v_list['거래처명']))
            mp = st.text_input("상품명")
            mt = st.number_input("금액", format="%.2f")
            m_cur = st.selectbox("통화", ["한화", "USD", "CNY"])
            
            if st.form_submit_button("수기 저장"):
                if mi and mv != "선택":
                    vt_query = pd.read_sql(f"SELECT 기본유형 FROM vendors WHERE 거래처명='{mv}'", conn)
                    vt = vt_query.iloc[0]['기본유형'] if not vt_query.empty else "기타"
                    
                    conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,0)", 
                                 (mi, md.strftime("%Y-%m-%d"), m_step, mv, mp, vt, m_cur, mt))
                    conn.commit()
                    st.success("수기 등록 완료!")
                    st.rerun()
                    
    st.divider()
    o_data = pd.read_sql("SELECT * FROM orders", conn)
    
    if not o_data.empty:
        st.subheader("발주 리스트 및 마감 관리")
        show_c = st.checkbox("마감된 발주 포함", value=True)
        disp_o = o_data if show_c else o_data[o_data['마감여부'] == 0]
        
        ev_o = st.data_editor(disp_o.sort_values('발주일', ascending=False), hide_index=True, use_container_width=True, disabled=["발주번호"], column_config={"마감여부": st.column_config.CheckboxColumn("마감")})
        
        if st.button("정보 소급 적용"):
            for _, r in ev_o.iterrows():
                conn.execute("UPDATE orders SET 발주일=?, 발주차수=?, 거래처명=?, 상품명=?, 유형=?, 통화=?, 발주총액=?, 마감여부=? WHERE 발주번호=?", 
                             (r['발주일'], r['발주차수'], r['거래처명'], r['상품명'], r['유형'], r['통화'], r['발주총액'], int(r['마감여부']), r['발주번호']))
                conn.execute("UPDATE payments SET 거래처명=?, 유형=?, 상품명=?, 통화=? WHERE 발주번호=?", 
                             (r['거래처명'], r['유형'], r['상품명'], r['통화'], r['발주번호']))
            conn.commit()
            st.success("동기화 완료!")
            st.rerun()

# ------------------------------------------
# [Tab 3] 상세내역 및 통합 정산 (★★ KeyError 및 Outer Merge 완벽 복구 ★★)
# ------------------------------------------
with tabs[3]:
    st.header("상세 내역 및 통합 정산")
    p_all = pd.read_sql("SELECT * FROM payments", conn)
    o_all = pd.read_sql("SELECT * FROM orders", conn)
    
    if not p_all.empty:
        p_all['dt'] = pd.to_datetime(p_all['입금일'])
        
        st.subheader("필터 및 검색")
        f1, f2, f3 = st.columns([1, 1, 2])
        y = f1.selectbox("연도", sorted(p_all['dt'].dt.year.unique(), reverse=True))
        m = f2.selectbox("월", ["전체"] + sorted(list(p_all[p_all['dt'].dt.year==y]['dt'].dt.month.unique())))
        
        search_col, step_col = f3.columns([2, 1])
        search = search_col.text_input("업체/상품 검색")
        search_step = step_col.text_input("발주차수 검색")
        
        fil_p = p_all[p_all['dt'].dt.year == y].copy()
        if m != "전체": 
            fil_p = fil_p[fil_p['dt'].dt.month == m]
        if search: 
            fil_p = fil_p[fil_p['거래처명'].str.contains(search, na=False) | fil_p['상품명'].str.contains(search, na=False)]
        
        fil_p = pd.merge(fil_p, o_all[['발주번호', '발주차수']], on='발주번호', how='left')
        if search_step: 
            fil_p = fil_p[fil_p['발주차수'].str.contains(search_step, na=False)]

        if not fil_p.empty:
            cat_sum = fil_p.groupby('유형').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
            cat_sum['총합계'] = cat_sum['실입금액'] + cat_sum['선급금액']
            st.write(f"#### {y}년 {m if m != '전체' else ''} 유형별 요약")
            st.table(cat_sum.style.format({'실입금액': '{:,.2f}', '선급금액': '{:,.2f}', '총합계': '{:,.2f}'}))
        
        st.divider()
        st.subheader("발주번호별 정산 및 미수금 현황")
        
        # ★ Outer Merge 복구: 발주가 등록되지 않은 입금 엑셀도 표에 표시되도록 outer 조인 사용
        p_agg = p_all.groupby('발주번호').agg({'실입금액':'sum'}).reset_index()
        sum_df = pd.merge(o_all, p_agg, on='발주번호', how='outer')
        
        # Outer 조인 시 발생하는 NaN 값들을 안전하게 채움
        sum_df['발주총액'] = sum_df['발주총액'].fillna(0)
        sum_df['실입금액'] = sum_df['실입금액'].fillna(0)
        sum_df['마감여부'] = sum_df['마감여부'].fillna(0)
        
        # 발주서가 없어서 빈 문자열이 되는 컬럼을 입금 내역(p_all)에서 끌어와서 채워줌
        p_latest = p_all.drop_duplicates('발주번호', keep='last').set_index('발주번호')
        sum_df = sum_df.set_index('발주번호')
        
        sum_df['거래처명'] = sum_df['거래처명'].fillna(p_latest['거래처명']).fillna('')
        sum_df['상품명'] = sum_df['상품명'].fillna(p_latest['상품명']).fillna('')
        sum_df['통화'] = sum_df['통화'].fillna(p_latest['통화']).fillna('한화')
        sum_df['발주차수'] = sum_df['발주차수'].fillna('미등록')
        
        sum_df = sum_df.reset_index()
        
        sum_df['잔액'] = sum_df['발주총액'] - sum_df['실입금액']
        sum_df['상태'] = sum_df['마감여부'].apply(lambda x: "마감완료" if x == 1 else "진행중")
        
        # ★ KeyError 완벽 해결: 컬럼을 빼기 '전에' 먼저 정렬 수행!
        sum_df = sum_df.sort_values(['마감여부', '발주번호'], ascending=[True, False])
        
        # 정렬이 완료된 후 화면에 보여줄 컬럼만 선택
        disp_sum = sum_df[['발주번호', '발주차수', '상태', '거래처명', '상품명', '발주총액', '실입금액', '잔액', '통화']]
        
        def highlight_closed(row):
            if row['상태'] == '마감완료':
                return ['background-color: #f0f2f6; color: #a0aab2'] * len(row)
            return [''] * len(row)
            
        st.dataframe(disp_sum.style.apply(highlight_closed, axis=1).format({'발주총액': '{:,.2f}', '실입금액': '{:,.2f}', '잔액': '{:,.2f}'}), use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("상세 리스트 편집 및 삭제")
        
        ex_db = pd.read_sql("SELECT * FROM exchange_rates", conn)
        ex_db['ym'] = pd.to_datetime(ex_db['날짜']).dt.strftime('%Y-%m')
        m_rates = ex_db.groupby('ym').agg({'usd': lambda x: x[x>0].mean(), 'cny': lambda x: x[x>0].mean()}).to_dict('index')

        def calc_krw_estimate(row):
            if row['통화'] == '한화': 
                return row['실입금액'] + row['선급금액']
            ym = str(row['입금일'])[:7]
            curr = row['통화'].lower()
            if ym in m_rates and m_rates[ym][curr] > 0:
                rate = m_rates[ym][curr]
            else:
                rate = get_realtime_rate(row['통화'])
            return (row['실입금액'] + row['선급금액']) * rate

        fil_p_merged = pd.merge(fil_p.drop(columns=['발주차수'], errors='ignore'), o_all[['발주번호', '발주차수', '발주총액']], on='발주번호', how='left')
        
        if fil_p_merged.empty: 
            fil_p_merged['예상환산액(KRW)'] = pd.Series(dtype=float)
        else: 
            fil_p_merged['예상환산액(KRW)'] = fil_p_merged.apply(calc_krw_estimate, axis=1)
        
        final_cols = ['id', '발주번호', '발주차수', '발주총액', '실입금액', '선급금액', '예상환산액(KRW)', '입금일', '유형', '거래처명', '상품명', '통화', '메모']
        fil_p_merged = fil_p_merged[[c for c in final_cols if c in fil_p_merged.columns]]
        
        ed_p = st.data_editor(fil_p_merged.sort_values('입금일', ascending=False), hide_index=True, use_container_width=True, 
                              disabled=["id", "발주차수", "발주총액", "예상환산액(KRW)"],
                              column_config={"예상환산액(KRW)": st.column_config.NumberColumn("예상환산액(KRW)", format="₩ %d")})
        
        eb1, eb2 = st.columns([1, 4])
        with eb1:
            if st.button("상세 수정 저장"):
                for _, r in ed_p.iterrows():
                    conn.execute("UPDATE payments SET 발주번호=?, 입금일=?, 유형=?, 거래처명=?, 상품명=?, 실입금액=?, 선급금액=?, 메모=? WHERE id=?", 
                                 (r['발주번호'], r['입금일'], r['유형'], r['거래처명'], r['상품명'], r['실입금액'], r['선급금액'], r['메모'], r['id']))
                conn.commit()
                st.success("저장 완료!")
                st.rerun()
        with eb2:
            with st.form("delete_v129", clear_on_submit=True):
                col_d1, col_d2 = st.columns([2, 1])
                del_id = col_d1.number_input("삭제 ID", min_value=0, step=1)
                if col_d2.form_submit_button("해당 ID 삭제"):
                    conn.execute(f"DELETE FROM payments WHERE id={del_id}")
                    conn.commit()
                    st.success("삭제 완료!")
                    st.rerun()

        st.divider()
        st.markdown(f"### 현재 검색 목록 총 합계 ({len(fil_p_merged)}건)")
        
        if not fil_p_merged.empty:
            tot_all = fil_p_merged['예상환산액(KRW)'].sum()
            tot_krw = fil_p_merged[fil_p_merged['통화']=='한화']['실입금액'].sum() + fil_p_merged[fil_p_merged['통화']=='한화']['선급금액'].sum()
            tot_usd = fil_p_merged[fil_p_merged['통화']=='USD']['실입금액'].sum() + fil_p_merged[fil_p_merged['통화']=='USD']['선급금액'].sum()
            tot_cny = fil_p_merged[fil_p_merged['통화']=='CNY']['실입금액'].sum() + fil_p_merged[fil_p_merged['통화']=='CNY']['선급금액'].sum()
        else:
            tot_all = tot_krw = tot_usd = tot_cny = 0.0
            
        mc1, mc2, mc3, mc4 = st.columns(4)
        mc1.metric("총 예상 환산액", f"₩ {tot_all:,.0f}")
        mc2.metric("KRW 합계", f"₩ {tot_krw:,.0f}")
        mc3.metric("USD 합계", f"$ {tot_usd:,.2f}")
        mc4.metric("CNY 합계", f"¥ {tot_cny:,.2f}")

# ------------------------------------------
# [Tab 4] 거래처 관리
# ------------------------------------------
with tabs[4]:
    st.header("거래처 관리")
    cv1, cv2 = st.columns([1.2, 0.8])
    with cv1:
        st.subheader("신규 등록")
        with st.form("vn_v129", clear_on_submit=True):
            vn = st.text_input("거래처명")
            vt = st.selectbox("유형", CATEGORIES)
            vc1, vc2, vc3 = st.columns(3)
            vb = vc1.text_input("은행")
            va = vc2.text_input("계좌번호")
            vh = vc3.text_input("예금주")
            if st.form_submit_button("저장"):
                if vn: 
                    conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (vn,vb,va,vh,vt))
                    conn.commit()
                    st.success("완료!")
                    st.rerun()
    with cv2:
        st.subheader("일괄 업로드")
        v_tmp = pd.DataFrame(columns=["거래처명", "은행", "계좌번호", "예금주", "기본유형"])
        st.download_button(label="양식 다운로드", data=v_tmp.to_csv(index=False).encode('utf-8-sig'), file_name='vendor_template.csv')
        vf = st.file_uploader("거래처 CSV", type=['csv'])
        if vf and st.button("업로드"):
            v_up = pd.read_csv(vf)
            for _, r in v_up.iterrows():
                conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", 
                             (r['거래처명'], r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
            conn.commit()
            st.success("완료!")
            st.rerun()
            
    st.divider()
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    if not v_data.empty:
        orig_v = v_data['거래처명'].tolist()
        ev_v = st.data_editor(v_data, hide_index=True, use_container_width=True)
        if st.button("거래처 동기화 저장"):
            for idx, r in ev_v.iterrows():
                old_n = orig_v[idx]
                new_n = r['거래처명']
                if old_n != new_n:
                    conn.execute(f"DELETE FROM vendors WHERE 거래처명='{old_n}'")
                    conn.execute("INSERT INTO vendors VALUES (?,?,?,?,?)", (new_n, r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
                    conn.execute("UPDATE orders SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_n, r['기본유형'], old_n))
                    conn.execute("UPDATE payments SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_n, r['기본유형'], old_n))
                else: 
                    conn.execute("UPDATE vendors SET 은행=?, 계좌번호=?, 예금주=?, 기본유형=? WHERE 거래처명=?", 
                                 (r['은행'], r['계좌번호'], r['예금주'], r['기본유형'], r['거래처명']))
            conn.commit()
            st.rerun()

# ------------------------------------------
# [Tab 5] 환율 관리
# ------------------------------------------
with tabs[5]:
    st.header("환율 정밀 분석")
    cu1, cu2 = st.columns(2)
    with cu1:
        f_usd = st.file_uploader("USD/KRW", type=['csv'], key="u")
        if f_usd and st.button("USD 업데이트"): 
            if process_exchange_csv(f_usd, "USD"): 
                st.success("완료")
                st.rerun()
    with cu2:
        f_cny = st.file_uploader("CNY/KRW", type=['csv'], key="c")
        if f_cny and st.button("CNY 업데이트"):
            if process_exchange_csv(f_cny, "CNY"): 
                st.success("완료")
                st.rerun()

    ex_db = pd.read_sql("SELECT * FROM exchange_rates ORDER BY 날짜 ASC", conn)
    if not ex_db.empty:
        ex_db['ym'] = pd.to_datetime(ex_db['날짜']).dt.strftime('%Y-%m')
        m_mean = ex_db.groupby('ym').agg({'usd': lambda x: x[x > 0].mean(), 'cny': lambda x: x[x > 0].mean()}).reset_index().fillna(0)
        m_mean['year'] = m_mean['ym'].str[:4].astype(int)
        m_mean['month'] = m_mean['ym'].str[5:].astype(int)
        
        st.subheader("월별 평균 추이")
        cc1, cc2 = st.columns(2)
        with cc1:
            fig_u = go.Figure()
            fig_u.add_trace(go.Scatter(x=m_mean['ym'], y=m_mean['usd'], mode='lines+markers', name='USD'))
            fig_u.update_layout(yaxis=dict(range=[1360, 1540], dtick=20), height=350, template="plotly_white")
            st.plotly_chart(fig_u, use_container_width=True)
            
        with cc2:
            fig_c = go.Figure()
            fig_c.add_trace(go.Scatter(x=m_mean['ym'], y=m_mean['cny'], mode='lines+markers', name='CNY', line=dict(color='orange')))
            fig_c.update_layout(yaxis=dict(range=[186, 226], dtick=2), height=350, template="plotly_white")
            st.plotly_chart(fig_c, use_container_width=True)

        def get_all_months_report(df, col):
            years = sorted(df['year'].unique(), reverse=True)
            if not years: 
                return pd.DataFrame()
                
            cy = years[0]
            py = years[1] if len(years) > 1 else None
            
            res = pd.DataFrame({'월': range(1, 13)})
            
            if py: 
                res[f'{py}년 평균'] = res['월'].map(df[df['year'] == py].set_index('month')[col])
                
            res[f'{cy}년 평균'] = res['월'].map(df[df['year'] == cy].set_index('month')[col])
            
            if py:
                def calc_yoy(row):
                    v1 = row[f'{cy}년 평균']
                    v2 = row[f'{py}년 평균']
                    if pd.notnull(v1) and pd.notnull(v2) and v2 > 0: 
                        return f"{(v1-v2):+.2f}({((v1-v2)/v2*100):+.1f}%)"
                    return "-"
                res['전년비(YoY)'] = res.apply(calc_yoy, axis=1)
                
            df_s = df.sort_values('ym').copy()
            df_s['diff'] = df_s[col].diff()
            df_s['prev'] = df_s[col].shift(1)
            
            mom_map = df_s[df_s['year'] == cy].set_index('month')
            
            def calc_mom(row):
                m = row['월']
                if m in mom_map.index:
                    d = mom_map.loc[m, 'diff']
                    v = mom_map.loc[m, 'prev']
                    if pd.notnull(d) and pd.notnull(v) and v > 0: 
                        return f"{d:+.2f}({(d/v*100):+.1f}%)"
                return "-"
                
            res['전월비(MoM)'] = res.apply(calc_mom, axis=1)
            
            return res[res[f'{cy}년 평균'].notnull() | (res[f'{py}년 평균'].notnull() if py else False)].reset_index(drop=True)

        st.divider()
        st.subheader("연도별 병렬 리포트")
        rc1, rc2 = st.columns(2)
        
        with rc1: 
            st.write("#### USD 환율")
            ur = get_all_months_report(m_mean, 'usd')
            if not ur.empty:
                st.table(ur.style.format({'월':'{:.0f}월'}, na_rep="-"))
                
        with rc2: 
            st.write("#### CNY 환율")
            cr = get_all_months_report(m_mean, 'cny')
            if not cr.empty:
                st.table(cr.style.format({'월':'{:.0f}월'}, na_rep="-"))
    else: 
        st.warning("데이터가 없습니다.")