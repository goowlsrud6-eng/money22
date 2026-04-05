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
    """안정적인 운영을 위해 매일 첫 접속 시 DB 백업 생성"""
    if not os.path.exists('backups'):
        os.makedirs('backups')
    db_file = 'finance_final_v124.db'
    today_str = datetime.now().strftime('%Y%m%d')
    backup_file = f"backups/backup_{today_str}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

st.set_page_config(page_title="자금 관리 시스템 v124", layout="wide", page_icon="💰")
run_backup()

@st.cache_resource
def get_db_connection():
    """모든 테이블 스키마 정의"""
    conn = sqlite3.connect('finance_final_v124.db', check_same_thread=False)
    c = conn.cursor()
    # [1] 거래처 마스터
    c.execute('''CREATE TABLE IF NOT EXISTS vendors 
                 (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT, 기본유형 TEXT)''')
    # [2] 발주 마스터
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (발주번호 TEXT PRIMARY KEY, 발주일 TEXT, 발주차수 TEXT, 거래처명 TEXT, 
                  상품명 TEXT, 유형 TEXT, 통화 TEXT, 발주총액 REAL, 마감여부 INTEGER DEFAULT 0)''')
    # [3] 입금 상세
    c.execute('''CREATE TABLE IF NOT EXISTS payments 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 발주번호 TEXT, 입금일 TEXT, 
                  유형 TEXT, 거래처명 TEXT, 상품명 TEXT, 통화 TEXT,
                  실입금액 REAL, 선급금액 REAL, 메모 TEXT, 한화환산액 REAL,
                  은행 TEXT, 계좌번호 TEXT, 예금주 TEXT)''')
    # [4] 환율 관리
    c.execute('''CREATE TABLE IF NOT EXISTS exchange_rates 
                 (날짜 TEXT PRIMARY KEY, usd REAL, cny REAL)''')
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비", "원단비", "기타"]

if 'order_up_key' not in st.session_state: st.session_state.order_up_key = 0
if 'pay_up_key' not in st.session_state: st.session_state.pay_up_key = 1000

# ==========================================
# 2. 유틸리티 함수 (날짜 정제 및 실시간 환율)
# ==========================================
def to_float(val):
    try:
        if val is None or pd.isna(val) or str(val).strip() == "": return 0.0
        return float(str(val).replace(',', '').strip())
    except: return 0.0

def to_str(val):
    if val is None or pd.isna(val): return ""
    s = str(val).strip()
    return "" if s.lower() in ["nan", "none", ""] else s

def smart_date(date_str):
    """날짜 형식 통일 (빈 값일 경우 오늘 날짜)"""
    try:
        if pd.isna(date_str) or str(date_str).strip() == "": 
            return datetime.now().strftime("%Y-%m-%d")
        ds = str(date_str).strip().replace(" ", "").replace(".", "-")
        return pd.to_datetime(ds).strftime("%Y-%m-%d")
    except: return datetime.now().strftime("%Y-%m-%d")

@st.cache_data(ttl=3600)
def get_realtime_rate(currency):
    """네이버 금융 실시간 환율 크롤링"""
    try:
        url = f"https://finance.naver.com/marketindex/exchangeDetail.naver?marketindexCd=FX_{currency}KRW"
        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        html = urllib.request.urlopen(req).read().decode('euc-kr', errors='ignore')
        match = re.search(r'<meta property="og:description" content=".*?\s+([\d,.]+)\s+전일대비', html)
        if match:
            return float(match.group(1).replace(',', ''))
    except: pass
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
                usd, cny = (price_val, 0.0) if currency_type == "USD" else (0.0, price_val)
                conn.execute("INSERT INTO exchange_rates VALUES (?,?,?)", (date_val, usd, cny))
            else:
                col = "usd" if currency_type == "USD" else "cny"
                conn.execute(f"UPDATE exchange_rates SET {col} = ? WHERE 날짜 = ?", (price_val, date_val))
        conn.commit(); return True
    except Exception as e:
        st.error(f"환율 파일 처리 오류: {e}"); return False

def process_ecount_v124(file):
    """발주서 분석기 (Syntax Error 방지를 위해 메시지 포맷팅 안전화)"""
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
            # ★ 이모지 충돌 에러 완전 방지
            err_msg = "미등록 업체입니다: [" + vendor_raw + "]"
            return False, err_msg
            
        v_type, v_fixed = match.iloc[0]['기본유형'], match.iloc[0]['거래처명']
        f6 = str(df.iloc[5, 5]) if len(df) > 5 else ""
        curr = "USD" if "USD" in f6 else ("CNY" if any(x in f6 for x in ["중국", "CNY"]) else "한화")
        p_col = 1 if curr == "한화" else 2
        prods = df.iloc[6:, p_col].dropna().astype(str).tolist()
        prod_n = (prods[0].split("[")[0].strip() + (f" 외 {len(prods)-1}건" if len(prods)>1 else "")) if prods else "품목미상"
        l_idx = df.iloc[:, 5].last_valid_index()
        total = to_float(df.iloc[l_idx, 5]) if curr != "한화" and l_idx else to_float(str(df.iloc[4, 0]).split(":")[-1])
        
        conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,0)", (raw_oid, odate, "", v_fixed, prod_n, v_type, curr, total))
        conn.commit(); return True, None
    except Exception as e: 
        err_msg = "분석 오류 발생: " + str(e)
        return False, err_msg

# ==========================================
# 4. 메인 UI 
# ==========================================
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", ⚙️ 거래처 관리", "📈 환율 관리"])

# ------------------------------------------
# [Tab 0] 입금 수기 입력
# ------------------------------------------
with tabs[0]:
    st.header("📝 입금 내역 수기 입력")
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    o_active = pd.read_sql("SELECT 발주번호 FROM orders WHERE 마감여부=0", conn)
    with st.form("pay_manual_v124", clear_on_submit=True):
        c1, c2 = st.columns(2)
        p_oid = c1.selectbox("🔗 발주번호 연동", ["없음"] + list(o_active['발주번호']))
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
        
        if st.form_submit_button("✅ 입금 내역 저장"):
            if p_vn == "선택":
                st.error("거래처를 선택하세요.")
            else:
                rate = 1350.0 if p_cur == "USD" else (190.0 if p_cur == "CNY" else 1.0)
                vi = v_data[v_data['거래처명']==p_vn].iloc[0]
                conn.execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                             (to_str(p_oid) if p_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), p_ct, p_vn, p_pr, p_cur, p_dep, p_pre, p_memo, (p_dep+p_pre)*rate, vi['은행'], vi['계좌번호'], vi['예금주']))
                conn.commit(); st.success("저장 완료!"); st.rerun()

# ------------------------------------------
# [Tab 1] 입금 엑셀 업로드
# ------------------------------------------
with tabs[1]:
    st.header("📂 통합 입금 엑셀 업로드")
    pay_tmp = pd.DataFrame(columns=["발주번호", "거래처", "유형", "상품명", "입금일", "실입금액", "선급금액", "송금사유"])
    st.download_button(label="📥 입금 업로드 양식 다운로드", data=pay_tmp.to_csv(index=False).encode('utf-8-sig'), file_name='payment_template.csv')
    
    f_p = st.file_uploader("입금 CSV 선택", type=['csv'], key=f"pay_up_{st.session_state.pay_up_key}")
    if f_p and st.button("🚀 데이터 일괄 저장"):
        try:
            df_p = pd.read_csv(f_p)
            df_p.columns = [str(c).strip().replace('\ufeff', '') for c in df_p.columns]
            
            v_l = pd.read_sql("SELECT * FROM vendors", conn)
            o_l = pd.read_sql("SELECT * FROM orders", conn)
            
            for _, r in df_p.iterrows():
                oid, vn_raw = to_str(r.get('발주번호')), to_str(r.get('거래처'))
                if not vn_raw and not oid: continue
                
                pd_s = smart_date(r.get('입금일'))
                
                if oid and not o_l[o_l['발주번호'] == oid].empty:
                    info = o_l[o_l['발주번호'] == oid].iloc[0]
                    vn, pc, pp, cur = info['거래처명'], info['유형'], info['상품명'], info['통화']
                else: 
                    vn, pc, pp, cur = vn_raw, to_str(r.get('유형')) or "사입", to_str(r.get('상품명')), "한화"
                    
                vi = v_l[v_l['거래처명'] == vn] if vn else pd.DataFrame()
                dep, pre = to_float(r.get('실입금액')), to_float(r.get('선급금액'))
                rt = 1350.0 if cur == "USD" else (190.0 if cur == "CNY" else 1.0)
                
                conn.execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                             (oid if oid else None, pd_s, pc, vn, pp, cur, dep, pre, to_str(r.get('송금사유')), (dep+pre)*rt, vi.iloc[0]['은행'] if not vi.empty else "", vi.iloc[0]['계좌번호'] if not vi.empty else "", vi.iloc[0]['예금주'] if not vi.empty else ""))
            conn.commit(); st.success("일괄 저장 완료!"); st.session_state.pay_up_key += 1; st.rerun()
        except Exception as e: st.error(f"오류: {e}")

# ------------------------------------------
# [Tab 2] 발주서 등록 및 마감
# ------------------------------------------
with tabs[2]:
    st.header("📥 발주서 등록 및 마감")
    c_o1, c_o2 = st.columns(2)
    with c_o1:
        st.subheader("⚡ 엑셀 일괄 등록")
        ord_tmp = pd.DataFrame(columns=["발주번호", "발주일", "발주차수", "거래처명", "상품명", "금액", "통화"])
        st.download_button(label="📥 수기용 발주 양식 다운로드", data=ord_tmp.to_csv(index=False).encode('utf-8-sig'), file_name='order_template.csv')
        of_list = st.file_uploader("발주서(xlsx) 선택", type=['xlsx'], accept_multiple_files=True, key=f"ord_{st.session_state.order_up_key}")
        
        if of_list and st.button("🚀 일괄 등록 실행"):
            success_cnt = 0
            error_messages = []
            for of in of_list: 
                is_success, err_msg = process_ecount_v124(of)
                if is_success:
                    success_cnt += 1
                else:
                    error_messages.append(f"[{of.name} 파일] ⚠️ {err_msg}")
            
            for em in error_messages: st.warning(em)
            if success_cnt > 0:
                st.success(f"✅ 총 {success_cnt}건 발주서 등록 성공!")
                if not error_messages: st.session_state.order_up_key += 1; st.rerun()

    with c_o2:
        st.subheader("✍️ 수기 발주 등록")
        v_list = pd.read_sql("SELECT 거래처명 FROM vendors", conn)
        with st.form("ord_manual_v124"):
            mi = st.text_input("발주번호")
            m_step = st.text_input("발주차수")
            md = st.date_input("발주일")
            mv = st.selectbox("거래처 선택", ["선택"] + list(v_list['거래처명']))
            mp = st.text_input("상품명")
            mt = st.number_input("금액", format="%.2f")
            m_cur = st.selectbox("통화", ["한화", "USD", "CNY"])
            
            if st.form_submit_button("✅ 수기 저장"):
                if mi and mv != "선택":
                    vt_res = pd.read_sql(f"SELECT 기본유형 FROM vendors WHERE 거래처명='{mv}'", conn)
                    vt = vt_res.iloc[0]['기본유형'] if not vt_res.empty else "사입"
                    conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,0)", 
                                 (mi, md.strftime("%Y-%m-%d"), m_step, mv, mp, vt, m_cur, mt))
                    conn.commit()
                    st.success("수기 등록 완료!")
                    st.rerun()
                    
    st.divider()
    o_data = pd.read_sql("SELECT * FROM orders", conn)
    if not o_data.empty:
        st.subheader("📄 발주 리스트 및 마감 관리")
        show_closed = st.checkbox("마감된 발주 포함해서 보기", value=False)
        disp_o = o_data if show_closed else o_data[o_data['마감여부'] == 0]
        
        ev_o = st.data_editor(disp_o.sort_values('발주일', ascending=False), hide_index=True, use_container_width=True, disabled=["발주번호"], column_config={"마감여부": st.column_config.CheckboxColumn("마감")})
        
        if st.button("💾 정보 업데이트 및 모든 상세내역 소급 적용"):
            for _, r in ev_o.iterrows():
                conn.execute("UPDATE orders SET 발주일=?, 발주차수=?, 거래처명=?, 상품명=?, 유형=?, 통화=?, 발주총액=?, 마감여부=? WHERE 발주번호=?", 
                             (r['발주일'], r['발주차수'], r['거래처명'], r['상품명'], r['유형'], r['통화'], r['발주총액'], int(r['마감여부']), r['발주번호']))
                conn.execute("UPDATE payments SET 거래처명=?, 유형=?, 상품명=?, 통화=? WHERE 발주번호=?", 
                             (r['거래처명'], r['유형'], r['상품명'], r['통화'], r['발주번호']))
            conn.commit()
            st.success("데이터 동기화 완료!")
            st.rerun()

# ------------------------------------------
# [Tab 3] 상세내역 및 통합 정산
# ------------------------------------------
with tabs[3]:
    st.header("🔍 상세 내역 및 통합 정산")
    p_all = pd.read_sql("SELECT * FROM payments", conn)
    o_all = pd.read_sql("SELECT * FROM orders", conn)
    
    if not p_all.empty:
        p_all['dt'] = pd.to_datetime(p_all['입금일'])
        st.subheader("📊 필터 및 검색")
        f1, f2, f3 = st.columns([1, 1, 2])
        y = f1.selectbox("기준 연도", sorted(p_all['dt'].dt.year.unique(), reverse=True))
        m = f2.selectbox("기준 월", ["전체"] + sorted(list(p_all[p_all['dt'].dt.year==y]['dt'].dt.month.unique())))
        search = f3.text_input("업체/상품 통합 검색")
        
        fil_p = p_all[p_all['dt'].dt.year == y].copy()
        if m != "전체": fil_p = fil_p[fil_p['dt'].dt.month == m]
        if search: fil_p = fil_p[fil_p['거래처명'].str.contains(search, na=False) | fil_p['상품명'].str.contains(search, na=False)]
        
        if not fil_p.empty:
            cat_sum = fil_p.groupby('유형').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
            cat_sum['총합계'] = cat_sum['실입금액'] + cat_sum['선급금액']
            st.write(f"#### 📈 {y}년 {m if m != '전체' else ''} 유형별 요약")
            st.table(cat_sum.style.format({'실입금액': '{:,.2f}', '선급금액': '{:,.2f}', '총합계': '{:,.2f}'}))
        
        st.divider()
        st.subheader("📊 발주번호별 정산 및 미수금 현황")
        p_agg = p_all.groupby('발주번호').agg({'실입금액':'sum'}).reset_index()
        sum_df = pd.merge(o_all, p_agg, on='발주번호', how='left').fillna(0)
        sum_df['잔액'] = sum_df['발주총액'] - sum_df['실입금액']
        
        sum_df['상태'] = sum_df['마감여부'].apply(lambda x: "✅ 마감완료" if x == 1 else "⏳ 진행중")
        
        def highlight_closed(row):
            if row['상태'] == '✅ 마감완료':
                return ['background-color: #f0f2f6; color: #a0aab2'] * len(row)
            return [''] * len(row)
            
        disp_sum = sum_df[['발주번호', '발주차수', '상태', '거래처명', '상품명', '발주총액', '실입금액', '잔액', '통화']].sort_values('상태', ascending=False)
        st.dataframe(disp_sum.style.apply(highlight_closed, axis=1).format({'발주총액': '{:,.2f}', '실입금액': '{:,.2f}', '잔액': '{:,.2f}'}), use_container_width=True, hide_index=True)

        st.divider()
        st.subheader("📑 상세 리스트 편집 및 삭제")
        
        ex_db = pd.read_sql("SELECT * FROM exchange_rates", conn)
        ex_db['ym'] = pd.to_datetime(ex_db['날짜']).dt.strftime('%Y-%m')
        m_rates = ex_db.groupby('ym').agg({'usd': lambda x: x[x>0].mean(), 'cny': lambda x: x[x>0].mean()}).to_dict('index')

        def calc_krw_estimate(row):
            if row['통화'] == '한화': return row['실입금액'] + row['선급금액']
            ym = str(row['입금일'])[:7]
            curr = row['통화'].lower()
            if ym in m_rates and m_rates[ym][curr] > 0:
                rate = m_rates[ym][curr]
            else:
                rate = get_realtime_rate(row['통화'])
            return (row['실입금액'] + row['선급금액']) * rate

        fil_p_merged = pd.merge(fil_p, o_all[['발주번호', '발주차수', '발주총액']], on='발주번호', how='left')
        fil_p_merged['예상환산액(KRW)'] = fil_p_merged.apply(calc_krw_estimate, axis=1)
        
        cols = list(fil_p_merged.columns)
        if 'dt' in cols: cols.remove('dt')
        
        for hide_col in ['은행', '계좌번호', '예금주']:
            if hide_col in cols: cols.remove(hide_col)
            
        if '발주차수' in cols: cols.remove('발주차수')
        if '발주총액' in cols: cols.remove('발주총액')
        if '예상환산액(KRW)' in cols: cols.remove('예상환산액(KRW)')
        
        idx = cols.index('발주번호') if '발주번호' in cols else 0
        cols.insert(idx+1, '발주차수')
        cols.insert(idx+2, '발주총액') 
        cols.append('예상환산액(KRW)')
        
        fil_p_merged = fil_p_merged[cols]
        
        ed_p = st.data_editor(fil_p_merged.sort_values('입금일', ascending=False), hide_index=True, use_container_width=True, 
                              disabled=["id", "발주차수", "발주총액", "예상환산액(KRW)"],
                              column_config={"예상환산액(KRW)": st.column_config.NumberColumn("예상환산액(KRW)", format="₩ %d")})
        
        eb1, eb2 = st.columns([1, 4])
        with eb1:
            if st.button("💾 상세 수정 저장"):
                for _, r in ed_p.iterrows():
                    conn.execute("UPDATE payments SET 발주번호=?, 입금일=?, 유형=?, 거래처명=?, 상품명=?, 실입금액=?, 선급금액=?, 메모=? WHERE id=?", 
                                 (r['발주번호'], r['입금일'], r['유형'], r['거래처명'], r['상품명'], r['실입금액'], r['선급금액'], r['메모'], r['id']))
                conn.commit()
                st.success("저장 완료!")
                st.rerun()
        with eb2:
            with st.form("delete_form_v124", clear_on_submit=True):
                col_d1, col_d2 = st.columns([2, 1])
                del_id = col_d1.number_input("삭제할 ID 번호 입력", min_value=0, step=1)
                if col_d2.form_submit_button("🗑️ 해당 ID 삭제"):
                    conn.execute(f"DELETE FROM payments WHERE id={del_id}")
                    conn.commit()
                    st.success("삭제 완료!")
                    st.rerun()

        st.divider()
        st.markdown(f"### 💰 현재 검색 목록 총 합계 ({len(fil_p_merged)}건)")
        
        tot_krw_all = fil_p_merged['예상환산액(KRW)'].sum()
        tot_krw = fil_p_merged[fil_p_merged['통화']=='한화']['실입금액'].sum() + fil_p_merged[fil_p_merged['통화']=='한화']['선급금액'].sum()
        tot_usd = fil_p_merged[fil_p_merged['통화']=='USD']['실입금액'].sum() + fil_p_merged[fil_p_merged['통화']=='USD']['선급금액'].sum()
        tot_cny = fil_p_merged[fil_p_merged['통화']=='CNY']['실입금액'].sum() + fil_p_merged[fil_p_merged['통화']=='CNY']['선급금액'].sum()
        
        mc1, mc2, mc3, mc4 = st.columns(4)
        mc1.metric("총 예상 환산액", f"₩ {tot_krw_all:,.0f}")
        mc2.metric("순수 한화(KRW) 합계", f"₩ {tot_krw:,.0f}")
        mc3.metric("USD 입금 합계", f"$ {tot_usd:,.2f}")
        mc4.metric("CNY 입금 합계", f"¥ {tot_cny:,.2f}")

# ------------------------------------------
# [Tab 4] 거래처 관리
# ------------------------------------------
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    cv1, cv2 = st.columns([1.2, 0.8])
    with cv1:
        st.subheader("➕ 신규 거래처 수기 등록")
        with st.form("vn_reg_v124", clear_on_submit=True):
            vn = st.text_input("거래처명")
            vt = st.selectbox("유형", CATEGORIES)
            vc1, vc2, vc3 = st.columns(3)
            vb = vc1.text_input("은행")
            va = vc2.text_input("계좌번호")
            vh = vc3.text_input("예금주")
            if st.form_submit_button("✅ 저장"):
                if vn: 
                    conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (vn,vb,va,vh,vt))
                    conn.commit()
                    st.success("등록 완료!")
                    st.rerun()
    with cv2:
        st.subheader("📂 일괄 업로드")
        v_tmp = pd.DataFrame(columns=["거래처명", "은행", "계좌번호", "예금주", "기본유형"])
        st.download_button(label="📥 거래처 양식 다운로드", data=v_tmp.to_csv(index=False).encode('utf-8-sig'), file_name='vendor_template.csv')
        vf = st.file_uploader("거래처 CSV", type=['csv'])
        if vf and st.button("🚀 업로드 실행"):
            v_up = pd.read_csv(vf)
            for _, r in v_up.iterrows(): 
                conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", 
                             (r['거래처명'], r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
            conn.commit()
            st.success("업로드 완료!")
            st.rerun()
            
    st.divider()
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    if not v_data.empty:
        st.subheader("🏢 거래처 정보 관리 및 소급 적용")
        orig_v = v_data['거래처명'].tolist()
        ev_v = st.data_editor(v_data, hide_index=True, use_container_width=True)
        if st.button("💾 거래처 동기화 저장"):
            for idx, r in ev_v.iterrows():
                old_n, new_n = orig_v[idx], r['거래처명']
                if old_n != new_n:
                    conn.execute(f"DELETE FROM vendors WHERE 거래처명='{old_n}'")
                    conn.execute("INSERT INTO vendors VALUES (?,?,?,?,?)", (new_n, r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
                    conn.execute("UPDATE orders SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_n, r['기본유형'], old_n))
                    conn.execute("UPDATE payments SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_n, r['기본유형'], old_n))
                else: 
                    conn.execute("UPDATE vendors SET 은행=?, 계좌번호=?, 예금주=?, 기본유형=? WHERE 거래처명=?", 
                                 (r['은행'], r['계좌번호'], r['예금주'], r['기본유형'], r['거래처명']))
            conn.commit()
            st.success("동기화 완료!")
            st.rerun()

# ------------------------------------------
# [Tab 5] 환율 관리
# ------------------------------------------
with tabs[5]:
    st.header("📈 환율 정밀 분석 (Investing.com 연동)")
    
    cu1, cu2 = st.columns(2)
    with cu1:
        f_usd = st.file_uploader("USD/KRW CSV 업로드", type=['csv'], key="usd_up")
        if f_usd and st.button("📥 USD 환율 업데이트"):
            if process_exchange_csv(f_usd, "USD"): 
                st.success("USD 데이터 반영 완료")
                st.rerun()
    with cu2:
        f_cny = st.file_uploader("CNY/KRW CSV 업로드", type=['csv'], key="cny_up")
        if f_cny and st.button("📥 CNY 환율 업데이트"):
            if process_exchange_csv(f_cny, "CNY"): 
                st.success("CNY 데이터 반영 완료")
                st.rerun()

    st.divider()
    ex_db = pd.read_sql("SELECT * FROM exchange_rates ORDER BY 날짜 ASC", conn)
    
    if not ex_db.empty:
        ex_db['dt'] = pd.to_datetime(ex_db['날짜'])
        ex_db['ym'] = ex_db['dt'].dt.strftime('%Y-%m')
        
        m_mean = ex_db.groupby('ym').agg({
            'usd': lambda x: x[x > 0].mean(),
            'cny': lambda x: x[x > 0].mean()
        }).reset_index().fillna(0)
        
        m_mean['year'] = m_mean['ym'].str[:4].astype(int)
        m_mean['month'] = m_mean['ym'].str[5:].astype(int)
        
        st.subheader("📉 월별 평균 환율 추이 (범위 고정)")
        cc1, cc2 = st.columns(2)
        with cc1:
            st.write("**[USD] 1360~1540 (20단위)**")
            fig_u = go.Figure()
            fig_u.add_trace(go.Scatter(x=m_mean['ym'], y=m_mean['usd'], mode='lines+markers', name='USD'))
            fig_u.update_layout(yaxis=dict(range=[1360, 1540], dtick=20), height=350, template="plotly_white", margin=dict(l=10,r=10,t=10,b=10))
            st.plotly_chart(fig_u, use_container_width=True)
            
        with cc2:
            st.write("**[CNY] 186~226 (2단위)**")
            fig_c = go.Figure()
            fig_c.add_trace(go.Scatter(x=m_mean['ym'], y=m_mean['cny'], mode='lines+markers', name='CNY', line=dict(color='orange')))
            fig_c.update_layout(yaxis=dict(range=[186, 226], dtick=2), height=350, template="plotly_white", margin=dict(l=10,r=10,t=10,b=10))
            st.plotly_chart(fig_c, use_container_width=True)

        def get_all_months_report(df, col):
            years_list = sorted(df['year'].unique(), reverse=True)
            if not years_list: return pd.DataFrame()
            
            curr_y = years_list[0]
            prev_y = years_list[1] if len(years_list) > 1 else None
            
            res = pd.DataFrame({'월': range(1, 13)})
            
            c_data = df[df['year'] == curr_y].set_index('month')[col]
            
            if prev_y:
                p_data = df[df['year'] == prev_y].set_index('month')[col]
                res[f'{prev_y}년 평균'] = res['월'].map(p_data)
            
            res[f'{curr_y}년 평균'] = res['월'].map(c_data)
            
            if prev_y:
                def calc_yoy(row):
                    cy, py = row[f'{curr_y}년 평균'], row[f'{prev_y}년 평균']
                    if pd.notnull(cy) and pd.notnull(py) and py > 0:
                        d = cy - py
                        p = (d / py) * 100
                        return f"{d:+.2f}({p:+.1f}%)"
                    return "-"
                res['전년비(YoY)'] = res.apply(calc_yoy, axis=1)
            
            df_sorted = df.sort_values('ym').copy()
            df_sorted['diff'] = df_sorted[col].diff()
            df_sorted['prev'] = df_sorted[col].shift(1)
            
            mom_map = df_sorted[df_sorted['year'] == curr_y].set_index('month')
            
            def calc_mom(row):
                m = row['월']
                if m in mom_map.index:
                    d, v = mom_map.loc[m, 'diff'], mom_map.loc[m, 'prev']
                    if pd.notnull(d) and pd.notnull(v) and v > 0:
                        p = (d / v) * 100
                        return f"{d:+.2f}({p:+.1f}%)"
                return "-"
            res['전월비(MoM)'] = res.apply(calc_mom, axis=1)
            
            return res[res[f'{curr_y}년 평균'].notnull() | (res[f'{prev_y}년 평균'].notnull() if prev_y else False)].reset_index(drop=True)

        st.divider()
        st.subheader("📅 연도별 병렬 분석 리포트 (YoY & MoM 통합)")
        rc1, rc2 = st.columns(2)
        with rc1:
            st.write("#### 💵 USD 환율 분석")
            u_res = get_all_months_report(m_mean, 'usd')
            if not u_res.empty: 
                st.table(u_res.style.format({'월':'{:.0f}월'}, na_rep="-"))
        with rc2:
            st.write("#### 💴 CNY 환율 분석")
            c_res = get_all_months_report(m_mean, 'cny')
            if not c_res.empty: 
                st.table(c_res.style.format({'월':'{:.0f}월'}, na_rep="-"))
    else:
        st.warning("등록된 환율 데이터가 없습니다. 상단에서 CSV를 업로드하세요.")