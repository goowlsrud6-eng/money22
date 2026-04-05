import streamlit as st
import pandas as pd
import sqlite3
import os
import shutil
import re
from datetime import datetime, timedelta

# --- 1. 백업 및 데이터베이스 설정 ---
def run_backup():
    if not os.path.exists('backups'):
        os.makedirs('backups')
    db_file = 'finance_final_v100.db'
    backup_file = f"backups/backup_{datetime.now().strftime('%Y%m%d')}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

st.set_page_config(page_title="자금 관리 v100", layout="wide", page_icon="💰")
run_backup()

@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_final_v100.db', check_same_thread=False)
    c = conn.cursor()
    # 거래처 마스터
    c.execute('CREATE TABLE IF NOT EXISTS vendors (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT, 기본유형 TEXT)')
    # 발주 마스터
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (발주번호 TEXT PRIMARY KEY, 발주일 TEXT, 발주차수 TEXT, 거래처명 TEXT, 
                  상품명 TEXT, 유형 TEXT, 통화 TEXT, 발주총액 REAL, 마감여부 INTEGER DEFAULT 0)''')
    # 입금 및 지출 상세 내역 (13개 컬럼 유지)
    c.execute('''CREATE TABLE IF NOT EXISTS payments 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 발주번호 TEXT, 입금일 TEXT, 
                  유형 TEXT, 거래처명 TEXT, 상품명 TEXT, 통화 TEXT,
                  실입금액 REAL, 선급금액 REAL, 메모 TEXT, 한화환산액 REAL,
                  은행 TEXT, 계좌번호 TEXT, 예금주 TEXT)''')
    # 환율 관리 테이블
    c.execute('CREATE TABLE IF NOT EXISTS exchange_rates (날짜 TEXT PRIMARY KEY, usd REAL, cny REAL)')
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비", "원단비", "기타"]

# --- 2. 세션 상태 관리 ---
if 'order_up_key' not in st.session_state: st.session_state.order_up_key = 0
if 'pay_up_key' not in st.session_state: st.session_state.pay_up_key = 1000

# --- 3. 유틸리티 함수 (Fix 유지) ---
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
        ds = to_str(date_str).replace(" ", "").replace(".", "-")
        if not ds: return datetime.now().strftime("%Y-%m-%d")
        return pd.to_datetime(ds).strftime("%Y-%m-%d")
    except: return datetime.now().strftime("%Y-%m-%d")

# --- 4. 분석 엔진 (Fix 유지) ---
def process_exchange_csv(file, currency_type):
    try:
        df = pd.read_csv(file)
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
    except: return False

def process_ecount_v100(file):
    try:
        df = pd.read_excel(file, header=None)
        raw_oid = str(df.iloc[1, 0]).split(":")[-1].strip() if ":" in str(df.iloc[1,0]) else str(df.iloc[1, 0])
        clean_oid = raw_oid.replace("-", "")
        odate = smart_date(clean_oid[:8])
        vendor_raw = ""
        for i in range(len(df)):
            if "수신" in str(df.iloc[i, 0]): vendor_raw = str(df.iloc[i, 0]).split(":")[-1].strip(); break
        v_master = pd.read_sql("SELECT 거래처명, 기본유형 FROM vendors", conn)
        v_master['clean'] = v_master['거래처명'].apply(lambda x: re.sub(r'\s+', '', str(x)))
        match = v_master[v_master['clean'] == re.sub(r'\s+', '', vendor_raw)]
        if match.empty: return False, f"⚠️ '{vendor_raw}' 미등록 업체"
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
    except: return False, "❗ 발주서 분석 오류"

# --- 5. 메인 UI 구성 ---
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리", "📈 환율 관리"])

# [Tab 0] 입금 수기 입력 (v91 Fix 로직)
with tabs[0]:
    st.header("📝 입금 내역 수기 입력")
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    o_active = pd.read_sql("SELECT 발주번호 FROM orders WHERE 마감여부=0", conn)
    with st.form("p_man_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        p_oid = col1.selectbox("🔗 진행중 발주번호 연동", ["없음"] + list(o_active['발주번호']))
        p_date = col2.date_input("입금일")
        col3, col4, col5 = st.columns(3)
        p_vn = col3.selectbox("거래처명", ["선택"] + list(v_data['거래처명']))
        p_ct = col4.selectbox("유형", CATEGORIES)
        p_pr = col5.text_input("상품명")
        col6, col7, col8 = st.columns(3)
        p_dep = col6.number_input("실입금액", format="%.2f")
        p_pre = col7.number_input("선급금액", format="%.2f")
        p_cur = col8.selectbox("통화", ["한화", "USD", "CNY"])
        p_memo = st.text_input("메모")
        if st.form_submit_button("✅ 입금 저장"):
            if p_vn != "선택":
                rate = 1350.0 if p_cur == "USD" else (190.0 if p_cur == "CNY" else 1.0)
                vi = v_data[v_data['거래처명']==p_vn].iloc[0]
                conn.execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                             (to_str(p_oid) if p_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), p_ct, p_vn, p_pr, p_cur, p_dep, p_pre, p_memo, (p_dep+p_pre)*rate, vi['은행'], vi['계좌번호'], vi['예금주']))
                conn.commit(); st.success("저장 완료!"); st.rerun()

# [Tab 1] 입금 엑셀 업로드 (v91 Fix 로직)
with tabs[1]:
    st.header("📂 통합 입금 엑셀 업로드")
    sample_csv = pd.DataFrame(columns=["발주번호", "거래처", "유형", "상품명", "입금일", "실입금액", "선급금액", "송금사유"])
    st.download_button("📥 샘플 양식 다운로드", sample_csv.to_csv(index=False).encode('utf-8-sig'), "payment_template.csv")
    f_p = st.file_uploader("입금 CSV 선택", type=['csv'], key=f"pay_up_{st.session_state.pay_up_key}")
    if f_p and st.button("🚀 데이터 일괄 저장"):
        try:
            df_p = pd.read_csv(f_p)
            v_l = pd.read_sql("SELECT * FROM vendors", conn); o_l = pd.read_sql("SELECT * FROM orders", conn)
            for _, r in df_p.iterrows():
                oid, vn_raw = to_str(r.get('발주번호')), to_str(r.get('거래처'))
                if not vn_raw and not oid: continue
                pd_s = smart_date(r.get('입금일'))
                if oid and not o_l[o_l['발주번호'] == oid].empty:
                    info = o_l[o_l['발주번호'] == oid].iloc[0]; vn, pc, pp, cur = info['거래처명'], info['유형'], info['상품명'], info['통화']
                else: vn, pc, pp, cur = vn_raw, to_str(r.get('유형')) or "사입", to_str(r.get('상품명')), "한화"
                vi = v_l[v_l['거래처명'] == vn] if vn else pd.DataFrame()
                dep, pre = to_float(r.get('실입금액')), to_float(r.get('선급금액'))
                rt = 1350.0 if cur == "USD" else (190.0 if cur == "CNY" else 1.0)
                conn.execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                             (oid if oid else None, pd_s, pc, vn, pp, cur, dep, pre, to_str(r.get('송금사유')), (dep+pre)*rt, vi.iloc[0]['은행'] if not vi.empty else "", vi.iloc[0]['계좌번호'] if not vi.empty else "", vi.iloc[0]['예금주'] if not vi.empty else ""))
            conn.commit(); st.success("일괄 저장 완료!"); st.session_state.pay_up_key += 1; st.rerun()
        except Exception as e: st.error(f"오류: {e}")

# [Tab 2] 발주서 등록 및 마감 (v91 Fix 로직)
with tabs[2]:
    st.header("📥 발주서 등록 및 마감 관리")
    col_o1, col_o2 = st.columns(2)
    with col_o1:
        of_list = st.file_uploader("발주서(xlsx) 일괄 선택", type=['xlsx'], accept_multiple_files=True, key=f"ord_{st.session_state.order_up_key}")
        if of_list and st.button("🚀 모든 발주서 등록"):
            for of in of_list: process_ecount_v100(of)
            st.success("등록 완료!"); st.session_state.order_up_key += 1; st.rerun()
    st.divider()
    o_data = pd.read_sql("SELECT * FROM orders", conn)
    if not o_data.empty:
        st.subheader("📄 발주 리스트 및 마감 관리")
        show_c = st.checkbox("마감된 발주 포함해서 보기", value=False)
        disp_o = o_data if show_c else o_data[o_data['마감여부'] == 0]
        ev_o = st.data_editor(disp_o[['발주번호', '발주차수', '거래처명', '상품명', '유형', '통화', '발주총액', '마감여부', '발주일']].sort_values('발주일', ascending=False), hide_index=True, use_container_width=True, disabled=["발주번호"])
        if st.button("💾 정보 업데이트 및 모든 상세내역 소급 적용"):
            for _, r in ev_o.iterrows():
                conn.execute("UPDATE orders SET 발주일=?, 발주차수=?, 거래처명=?, 상품명=?, 유형=?, 통화=?, 발주총액=?, 마감여부=? WHERE 발주번호=?", (r['발주일'], r['발주차수'], r['거래처명'], r['상품명'], r['유형'], r['통화'], r['발주총액'], int(r['마감여부']), r['발주번호']))
                conn.execute("UPDATE payments SET 거래처명=?, 유형=?, 상품명=?, 통화=? WHERE 발주번호=?", (r['거래처명'], r['유형'], r['상품명'], r['통화'], r['발주번호']))
            conn.commit(); st.success("동기화 완료!"); st.rerun()

# [Tab 3] 상세내역 및 통합 정산 (v91 Fix 로직)
with tabs[3]:
    st.header("🔍 상세 내역 및 통합 정산")
    p_all = pd.read_sql("SELECT * FROM payments", conn); o_all = pd.read_sql("SELECT * FROM orders", conn)
    if not p_all.empty:
        p_all['dt'] = pd.to_datetime(p_all['입금일'])
        st.subheader("📊 필터 및 검색")
        f1, f2, f3 = st.columns([1, 1, 2])
        y = f1.selectbox("기준 연도", sorted(p_all['dt'].dt.year.unique(), reverse=True))
        m = f2.selectbox("기준 월", ["전체"] + sorted(list(p_all[p_all['dt'].dt.year==y]['dt'].dt.month.unique())))
        search = f3.text_input("업체/상품 통합 검색")
        fil_p = p_all[p_all['dt'].dt.year == y]
        if m != "전체": fil_p = fil_p[fil_p['dt'].dt.month == m]
        if search: fil_p = fil_p[fil_p['거래처명'].str.contains(search, na=False) | fil_p['상품명'].str.contains(search, na=False)]
        
        if not fil_p.empty:
            st.write(f"#### 📈 {y}년 {m if m != '전체' else ''} 유형별 요약")
            cat_sum = fil_p.groupby('유형').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
            cat_sum['총합계'] = cat_sum['실입금액'] + cat_sum['선급금액']
            st.table(cat_sum.style.format({'실입금액':'{:,.2f}', '선급금액':'{:,.2f}', '총합계':'{:,.2f}'}))
        
        st.divider(); st.subheader("📊 발주번호별 정산 및 미수금 현황")
        p_agg = p_all.groupby('발주번호').agg({'실입금액':'sum'}).reset_index()
        sum_df = pd.merge(o_all, p_agg, on='발주번호', how='left').fillna(0)
        sum_df['잔액'] = sum_df['발주총액'] - sum_df['실입금액']
        st.dataframe(sum_df[['발주번호', '거래처명', '상품명', '발주총액', '실입금액', '잔액', '통화']], use_container_width=True)

        st.divider(); st.subheader("📑 상세 리스트 편집")
        ed_p = st.data_editor(fil_p.drop(columns=['dt']).sort_values('입금일', ascending=False), hide_index=True, use_container_width=True, disabled=["id"])
        if st.button("💾 상세 내역 개별 수정 저장"):
            for _, r in ed_p.iterrows():
                conn.execute("UPDATE payments SET 발주번호=?, 입금일=?, 유형=?, 거래처명=?, 상품명=?, 실입금액=?, 선급금액=?, 메모=? WHERE id=?", (r['발주번호'], r['입금일'], r['유형'], r['거래처명'], r['상품명'], r['실입금액'], r['선급금액'], r['메모'], r['id']))
            conn.commit(); st.success("저장 완료!"); st.rerun()

# [Tab 4] 거래처 관리 (v91 Fix 로직)
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    if not v_data.empty:
        orig_v = v_data['거래처명'].tolist()
        ev_v = st.data_editor(v_data, hide_index=True, use_container_width=True)
        if st.button("💾 거래처 정보 통합 동기화 저장"):
            for idx, r in ev_v.iterrows():
                old_n, new_n = orig_v[idx], r['거래처명']
                if old_n != new_n:
                    conn.execute(f"DELETE FROM vendors WHERE 거래처명='{old_n}'")
                    conn.execute("INSERT INTO vendors VALUES (?,?,?,?,?)", (new_n, r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
                    conn.execute("UPDATE orders SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_n, r['기본유형'], old_n))
                    conn.execute("UPDATE payments SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_n, r['기본유형'], old_n))
                else:
                    conn.execute("UPDATE vendors SET 은행=?, 계좌번호=?, 예금주=?, 기본유형=? WHERE 거래처명=?", (r['은행'], r['계좌번호'], r['예금주'], r['기본유형'], r['거래처명']))
            conn.commit(); st.success("동기화 완료!"); st.rerun()

# [Tab 5] 환율 관리 (사용자 요청: 차트 분리 및 병렬 분석 리포트)
with tabs[5]:
    st.header("📈 환율 정밀 분석 (Investing.com 연동)")
    st.info("USD와 CNY를 분리하여 가시성을 높였으며, 연도별 병렬 비교표를 통해 흐름을 분석합니다.")
    
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        f_usd = st.file_uploader("USD/KRW 과거 데이터 CSV", type=['csv'], key="usd_up")
        if f_usd and st.button("📥 USD 데이터 업데이트"):
            if process_exchange_csv(f_usd, "USD"): st.success("USD 반영 완료"); st.rerun()
    with col_up2:
        f_cny = st.file_uploader("CNY/KRW 과거 데이터 CSV", type=['csv'], key="cny_up")
        if f_cny and st.button("📥 CNY 데이터 업데이트"):
            if process_exchange_csv(f_cny, "CNY"): st.success("CNY 반영 완료"); st.rerun()

    st.divider()
    ex_db = pd.read_sql("SELECT * FROM exchange_rates", conn)
    if not ex_db.empty:
        ex_db['dt'] = pd.to_datetime(ex_db['날짜'])
        ex_db['연도'] = ex_db['dt'].dt.year
        ex_db['월'] = ex_db['dt'].dt.month
        
        # 월별 평균 가공
        monthly_mean = ex_db.groupby(['연도', '월']).agg({
            'usd': lambda x: x[x > 0].mean(),
            'cny': lambda x: x[x > 0].mean()
        }).reset_index().fillna(0)
        monthly_mean['연월'] = monthly_mean['연도'].astype(str) + "-" + monthly_mean['월'].astype(str).str.zfill(2)
        
        # 1. 차트 분리 표시
        st.subheader("📉 월별 평균 환율 추이 (단위별 개별 표시)")
        chart_c1, chart_c2 = st.columns(2)
        with chart_c1:
            st.write("**[USD] 월별 평균 흐름**")
            st.line_chart(monthly_mean.set_index('연월').sort_index()[['usd']])
        with chart_c2:
            st.write("**[CNY] 월별 평균 흐름**")
            st.line_chart(monthly_mean.set_index('연월').sort_index()[['cny']])

        # 2. 병렬 비교 분석표 생성 함수
        def get_enhanced_report(df, col_name):
            years = sorted(df['연도'].unique(), reverse=True)
            pivot_df = df.pivot(index='월', columns='연도', values=col_name).sort_index()
            curr_y = years[0]
            prev_y = years[1] if len(years) > 1 else None
            
            report = pd.DataFrame(index=pivot_df.index)
            if prev_y: report[f'{prev_y}년 평균'] = pivot_df[prev_y]
            report[f'{curr_y}년 평균'] = pivot_df[curr_y]
            if prev_y: report['전년비(YoY)'] = report[f'{curr_y}년 평균'] - report[f'{prev_y}년 평균']
            report['전월비(MoM)'] = report[f'{curr_y}년 평균'].diff()
            return report.reset_index()

        st.divider()
        st.subheader("📅 연도별 병렬 분석 리포트 (YoY & MoM 통합)")
        rep_c1, rep_c2 = st.columns(2)
        with rep_c1:
            st.write("#### 💵 USD 환율 분석")
            usd_rep = get_enhanced_report(monthly_mean, 'usd')
            if not usd_rep.empty:
                st.table(usd_rep.style.format({'월':'{:.0f}월','전년비(YoY)':'{:+.2f}','전월비(MoM)':'{:+.2f}'}, na_rep="-"))
        with rep_c2:
            st.write("#### 💴 CNY 환율 분석")
            cny_rep = get_enhanced_report(monthly_mean, 'cny')
            if not cny_rep.empty:
                st.table(cny_rep.style.format({'월':'{:.0f}월','전년비(YoY)':'{:+.2f}','전월비(MoM)':'{:+.2f}'}, na_rep="-"))
    else:
        st.warning("데이터가 없습니다. 상단 CSV 업로드 기능을 이용해 주세요.")