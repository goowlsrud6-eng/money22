import streamlit as st
import pandas as pd
import sqlite3
import os
import shutil
import re
from datetime import datetime, timedelta

# --- 1. 백업 및 데이터베이스 설정 ---
def run_backup():
    """시스템 시작 시 데이터베이스 백업 생성"""
    if not os.path.exists('backups'):
        os.makedirs('backups')
    db_file = 'finance_final_v98.db'
    backup_file = f"backups/backup_{datetime.now().strftime('%Y%m%d')}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

st.set_page_config(page_title="자금 관리 v98", layout="wide", page_icon="💰")
run_backup()

@st.cache_resource
def get_db_connection():
    """DB 연결 및 테이블 스키마 초기화 (누락 없음)"""
    conn = sqlite3.connect('finance_final_v98.db', check_same_thread=False)
    c = conn.cursor()
    # 거래처 마스터
    c.execute('CREATE TABLE IF NOT EXISTS vendors (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT, 기본유형 TEXT)')
    # 발주 마스터 (마감여부 필드 포함)
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (발주번호 TEXT PRIMARY KEY, 발주일 TEXT, 발주차수 TEXT, 거래처명 TEXT, 
                  상품명 TEXT, 유형 TEXT, 통화 TEXT, 발주총액 REAL, 마감여부 INTEGER DEFAULT 0)''')
    # 입금 및 지출 상세 내역 (13개 핵심 컬럼)
    c.execute('''CREATE TABLE IF NOT EXISTS payments 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 발주번호 TEXT, 입금일 TEXT, 
                  유형 TEXT, 거래처명 TEXT, 상품명 TEXT, 통화 TEXT,
                  실입금액 REAL, 선급금액 REAL, 메모 TEXT, 한화환산액 REAL,
                  은행 TEXT, 계좌번호 TEXT, 예금주 TEXT)''')
    # 환율 관리 테이블 (Investing.com 데이터 저장용)
    c.execute('CREATE TABLE IF NOT EXISTS exchange_rates (날짜 TEXT PRIMARY KEY, usd REAL, cny REAL)')
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비", "원단비", "기타"]

# --- 2. 세션 상태 관리 (업로드 리스트 초기화용 키) ---
if 'order_up_key' not in st.session_state: st.session_state.order_up_key = 0
if 'pay_up_key' not in st.session_state: st.session_state.pay_up_key = 1000

# --- 3. 유틸리티 함수 (빈칸 대응 및 데이터 정제) ---
def to_float(val):
    """숫자 형식 정제 (쉼표 제거 및 NaN 처리)"""
    try:
        if val is None or pd.isna(val) or str(val).strip() == "": return 0.0
        return float(str(val).replace(',', ''))
    except: return 0.0

def to_str(val):
    """문자열 정제 (공백 및 특수값 처리)"""
    if val is None or pd.isna(val): return ""
    s = str(val).strip()
    return "" if s.lower() in ["nan", "none", ""] else s

def smart_date(date_str):
    """Investing.com 및 엑셀 날짜 형식을 표준(YYYY-MM-DD)으로 변환"""
    try:
        ds = to_str(date_str).replace(" ", "").replace(".", "-")
        if not ds: return datetime.now().strftime("%Y-%m-%d")
        return pd.to_datetime(ds).strftime("%Y-%m-%d")
    except: return datetime.now().strftime("%Y-%m-%d")

# --- 4. 핵심 분석 엔진 (환율 및 발주서) ---
def process_exchange_csv(file, currency_type):
    """Investing.com CSV를 읽어 날짜별 종가 업데이트"""
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
        conn.commit()
        return True
    except: return False

def process_ecount_v98(file):
    """이카운트 발주서 엑셀을 분석하여 마스터 테이블에 저장"""
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
        if match.empty: return False, f"⚠️ '{vendor_raw}'은(는) 미등록 업체입니다."
        
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

# [Tab 0] 입금 수기 입력
with tabs[0]:
    st.header("📝 입금 내역 수기 입력")
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    o_active = pd.read_sql("SELECT 발주번호 FROM orders WHERE 마감여부=0", conn)
    with st.form("manual_pay_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        p_oid = c1.selectbox("🔗 진행중 발주번호 연동", ["없음"] + list(o_active['발주번호']))
        p_date = c2.date_input("입금일")
        c3, c4, c5 = st.columns(3)
        p_vn = c3.selectbox("거래처명", ["선택"] + list(v_data['거래처명']))
        p_ct, p_pr = c4.selectbox("유형", CATEGORIES), c5.text_input("상품명")
        c6, c7, c8 = st.columns(3)
        p_dep, p_pre = c6.number_input("실입금액", format="%.2f"), c7.number_input("선급금액", format="%.2f")
        p_cur = c8.selectbox("통화", ["한화", "USD", "CNY"])
        if st.form_submit_button("✅ 입금 저장"):
            rate = 1350.0 if p_cur == "USD" else (190.0 if p_cur == "CNY" else 1.0)
            vi = v_data[v_data['거래처명']==p_vn].iloc[0] if p_vn != "선택" else {}
            conn.execute("INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                         (to_str(p_oid) if p_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), p_ct, p_vn if p_vn != "선택" else "", p_pr, p_cur, p_dep, p_pre, "", (p_dep+p_pre)*rate, vi.get('은행',''), vi.get('계좌번호',''), vi.get('예금주','')))
            conn.commit(); st.success("저장 완료!"); st.rerun()

# [Tab 1] 입금 엑셀 업로드 (샘플 다운로드 포함)
with tabs[1]:
    st.header("📂 통합 입금 엑셀 업로드")
    sample_csv = pd.DataFrame(columns=["발주번호", "거래처", "유형", "상품명", "입금일", "실입금액", "선급금액", "송금사유"])
    st.download_button("📥 입금 샘플 양식 다운로드", sample_csv.to_csv(index=False).encode('utf-8-sig'), "payment_template.csv")
    f_p = st.file_uploader("입금 CSV 선택", type=['csv'], key=f"pay_up_{st.session_state.pay_up_key}")
    if f_p and st.button("🚀 데이터 일괄 저장 실행"):
        df_p = pd.read_csv(f_p)
        o_l = pd.read_sql("SELECT * FROM orders", conn); v_l = pd.read_sql("SELECT * FROM vendors", conn)
        for _, r in df_p.iterrows():
            oid, vn_raw = to_str(r.get('발주번호')), to_str(r.get('거래처'))
            if not vn_raw and not oid: continue
            pd_s = smart_date(r.get('입금일'))
            if oid and not o_l[o_l['발주번호'] == oid].empty:
                i = o_l[o_l['발주번호'] == oid].iloc[0]; vn, pc, pp, cur = i['거래처명'], i['유형'], i['상품명'], i['통화']
            else: vn, pc, pp, cur = vn_raw, to_str(r.get('유형')) or "사입", to_str(r.get('상품명')), "한화"
            vi = v_l[v_l['거래처명'] == vn] if vn else pd.DataFrame()
            dep, pre = to_float(r.get('실입금액')), to_float(r.get('선급금액'))
            rt = 1350.0 if cur == "USD" else (190.0 if cur == "CNY" else 1.0)
            conn.execute("INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)",
                         (oid if oid else None, pd_s, pc, vn, pp, cur, dep, pre, to_str(r.get('송금사유')), (dep+pre)*rt, vi.iloc[0]['은행'] if not vi.empty else "", vi.iloc[0]['계좌번호'] if not vi.empty else "", vi.iloc[0]['예금주'] if not vi.empty else ""))
        conn.commit(); st.success("일괄 저장 완료!"); st.session_state.pay_up_key += 1; st.rerun()

# [Tab 2] 발주서 등록 및 마감 (소급 업데이트 로직 포함)
with tabs[2]:
    st.header("📥 발주서 등록 및 마감 관리")
    col_o1, col_o2 = st.columns(2)
    with col_o1:
        of_list = st.file_uploader("발주서(xlsx) 일괄 선택", type=['xlsx'], accept_multiple_files=True, key=f"ord_{st.session_state.order_up_key}")
        if of_list and st.button("🚀 모든 발주서 일괄 등록"):
            for of in of_list: process_ecount_v98(of)
            st.success("등록 완료!"); st.session_state.order_up_key += 1; st.rerun()
    st.divider()
    o_data = pd.read_sql("SELECT * FROM orders", conn)
    if not o_data.empty:
        st.subheader("📄 발주 리스트 및 마감 관리")
        show_c = st.checkbox("마감된 발주 건 포함해서 보기", value=False)
        disp_o = o_data if show_c else o_data[o_data['마감여부'] == 0]
        ev_o = st.data_editor(disp_o.sort_values('발주일', ascending=False), hide_index=True, use_container_width=True, disabled=["발주번호"])
        if st.button("💾 정보 업데이트 및 모든 상세내역 소급 적용"):
            for _, r in ev_o.iterrows():
                conn.execute("UPDATE orders SET 발주일=?, 발주차수=?, 거래처명=?, 상품명=?, 유형=?, 통화=?, 발주총액=?, 마감여부=? WHERE 발주번호=?", (r['발주일'], r['발주차수'], r['거래처명'], r['상품명'], r['유형'], r['통화'], r['발주총액'], int(r['마감여부']), r['발주번호']))
                conn.execute("UPDATE payments SET 거래처명=?, 유형=?, 상품명=?, 통화=? WHERE 발주번호=?", (r['거래처명'], r['유형'], r['상품명'], r['통화'], r['발주번호']))
            conn.commit(); st.success("동기화 완료!"); st.rerun()

# [Tab 3] 상세내역 및 통합 정산 (필터/검색 및 잔액 현황)
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
            st.write(f"#### 📈 {y}년 {m if m != '전체' else ''} 유형별 입금 요약")
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

# [Tab 4] 거래처 관리 (이름 변경 시 전체 소급 업데이트)
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    if not v_data.empty:
        orig_v = v_data['거래처명'].tolist()
        ev_v = st.data_editor(v_data, hide_index=True, use_container_width=True)
        if st.button("💾 거래처 정보 변경 및 전체 소급 동기화"):
            for idx, r in ev_v.iterrows():
                old_n, new_n = orig_v[idx], r['거래처명']
                if old_n != new_n:
                    conn.execute(f"DELETE FROM vendors WHERE 거래처명='{old_n}'")
                    conn.execute("INSERT INTO vendors VALUES (?,?,?,?,?)", (new_n, r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
                    conn.execute("UPDATE orders SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_n, r['기본유형'], old_n))
                    conn.execute("UPDATE payments SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_n, r['기본유형'], old_n))
                else:
                    conn.execute("UPDATE vendors SET 은행=?, 계좌번호=?, 예금주=?, 기본유형=? WHERE 거래처명=?", (r['은행'], r['계좌번호'], r['예금주'], r['기본유형'], r['거래처명']))
            conn.commit(); st.success("거래처 정보가 전체 업데이트되었습니다."); st.rerun()

# [Tab 5] 환율 관리 (Investing.com 데이터 기반 분석 및 병렬 비교)
with tabs[5]:
    st.header("📈 환율 분석 및 월별 추이 (Investing.com CSV)")
    c_u1, c_u2 = st.columns(2)
    with c_u1:
        f_usd = st.file_uploader("USD/KRW 과거 데이터 CSV 업로드", type=['csv'], key="u_up")
        if f_usd and st.button("📥 USD 환율 일괄 업데이트"):
            if process_exchange_csv(f_usd, "USD"): st.success("USD 데이터 반영 완료"); st.rerun()
    with c_u2:
        f_cny = st.file_uploader("CNY/KRW 과거 데이터 CSV 업로드", type=['csv'], key="c_up")
        if f_cny and st.button("📥 CNY 환율 일괄 업데이트"):
            if process_exchange_csv(f_cny, "CNY"): st.success("CNY 데이터 반영 완료"); st.rerun()

    st.divider()
    ex_db = pd.read_sql("SELECT * FROM exchange_rates", conn)
    if not ex_db.empty:
        ex_db['dt'] = pd.to_datetime(ex_db['날짜'])
        ex_db['연도'] = ex_db['dt'].dt.year
        ex_db['월'] = ex_db['dt'].dt.month
        
        monthly_mean = ex_db.groupby(['연도', '월']).agg({
            'usd': lambda x: x[x > 0].mean(),
            'cny': lambda x: x[x > 0].mean()
        }).reset_index().fillna(0)
        monthly_mean['연월'] = monthly_mean['연도'].astype(str) + "-" + monthly_mean['월'].astype(str).str.zfill(2)
        
        st.subheader("📉 월별 평균 환율 변동 추이 (Trend Chart)")
        st.line_chart(monthly_mean.set_index('연월').sort_index()[['usd', 'cny']])

        def get_analysis_table(df, col):
            res = []
            sorted_df = df.sort_values(['연도', '월'], ascending=False)
            for _, r in sorted_df.iterrows():
                y, m, v = int(r['연도']), int(r['월']), r[col]
                if v == 0: continue
                # MoM / YoY 계산
                pm, pym = (m - 1, y) if m > 1 else (12, y - 1)
                pmv = df[(df['연도'] == pym) & (df['월'] == pm)][col].values
                mom = f"{(v - pmv[0]):+.2f}" if len(pmv) > 0 and pmv[0] > 0 else "-"
                pyv = df[(df['연도'] == y - 1) & (df['월'] == m)][col].values
                yoy = f"{(v - pyv[0]):+.2f}" if len(pyv) > 0 and pyv[0] > 0 else "-"
                res.append({"연도": y, "월": f"{m}월", "평균 환율": v, "전월대비(MoM)": mom, "전년대비(YoY)": yoy})
            return pd.DataFrame(res)

        st.divider()
        at1, at2 = st.columns(2)
        with at1:
            st.write("#### 💵 USD 월별 평균 및 분석")
            usd_tab = get_analysis_table(monthly_mean, 'usd')
            if not usd_tab.empty: st.table(usd_tab.style.format({'평균 환율': '{:,.2f}'}))
        with at2:
            st.write("#### 💴 CNY 월별 평균 및 분석")
            cny_tab = get_analysis_table(monthly_mean, 'cny')
            if not cny_tab.empty: st.table(cny_tab.style.format({'평균 환율': '{:,.2f}'}))

        st.divider(); st.subheader("📅 연도별 환율 병렬 비교 (Yearly Comparison)")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            st.write("**USD 연도별 비교**")
            st.dataframe(monthly_mean.pivot(index='월', columns='연도', values='usd').sort_index().style.format('{:,.2f}'), use_container_width=True)
        with col_p2:
            st.write("**CNY 연도별 비교**")
            st.dataframe(monthly_mean.pivot(index='월', columns='연도', values='cny').sort_index().style.format('{:,.2f}'), use_container_width=True)
    else:
        st.warning("등록된 환율 데이터가 없습니다. 상단 CSV 업로드 기능을 이용해 주세요.")