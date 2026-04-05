import streamlit as st
import pandas as pd
import sqlite3
import os
import shutil
import re
from datetime import datetime, timedelta
import plotly.graph_objects as go

# --- 1. 백업 및 데이터베이스 설정 ---
def run_backup():
    """안정적인 운영을 위해 매일 첫 접속 시 데이터베이스 백업 생성"""
    if not os.path.exists('backups'):
        os.makedirs('backups')
    db_file = 'finance_final_v110.db'
    today_str = datetime.now().strftime('%Y%m%d')
    backup_file = f"backups/backup_{today_str}.db"
    
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

# 레이아웃 및 페이지 설정
st.set_page_config(page_title="자금 관리 시스템 v110", layout="wide", page_icon="💰")
run_backup()

@st.cache_resource
def get_db_connection():
    """모든 테이블 스키마 정의 (단 하나도 생략 없음)"""
    conn = sqlite3.connect('finance_final_v110.db', check_same_thread=False)
    c = conn.cursor()
    # [1] 거래처 마스터
    c.execute('''CREATE TABLE IF NOT EXISTS vendors 
                 (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT, 기본유형 TEXT)''')
    # [2] 발주 마스터
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (발주번호 TEXT PRIMARY KEY, 발주일 TEXT, 발주차수 TEXT, 거래처명 TEXT, 
                  상품명 TEXT, 유형 TEXT, 통화 TEXT, 발주총액 REAL, 마감여부 INTEGER DEFAULT 0)''')
    # [3] 입금 및 지출 상세 내역 (13개 핵심 컬럼)
    c.execute('''CREATE TABLE IF NOT EXISTS payments 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 발주번호 TEXT, 입금일 TEXT, 
                  유형 TEXT, 거래처명 TEXT, 상품명 TEXT, 통화 TEXT,
                  실입금액 REAL, 선급금액 REAL, 메모 TEXT, 한화환산액 REAL,
                  은행 TEXT, 계좌번호 TEXT, 예금주 TEXT)''')
    # [4] 환율 관리 테이블
    c.execute('''CREATE TABLE IF NOT EXISTS exchange_rates 
                 (날짜 TEXT PRIMARY KEY, usd REAL, cny REAL)''')
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비", "원단비", "기타"]

# --- 2. 세션 상태 관리 (업로드 리셋용) ---
if 'order_up_key' not in st.session_state:
    st.session_state.order_up_key = 0
if 'pay_up_key' not in st.session_state:
    st.session_state.pay_up_key = 1000

# --- 3. 유틸리티 함수 (데이터 정제) ---
def to_float(val):
    """숫자 변환 (쉼표 제거 및 예외처리)"""
    try:
        if val is None or pd.isna(val) or str(val).strip() == "":
            return 0.0
        return float(str(val).replace(',', ''))
    except:
        return 0.0

def to_str(val):
    """문자열 정제 (공백 및 특수값 처리)"""
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    return "" if s.lower() in ["nan", "none", ""] else s

def smart_date(date_str):
    """다양한 날짜 형식을 YYYY-MM-DD로 통일"""
    try:
        ds = to_str(date_str).replace(" ", "").replace(".", "-")
        if not ds:
            return datetime.now().strftime("%Y-%m-%d")
        return pd.to_datetime(ds).strftime("%Y-%m-%d")
    except:
        return datetime.now().strftime("%Y-%m-%d")

# --- 4. 분석 및 데이터 처리 엔진 ---
def process_exchange_csv(file, currency_type):
    """Investing.com CSV 환율 데이터를 DB에 반영"""
    try:
        df = pd.read_csv(file)
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

def process_ecount_v110(file):
    """ERP 발주서 엑셀 분석 및 마스터 등록"""
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
            return False, f"⚠️ '{vendor_raw}' 미등록 업체"
        
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
    except:
        return False, "❗ 발주서 분석 오류"

# --- 5. 메인 UI 구성 ---
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리", "📈 환율 관리"])

# [Tab 0] 입금 수기 입력 (v91 기반 상세 필드 복구)
with tabs[0]:
    st.header("📝 입금 내역 수기 입력")
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    o_active = pd.read_sql("SELECT 발주번호 FROM orders WHERE 마감여부=0", conn)
    with st.form("manual_pay_form_v110", clear_on_submit=True):
        c1, c2 = st.columns(2)
        p_oid = c1.selectbox("🔗 진행중 발주번호 연동", ["없음"] + list(o_active['발주번호']))
        p_date = c2.date_input("입금일", value=datetime.now())
        
        c3, c4, c5 = st.columns(3)
        p_vn = c3.selectbox("거래처명", ["선택"] + list(v_data['거래처명']))
        p_ct = c4.selectbox("유형", CATEGORIES)
        p_pr = c5.text_input("상품명")
        
        c6, c7, c8 = st.columns(3)
        p_dep = c6.number_input("실입금액", format="%.2f")
        p_pre = c7.number_input("선급금액", format="%.2f")
        p_cur = c8.selectbox("통화", ["한화", "USD", "CNY"])
        
        p_memo = st.text_input("메모(송금사유 등)")
        
        if st.form_submit_button("✅ 입금 내역 저장"):
            if p_vn == "선택":
                st.error("거래처를 선택하세요.")
            else:
                rate = 1350.0 if p_cur == "USD" else (190.0 if p_cur == "CNY" else 1.0)
                vi = v_data[v_data['거래처명']==p_vn].iloc[0]
                conn.execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                             (to_str(p_oid) if p_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), p_ct, p_vn, p_pr, p_cur, p_dep, p_pre, p_memo, (p_dep+p_pre)*rate, vi['은행'], vi['계좌번호'], vi['예금주']))
                conn.commit(); st.success("성공적으로 저장되었습니다."); st.rerun()

# [Tab 1] 입금 엑셀 업로드 (양식 다운로드 포함)
with tabs[1]:
    st.header("📂 통합 입금 엑셀 업로드")
    pay_template = pd.DataFrame(columns=["발주번호", "거래처", "유형", "상품명", "입금일", "실입금액", "선급금액", "송금사유"])
    st.download_button(label="📥 입금 업로드 샘플 양식 다운로드", data=pay_template.to_csv(index=False).encode('utf-8-sig'), file_name='payment_template.csv', mime='text/csv')
    
    f_p = st.file_uploader("입금 CSV 선택", type=['csv'], key=f"pay_up_{st.session_state.pay_up_key}")
    if f_p and st.button("🚀 데이터 일괄 저장 실행"):
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

# [Tab 2] 발주서 등록 및 마감 (v91 원본 상세 로직 복구)
with tabs[2]:
    st.header("📥 발주서 등록 및 마감 관리")
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("⚡ 엑셀 일괄 등록")
        ord_tmp = pd.DataFrame(columns=["발주번호", "발주일", "발주차수", "거래처명", "상품명", "금액", "통화"])
        st.download_button("📥 수기용 발주 양식 다운로드", ord_tmp.to_csv(index=False).encode('utf-8-sig'), "order_manual_template.csv")
        of_list = st.file_uploader("발주서(xlsx) 선택", type=['xlsx'], accept_multiple_files=True, key=f"ord_{st.session_state.order_up_key}")
        if of_list and st.button("🚀 모든 발주서 일괄 등록"):
            for of in of_list: process_ecount_v110(of)
            st.success("등록 완료!"); st.session_state.order_up_key += 1; st.rerun()
    with col2:
        st.subheader("✍️ 수기 발주 등록")
        v_list = pd.read_sql("SELECT 거래처명 FROM vendors", conn)
        with st.form("manual_order_v110"):
            mi, m_step = st.text_input("발주번호"), st.text_input("발주차수")
            md = st.date_input("발주일")
            mv = st.selectbox("거래처 선택", ["선택"] + list(v_list['거래처명']))
            mp = st.text_input("상품명")
            mt, m_cur = st.number_input("금액", format="%.2f"), st.selectbox("통화", ["한화", "USD", "CNY"])
            if st.form_submit_button("✅ 수기 저장"):
                if mi and mv != "선택":
                    vt = pd.read_sql(f"SELECT 기본유형 FROM vendors WHERE 거래처명='{mv}'", conn).iloc[0]['기본유형']
                    conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,0)", (mi, md.strftime("%Y-%m-%d"), m_step, mv, mp, vt, m_cur, mt))
                    conn.commit(); st.success("수기 등록 완료!"); st.rerun()
    st.divider()
    o_data = pd.read_sql("SELECT * FROM orders", conn)
    if not o_data.empty:
        st.subheader("📄 발주 리스트 및 마감 관리")
        show_c = st.checkbox("마감된 발주 포함해서 보기", value=False)
        disp_o = o_data if show_c else o_data[o_data['마감여부'] == 0]
        ev_o = st.data_editor(disp_o.sort_values('발주일', ascending=False), hide_index=True, use_container_width=True, disabled=["발주번호"], column_config={"마감여부": st.column_config.CheckboxColumn("마감")})
        if st.button("💾 정보 업데이트 및 모든 상세내역 소급 적용"):
            for _, r in ev_o.iterrows():
                conn.execute("UPDATE orders SET 발주일=?, 발주차수=?, 거래처명=?, 상품명=?, 유형=?, 통화=?, 발주총액=?, 마감여부=? WHERE 발주번호=?", (r['발주일'], r['발주차수'], r['거래처명'], r['상품명'], r['유형'], r['통화'], r['발주총액'], int(r['마감여부']), r['발주번호']))
                conn.execute("UPDATE payments SET 거래처명=?, 유형=?, 상품명=?, 통화=? WHERE 발주번호=?", (r['거래처명'], r['유형'], r['상품명'], r['통화'], r['발주번호']))
            conn.commit(); st.success("동기화 완료!"); st.rerun()

# [Tab 3] 상세내역 및 통합 정산 (v91 원본 필터/검색 복구)
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
            cat_sum = fil_p.groupby('유형').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
            cat_sum['총합계'] = cat_sum['실입금액'] + cat_sum['선급금액']
            st.table(cat_sum.style.format('{:,.2f}'))
        
        st.divider(); st.subheader("📊 발주번호별 정산 및 미수금 현황")
        p_agg = p_all.groupby('발주번호').agg({'실입금액':'sum'}).reset_index()
        sum_df = pd.merge(o_all, p_agg, on='발주번호', how='left').fillna(0)
        sum_df['잔액'] = sum_df['발주총액'] - sum_df['실입금액']
        st.dataframe(sum_df[['발주번호', '거래처명', '상품명', '발주총액', '실입금액', '잔액', '통화']], use_container_width=True)

        st.divider(); st.subheader("📑 상세 리스트 편집")
        ed_p = st.data_editor(fil_p.drop(columns=['dt']).sort_values('입금일', ascending=False), hide_index=True, use_container_width=True, disabled=["id"])
        if st.button("💾 상세 내역 수정 저장"):
            for _, r in ed_p.iterrows():
                conn.execute("UPDATE payments SET 발주번호=?, 입금일=?, 유형=?, 거래처명=?, 상품명=?, 실입금액=?, 선급금액=?, 메모=? WHERE id=?", (r['발주번호'], r['입금일'], r['유형'], r['거래처명'], r['상품명'], r['실입금액'], r['선급금액'], r['메모'], r['id']))
            conn.commit(); st.success("저장 완료!"); st.rerun()

# [Tab 4] 거래처 관리 (수기 등록 폼 복구)
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    cv1, cv2 = st.columns([1.2, 0.8])
    with cv1:
        st.subheader("➕ 신규 거래처 수기 등록")
        with st.form("vendor_reg_form_v110", clear_on_submit=True):
            vn, vt = st.text_input("거래처명"), st.selectbox("유형", CATEGORIES)
            vc1, vc2, vc3 = st.columns(3)
            vb, va, vh = vc1.text_input("은행"), vc2.text_input("계좌"), vc3.text_input("예금주")
            if st.form_submit_button("✅ 거래처 저장"):
                if vn: conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (vn,vb,va,vh,vt)); conn.commit(); st.success("완료!"); st.rerun()
    with cv2:
        st.subheader("📂 거래처 일괄 업로드")
        v_tmp = pd.DataFrame(columns=["거래처명", "은행", "계좌번호", "예금주", "기본유형"])
        st.download_button(label="📥 거래처 양식 다운로드", data=v_tmp.to_csv(index=False).encode('utf-8-sig'), file_name='vendor_template.csv')
        vf = st.file_uploader("거래처 CSV 업로드", type=['csv'])
        if vf and st.button("🚀 업로드"):
            v_up = pd.read_csv(vf)
            for _, r in v_up.iterrows(): conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (r['거래처명'], r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
            conn.commit(); st.success("완료"); st.rerun()
    st.divider(); v_data = pd.read_sql("SELECT * FROM vendors", conn)
    if not v_data.empty:
        orig_v = v_data['거래처명'].tolist(); ev_v = st.data_editor(v_data, hide_index=True, use_container_width=True)
        if st.button("💾 거래처명 동기화 저장"):
            for idx, r in ev_v.iterrows():
                old_n, new_n = orig_v[idx], r['거래처명']
                if old_n != new_n:
                    conn.execute(f"DELETE FROM vendors WHERE 거래처명='{old_n}'")
                    conn.execute("INSERT INTO vendors VALUES (?,?,?,?,?)", (new_n, r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
                    conn.execute("UPDATE orders SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_n, r['기본유형'], old_n))
                    conn.execute("UPDATE payments SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_n, r['기본유형'], old_n))
                else: conn.execute("UPDATE vendors SET 은행=?, 계좌번호=?, 예금주=?, 기본유형=? WHERE 거래처명=?", (r['은행'], r['계좌번호'], r['예금주'], r['기본유형'], r['거래처명']))
            conn.commit(); st.success("동기화 완료!"); st.rerun()

# [Tab 5] 환율 관리 (오류 수정 및 % 표시 반영 최종판)
with tabs[5]:
    st.header("📈 환율 정밀 분석 (Investing.com 연동)")
    cu1, cu2 = st.columns(2)
    with cu1:
        f_usd = st.file_uploader("USD/KRW CSV", type=['csv'], key="usd_up")
        if f_usd and st.button("📥 USD 업데이트"):
            if process_exchange_csv(f_usd, "USD"): st.success("USD 데이터 반영 완료"); st.rerun()
    with cu2:
        f_cny = st.file_uploader("CNY/KRW CSV", type=['csv'], key="cny_up")
        if f_cny and st.button("📥 CNY 업데이트"):
            if process_exchange_csv(f_cny, "CNY"): st.success("CNY 데이터 반영 완료"); st.rerun()

    st.divider()
    ex_db = pd.read_sql("SELECT * FROM exchange_rates ORDER BY 날짜 ASC", conn)
    if not ex_db.empty:
        ex_db['dt'] = pd.to_datetime(ex_db['날짜'])
        ex_db['연월'] = ex_db['dt'].dt.strftime('%Y-%m')
        
        # [1] 월별 평균 가공 (0 제외 평균)
        monthly_mean = ex_db.groupby('연월').agg({
            'usd': lambda x: x[x > 0].mean(),
            'cny': lambda x: x[x > 0].mean()
        }).reset_index().fillna(0)
        
        # [2] 차트 분리 표시
        st.subheader("📉 월별 평균 환율 추이 (범위 고정)")
        cc1, cc2 = st.columns(2)
        with cc1:
            fig_u = go.Figure(); fig_u.add_trace(go.Scatter(x=monthly_mean['연월'], y=monthly_mean['usd'], mode='lines+markers', name='USD'))
            fig_u.update_layout(yaxis=dict(range=[1360, 1540], dtick=20), height=350, template="plotly_white")
            st.plotly_chart(fig_u, use_container_width=True)
        with cc2:
            fig_c = go.Figure(); fig_c.add_trace(go.Scatter(x=monthly_mean['연월'], y=monthly_mean['cny'], mode='lines+markers', name='CNY', line=dict(color='orange')))
            fig_c.update_layout(yaxis=dict(range=[186, 226], dtick=2), height=350, template="plotly_white")
            st.plotly_chart(fig_c, use_container_width=True)

        # [3] 리포트용 증감 분석 로직 (전체 시계열 기반)
        monthly_mean['year'] = monthly_mean['연월'].str[:4].astype(int)
        monthly_mean['month'] = monthly_mean['연월'].str[5:].astype(int)
        monthly_mean['usd_mom_val'] = monthly_mean['usd'].diff()
        monthly_mean['cny_mom_val'] = monthly_mean['cny'].diff()

        def get_final_robust_report(df, col_name):
            years_list = sorted(df['year'].unique(), reverse=True)
            if not years_list: return pd.DataFrame()
            
            # 피벗을 통해 연도별 매칭
            pivot_df = df.pivot(index='month', columns='year', values=col_name).sort_index()
            curr_year = years_list[0]
            prev_year = years_list[1] if len(years_list) > 1 else None
            
            report = pd.DataFrame(index=pivot_df.index)
            if prev_year: report[f'{prev_year}년 평균'] = pivot_df[prev_year]
            report[f'{curr_year}년 평균'] = pivot_df[curr_year]
            
            # YoY % 계산
            if prev_year:
                y_diff = report[f'{curr_year}년 평균'] - report[f'{prev_year}년 평균']
                y_pct = (y_diff / report[f'{prev_year}년 평균']) * 100
                report['전년비(YoY)'] = [f"{d:+.2f}({p:+.1f}%)" if pd.notnull(d) and pd.notnull(p) else "-" for d, p in zip(y_diff, y_pct)]
            
            # MoM % 계산 (전체 시계열 맵핑으로 1월 오류 해결)
            mom_map = df[df['year'] == curr_year].set_index('month')
            m_diff = mom_map[f'{col_name}_mom_val']
            prev_val = mom_map[col_name] - m_diff
            m_pct = (m_diff / prev_val) * 100
            report['전월비(MoM)'] = [f"{d:+.2f}({p:+.1f}%)" if pd.notnull(d) and v > 0 else "-" for d, p, v in zip(m_diff, m_pct, prev_val)]
            
            return report.reset_index()

        st.divider(); st.subheader("📅 연도별 병렬 분석 리포트 (YoY & MoM)")
        rc1, rc2 = st.columns(2)
        with rc1:
            st.write("#### 💵 USD 환율 분석")
            usd_res = get_final_robust_report(monthly_mean, 'usd')
            if not usd_res.empty: st.table(usd_res.style.format({'month':'{:.0f}월', f'{sorted(monthly_mean["year"].unique(), reverse=True)[0]}년 평균':'{:,.2f}'}, na_rep="-"))
        with rc2:
            st.write("#### 💴 CNY 환율 분석")
            cny_res = get_final_robust_report(monthly_mean, 'cny')
            if not cny_res.empty: st.table(cny_res.style.format({'month':'{:.0f}월', f'{sorted(monthly_mean["year"].unique(), reverse=True)[0]}년 평균':'{:,.2f}'}, na_rep="-"))
    else:
        st.warning("등록된 환율 데이터가 없습니다. 상단 CSV 업로드 기능을 이용해 주세요.")