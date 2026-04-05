import streamlit as st
import pandas as pd
import sqlite3
import os
import shutil
import re
from datetime import datetime, timedelta
import plotly.graph_objects as go

# ==========================================
# 1. 시스템 환경 설정 및 데이터베이스 초기화
# ==========================================
def run_backup():
    """안정적인 운영을 위해 매일 첫 접속 시 DB 백업 파일 생성"""
    if not os.path.exists('backups'):
        os.makedirs('backups')
    db_file = 'finance_final_v107.db'
    # 날짜별 백업 파일명 생성
    today_str = datetime.now().strftime('%Y%m%d')
    backup_file = f"backups/backup_{today_str}.db"
    
    # 원본 파일이 있고 백업이 아직 없는 경우 복사 실행
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

# 페이지 설정
st.set_page_config(
    page_title="자금 관리 시스템 v107", 
    layout="wide", 
    page_icon="💰"
)
run_backup()

@st.cache_resource
def get_db_connection():
    """데이터베이스 연결 및 모든 필수 테이블 스키마 정의"""
    conn = sqlite3.connect('finance_final_v107.db', check_same_thread=False)
    c = conn.cursor()
    
    # [1] 거래처 마스터 테이블
    c.execute('''
        CREATE TABLE IF NOT EXISTS vendors (
            거래처명 TEXT PRIMARY KEY, 
            은행 TEXT, 
            계좌번호 TEXT, 
            예금주 TEXT, 
            기본유형 TEXT
        )
    ''')
    
    # [2] 발주서 마스터 테이블
    c.execute('''
        CREATE TABLE IF NOT EXISTS orders (
            발주번호 TEXT PRIMARY KEY, 
            발주일 TEXT, 
            발주차수 TEXT, 
            거래처명 TEXT, 
            상품명 TEXT, 
            유형 TEXT, 
            통화 TEXT, 
            발주총액 REAL, 
            마감여부 INTEGER DEFAULT 0
        )
    ''')
    
    # [3] 입금 및 지출 상세 내역 테이블 (13개 핵심 컬럼)
    c.execute('''
        CREATE TABLE IF NOT EXISTS payments (
            id INTEGER PRIMARY KEY AUTOINCREMENT, 
            발주번호 TEXT, 
            입금일 TEXT, 
            유형 TEXT, 
            거래처명 TEXT, 
            상품명 TEXT, 
            통화 TEXT,
            실입금액 REAL, 
            선급금액 REAL, 
            메모 TEXT, 
            한화환산액 REAL,
            은행 TEXT, 
            계좌번호 TEXT, 
            예금주 TEXT
        )
    ''')
    
    # [4] 환율 관리 테이블 (Investing.com 데이터 저장용)
    c.execute('''
        CREATE TABLE IF NOT EXISTS exchange_rates (
            날짜 TEXT PRIMARY KEY, 
            usd REAL, 
            cny REAL
        )
    ''')
    
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비", "원단비", "기타"]

# ==========================================
# 2. 세션 상태 관리 (업로드 목록 리셋용)
# ==========================================
if 'order_up_key' not in st.session_state:
    st.session_state.order_up_key = 0
if 'pay_up_key' not in st.session_state:
    st.session_state.pay_up_key = 1000

# ==========================================
# 3. 유틸리티 함수 (데이터 정제 및 변환)
# ==========================================
def to_float(val):
    """문자열/쉼표 포함 숫자를 float로 안전하게 변환"""
    try:
        if val is None or pd.isna(val) or str(val).strip() == "":
            return 0.0
        return float(str(val).replace(',', ''))
    except:
        return 0.0

def to_str(val):
    """빈 값(NaN, None)을 빈 문자열로 정제"""
    if val is None or pd.isna(val):
        return ""
    s = str(val).strip()
    if s.lower() in ["nan", "none", ""]:
        return ""
    return s

def smart_date(date_str):
    """Investing.com 및 엑셀의 다양한 날짜 형식을 YYYY-MM-DD로 통일"""
    try:
        ds = to_str(date_str).replace(" ", "").replace(".", "-")
        if not ds:
            return datetime.now().strftime("%Y-%m-%d")
        
        # "01월07일"과 같은 특수 형식 대응
        if "월" in ds and "일" in ds:
            return datetime.strptime(f"2026 {ds}", "%Y %m월 %d일").strftime("%Y-%m-%d")
            
        return pd.to_datetime(ds).strftime("%Y-%m-%d")
    except:
        return datetime.now().strftime("%Y-%m-%d")

# ==========================================
# 4. 분석 및 데이터 처리 엔진
# ==========================================
def process_exchange_csv(file, currency_type):
    """Investing.com CSV 환율 데이터를 DB에 일괄 반영"""
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

def process_ecount_v107(file):
    """ERP 발주서(xlsx) 분석 및 마스터 등록"""
    try:
        df = pd.read_excel(file, header=None)
        # 발주번호 추출
        raw_oid = str(df.iloc[1, 0]).split(":")[-1].strip() if ":" in str(df.iloc[1,0]) else str(df.iloc[1, 0])
        clean_oid = raw_oid.replace("-", "")
        # 발주번호 날짜 기반 자동 날짜 생성
        odate = smart_date(clean_oid[:8])
        
        # 거래처명 추출 (수신 항목 탐색)
        vendor_raw = ""
        for i in range(len(df)):
            if "수신" in str(df.iloc[i, 0]):
                vendor_raw = str(df.iloc[i, 0]).split(":")[-1].strip()
                break
        
        # 거래처 마스터와 비교 연동
        v_master = pd.read_sql("SELECT 거래처명, 기본유형 FROM vendors", conn)
        v_master['clean'] = v_master['거래처명'].apply(lambda x: re.sub(r'\s+', '', str(x)))
        match = v_master[v_master['clean'] == re.sub(r'\s+', '', vendor_raw)]
        
        if match.empty:
            return False, f"⚠️ '{vendor_raw}' 미등록 업체입니다."
        
        v_type, v_fixed = match.iloc[0]['기본유형'], match.iloc[0]['거래처명']
        f6 = str(df.iloc[5, 5]) if len(df) > 5 else ""
        curr = "USD" if "USD" in f6 else ("CNY" if any(x in f6 for x in ["중국", "CNY"]) else "한화")
        
        # 품목명 및 금액 추출 로직
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
        return False, f"❗ 발주서 형식 분석 중 오류: {e}"

# ==========================================
# 5. 메인 UI (6단 탭 구성)
# ==========================================
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리", "📈 환율 관리"])

# ------------------------------------------
# [Tab 0] 입금 수기 입력 (수정 없음)
# ------------------------------------------
with tabs[0]:
    st.header("📝 입금 내역 수기 입력")
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    o_active = pd.read_sql("SELECT 발주번호 FROM orders WHERE 마감여부=0", conn)
    
    with st.form("manual_payment_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        p_oid = col1.selectbox("🔗 진행중인 발주번호 연동", ["없음"] + list(o_active['발주번호']))
        p_date = col2.date_input("입금일", value=datetime.now())
        
        col3, col4, col5 = st.columns(3)
        p_vn = col3.selectbox("거래처명", ["선택"] + list(v_data['거래처명']))
        p_ct = col4.selectbox("유형", CATEGORIES)
        p_pr = col5.text_input("상품명")
        
        col6, col7, col8 = st.columns(3)
        p_dep = col6.number_input("실입금액", format="%.2f")
        p_pre = col7.number_input("선급금액", format="%.2f")
        p_cur = col8.selectbox("통화", ["한화", "USD", "CNY"])
        
        p_memo = st.text_input("메모(송금사유 등)")
        
        if st.form_submit_button("✅ 입금 내역 저장"):
            if p_vn == "선택":
                st.error("거래처를 반드시 선택해주세요.")
            else:
                rate = 1350.0 if p_cur == "USD" else (190.0 if p_cur == "CNY" else 1.0)
                vi = v_data[v_data['거래처명']==p_vn].iloc[0]
                
                conn.execute('''
                    INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                ''', (to_str(p_oid) if p_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), p_ct, p_vn, p_pr, p_cur, 
                      p_dep, p_pre, p_memo, (p_dep+p_pre)*rate, vi['은행'], vi['계좌번호'], vi['예금주']))
                conn.commit()
                st.success("입금 내역이 저장되었습니다.")
                st.rerun()

# ------------------------------------------
# [Tab 1] 입금 엑셀 업로드 (양식 다운로드 포함)
# ------------------------------------------
with tabs[1]:
    st.header("📂 통합 입금 엑셀 업로드")
    
    # 샘플 양식 데이터 생성
    sample_pay = pd.DataFrame(columns=["발주번호", "거래처", "유형", "상품명", "입금일", "실입금액", "선급금액", "송금사유"])
    st.download_button(
        label="📥 입금 업로드 샘플 양식 다운로드", 
        data=sample_pay.to_csv(index=False).encode('utf-8-sig'), 
        file_name='payment_template.csv',
        mime='text/csv'
    )
    
    pay_file = st.file_uploader("입금 CSV 파일 선택", type=['csv'], key=f"pay_up_{st.session_state.pay_up_key}")
    if pay_file and st.button("🚀 입금 데이터 일괄 저장"):
        try:
            df_p = pd.read_csv(pay_file)
            v_l = pd.read_sql("SELECT * FROM vendors", conn)
            o_l = pd.read_sql("SELECT * FROM orders", conn)
            
            for _, r in df_p.iterrows():
                oid, vn_raw = to_str(r.get('발주번호')), to_str(r.get('거래처'))
                if not vn_raw and not oid: 
                    continue
                
                pd_s = smart_date(r.get('입금일'))
                # 발주번호 연동 로직
                if oid and not o_l[o_l['발주번호'] == oid].empty:
                    info = o_l[o_l['발주번호'] == oid].iloc[0]
                    vn, pc, pp, cur = info['거래처명'], info['유형'], info['상품명'], info['통화']
                else:
                    vn, pc, pp, cur = vn_raw, to_str(r.get('유형')) or "사입", to_str(r.get('상품명')), "한화"
                
                vi = v_l[v_l['거래처명'] == vn] if vn else pd.DataFrame()
                dep, pre = to_float(r.get('실입금액')), to_float(r.get('선급금액'))
                rt = 1350.0 if cur == "USD" else (190.0 if cur == "CNY" else 1.0)
                
                conn.execute('''
                    INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
                ''', (oid if oid else None, pd_s, pc, vn, pp, cur, dep, pre, to_str(r.get('송금사유')), (dep+pre)*rt, 
                      vi.iloc[0]['은행'] if not vi.empty else "", vi.iloc[0]['계좌번호'] if not vi.empty else "", vi.iloc[0]['예금주'] if not vi.empty else ""))
            
            conn.commit()
            st.success("입금 엑셀 일괄 저장 완료!")
            st.session_state.pay_up_key += 1
            st.rerun()
        except Exception as e:
            st.error(f"엑셀 처리 중 오류: {e}")

# ------------------------------------------
# [Tab 2] 발주서 등록 및 마감 (샘플 + 수기 + 동기화)
# ------------------------------------------
with tabs[2]:
    st.header("📥 발주서 등록 및 마감 관리")
    col_o1, col_o2 = st.columns(2)
    
    with col_o1:
        st.subheader("⚡ 엑셀 일괄 등록")
        # 수기용 양식 다운로드
        ord_tmp = pd.DataFrame(columns=["발주번호", "발주일", "발주차수", "거래처명", "상품명", "금액", "통화"])
        st.download_button("📥 수기용 발주 양식 다운로드", ord_tmp.to_csv(index=False).encode('utf-8-sig'), "order_manual_template.csv")
        
        of_list = st.file_uploader("발주서(xlsx) 일괄 선택", type=['xlsx'], accept_multiple_files=True, key=f"ord_{st.session_state.order_up_key}")
        if of_list and st.button("🚀 모든 발주서 일괄 등록"):
            for of in of_list:
                process_ecount_v107(of)
            st.success("발주서 등록 작업 완료!")
            st.session_state.order_up_key += 1
            st.rerun()
            
    with col_o2:
        st.subheader("✍️ 수기 발주 등록")
        v_list = pd.read_sql("SELECT 거래처명 FROM vendors", conn)
        with st.form("order_manual_form"):
            o_id = st.text_input("발주번호")
            o_step = st.text_input("발주차수")
            o_date = st.date_input("발주일")
            o_vendor = st.selectbox("거래처 선택", ["선택"] + list(v_list['거래처명']))
            o_prod = st.text_input("상품명")
            o_amt = st.number_input("발주금액", format="%.2f")
            o_cur = st.selectbox("통화", ["한화", "USD", "CNY"])
            
            if st.form_submit_button("✅ 수기 발주 저장"):
                if o_id and o_vendor != "선택":
                    vt_row = pd.read_sql(f"SELECT 기본유형 FROM vendors WHERE 거래처명='{o_vendor}'", conn)
                    vt = vt_row.iloc[0]['기본유형'] if not vt_row.empty else "사입"
                    conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,0)", 
                                 (o_id, o_date.strftime("%Y-%m-%d"), o_step, o_vendor, o_prod, vt, o_cur, o_amt))
                    conn.commit()
                    st.success("수기 발주서가 등록되었습니다.")
                    st.rerun()
    
    st.divider()
    # 발주 리스트 및 소급 적용
    o_data = pd.read_sql("SELECT * FROM orders", conn)
    if not o_data.empty:
        st.subheader("📄 발주 리스트 및 마감 관리")
        show_closed = st.checkbox("마감된 발주 포함해서 보기", value=False)
        disp_o = o_data if show_closed else o_data[o_data['마감여부'] == 0]
        
        # 체크박스 마감여부 처리
        ev_o = st.data_editor(
            disp_o.sort_values('발주일', ascending=False), 
            hide_index=True, 
            use_container_width=True, 
            disabled=["발주번호"],
            column_config={"마감여부": st.column_config.CheckboxColumn("마감", help="체크 시 정산완료")}
        )
        
        if st.button("💾 정보 업데이트 및 모든 상세내역 소급 적용"):
            for _, r in ev_o.iterrows():
                # [1] 발주 마스터 업데이트
                conn.execute('''
                    UPDATE orders SET 발주일=?, 발주차수=?, 거래처명=?, 상품명=?, 유형=?, 통화=?, 발주총액=?, 마감여부=? 
                    WHERE 발주번호=?
                ''', (r['발주일'], r['발주차수'], r['거래처명'], r['상품명'], r['유형'], r['통화'], r['발주총액'], int(r['마감여부']), r['발주번호']))
                # [2] 입금 상세내역(payments) 강제 동기화
                conn.execute('''
                    UPDATE payments SET 거래처명=?, 유형=?, 상품명=?, 통화=? 
                    WHERE 발주번호=?
                ''', (r['거래처명'], r['유형'], r['상품명'], r['통화'], r['발주번호']))
            conn.commit()
            st.success("모든 데이터 동기화 및 업데이트가 완료되었습니다.")
            st.rerun()

# ------------------------------------------
# [Tab 3] 상세내역 및 통합 정산 (필터/검색/잔액)
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
        search = f3.text_input("업체명 또는 상품명 검색", placeholder="검색어를 입력하고 엔터를 치세요")
        
        # 필터링 엔진
        fil_p = p_all[p_all['dt'].dt.year == y]
        if m != "전체": 
            fil_p = fil_p[fil_p['dt'].dt.month == m]
        if search: 
            fil_p = fil_p[fil_p['거래처명'].str.contains(search, na=False) | fil_p['상품명'].str.contains(search, na=False)]
        
        # [1] 유형별 요약
        if not fil_p.empty:
            st.write(f"#### 📈 {y}년 {m if m != '전체' else ''} 유형별 요약")
            cat_sum = fil_p.groupby('유형').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
            cat_sum['총합계'] = cat_sum['실입금액'] + cat_sum['선급금액']
            st.table(cat_sum.style.format({'실입금액':'{:,.2f}', '선급금액':'{:,.2f}', '총합계':'{:,.2f}'}))
        
        # [2] 정산 현황
        st.divider()
        st.subheader("📊 발주번호별 정산 및 미수금 현황")
        p_agg = p_all.groupby('발주번호').agg({'실입금액':'sum'}).reset_index()
        sum_df = pd.merge(o_all, p_agg, on='발주번호', how='left').fillna(0)
        sum_df['잔액'] = sum_df['발주총액'] - sum_df['실입금액']
        sum_df['상태'] = sum_df['마감여부'].apply(lambda x: "✅ 마감완료" if x == 1 else "⏳ 진행중")
        
        st.dataframe(
            sum_df[['발주번호', '상태', '거래처명', '상품명', '발주총액', '실입금액', '잔액', '통화']].sort_values('상태'), 
            use_container_width=True,
            hide_index=True
        )

        # [3] 상세 편집 리스트
        st.divider()
        st.subheader("📑 상세 리스트 편집")
        ed_p = st.data_editor(
            fil_p.drop(columns=['dt']).sort_values('입금일', ascending=False), 
            hide_index=True, 
            use_container_width=True, 
            disabled=["id"]
        )
        
        eb1, eb2 = st.columns([1, 4])
        if eb1.button("💾 상세 내역 수정 저장"):
            for _, r in ed_p.iterrows():
                conn.execute('''
                    UPDATE payments SET 발주번호=?, 입금일=?, 유형=?, 거래처명=?, 상품명=?, 실입금액=?, 선급금액=?, 메모=? 
                    WHERE id=?
                ''', (r['발주번호'], r['입금일'], r['유형'], r['거래처명'], r['상품명'], r['실입금액'], r['선급금액'], r['메모'], r['id']))
            conn.commit()
            st.success("수정 완료!")
            st.rerun()
            
        did = eb2.number_input("삭제할 데이터 ID 입력", min_value=0, step=1)
        if eb2.button("🗑️ 해당 ID 행 삭제"):
            conn.execute(f"DELETE FROM payments WHERE id={did}")
            conn.commit()
            st.rerun()

# ------------------------------------------
# [Tab 4] 거래처 관리 (샘플 + 수기 + 소급)
# ------------------------------------------
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    cv1, cv2 = st.columns([1.2, 0.8])
    
    with cv1:
        st.subheader("➕ 신규 거래처 등록")
        with st.form("vendor_reg_form", clear_on_submit=True):
            vn = st.text_input("거래처명")
            vt = st.selectbox("유형", CATEGORIES)
            vcol1, vcol2, vcol3 = st.columns(3)
            vb = vcol1.text_input("은행")
            va = vcol2.text_input("계좌번호")
            vh = vcol3.text_input("예금주")
            
            if st.form_submit_button("✅ 거래처 저장"):
                if vn:
                    conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (vn,vb,va,vh,vt))
                    conn.commit()
                    st.success(f"'{vn}' 거래처가 등록되었습니다.")
                    st.rerun()
                    
    with cv2:
        st.subheader("📂 거래처 일괄 업로드")
        v_tmp = pd.DataFrame(columns=["거래처명", "은행", "계좌번호", "예금주", "기본유형"])
        st.download_button("📥 거래처 양식 다운로드", v_tmp.to_csv(index=False).encode('utf-8-sig'), "vendor_template.csv")
        
        vf = st.file_uploader("거래처 CSV 업로드", type=['csv'])
        if vf and st.button("🚀 거래처 리스트 업로드"):
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
        st.subheader("🏢 거래처 마스터 리스트 (이름 수정 시 전체 소급)")
        orig_v = v_data['거래처명'].tolist()
        ev_v = st.data_editor(v_data, hide_index=True, use_container_width=True)
        
        if st.button("💾 거래처 정보 및 전체 데이터 동기화"):
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
            st.success("전체 데이터 동기화 완료!")
            st.rerun()

# ------------------------------------------
# [Tab 5] 환율 관리 (Plotly 차트 + % 분석 리포트)
# ------------------------------------------
with tabs[5]:
    st.header("📈 환율 정밀 분석 (절대치 및 증감율 병기)")
    
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        f_usd = st.file_uploader("USD/KRW CSV 업로드", type=['csv'], key="usd_up")
        if f_usd and st.button("📥 USD 환율 업데이트"):
            if process_exchange_csv(f_usd, "USD"):
                st.success("USD 데이터 반영 완료")
                st.rerun()
    with col_up2:
        f_cny = st.file_uploader("CNY/KRW CSV 업로드", type=['csv'], key="cny_up")
        if f_cny and st.button("📥 CNY 환율 업데이트"):
            if process_exchange_csv(f_cny, "CNY"):
                st.success("CNY 데이터 반영 완료")
                st.rerun()

    st.divider()
    ex_db = pd.read_sql("SELECT * FROM exchange_rates ORDER BY 날짜 ASC", conn)
    if not ex_db.empty:
        ex_db['dt'] = pd.to_datetime(ex_db['날짜'])
        ex_db['연월'] = ex_db['dt'].dt.strftime('%Y-%m')
        
        # 월별 평균 데이터 가공
        monthly_mean = ex_db.groupby('연월').agg({
            'usd': lambda x: x[x > 0].mean(),
            'cny': lambda x: x[x > 0].mean()
        }).reset_index().fillna(0)
        
        # [1] 시각화 차트 (Y축 고정)
        st.subheader("📉 월별 평균 환율 추이 (정밀 범위)")
        c1, c2 = st.columns(2)
        with c1:
            st.write("**[USD] 1360~1540 (20단위)**")
            fig_u = go.Figure()
            fig_u.add_trace(go.Scatter(x=monthly_mean['연월'], y=monthly_mean['usd'], mode='lines+markers', name='USD'))
            fig_u.update_layout(yaxis=dict(range=[1360, 1540], dtick=20), margin=dict(l=10,r=10,t=10,b=10), height=350, template="plotly_white")
            st.plotly_chart(fig_u, use_container_width=True)
        with c2:
            st.write("**[CNY] 186~226 (2단위)**")
            fig_c = go.Figure()
            fig_c.add_trace(go.Scatter(x=monthly_mean['연월'], y=monthly_mean['cny'], mode='lines+markers', name='CNY', line=dict(color='orange')))
            fig_c.update_layout(yaxis=dict(range=[186, 226], dtick=2), margin=dict(l=10,r=10,t=10,b=10), height=350, template="plotly_white")
            st.plotly_chart(fig_c, use_container_width=True)

        # [2] 분석 로직 (YoY / MoM %)
        monthly_mean['year'] = monthly_mean['연월'].str[:4].astype(int)
        monthly_mean['month'] = monthly_mean['연월'].str[5:].astype(int)
        
        # 전월비(MoM) 계산 (이전 행과의 차이)
        monthly_mean['usd_mom_val'] = monthly_mean['usd'].diff()
        monthly_mean['cny_mom_val'] = monthly_mean['cny'].diff()
        
        def get_full_report(df, col):
            years = sorted(df['year'].unique(), reverse=True)
            pivot_m = df.pivot(index='month', columns='year', values=col).sort_index()
            curr_y = years[0]
            prev_y = years[1] if len(years) > 1 else None
            
            rep = pd.DataFrame(index=pivot_m.index)
            if prev_y: 
                rep[f'{prev_y}년 평균'] = pivot_m[prev_y]
            rep[f'{curr_y}년 평균'] = pivot_m[curr_y]
            
            # [YoY %]
            if prev_y:
                y_diff = rep[f'{curr_y}년 평균'] - rep[f'{prev_y}년 평균']
                y_pct = (y_diff / rep[f'{prev_y}년 평균']) * 100
                rep['전년비(YoY)'] = [f"{d:+.2f}({p:+.1f}%)" if pd.notnull(d) and pd.notnull(p) else "-" for d, p in zip(y_diff, y_pct)]
            
            # [MoM %] - 전체 시계열에서 참조하여 1월도 계산
            mom_data = df[df['year'] == curr_y].set_index('month')
            m_diff = mom_data[f'{col}_mom_val']
            p_m_val = mom_data[col] - m_diff
            m_pct = (m_diff / p_m_val) * 100
            
            rep['전월비(MoM)'] = [f"{d:+.2f}({p:+.1f}%)" if pd.notnull(d) and v > 0 else "-" for d, p, v in zip(m_diff, m_pct, p_m_val)]
            
            return rep.reset_index()

        st.divider()
        st.subheader("📅 연도별 병렬 분석 리포트 (YoY & MoM 통합)")
        rep_col1, rep_col2 = st.columns(2)
        with rep_col1:
            st.write("#### 💵 USD 환율 리포트")
            u_r = get_full_report(monthly_mean, 'usd')
            if not u_r.empty: 
                st.table(u_r.style.format({'month':'{:.0f}월', f'{sorted(monthly_mean["year"].unique(), reverse=True)[0]}년 평균':'{:,.2f}'}, na_rep="-"))
        with rep_col2:
            st.write("#### 💴 CNY 환율 리포트")
            c_r = get_full_report(monthly_mean, 'cny')
            if not c_r.empty: 
                st.table(c_r.style.format({'month':'{:.0f}월', f'{sorted(monthly_mean["year"].unique(), reverse=True)[0]}년 평균':'{:,.2f}'}, na_rep="-"))
    else:
        st.warning("환율 데이터가 없습니다. 상단에서 Investing.com CSV를 업로드하세요.")