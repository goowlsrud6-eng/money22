import streamlit as st
import pandas as pd
import sqlite3
import os
import shutil
import re
from datetime import datetime, timedelta

# --- 1. 백업 및 데이터베이스 설정 ---
def run_backup():
    """데이터베이스 백업 생성"""
    if not os.path.exists('backups'):
        os.makedirs('backups')
    db_file = 'finance_final_v97.db'
    backup_file = f"backups/backup_{datetime.now().strftime('%Y%m%d')}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)

st.set_page_config(page_title="자금 관리 v97", layout="wide", page_icon="💰")
run_backup()

@st.cache_resource
def get_db_connection():
    """DB 연결 및 테이블 생성 (누락 방지)"""
    conn = sqlite3.connect('finance_final_v97.db', check_same_thread=False)
    c = conn.cursor()
    # 1. 거래처 마스터
    c.execute('CREATE TABLE IF NOT EXISTS vendors (거래처명 TEXT PRIMARY KEY, 은행 TEXT, 계좌번호 TEXT, 예금주 TEXT, 기본유형 TEXT)')
    # 2. 발주 마스터 (마감여부 INTEGER: 0-진행, 1-마감)
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (발주번호 TEXT PRIMARY KEY, 발주일 TEXT, 발주차수 TEXT, 거래처명 TEXT, 
                  상품명 TEXT, 유형 TEXT, 통화 TEXT, 발주총액 REAL, 마감여부 INTEGER DEFAULT 0)''')
    # 3. 입금 상세 내역 (정확히 13개 컬럼 유지)
    c.execute('''CREATE TABLE IF NOT EXISTS payments 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 발주번호 TEXT, 입금일 TEXT, 
                  유형 TEXT, 거래처명 TEXT, 상품명 TEXT, 통화 TEXT,
                  실입금액 REAL, 선급금액 REAL, 메모 TEXT, 한화환산액 REAL,
                  은행 TEXT, 계좌번호 TEXT, 예금주 TEXT)''')
    # 4. 환율 관리 테이블
    c.execute('CREATE TABLE IF NOT EXISTS exchange_rates (날짜 TEXT PRIMARY KEY, usd REAL, cny REAL)')
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비", "원단비", "기타"]

# --- 2. 세션 상태 관리 (파일 업로더 초기화용 키) ---
if 'order_up_key' not in st.session_state: st.session_state.order_up_key = 0
if 'pay_up_key' not in st.session_state: st.session_state.pay_up_key = 1000

# --- 3. 유틸리티 함수 (에러 방지 및 데이터 정제) ---
def to_float(val):
    """문자열/NaN을 숫자로 안전하게 변환"""
    try:
        if val is None or pd.isna(val) or str(val).strip() == "": return 0.0
        return float(str(val).replace(',', ''))
    except: return 0.0

def to_str(val):
    """빈 값을 None 또는 빈 문자열로 정제"""
    if val is None or pd.isna(val): return ""
    s = str(val).strip()
    return "" if s.lower() in ["nan", "none", ""] else s

def smart_date(date_str):
    """다양한 날짜 형식을 YYYY-MM-DD로 통일"""
    try:
        ds = to_str(date_str).replace(" ", "").replace(".", "-")
        if not ds: return datetime.now().strftime("%Y-%m-%d")
        return pd.to_datetime(ds).strftime("%Y-%m-%d")
    except: return datetime.now().strftime("%Y-%m-%d")

# --- 4. 핵심 로직: 환율 CSV 분석 엔진 ---
def process_exchange_csv(file, currency_type):
    """Investing.com CSV 파일을 읽어 DB에 업데이트"""
    try:
        df = pd.read_csv(file)
        # Investing.com 필수 컬럼: 날짜, 종가
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
    except Exception as e:
        st.error(f"환율 분석 중 오류: {e}")
        return False

# --- 5. ERP 발주서 분석 로직 ---
def process_ecount_v97(file):
    """이카운트 발주서 엑셀 분석 및 저장"""
    try:
        df = pd.read_excel(file, header=None)
        # 발주번호 추출
        raw_oid = str(df.iloc[1, 0]).split(":")[-1].strip() if ":" in str(df.iloc[1,0]) else str(df.iloc[1, 0])
        clean_oid = raw_oid.replace("-", "")
        odate = smart_date(clean_oid[:8])
        
        # 거래처명 추출 (수신: 항목 찾기)
        vendor_raw = ""
        for i in range(len(df)):
            if "수신" in str(df.iloc[i, 0]):
                vendor_raw = str(df.iloc[i, 0]).split(":")[-1].strip()
                break
        
        # 거래처 마스터와 대조 (공백 제거 후 매칭)
        v_master = pd.read_sql("SELECT 거래처명, 기본유형 FROM vendors", conn)
        v_master['clean'] = v_master['거래처명'].apply(lambda x: re.sub(r'\s+', '', str(x)))
        target_key = re.sub(r'\s+', '', vendor_raw)
        
        match = v_master[v_master['clean'] == target_key]
        if match.empty:
            return False, f"⚠️ '{vendor_raw}'은(는) 미등록 업체입니다. 거래처 관리에서 먼저 등록하세요."
        
        v_type, vendor_fixed = match.iloc[0]['기본유형'], match.iloc[0]['거래처명']
        
        # 통화 및 품목 정보 추출
        f6 = str(df.iloc[5, 5]) if len(df) > 5 else ""
        curr = "USD" if "USD" in f6 else ("CNY" if any(x in f6 for x in ["중국", "CNY"]) else "한화")
        p_col = 1 if curr == "한화" else 2
        prods = df.iloc[6:, p_col].dropna().astype(str).tolist()
        prod_n = (prods[0].split("[")[0].strip() + (f" 외 {len(prods)-1}건" if len(prods)>1 else "")) if prods else "품목미상"
        
        # 금액 추출
        l_idx = df.iloc[:, 5].last_valid_index()
        total = to_float(df.iloc[l_idx, 5]) if curr != "한화" and l_idx else to_float(str(df.iloc[4, 0]).split(":")[-1])

        # 저장
        conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,0)", 
                     (raw_oid, odate, "", vendor_fixed, prod_n, v_type, curr, total))
        conn.commit()
        return True, None
    except Exception as e:
        return False, f"❗ 발주서 분석 오류: {e}"

# --- 6. 메인 UI 구성 ---
tabs = st.tabs(["📝 입금 입력", "📂 입금 엑셀 업로드", "📥 발주서 등록", "🔍 상세내역 및 정산", "⚙️ 거래처 관리", "📈 환율 관리"])

# [Tab 0] 입금 수기 입력
with tabs[0]:
    st.header("📝 입금 내역 수기 입력")
    v_data = pd.read_sql("SELECT * FROM vendors", conn)
    o_active = pd.read_sql("SELECT 발주번호 FROM orders WHERE 마감여부=0", conn)
    
    with st.form("manual_payment_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        p_oid = col1.selectbox("🔗 진행중인 발주번호 연동", ["없음"] + list(o_active['발주번호']))
        p_date = col2.date_input("입금일")
        
        col3, col4, col5 = st.columns(3)
        p_vn = col3.selectbox("거래처명", ["선택"] + list(v_data['거래처명']))
        p_ct = col4.selectbox("유형", CATEGORIES)
        p_pr = col5.text_input("상품명")
        
        col6, col7, col8 = st.columns(3)
        p_dep = col6.number_input("실입금액", format="%.2f")
        p_pre = col7.number_input("선급금액", format="%.2f")
        p_cur = col8.selectbox("통화", ["한화", "USD", "CNY"])
        
        p_memo = st.text_input("메모 (기타사항)")
        
        if st.form_submit_button("✅ 입금 내역 저장"):
            if p_vn == "선택":
                st.error("거래처를 선택해주세요.")
            else:
                # 환율 기본값 설정 (환율 관리 탭 데이터와 연동 전 고정값)
                rate = 1350.0 if p_cur == "USD" else (190.0 if p_cur == "CNY" else 1.0)
                vi = v_data[v_data['거래처명']==p_vn].iloc[0]
                
                conn.execute('''INSERT INTO payments 
                                (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                             (to_str(p_oid) if p_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), p_ct, p_vn, p_pr, p_cur, 
                              p_dep, p_pre, p_memo, (p_dep+p_pre)*rate, vi['은행'], vi['계좌번호'], vi['예금주']))
                conn.commit()
                st.success("입금 내역이 성공적으로 저장되었습니다.")
                st.rerun()

# [Tab 1] 입금 엑셀 업로드
with tabs[1]:
    st.header("📂 통합 입금 엑셀 업로드")
    # 샘플 양식 버튼
    sample_pay = pd.DataFrame(columns=["발주번호", "거래처", "유형", "상품명", "입금일", "실입금액", "선급금액", "송금사유"])
    st.download_button("📥 입금 업로드 샘플 다운로드", sample_pay.to_csv(index=False).encode('utf-8-sig'), "pay_sample.csv")
    
    f_p = st.file_uploader("입금 CSV 파일 선택", type=['csv'], key=f"pay_up_{st.session_state.pay_up_key}")
    if f_p and st.button("🚀 데이터 일괄 저장 실행"):
        try:
            df_p = pd.read_csv(f_p)
            v_master_l = pd.read_sql("SELECT * FROM vendors", conn)
            o_master_l = pd.read_sql("SELECT * FROM orders", conn)
            
            for _, row in df_p.iterrows():
                oid, vn_raw = to_str(row.get('발주번호')), to_str(row.get('거래처'))
                if not vn_raw and not oid: continue
                
                pd_s = smart_date(row.get('입금일'))
                
                # 연동 로직: 발주번호가 있으면 발주서 정보를 우선 채움
                if oid and not o_master_l[o_master_l['발주번호'] == oid].empty:
                    info = o_master_l[o_master_l['발주번호'] == oid].iloc[0]
                    vn, pc, pp, cur = info['거래처명'], info['유형'], info['상품명'], info['통화']
                else:
                    vn, pc, pp, cur = vn_raw, to_str(row.get('유형')) or "사입", to_str(row.get('상품명')), "한화"
                
                vi = v_master_l[v_master_l['거래처명'] == vn] if vn else pd.DataFrame()
                dep, pre = to_float(row.get('실입금액')), to_float(row.get('선급금액'))
                rt = 1350.0 if cur == "USD" else (190.0 if cur == "CNY" else 1.0)
                
                conn.execute('''INSERT INTO payments (발주번호, 입금일, 유형, 거래처명, 상품명, 통화, 실입금액, 선급금액, 메모, 한화환산액, 은행, 계좌번호, 예금주) 
                                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)''',
                             (oid if oid else None, pd_s, pc, vn, pp, cur, dep, pre, to_str(row.get('송금사유')), (dep+pre)*rt, 
                              vi.iloc[0]['은행'] if not vi.empty else "", vi.iloc[0]['계좌번호'] if not vi.empty else "", vi.iloc[0]['예금주'] if not vi.empty else ""))
            
            conn.commit()
            st.success("입금 엑셀 데이터가 일괄 저장되었습니다.")
            st.session_state.pay_up_key += 1
            st.rerun()
        except Exception as e:
            st.error(f"입금 엑셀 처리 중 오류: {e}")

# [Tab 2] 발주서 등록 및 마감 관리
with tabs[2]:
    st.header("📥 발주서 등록 및 마감 관리")
    col_o1, col_o2 = st.columns(2)
    with col_o1:
        st.subheader("⚡ 엑셀 일괄 등록")
        of_list = st.file_uploader("발주서(xlsx) 선택", type=['xlsx'], accept_multiple_files=True, key=f"ord_{st.session_state.order_up_key}")
        if of_list and st.button("🚀 모든 발주서 일괄 등록"):
            scnt, errs = 0, []
            for of in of_list:
                res, msg = process_ecount_v97(of)
                if res: scnt += 1
                else: errs.append(msg)
            for em in errs: st.warning(em)
            if scnt > 0: 
                st.success(f"✅ {scnt}건의 발주서가 성공적으로 등록되었습니다.")
                if not errs: st.session_state.order_up_key += 1; st.rerun()
    with col_o2:
        st.subheader("✍️ 수기 발주 등록")
        with st.form("manual_order_form"):
            mi, m_step = st.text_input("발주번호"), st.text_input("발주차수")
            md, mv = st.date_input("발주일"), st.selectbox("거래처 선택", ["선택"] + list(pd.read_sql("SELECT 거래처명 FROM vendors", conn)['거래처명']))
            mp = st.text_input("상품명")
            mt, m_cur = st.number_input("발주금액", format="%.2f"), st.selectbox("통화", ["한화", "USD", "CNY"])
            if st.form_submit_button("✅ 수기 저장"):
                if mi and mv != "선택":
                    v_type_info = pd.read_sql(f"SELECT 기본유형 FROM vendors WHERE 거래처명='{mv}'", conn)
                    vt = v_type_info.iloc[0]['기본유형'] if not v_type_info.empty else "사입"
                    conn.execute("INSERT OR REPLACE INTO orders VALUES (?,?,?,?,?,?,?,?,0)", 
                                 (mi, md.strftime("%Y-%m-%d"), m_step, mv, mp, vt, m_cur, mt))
                    conn.commit()
                    st.success("수기 발주서가 등록되었습니다.")
                    st.rerun()
    
    st.divider()
    order_master_data = pd.read_sql("SELECT * FROM orders", conn)
    if not order_master_data.empty:
        st.subheader("📄 발주 리스트 및 마감 관리")
        show_closed = st.checkbox("마감된 발주 건 포함해서 보기", value=False)
        display_orders = order_master_data if show_closed else order_master_data[order_master_data['마감여부'] == 0]
        
        # 데이터 에디터 (마감 체크 기능)
        edited_orders = st.data_editor(
            display_orders[['발주번호', '발주차수', '거래처명', '상품명', '유형', '통화', '발주총액', '마감여부', '발주일']].sort_values('발주일', ascending=False),
            hide_index=True, use_container_width=True, disabled=["발주번호"],
            column_config={"마감여부": st.column_config.CheckboxColumn("마감", help="정산 완료 시 체크")}
        )
        
        if st.button("💾 정보 업데이트 및 모든 상세내역 소급 적용"):
            for _, r in edited_orders.iterrows():
                # 1. 발주 마스터 업데이트
                conn.execute('''UPDATE orders SET 발주일=?, 발주차수=?, 거래처명=?, 상품명=?, 유형=?, 통화=?, 발주총액=?, 마감여부=? 
                                WHERE 발주번호=?''', 
                             (r['발주일'], r['발주차수'], r['거래처명'], r['상품명'], r['유형'], r['통화'], r['발주총액'], int(r['마감여부']), r['발주번호']))
                # 2. 입금 상세내역(payments) 소급 적용 (연동의 핵심)
                conn.execute('''UPDATE payments SET 거래처명=?, 유형=?, 상품명=?, 통화=? 
                                WHERE 발주번호=?''', 
                             (r['거래처명'], r['유형'], r['상품명'], r['통화'], r['발주번호']))
            conn.commit()
            st.success("✅ 발주 정보 및 입금 내역의 동기화가 완료되었습니다.")
            st.rerun()

# [Tab 3] 상세내역 및 통합 정산 (사용자 요청: 필터/검색 강화 버전 복구)
with tabs[3]:
    st.header("🔍 상세 내역 및 통합 정산")
    p_all_data = pd.read_sql("SELECT * FROM payments", conn)
    o_all_data = pd.read_sql("SELECT * FROM orders", conn)
    
    if not p_all_data.empty:
        p_all_data['dt'] = pd.to_datetime(p_all_data['입금일'])
        
        st.subheader("📊 월별 필터 및 데이터 검색")
        filt_c1, filt_c2, filt_c3 = st.columns([1, 1, 2])
        sel_year = filt_c1.selectbox("기준 연도", sorted(p_all_data['dt'].dt.year.unique(), reverse=True))
        sel_month = filt_c2.selectbox("기준 월", ["전체"] + sorted(list(p_all_data[p_all_data['dt'].dt.year == sel_year]['dt'].dt.month.unique())))
        search_query = filt_c3.text_input("업체명 또는 상품명 통합 검색", placeholder="검색어를 입력하고 엔터를 치세요")
        
        # 필터링 엔진
        filtered_pay = p_all_data[p_all_data['dt'].dt.year == sel_year]
        if sel_month != "전체": 
            filtered_pay = filtered_pay[filtered_pay['dt'].dt.month == sel_month]
        if search_query:
            filtered_pay = filtered_pay[filtered_pay['거래처명'].str.contains(search_query, na=False) | filtered_pay['상품명'].str.contains(search_query, na=False)]
        
        # 1. 유형별 요약 테이블
        if not filtered_pay.empty:
            st.write(f"### 📈 {sel_year}년 {sel_month if sel_month != '전체' else ''} 유형별 입금 요약")
            category_summary = filtered_pay.groupby('유형').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
            category_summary['총 합계액'] = category_summary['실입금액'] + category_summary['선급금액']
            st.table(category_summary.style.format({'실입금액': '{:,.2f}', '선급금액': '{:,.2f}', '총 합계액': '{:,.2f}'}))
        
        # 2. 발주 대비 정산 현황 (미수금 관리용)
        st.divider()
        st.subheader("📊 발주번호별 정산 및 미수금 현황")
        pay_agg_by_oid = p_all_data.groupby('발주번호').agg({'실입금액':'sum'}).reset_index()
        settle_status = pd.merge(o_all_data, pay_agg_by_oid, on='발주번호', how='left').fillna(0)
        settle_status['미수잔액'] = settle_status['발주총액'] - settle_status['실입금액']
        settle_status['상태'] = settle_status['마감여부'].apply(lambda x: "✅ 마감완료" if x == 1 else "⏳ 진행중")
        
        st.dataframe(
            settle_status[['발주번호', '상태', '거래처명', '상품명', '발주총액', '실입금액', '미수잔액', '통화']].sort_values('상태'),
            use_container_width=True, hide_index=True
        )

        # 3. 상세 입금 내역 편집 리스트
        st.divider()
        st.subheader("📑 상세 입금 내역 리스트 (수정 및 삭제)")
        edited_pay_list = st.data_editor(
            filtered_pay.drop(columns=['dt']).sort_values('입금일', ascending=False),
            hide_index=True, use_container_width=True, disabled=["id"]
        )
        
        col_btn1, col_btn2 = st.columns([1, 4])
        if col_btn1.button("💾 상세 내역 수정 저장"):
            for _, r in edited_pay_list.iterrows():
                conn.execute('''UPDATE payments SET 발주번호=?, 입금일=?, 유형=?, 거래처명=?, 상품명=?, 실입금액=?, 선급금액=?, 메모=? 
                                WHERE id=?''', 
                             (r['발주번호'], r['입금일'], r['유형'], r['거래처명'], r['상품명'], r['실입금액'], r['선급금액'], r['메모'], r['id']))
            conn.commit()
            st.success("상세 내역이 수정되었습니다.")
            st.rerun()
            
        del_target_id = col_btn2.number_input("삭제할 데이터 ID", min_value=0, step=1)
        if st.button("🗑️ 선택한 ID 행 삭제"):
            conn.execute(f"DELETE FROM payments WHERE id={del_target_id}")
            conn.commit()
            st.rerun()

# [Tab 4] 거래처 관리
with tabs[4]:
    st.header("⚙️ 거래처 관리")
    col_v1, col_v2 = st.columns([1.2, 0.8])
    with col_v1:
        st.subheader("➕ 신규 거래처 등록")
        with st.form("vendor_register_form", clear_on_submit=True):
            vn_new = st.text_input("거래처명 (정식명칭)")
            vt_new = st.selectbox("거래처 기본 유형", CATEGORIES)
            vc1, vc2, vc3 = st.columns(3)
            vb_new = vc1.text_input("입금 은행")
            va_new = vc2.text_input("계좌 번호")
            vh_new = vc3.text_input("예금주 명")
            if st.form_submit_button("✅ 거래처 저장"):
                if vn_new:
                    conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", (vn_new, vb_new, va_new, vh_new, vt_new))
                    conn.commit()
                    st.success(f"'{vn_new}' 거래처가 등록되었습니다.")
                    st.rerun()
    with col_v2:
        st.subheader("📂 거래처 일괄 업로드 (CSV)")
        v_file = st.file_uploader("거래처 CSV 업로드", type=['csv'])
        if v_file and st.button("🚀 거래처 리스트 반영"):
            v_csv_data = pd.read_csv(v_file)
            for _, r in v_csv_data.iterrows():
                conn.execute("INSERT OR REPLACE INTO vendors VALUES (?,?,?,?,?)", 
                             (r['거래처명'], r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
            conn.commit()
            st.success("거래처 목록이 업데이트되었습니다.")
            st.rerun()
    
    st.divider()
    vendor_list_data = pd.read_sql("SELECT * FROM vendors", conn)
    if not vendor_list_data.empty:
        st.subheader("🏢 등록 거래처 마스터 리스트 (이름 수정 시 자동 소급)")
        orig_vendor_names = vendor_list_data['거래처명'].tolist()
        edited_vendors = st.data_editor(vendor_list_data, hide_index=True, use_container_width=True)
        
        if st.button("💾 거래처 정보 변경 및 전체 데이터 동기화"):
            for idx, r in edited_vendors.iterrows():
                old_name, new_name = orig_vendor_names[idx], r['거래처명']
                if old_name != new_name:
                    # 이름이 바뀐 경우 DB 모든 테이블 업데이트
                    conn.execute(f"DELETE FROM vendors WHERE 거래처명='{old_name}'")
                    conn.execute("INSERT INTO vendors VALUES (?,?,?,?,?)", (new_name, r['은행'], r['계좌번호'], r['예금주'], r['기본유형']))
                    conn.execute("UPDATE orders SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_name, r['기본유형'], old_name))
                    conn.execute("UPDATE payments SET 거래처명=?, 유형=? WHERE 거래처명=?", (new_name, r['기본유형'], old_name))
                else:
                    # 이름은 같고 계좌 정보만 바뀐 경우
                    conn.execute("UPDATE vendors SET 은행=?, 계좌번호=?, 예금주=?, 기본유형=? WHERE 거래처명=?", 
                                 (r['은행'], r['계좌번호'], r['예금주'], r['기본유형'], new_name))
            conn.commit()
            st.success("✅ 거래처 및 관련 모든 데이터가 업데이트되었습니다.")
            st.rerun()

# [Tab 5] 환율 관리 (Investing.com 데이터 분석 핵심 탭)
with tabs[5]:
    st.header("📈 환율 분석 및 월별 통계")
    st.info("Investing.com에서 받은 '과거 데이터(CSV)'를 업로드하세요. 자동으로 전월/전년 대비 변동폭을 계산합니다.")
    
    col_x1, col_x2 = st.columns(2)
    with col_x1:
        f_usd_csv = st.file_uploader("USD/KRW 과거 데이터 CSV 업로드", type=['csv'])
        if f_usd_csv and st.button("📥 USD 환율 일괄 업데이트"):
            if process_exchange_csv(f_usd_csv, "USD"):
                st.success("USD 환율 데이터 반영 완료")
                st.rerun()
    with col_x2:
        f_cny_csv = st.file_uploader("CNY/KRW 과거 데이터 CSV 업로드", type=['csv'])
        if f_cny_csv and st.button("📥 CNY 환율 일괄 업데이트"):
            if process_exchange_csv(f_cny_csv, "CNY"):
                st.success("CNY 환율 데이터 반영 완료")
                st.rerun()

    st.divider()
    exchange_raw_data = pd.read_sql("SELECT * FROM exchange_rates ORDER BY 날짜 DESC", conn)
    if not exchange_raw_data.empty:
        exchange_raw_data['연월'] = exchange_raw_data['날짜'].apply(lambda x: x[:7])
        exchange_raw_data['연도'] = exchange_raw_data['날짜'].apply(lambda x: x[:4])
        exchange_raw_data['월'] = exchange_raw_data['날짜'].apply(lambda x: x[5:7])

        # 월별 평균 산출 (0 제외 평균)
        monthly_exchange_avg = exchange_raw_data.groupby(['연도', '월']).agg({
            'usd': lambda x: x[x > 0].mean(),
            'cny': lambda x: x[x > 0].mean()
        }).reset_index().fillna(0)
        
        monthly_exchange_avg['연월'] = monthly_exchange_avg['연도'] + "-" + monthly_exchange_avg['월']
        monthly_exchange_avg = monthly_exchange_avg.sort_values('연월', ascending=False)

        st.subheader("📊 월별 평균 환율 및 증감 분석 (YoY / MoM)")
        
        # 비교 분석표 생성
        exchange_analysis_list = []
        for i, row in monthly_exchange_avg.iterrows():
            c_y, c_m = row['연도'], row['월']
            c_usd, c_cny = row['usd'], row['cny']
            
            # 전월(MoM) 찾기
            prev_month_date = (datetime.strptime(f"{c_y}-{c_m}-01", "%Y-%m-%d") - timedelta(days=1)).strftime("%Y-%m")
            prev_m_row = monthly_exchange_avg[monthly_exchange_avg['연월'] == prev_month_date]
            mom_diff = f"{(c_usd - prev_m_row.iloc[0]['usd']):+.2f}" if not prev_m_row.empty and prev_m_row.iloc[0]['usd'] > 0 else "-"
            
            # 전년(YoY) 동월 찾기
            prev_year_date = f"{int(c_y)-1}-{c_m}"
            prev_y_row = monthly_exchange_avg[monthly_exchange_avg['연월'] == prev_year_date]
            yoy_diff = f"{(c_usd - prev_y_row.iloc[0]['usd']):+.2f}" if not prev_y_row.empty and prev_y_row.iloc[0]['usd'] > 0 else "-"
            
            exchange_analysis_list.append({
                "기준 연월": row['연월'],
                "USD 평균": c_usd,
                "전월대비(MoM)": mom_diff,
                "전년대비(YoY)": yoy_diff,
                "CNY 평균": c_cny
            })
        
        st.table(pd.DataFrame(exchange_analysis_list).style.format({'USD 평균': '{:,.2f}', 'CNY 평균': '{:,.2f}'}))
        
        st.subheader("📈 환율 변동 추이 차트")
        st.line_chart(exchange_raw_data.set_index('날짜')[['usd', 'cny']])
    else:
        st.write("등록된 환율 데이터가 없습니다. CSV 파일을 업로드해주세요.")