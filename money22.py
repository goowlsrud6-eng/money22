import streamlit as st
import pandas as pd
import sqlite3
import os
import shutil
from datetime import datetime
import plotly.express as px

# --- 1. 데이터 안전장치 (자동 백업) ---
def run_backup():
    if not os.path.exists('backups'):
        os.makedirs('backups')
    today_str = datetime.now().strftime("%Y%m%d")
    db_file = 'finance_v11.db'
    backup_file = f"backups/backup_{today_str}.db"
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)
        return True
    return False

# 2. 페이지 설정
st.set_page_config(page_title="자금 관리 시스템 v11", layout="wide", page_icon="💰")
run_backup()

# 3. DB 연결 및 테이블 구조 (발주번호 중심)
@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_v11.db', check_same_thread=False)
    c = conn.cursor()
    # 발주 마스터 (is_closed로 진행/마감 구분)
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (order_id TEXT PRIMARY KEY, order_date TEXT, vendor TEXT, 
                  product TEXT, category TEXT, currency TEXT, total_amt REAL, is_closed INTEGER DEFAULT 0)''')
    # 입금 내역 (발주번호 연동)
    c.execute('''CREATE TABLE IF NOT EXISTS payments 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, order_id TEXT, pay_date TEXT, 
                  deposit REAL, advance REAL, note TEXT, krw_val REAL)''')
    conn.commit()
    return conn

conn = get_db_connection()

# --- 환율 로직 ---
def get_exchange_rate(currency):
    rates = {"USD": 1350.0, "CNY": 190.0, "한화": 1.0, "KRW": 1.0}
    return rates.get(currency, 1.0)

# --- 메인 화면 ---
st.title("💰 자금 관리 시스템 v11")
tabs = st.tabs(["📥 발주서 등록(Master)", "📂 입금 내역 업로드", "🔍 상세 정산 관리", "📊 보고 대시보드"])

# --- Tab 1: 발주서 등록 (다품목 대응) ---
with tabs[0]:
    st.header("📄 발주 마스터 등록 (이카운트 기준)")
    with st.form("order_reg"):
        c1, c2, c3 = st.columns(3)
        oid = c1.text_input("발주번호 (ERP 전표번호)", help="이 번호가 모든 데이터의 연결 고리가 됩니다.")
        odate = c2.date_input("발주일", datetime.now())
        ocat = c3.selectbox("유형", ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비"])
        
        c4, c5, c6 = st.columns(3)
        ovendor = c4.text_input("거래처명")
        oprod = c5.text_input("대표 상품명 (예: 품목 외 n건)")
        ocurr = c6.selectbox("통화", ["한화", "USD", "CNY"])
        
        ototal = st.number_input("발주 합계 금액 (전표 총액)", min_value=0.0)
        
        if st.form_submit_button("🚀 발주서 정보 저장"):
            if oid and ovendor:
                conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?, ?, ?, ?, ?, ?, ?, 0)", 
                                     (oid, odate.strftime("%Y-%m-%d"), ovendor, oprod, ocat, ocurr, ototal))
                conn.commit()
                st.success(f"발주번호 {oid} 등록 완료!")

# --- Tab 2: 입금 내역 업로드 (최적화 양식) ---
with tabs[1]:
    st.header("📂 입금 내역 엑셀 업로드")
    st.markdown("💡 **발주번호**만 정확하면 업체명, 상품명은 자동으로 매칭됩니다.")
    
    # 수정된 양식 제공
    template = pd.DataFrame(columns=["발주번호", "입금일", "입금액", "통화", "선급금변동", "송금사유"])
    st.download_button("📥 입금 업로드 양식(CSV) 다운로드", template.to_csv(index=False).encode('utf-8-sig'), "pay_template.csv")
    
    up_file = st.file_uploader("파일 선택", type=['csv'])
    if up_file:
        df_up = pd.read_csv(up_file)
        if st.button("✅ 입금 내역 저장/동기화"):
            for _, r in df_up.iterrows():
                rate = get_exchange_rate(r['통화'])
                conn.cursor().execute('''INSERT INTO payments (order_id, pay_date, deposit, advance, note, krw_val) 
                                       VALUES (?, ?, ?, ?, ?, ?)''',
                                    (r['발주번호'], r['입금일'], r['입금액'], r['선급금변동'], r['송금사유'], float(r['입금액']) * rate))
            conn.commit()
            st.success(f"총 {len(df_up)}건의 입금 내역이 저장되었습니다.")

# --- Tab 3: 상세 정산 관리 (엑셀 대장 스타일) ---
with tabs[2]:
    st.header("🔍 상세 정산 관리 (진행/마감 분리)")
    
    # 데이터 로드 및 병합
    o_df = pd.read_sql("SELECT * FROM orders", conn)
    p_df = pd.read_sql("SELECT * FROM payments", conn)
    
    if not o_df.empty:
        # 발주번호별 입금 합계 계산
        p_sum = p_df.groupby('order_id').agg({'deposit': 'sum', 'advance': 'sum'}).reset_index()
        main_df = o_df.merge(p_sum, on='order_id', how='left').fillna(0)
        main_df['잔금'] = main_df['total_amt'] - main_df['deposit']
        
        # 상태 필터
        view_status = st.radio("보기 설정", ["진행 중인 발주", "마감된 발주"], horizontal=True)
        is_closed_val = 1 if view_status == "마감된 발주" else 0
        
        disp_df = main_df[main_df['is_closed'] == is_closed_val]
        
        # 가시화
        st.subheader(f"📋 {view_status} 목록")
        # 컬럼 순서 재배치 (사용자 편의)
        col_order = ['order_id', 'order_date', 'vendor', 'category', 'product', 'currency', 'total_amt', 'deposit', '잔금', 'advance']
        st.data_editor(disp_df[col_order], use_container_width=True, key="main_editor")
        
        # 마감 기능
        if is_closed_val == 0:
            to_close = st.selectbox("마감 처리할 발주번호", disp_df['order_id'].unique())
            if st.button("🚩 선택한 발주 마감 (회색 처리)"):
                conn.cursor().execute("UPDATE orders SET is_closed = 1 WHERE order_id = ?", (to_close,))
                conn.commit()
                st.rerun()
    else:
        st.info("먼저 '발주서 등록' 탭에서 정보를 입력해주세요.")

# --- Tab 4: 보고 대시보드 ---
with tabs[3]:
    st.header("📊 지출 현황 요약")
    if not p_df.empty:
        # 한화 환산액 기준 차트
        fig = px.pie(p_df, values='krw_val', names='order_id', title="발주번호별 집행 비중")
        st.plotly_chart(fig, use_container_width=True)