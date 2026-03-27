import streamlit as st
import pandas as pd
import sqlite3
import os
import shutil
from datetime import datetime
import plotly.express as px

# --- 1. 데이터 안전장치 (자동 백업 로직) ---
def run_backup():
    if not os.path.exists('backups'):
        os.makedirs('backups')
    
    # 오늘 날짜로 백업 파일명 생성
    today_str = datetime.now().strftime("%Y%m%d")
    db_file = 'finance_v10.db'
    backup_file = f"backups/backup_{today_str}.db"
    
    if os.path.exists(db_file) and not os.path.exists(backup_file):
        shutil.copy2(db_file, backup_file)
        return True
    return False

# 2. 페이지 설정 및 백업 실행
st.set_page_config(page_title="자금 관리 시스템 v10", layout="wide", page_icon="💰")
backup_status = run_backup()

# 3. DB 연결
@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_v10.db', check_same_thread=False)
    c = conn.cursor()
    # 발주 마스터
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (order_id TEXT PRIMARY KEY, order_date TEXT, vendor TEXT, 
                  product TEXT, category TEXT, currency TEXT, total_amt REAL, is_closed INTEGER DEFAULT 0)''')
    # 입금 내역 (기존 엑셀 양식 컬럼 준수)
    c.execute('''CREATE TABLE IF NOT EXISTS payments 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, order_id TEXT, pay_date TEXT, 
                  category TEXT, vendor TEXT, product TEXT, currency TEXT,
                  deposit REAL, advance REAL, note TEXT, krw_val REAL)''')
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비"]
CURRENCIES = ["한화", "USD", "CNY"]

# --- 공통 데이터 로드 ---
def load_df(table_name):
    return pd.read_sql(f"SELECT * FROM {table_name}", conn)

# 상단 알림
if backup_status:
    st.toast("✅ 오늘자 데이터 자동 백업 완료!", icon="💾")

# --- 메인 화면 ---
st.title("💰 자금 관리 및 보고 시스템 v10")

tabs = st.tabs(["📥 발주서 등록", "📂 입금 엑셀 업로드", "🔍 상세 내역 관리", "📊 보고용 대시보드"])

# --- Tab 1: 발주서 등록 (이카운트 기반) ---
with tabs[0]:
    st.header("📄 발주서(Master) 등록")
    with st.form("order_form"):
        c1, c2, c3 = st.columns(3)
        oid = c1.text_input("발주번호 (ERP 전표번호)")
        odate = c2.date_input("발주일", datetime.now())
        ocat = c3.selectbox("유형", CATEGORIES)
        
        c4, c5, c6 = st.columns(3)
        ovendor = c4.text_input("거래처명")
        oprod = c5.text_input("상품명")
        ocurr = c6.selectbox("발주 통화", CURRENCIES)
        
        ototal = st.number_input("발주 총액", min_value=0.0)
        if st.form_submit_button("🚀 발주 정보 저장"):
            if oid and ovendor:
                conn.cursor().execute("INSERT OR REPLACE INTO orders VALUES (?, ?, ?, ?, ?, ?, ?, 0)", 
                                     (oid, odate.strftime("%Y-%m-%d"), ovendor, oprod, ocat, ocurr, ototal))
                conn.commit()
                st.success(f"발주번호 {oid} 등록 완료")

# --- Tab 2: 입금 엑셀 업로드 (기존 방식 유지) ---
with tabs[1]:
    st.header("📂 입금 내역 일괄 업로드")
    st.info("엑셀에서 작업한 내용을 CSV로 저장하여 업로드하세요.")
    
    # 양식 다운로드 (사용자 요청 컬럼 반영)
    template = pd.DataFrame(columns=["입금일", "거래처", "유형", "통화", "상품명", "입금액", "선급금", "송금사유", "발주번호"])
    st.download_button("📥 업로드 양식(CSV) 받기", template.to_csv(index=False).encode('utf-8-sig'), "upload_template.csv")
    
    up_file = st.file_uploader("파일 선택", type=['csv'])
    if up_file:
        df_up = pd.read_csv(up_file)
        if st.button("✅ 입금 내역 동기화"):
            for _, r in df_up.iterrows():
                # 외화 환산 로직 (임시 1350/190 적용)
                rate = 1350.0 if r['통화'] == "USD" else (190.0 if r['통화'] == "CNY" else 1.0)
                conn.cursor().execute('''INSERT INTO payments (order_id, pay_date, category, vendor, product, currency, deposit, advance, note, krw_val) 
                                       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                                    (r['발주번호'], r['입금일'], r['유형'], r['거래처'], r['상품명'], r['통화'], r['입금액'], r['선급금'], r['송금사유'], float(r['입금액']) * rate))
            conn.commit()
            st.success(f"{len(df_up)}건의 내역이 안전하게 저장되었습니다.")

# --- Tab 3: 상세 내역 관리 (필터 및 마감) ---
with tabs[2]:
    st.header("🔍 상세 내역 조회 및 수정")
    p_data = load_df("payments")
    o_data = load_df("orders")
    
    if not p_data.empty:
        # 필터 구역
        f1, f2, f3 = st.columns(3)
        sel_vendor = f1.multiselect("거래처", p_data['vendor'].unique())
        sel_cat = f2.multiselect("유형", CATEGORIES)
        sel_status = f3.radio("마감 상태", ["진행 중", "마감 건"], horizontal=True)
        
        # 데이터 병합 및 필터링
        df_merged = p_data.merge(o_data[['order_id', 'is_closed']], on='order_id', how='left').fillna(0)
        
        target_val = 1 if sel_status == "마감 건" else 0
        df_final = df_merged[df_merged['is_closed'] == target_val]
        if sel_vendor: df_final = df_final[df_final['vendor'].isin(sel_vendor)]
        if sel_cat: df_final = df_final[df_final['category'].isin(sel_cat)]
        
        # 가시화 및 수정 (마감 건은 회색 음영 느낌으로 제공)
        st.subheader(f"📋 {sel_status} 목록")
        st.data_editor(df_final, use_container_width=True, key="editor_v10")
        
        # 마감 처리 버튼
        if sel_status == "진행 중":
            to_close = st.selectbox("마감할 발주번호", df_final['order_id'].unique())
            if st.button("🚩 해당 발주 전체 마감"):
                conn.cursor().execute("UPDATE orders SET is_closed = 1 WHERE order_id = ?", (to_close,))
                conn.commit()
                st.rerun()

# --- Tab 4: 보고용 대시보드 (대표님 보고용) ---
with tabs[3]:
    st.header("📊 지출 및 정산 요약 보고")
    p_all = load_df("payments")
    if not p_all.empty:
        c1, c2 = st.columns(2)
        # 1. 유형별 지출
        fig1 = px.pie(p_all, values='krw_val', names='category', title="업무 유형별 지출 비중")
        c1.plotly_chart(fig1)
        
        # 2. 거래처별 미결제 (발주총액 대비)
        o_all = load_df("orders")
        pay_sum = p_all.groupby('order_id')['deposit'].sum().reset_index()
        report_df = o_all.merge(pay_sum, on='order_id', how='left').fillna(0)
        report_df['미결제잔액'] = report_df['total_amt'] - report_df['deposit']
        
        fig2 = px.bar(report_df[report_df['is_closed']==0], x='vendor', y='미결제잔액', color='category', title="거래처별 남은 잔금(진행중)")
        c2.plotly_chart(fig2)
        
        st.divider()
        st.subheader("📝 업체별 상세 정산표")
        st.table(report_df[['vendor', 'product', 'category', 'total_amt', 'deposit', '미결제잔액']])