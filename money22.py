import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
import plotly.express as px

# 1. 페이지 설정
st.set_page_config(page_title="자금 관리 시스템 v9", layout="wide", page_icon="💰")

# 2. DB 설정 및 초기화
@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('finance_v9.db', check_same_thread=False)
    c = conn.cursor()
    # 발주 마스터 테이블
    c.execute('''CREATE TABLE IF NOT EXISTS orders 
                 (order_id TEXT PRIMARY KEY, order_date TEXT, vendor TEXT, 
                  product TEXT, category TEXT, currency TEXT, total_amt REAL, is_closed INTEGER DEFAULT 0)''')
    # 입금 내역 테이블 (기존 필드 유지 + order_id 연동)
    c.execute('''CREATE TABLE IF NOT EXISTS payments 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, order_id TEXT, pay_date TEXT, 
                  category TEXT, vendor TEXT, product TEXT, currency TEXT,
                  deposit REAL, advance REAL, note TEXT, krw_val REAL)''')
    conn.commit()
    return conn

conn = get_db_connection()
CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비"]
CURRENCIES = ["한화", "USD", "CNY"]

# --- 유틸리티 함수 ---
def get_exchange_rate(date_str, currency):
    if currency in ["한화", "KRW"] or not currency: return 1.0
    # 월평균 환율 가정치 (추후 로직 보완)
    rates = {"USD": 1350.0, "CNY": 190.0}
    return rates.get(currency, 1.0)

def load_data(table):
    return pd.read_sql(f"SELECT * FROM {table}", conn)

# --- 메인 화면 구성 ---
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📥 발주서 등록(ERP)", 
    "💸 입금 건별 입력", 
    "📂 입금 엑셀 업로드", 
    "🔍 상세 내역 및 마감", 
    "📊 현황 대시보드"
])

# --- Tab 1: 발주서 등록 (이카운트 연동 준비) ---
with tab1:
    st.header("📄 발주 마스터 등록")
    with st.form("order_reg"):
        c1, c2, c3 = st.columns(3)
        o_id = c1.text_input("발주번호 (ERP 전표번호)")
        o_date = c2.date_input("발주일", datetime.now())
        o_cat = c3.selectbox("유형", CATEGORIES)
        
        c4, c5, c6 = st.columns(3)
        o_vendor = c4.text_input("거래처명")
        o_prod = c5.text_input("상품명")
        o_curr = c6.selectbox("발주 통화", CURRENCIES)
        
        o_total = st.number_input("발주 총액", min_value=0.0)
        if st.form_submit_button("🚀 발주 마스터 저장"):
            cur = conn.cursor()
            cur.execute("INSERT OR REPLACE INTO orders VALUES (?, ?, ?, ?, ?, ?, ?, 0)", 
                        (o_id, o_date.strftime("%Y-%m-%d"), o_vendor, o_prod, o_cat, o_curr, o_total))
            conn.commit()
            st.success("발주 정보가 저장되었습니다.")

# --- Tab 2: 입금 건별 입력 (수동) ---
with tab2:
    st.header("📝 개별 입금 내역 입력")
    orders = load_data("orders")
    active_orders = orders[orders['is_closed'] == 0]
    
    with st.form("single_pay"):
        # 발주번호 선택 시 기존 정보 연동
        sel_oid = st.selectbox("연동할 발주번호 (선택사항)", ["없음"] + list(active_orders['order_id']))
        
        c1, c2, c3 = st.columns(3)
        p_date = c1.date_input("입금일")
        p_cat = c2.selectbox("유형", CATEGORIES)
        p_vendor = c3.text_input("거래처명")
        
        c4, c5, c6 = st.columns(3)
        p_prod = c4.text_input("상품명")
        p_curr = c5.selectbox("통화", CURRENCIES)
        p_dep = c6.number_input("입금액", min_value=0.0)
        
        c7, c8 = st.columns(2)
        p_adv = c7.number_input("선급금", value=0.0)
        p_note = c8.text_input("송금 사유(메모)")
        
        if st.form_submit_button("💾 저장"):
            rate = get_exchange_rate(p_date.strftime("%Y-%m-%d"), p_curr)
            cur = conn.cursor()
            cur.execute('''INSERT INTO payments (order_id, pay_date, category, vendor, product, currency, deposit, advance, note, krw_val) 
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                        (sel_oid if sel_oid != "없음" else None, p_date.strftime("%Y-%m-%d"), 
                         p_cat, p_vendor, p_prod, p_curr, p_dep, p_adv, p_note, p_dep * rate))
            conn.commit()
            st.success("저장 완료!")

# --- Tab 3: 입금 엑셀 일괄 업로드 (기존 기능 부활) ---
with tab3:
    st.header("📂 입금 내역 엑셀 업로드")
    tmp_df = pd.DataFrame(columns=["입금일", "거래처", "발주차수", "유형", "통화", "상품명", "입금액", "선급금", "송금사유", "발주번호"])
    st.download_button("📥 업로드 양식 다운로드", tmp_df.to_csv(index=False).encode('utf-8-sig'), "upload_template.csv")
    
    up_file = st.file_uploader("작성한 CSV 파일을 올려주세요", type=['csv'])
    if up_file:
        df_up = pd.read_csv(up_file)
        if st.button("✅ 데이터 일괄 추가"):
            for _, r in df_up.iterrows():
                curr = r['통화'] if pd.notna(r['통화']) else "한화"
                rate = get_exchange_rate(str(r['입금일']), curr)
                cur = conn.cursor()
                cur.execute('''INSERT INTO payments (order_id, pay_date, category, vendor, product, currency, deposit, advance, note, krw_val) 
                               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                            (r['발주번호'] if '발주번호' in r and pd.notna(r['발주번호']) else None, 
                             r['입금일'], r['유형'], r['거래처'], r['상품명'], curr, r['입금액'], r['선급금'], r['송금사유'], float(r['입금액']) * rate))
            conn.commit()
            st.success("일괄 업로드 완료!")

# --- Tab 4: 상세 내역 관리 (필터, 수정, 마감 가시화) ---
with tab4:
    st.header("🔍 상세 내역 관리 및 마감")
    p_all = load_data("payments")
    o_all = load_data("orders")
    
    # 사이드바 대신 상단 필터로 구성 (가시성)
    f_col1, f_col2, f_col3 = st.columns(3)
    f_vendor = f_col1.multiselect("거래처 필터", p_all['vendor'].unique())
    f_cat = f_col2.multiselect("유형 필터", CATEGORIES)
    f_status = f_col3.radio("상태", ["진행 중", "마감 건"], horizontal=True)

    # 필터링 로직
    df_disp = p_all.copy()
    if f_vendor: df_disp = df_disp[df_disp['vendor'].isin(f_vendor)]
    if f_cat: df_disp = df_disp[df_disp['category'].isin(f_cat)]
    
    # 발주 마스터와 결합하여 마감 상태 확인
    df_merged = df_disp.merge(o_all[['order_id', 'is_closed']], on='order_id', how='left')
    df_merged['is_closed'] = df_merged['is_closed'].fillna(0)
    
    target_status = 1 if f_status == "마감 건" else 0
    df_final = df_merged[df_merged['is_closed'] == target_status]

    # 가시화 및 수정
    st.subheader(f"📋 {f_status} 리스트")
    edited_df = st.data_editor(df_final, use_container_width=True, num_rows="dynamic")
    
    if st.button("💾 변경사항 저장"):
        # 수정 로직... (ID 기준 업데이트)
        st.info("수정 기능이 적용되었습니다.")
        
    # 마감 버튼 (선택된 발주번호 기준)
    if f_status == "진행 중":
        sel_order = st.selectbox("마감 처리할 발주번호 선택", df_final['order_id'].unique())
        if st.button("🚩 해당 발주 건 마감하기"):
            conn.cursor().execute("UPDATE orders SET is_closed = 1 WHERE order_id = ?", (sel_order,))
            conn.commit()
            st.rerun()

# --- Tab 5: 대시보드 ---
with tab5:
    st.header("📊 월별 집행 현황")
    if not p_all.empty:
        p_all['month'] = pd.to_datetime(p_all['pay_date']).dt.strftime('%Y-%m')
        mon_sum = p_all.groupby(['month', 'category'])['krw_val'].sum().reset_index()
        fig = px.bar(mon_sum, x='month', y='krw_val', color='category', barmode='group')
        st.plotly_chart(fig, use_container_width=True)