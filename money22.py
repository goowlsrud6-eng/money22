import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
import numpy as np

# 1. 페이지 설정
st.set_page_config(page_title="자금 관리 시스템 v6", layout="wide", page_icon="💰")

# 2. DB 연결 및 초기화
@st.cache_resource
def get_db_connection():
    conn = sqlite3.connect('money_management_v6.db', check_same_thread=False)
    c = conn.cursor()
    # 테이블 구조 생성 (마감여부, 환산금액 등 포함)
    c.execute('''CREATE TABLE IF NOT EXISTS history 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                  거래처 TEXT, 발주차수 TEXT, 유형 TEXT, 통화 TEXT, 상품명 TEXT, 
                  입금일 TEXT, 입금액 REAL, 선급금 REAL, 송금사유 TEXT, 
                  한화환산액 REAL, 마감여부 INTEGER DEFAULT 0)''')
    conn.commit()
    return conn

conn = get_db_connection()

# 3. 가상 환율 데이터 (실제 운영 시 API 연동 가능, 현재는 월평균 가정치)
def get_monthly_exchange_rate(date_str, currency):
    if currency == "한화" or not currency: return 1.0
    # 예시: 월별 고정 환율 (실제로는 로직에 따라 계산된 값 입력)
    rates = {
        "달러": {"01": 1320, "02": 1340, "03": 1350},
        "위안": {"01": 185, "02": 188, "03": 190}
    }
    month = date_str.split("-")[1]
    return rates.get(currency, {}).get(month, 1300 if currency == "달러" else 180)

# 4. 데이터 로드
def load_data():
    df = pd.read_sql("SELECT * FROM history", conn)
    df['입금일'] = pd.to_datetime(df['입금일'])
    return df

df_all = load_data()

# --- 사이드바 필터 영역 ---
st.sidebar.header("🔍 상세 필터 및 구분")
view_mode = st.sidebar.radio("구분 모드", ["업체별 구분", "유형별 구분"])

# 필터 공통 요소
all_vendors = sorted(df_all['거래처'].unique()) if not df_all.empty else []
all_types = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비"]

if view_mode == "업체별 구분":
    sel_vendor = st.sidebar.selectbox("거래처 선택", ["전체"] + all_vendors)
    filtered_df = df_all if sel_vendor == "전체" else df_all[df_all['거래처'] == sel_vendor]
else:
    sel_type = st.sidebar.selectbox("유형 선택", ["전체"] + all_types)
    filtered_df = df_all if sel_type == "전체" else df_all[df_all['유형'] == sel_type]

# 날짜 및 품목 추가 필터
if not filtered_df.empty:
    search_product = st.sidebar.text_input("상품명 검색")
    sort_order = st.sidebar.selectbox("날짜 정렬", ["내림차순", "오름차순"])
    
    if search_product:
        filtered_df = filtered_df[filtered_df['상품명'].str.contains(search_product, na=False)]
    
    filtered_df = filtered_df.sort_values(by="입금일", ascending=(sort_order == "오름차순"))

# --- 메인 화면 ---
tab1, tab2, tab3 = st.tabs(["📂 엑셀 일괄 업로드", "📋 상세 내역 관리", "📝 개별 입력"])

# --- Tab 1: 엑셀 업로드 ---
with tab1:
    st.subheader("📂 한글 양식 업로드")
    
    # 양식 다운로드
    template_cols = ["입금일", "거래처", "발주차수", "유형", "통화", "상품명", "입금액", "선급금", "송금사유"]
    tmp_df = pd.DataFrame(columns=template_cols)
    st.download_button("📥 한글 업로드 양식 받기", tmp_df.to_csv(index=False).encode('utf-8-sig'), "양식.csv")
    
    up_file = st.file_uploader("CSV 파일 업로드", type=['csv'])
    if up_file:
        up_df = pd.read_csv(up_file)
        if st.button("🚀 데이터 일괄 저장"):
            for _, row in up_df.iterrows():
                curr = row['통화'] if pd.notna(row['통화']) else "한화"
                rate = get_monthly_exchange_rate(str(row['입금일']), curr)
                krw_val = float(row['입금액']) * rate
                
                cur = conn.cursor()
                cur.execute('''INSERT INTO history (입금일, 거래처, 발주차수, 유형, 통화, 상품명, 입금액, 선급금, 송금사유, 한화환산액) 
                               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                            (row['입금일'], row['거래처'], row['발주차수'], row['유형'], curr, row['상품명'], 
                             row['입금액'], row['선급금'], row['송금사유'], krw_val))
            conn.commit()
            st.success("업로드 완료!")
            st.rerun()

# --- Tab 2: 상세 내역 관리 (핵심 기능) ---
with tab2:
    if not filtered_df.empty:
        # 요약 정보 (업체별일 때 잔여 선급금 등 계산)
        if view_mode == "업체별 구분" and sel_vendor != "전체":
            col1, col2 = st.columns(2)
            total_dep = filtered_df['입금액'].sum()
            total_adv = filtered_df['선급금'].sum()
            col1.metric("총 입금액", f"{total_dep:,.0f}")
            col2.metric("잔여 선급금 합계", f"{total_adv:,.0f}")

        st.markdown("---")
        
        # 마감 처리를 위한 가시화 (스타일링 함수)
        def highlight_closed(row):
            if row['마감여부'] == 1:
                return ['background-color: #e0e0e0; color: #9e9e9e'] * len(row)
            return [''] * len(row)

        # 데이터 에디터
        st.write("💡 **마감여부**를 체크(1)하면 해당 행이 회색으로 표시됩니다.")
        edited_df = st.data_editor(
            filtered_df,
            column_config={
                "마감여부": st.column_config.CheckboxColumn("마감", default=False),
                "입금일": st.column_config.DateColumn("입금일"),
                "입금액": st.column_config.NumberColumn("입금액", format="%f"),
                "한화환산액": st.column_config.NumberColumn("한화 환산(월평균)", disabled=True)
            },
            use_container_width=True,
            num_rows="dynamic",
            key="main_editor"
        )

        # 스타일 적용된 테이블 보기 (가시화 전용)
        st.subheader("👀 가시화 뷰 (마감건 확인)")
        st.dataframe(edited_df.style.apply(highlight_closed, axis=1), use_container_width=True)

        if st.button("💾 변경사항 최종 저장"):
            # 기존 필터링된 데이터 기반으로 전체 DB 업데이트 로직 (id 기준)
            for _, row in edited_df.iterrows():
                cur = conn.cursor()
                cur.execute('''UPDATE history SET 
                               거래처=?, 발주차수=?, 유형=?, 통화=?, 상품명=?, 
                               입금일=?, 입금액=?, 선급금=?, 송금사유=?, 마감여부=? 
                               WHERE id=?''', 
                            (row['거래처'], row['발주차수'], row['유형'], row['통화'], row['상품명'], 
                             str(row['입금일'])[:10], row['입금액'], row['선급금'], row['송금사유'], row['마감여부'], row['id']))
            conn.commit()
            st.success("저장되었습니다.")
            st.rerun()
            
        # 엑셀 다운로드
        csv_download = edited_df.to_csv(index=False).encode('utf-8-sig')
        st.download_button("📥 현재 리스트 엑셀(CSV) 다운로드", csv_download, f"detail_{datetime.now().strftime('%m%d')}.csv")
    else:
        st.info("조건에 맞는 데이터가 없습니다.")

# --- Tab 3: 개별 입력 ---
with tab3:
    with st.form("single_input"):
        c1, c2, c3 = st.columns(3)
        v = c1.text_input("거래처명")
        o = c2.text_input("발주차수 (예: 초도2차)")
        t = c3.selectbox("유형", all_types)
        
        c4, c5, c6 = st.columns(3)
        curr = c4.selectbox("통화", ["한화", "달러", "위안"])
        prod = c5.text_input("상품명")
        dt = c6.date_input("입금일")
        
        c7, c8, c9 = st.columns(3)
        amt = c7.number_input("입금액", min_value=0.0)
        adv = c8.number_input("선급금", min_value=0.0)
        memo = c9.text_input("송금사유(메모)")
        
        if st.form_submit_button("저장"):
            rate = get_monthly_exchange_rate(dt.strftime("%Y-%m-%d"), curr)
            krw_val = amt * rate
            cur = conn.cursor()
            cur.execute('''INSERT INTO history (입금일, 거래처, 발주차수, 유형, 통화, 상품명, 입금액, 선급금, 송금사유, 한화환산액) 
                           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                        (dt.strftime("%Y-%m-%d"), v, o, t, curr, prod, amt, adv, memo, krw_val))
            conn.commit()
            st.rerun()