import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime
import plotly.express as px
import io

# 1. DB 설정 및 초기화 (데이터 영구 저장)
def init_db():
    conn = sqlite3.connect('finance_management.db', check_same_thread=False)
    c = conn.cursor()
    # 테이블 구조: id, 날짜, 월, 대분류, 업체명, 상품명, 발주유형, 통화, 발주총액, 실제입금액, 선급금변동, 비고
    c.execute('''CREATE TABLE IF NOT EXISTS history 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, 
                  pay_date TEXT, 
                  month TEXT, 
                  category TEXT, 
                  vendor TEXT, 
                  product TEXT, 
                  order_type TEXT, 
                  currency TEXT, 
                  total_order_amt REAL, 
                  deposit REAL, 
                  advance_change REAL, 
                  note TEXT)''')
    conn.commit()
    return conn

conn = init_db()

# 페이지 설정
st.set_page_config(page_title="자금 관리 시스템 v5", layout="wide", page_icon="💰")

# 데이터 로드 함수
def load_data():
    return pd.read_sql("SELECT * FROM history", conn)

df_all = load_data()
CATEGORIES = ["건기식", "사입", "제작(국내)", "수입(외화)", "물류비", "물품대"]
CURRENCIES = ["KRW", "USD", "CNY"]

# --- 사이드바: 퀵 서머리 ---
st.sidebar.header("📊 Quick Summary")
if not df_all.empty:
    for curr in CURRENCIES:
        curr_sum = df_all[df_all['currency'] == curr]['deposit'].sum()
        if curr_sum > 0:
            st.sidebar.metric(f"누적 집행 ({curr})", f"{curr_sum:,.2f}")

# --- 메인 탭 구성 ---
tab1, tab2, tab3, tab4 = st.tabs(["📝 입금/정산 입력", "📊 실시간 대시보드", "🔍 상세 내역 관리", "📂 엑셀 일괄 업로드"])

# --- Tab 1: 입금/정산 입력 ---
with tab1:
    st.header("✨ 새로운 내역 등록")
    with st.form("input_form", clear_on_submit=True):
        c1, c2, c3, c4 = st.columns(4)
        category = c1.selectbox("대분류", CATEGORIES)
        vendor = c2.text_input("업체명 (예: 우일코리아)")
        product = c3.text_input("상품명 (예: 에어메쉬)")
        order_type = c4.text_input("발주유형 (예: 초도, 리오더1)")

        st.divider()
        c5, c6, c7 = st.columns(3)
        currency = c5.selectbox("통화", CURRENCIES)
        total_order_amt = c6.number_input("해당 건 발주 총액", min_value=0.0, step=100.0)
        pay_date = c7.date_input("입금일", datetime.now())

        c8, c9, c10 = st.columns(3)
        deposit = c8.number_input("이번 실제 입금액", min_value=0.0, step=100.0)
        advance_change = c9.number_input("선급금 변동 (+적립/-차감)", value=0.0, step=100.0)
        note = c10.text_area("메모/사유", height=68)

        submit = st.form_submit_button("🚀 데이터 저장하기", use_container_width=True)
        
        if submit:
            if vendor and product:
                cur = conn.cursor()
                cur.execute('''INSERT INTO history (pay_date, month, category, vendor, product, 
                               order_type, currency, total_order_amt, deposit, advance_change, note) 
                               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)''',
                            (pay_date.strftime("%Y-%m-%d"), pay_date.strftime("%Y-%m"), 
                             category, vendor, product, order_type, currency, 
                             total_order_amt, deposit, advance_change, note))
                conn.commit()
                st.success(f"✅ {vendor} 내역 저장 완료!")
                st.rerun()
            else:
                st.warning("⚠️ 업체명과 상품명은 필수 입력 사항입니다.")

# --- Tab 2: 실시간 대시보드 ---
with tab2:
    if not df_all.empty:
        st.header("📈 자금 흐름 분석")
        
        # 월별 지출 추이 차트
        chart_df = df_all.groupby(['month', 'currency'])['deposit'].sum().reset_index()
        fig = px.bar(chart_df, x='month', y='deposit', color='currency', 
                     title="월별/통화별 실제 입금 합계", barmode='group',
                     labels={'deposit': '입금액', 'month': '기준월'})
        st.plotly_chart(fig, use_container_width=True)
        
        st.divider()
        
        # 업체별 정산 현황 테이블
        st.subheader("🚩 업체별/품목별 정산 요약")
        summary = df_all.groupby(['업체명', '상품명', 'currency']).agg({
            'total_order_amt': 'max', # 발주총액은 해당 건의 고정값으로 가정
            'deposit': 'sum',
            'advance_change': 'sum'
        }).reset_index()
        summary['미결제 잔액'] = summary['total_order_amt'] - summary['deposit']
        
        st.dataframe(summary.style.format({
            'total_order_amt': '{:,.0f}',
            'deposit': '{:,.0f}',
            'advance_change': '{:,.0f}',
            '미결제 잔액': '{:,.0f}'
        }), use_container_width=True)
    else:
        st.info("데이터가 없습니다. 첫 번째 탭에서 내역을 입력해주세요.")

# --- Tab 3: 상세 내역 관리 ---
with tab3:
    st.header("🔍 데이터 수정 및 삭제")
    st.caption("표 안의 셀을 더블클릭하여 수정한 후 하단의 저장 버튼을 누르세요.")
    
    if not df_all.empty:
        # 수정 가능한 데이터 에디터
        edited_df = st.data_editor(
            df_all, 
            use_container_width=True, 
            num_rows="dynamic",
            column_config={
                "id": st.column_config.NumberColumn("ID", disabled=True),
                "pay_date": st.column_config.DateColumn("입금일"),
                "total_order_amt": st.column_config.NumberColumn("발주총액"),
                "deposit": st.column_config.NumberColumn("실제입금액")
            }
        )
        
        c_save, c_empty = st.columns([1, 4])
        if c_save.button("💾 변경사항 DB 최종 저장", type="primary", use_container_width=True):
            edited_df.to_sql('history', conn, if_exists='replace', index=False)
            st.success("수정사항이 반영되었습니다.")
            st.rerun()
    else:
        st.info("관리할 데이터가 없습니다.")

# --- Tab 4: 엑셀 일괄 업로드 ---
with tab4:
    st.header("📂 엑셀(CSV) 일괄 업로드")
    
    col_guide, col_upload = st.columns([1, 1])
    
    with col_guide:
        st.subheader("1. 양식 다운로드")
        st.write("정해진 양식에 맞춰 데이터를 작성해야 오류 없이 업로드됩니다.")
        
        # 양식 데이터 생성 (유저가 참고할 샘플)
        template = pd.DataFrame({
            "pay_date": ["2024-05-01"],
            "category": ["사입"],
            "vendor": ["업체명"],
            "product": ["상품명"],
            "order_type": ["초도"],
            "currency": ["KRW"],
            "total_order_amt": [1000000],
            "deposit": [1000000],
            "advance_change": [0],
            "note": ["메모"]
        })
        
        csv_template = template.to_csv(index=False).encode('utf-8-sig')
        st.download_button(
            label="📥 업로드용 CSV 양식 다운로드",
            data=csv_template,
            file_name="finance_template.csv",
            mime="text/csv",
            use_container_width=True
        )
        
        with st.expander("💡 작성 가이드"):
            st.markdown("""
            - **pay_date**: `YYYY-MM-DD` 형식으로 작성 (예: 2024-05-25)
            - **currency**: `KRW`, `USD`, `CNY` 중 대문자로 입력
            - **숫자 항목**: 콤마(,) 없이 숫자만 입력
            - **파일 형식**: 반드시 CSV(UTF-8) 형식으로 저장
            """)

    with col_upload:
        st.subheader("2. 파일 업로드")
        uploaded_file = st.file_uploader("작성한 CSV 파일을 선택하세요.", type=['csv'])
        
        if uploaded_file:
            try:
                up_df = pd.read_csv(uploaded_file)
                # 필수 컬럼 검증
                required = ["pay_date", "category", "vendor", "product", "currency", "total_order_amt", "deposit"]
                if all(col in up_df.columns for col in required):
                    st.success("✅ 양식 검증 완료")
                    st.dataframe(up_df.head(3), use_container_width=True)
                    
                    if st.button("🚀 DB에 데이터 일괄 추가", use_container_width=True):
                        # month 컬럼 자동 생성
                        up_df['month'] = pd.to_datetime(up_df['pay_date']).dt.strftime('%Y-%m')
                        # DB 추가
                        up_df.to_sql('history', conn, if_exists='append', index=False)
                        st.success(f"🎉 총 {len(up_df)}건의 데이터가 성공적으로 추가되었습니다!")
                        st.rerun()
                else:
                    st.error("❌ 필수 항목이 누락되었습니다. 양식을 다시 확인해주세요.")
            except Exception as e:
                st.error(f"오류 발생: {e}")

# 전체 데이터 백업 버튼 (하단)
st.sidebar.divider()
if not df_all.empty:
    full_csv = df_all.to_csv(index=False).encode('utf-8-sig')
    st.sidebar.download_button("📥 전체 데이터 백업(CSV)", full_csv, f"full_backup_{datetime.now().strftime('%Y%m%d')}.csv", use_container_width=True)