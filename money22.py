import streamlit as st
import pandas as pd
import re
import os
from datetime import datetime, timedelta
import plotly.graph_objects as go
from supabase import create_client, Client

# ==============================================================================
# 1. 초기 설정 및 Supabase 연결
# ==============================================================================
st.set_page_config(page_title="자금 관리 시스템 v136_Full_Cloud", layout="wide")

SUPABASE_URL = "https://nbpeuxblyphzmbktcqtq.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im5icGV1eGJseXBoem1ia3RjcXRxIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzcwMDc1NTEsImV4cCI6MjA5MjU4MzU1MX0.Q6A8T6_JiPIOBnjf8wKtjWTsRAk-pzvKdSqbfPp-3w4"
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

CATEGORIES = ["제작(국내)", "제작(수입)", "사입", "건기식", "물품대", "물류비", "원단비", "기타"]

if 'order_up_key' not in st.session_state: st.session_state.order_up_key = 0
if 'pay_up_key' not in st.session_state: st.session_state.pay_up_key = 1000

# ==============================================================================
# 2. 데이터 엔진 (v136 로직 복구)
# ==============================================================================

def get_supabase_data(table_name):
    try:
        res = supabase.table(table_name).select("*").execute()
        return pd.DataFrame(res.data)
    except: return pd.DataFrame()

def upsert_supabase_data(table_name, data):
    try:
        if not data: return True
        supabase.table(table_name).upsert(data).execute()
        return True
    except Exception as e:
        st.error(f"{table_name} 저장 실패: {e}")
        return False

# [복구] ID 재사용 로직 (벌크 업로드 시 충돌 방지용 리스트 반환)
def get_multiple_available_ids(count):
    df = get_supabase_data("payments")
    if df.empty: return list(range(1, count + 1))
    ids = sorted(df['id'].unique().tolist())
    available = []
    current = 1
    while len(available) < count:
        if current not in ids:
            available.append(current)
        current += 1
    return available

# [복구] 이카운트 엑셀 정밀 분석 (money22.py 로직 100% 동일)
def process_ecount_v136_cloud(file):
    try:
        df = pd.read_excel(file, header=None)
        raw_oid = str(df.iloc[1, 0]).split(":")[-1].strip() if ":" in str(df.iloc[1,0]) else str(df.iloc[1, 0])
        odate = smart_date(raw_oid.replace("-", "")[:8])
        vendor_raw = ""
        for i in range(len(df)):
            if "수신" in str(df.iloc[i, 0]): 
                vendor_raw = str(df.iloc[i, 0]).split(":")[-1].strip()
                break
        v_master = get_supabase_data("vendors")
        v_master['clean'] = v_master['거래처명'].apply(lambda x: re.sub(r'\s+', '', str(x)).lower())
        match = v_master[v_master['clean'] == re.sub(r'\s+', '', vendor_raw).lower()]
        if match.empty: return False, f"미등록 업체: [{vendor_raw}]"
        v_fixed, v_type = match.iloc[0]['거래처명'], match.iloc[0]['기본유형']
        f6_val = str(df.iloc[5, 5]) if len(df) > 5 else ""
        curr = "USD" if "USD" in f6_val else ("CNY" if any(x in f6_val for x in ["중국", "CNY"]) else "한화")
        prods = df.iloc[6:, 1 if curr == "한화" else 2].dropna().astype(str).tolist()
        prod_n = (prods[0].split("[")[0].strip() + (f" 외 {len(prods)-1}건" if len(prods)>1 else "")) if prods else "품목미상"
        l_idx = df.iloc[:, 5].last_valid_index()
        total = to_float(df.iloc[l_idx, 5]) if curr != "한화" and l_idx else to_float(str(df.iloc[4, 0]).split(":")[-1])
        upsert_supabase_data("orders", {"발주번호": raw_oid, "발주일": odate, "거래처명": v_fixed, "상품명": prod_n, "유형": v_type, "통화": curr, "발주총액": total})
        return True, None
    except Exception as e: return False, str(e)

# ==============================================================================
# 3. 유틸리티 (원본 보존)
# ==============================================================================
def to_float(val):
    try:
        if val is None or pd.isna(val) or str(val).strip() == "": return 0.0
        return float(str(val).replace(',', '').strip())
    except: return 0.0

def to_str(val):
    if val is None or pd.isna(val): return ""
    s = str(val).strip()
    return "" if s.lower() in ["nan", "none", ""] else s

def smart_date(date_val):
    try:
        if pd.isna(date_val) or str(date_val).strip() == "": return datetime.now().strftime("%Y-%m-%d")
        if isinstance(date_val, (datetime, pd.Timestamp)): return date_val.strftime("%Y-%m-%d")
        ds = str(date_val).strip()
        ds = re.sub(r'(\d{1,2})월\s*(\d{1,2})일', r'\1-\2', ds)
        if re.match(r'^\d{1,2}[/-]\d{1,2}$', ds): ds = f"{datetime.now().year}-{ds.replace('/', '-')}"
        ds = ds.replace(".", "-").replace("/", "-").replace(" ", "")
        return pd.to_datetime(ds).strftime("%Y-%m-%d")
    except: return datetime.now().strftime("%Y-%m-%d")
# ==============================================================================
# 4. 메인 UI (Tab 0 ~ Tab 4)
# ==============================================================================
tabs = st.tabs(["입금 등록", "발주서 등록", "상세내역 및 정산", "거래처 관리", "환율 분석"])

# --- [Tab 0] 입금 등록 (CSV 지능형 업로드 복구) ---
with tabs[0]:
    v_master = get_supabase_data("vendors")
    o_data = get_supabase_data("orders")
    o_active = o_data[o_data['마감여부'] == 0] if not o_data.empty else pd.DataFrame()
    col1, col2 = st.columns([1.5, 1])
    with col1:
        with st.form("manual_pay", clear_on_submit=True):
            p_oid = st.selectbox("발주번호 연동", ["없음"] + (list(o_active['발주번호']) if not o_active.empty else []))
            p_date = st.date_input("입금일자", value=datetime.now())
            auto_prod = o_active[o_active['발주번호'] == p_oid]['상품명'].values[0] if p_oid != "없음" else ""
            p_vn = st.selectbox("거래처 선택", ["선택"] + (list(v_master['거래처명']) if not v_master.empty else []))
            p_ct = st.selectbox("유형 분류", CATEGORIES)
            p_pr = st.text_input("상품명", value=auto_prod)
            r3c1, r3c2, r3c3 = st.columns(3)
            p_dep, p_pre, p_cur = r3c1.number_input("실입금액"), r3c2.number_input("선급금액"), r3c3.selectbox("거래통화", ["한화", "USD", "CNY"])
            p_memo = st.text_input("비고")
            if st.form_submit_button("저장"):
                if p_vn != "선택":
                    vi = v_master[v_master['거래처명']==p_vn].iloc[0]
                    upsert_supabase_data("payments", {"id": get_multiple_available_ids(1)[0], "발주번호": p_oid if p_oid != "없음" else None, "입금일": p_date.strftime("%Y-%m-%d"), "유형": p_ct, "거래처명": p_vn, "상품명": p_pr, "통화": p_cur, "실입금액": p_dep, "선급금액": p_pre, "메모": p_memo, "은행": vi['은행'], "계좌번호": vi['계좌번호'], "예금주": vi['예금주']})
                    st.rerun()

    with col2:
        csv_template = pd.DataFrame(columns=["발주번호", "거래처", "유형", "상품명", "입금일", "실입금액", "선급금액", "송금사유"])
        st.download_button("CSV 양식", csv_template.to_csv(index=False).encode('utf-8-sig'), "template.csv")
        up_pay = st.file_uploader("CSV 업로드", type=['csv'], key=f"pay_up_{st.session_state.pay_up_key}")
        if up_pay and st.button("일괄 업로드 실행"):
            df_up = pd.read_csv(up_pay)
            df_up.columns = [str(c).strip().replace('\ufeff', '') for c in df_up.columns]
            ids = get_multiple_available_ids(len(df_up))
            up_list = []
            for i, r in df_up.iterrows():
                oid_v, vn_v = to_str(r.get('발주번호')), to_str(r.get('거래처'))
                match_o = o_data[o_data['발주번호'] == oid_v].iloc[0] if oid_v and not o_data[o_data['발주번호'] == oid_v].empty else None
                vn_f = match_o['거래처명'] if match_o is not None else vn_v
                vi = v_master[v_master['거래처명'].str.lower() == vn_f.lower()].iloc[0] if not v_master[v_master['거래처명'].str.lower() == vn_f.lower()].empty else None
                up_list.append({"id": ids[i], "발주번호": oid_v or None, "입금일": smart_date(r.get('입금일')), "유형": match_o['유형'] if match_o is not None else (to_str(r.get('유형')) or "사입"), "거래처명": vn_f, "상품명": match_o['상품명'] if match_o is not None else to_str(r.get('상품명')), "통화": match_o['통화'] if match_o is not None else "한화", "실입금액": to_float(r.get('실입금액')), "선급금액": to_float(r.get('선급금액')), "메모": to_str(r.get('송금사유')), "은행": vi['은행'] if vi is not None else "", "계좌번호": vi['계좌번호'] if vi is not None else "", "예금주": vi['예금주'] if vi is not None else ""})
            if upsert_supabase_data("payments", up_list): st.session_state.pay_up_key += 1; st.rerun()

# --- [Tab 2] 상세 및 정산 (v136 한화환산액 로직 완벽 복구) ---
with tabs[2]:
    st.header("상세 내역 및 통합 정산")
    p_all = get_supabase_data("payments")
    o_all = get_supabase_data("orders")
    ex_rates = get_supabase_data("exchange_rates")
    
    if not p_all.empty:
        p_all['dt'] = pd.to_datetime(p_all['입금일'])
        
        # [복구] v136 정밀 필터 UI
        f_c1, f_c2, f_c3, f_c4 = st.columns(4)
        years = sorted(p_all['dt'].dt.year.unique())
        start_y = f_c1.selectbox("시작 연도", years, index=0)
        end_y = f_c1.selectbox("종료 연도", years, index=len(years)-1)
        target_m = f_c2.selectbox("조회 월", ["전체"] + list(range(1, 13)))
        filter_cat = f_c3.selectbox("유형 필터", ["전체 유형"] + CATEGORIES)
        search_key = f_c4.text_input("업체/상품 통합 검색")
        
        # 필터링 적용
        filtered = p_all[(p_all['dt'].dt.year >= start_y) & (p_all['dt'].dt.year <= end_y)]
        if target_m != "전체": 
            filtered = filtered[filtered['dt'].dt.month == int(target_m)]
        if filter_cat != "전체 유형": 
            filtered = filtered[filtered['유형'] == filter_cat]
        if search_key: 
            filtered = filtered[filtered['거래처명'].str.contains(search_key, case=False, na=False) | 
                                filtered['상품명'].str.contains(search_key, case=False, na=False)]
        
        # [핵심 복구] v136 월별 평균 환율 참조 및 환산액 계산 로직
        def get_v136_conversion(row):
            if row['통화'] == '한화': 
                return to_float(row['실입금액'])
            
            ym_key = str(row['입금일'])[:7] # YYYY-MM
            curr_key = row['통화'].lower() # usd or cny
            
            # 클라우드 환율 데이터에서 해당 월의 평균값 추출
            if not ex_rates.empty:
                ex_rates['ym'] = pd.to_datetime(ex_rates['날짜']).dt.strftime('%Y-%m')
                monthly_avg = ex_rates[ex_rates['ym'] == ym_key][curr_key].mean()
                if not pd.isna(monthly_avg) and monthly_avg > 0:
                    return to_float(row['실입금액']) * monthly_avg
            
            # 환율 데이터 없을 시 v136 fallback 수치 적용
            fallback = 1350.0 if row['통화'] == 'USD' else 190.0
            return to_float(row['실입금액']) * fallback

        # 개별 행마다 한화환산액 계산하여 컬럼 추가
        filtered['한화환산액'] = filtered.apply(get_v136_conversion, axis=1)
        
        # [복구] 유형별 요약 테이블 (환산액 포함)
        st.subheader("📊 유형별 지출 요약")
        summary = filtered.groupby('유형').agg({
            '실입금액': 'sum', 
            '선급금액': 'sum', 
            '한화환산액': 'sum'
        }).reset_index()
        st.table(summary.style.format({
            '실입금액': '{:,.2f}', 
            '선급금액': '{:,.2f}', 
            '한화환산액': '{:,.0f}'
        }))

        # [복구] 발주번호별 정산 및 잔액 현황
        st.subheader("🔍 발주별 정산 및 미수금 현황")
        pay_agg = p_all.groupby('발주번호').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
        settle_df = pd.merge(o_all, pay_agg, on='발주번호', how='left').fillna(0)
        settle_df['잔액'] = settle_df['발주총액'] - (settle_df['실입금액'] + settle_df['선급금액'])
        settle_df['진행상태'] = settle_df['마감여부'].apply(lambda x: "✅ 마감" if x == 1 else "⏳ 진행")
        
        display_cols = ['발주번호', '진행상태', '거래처명', '상품명', '발주총액', '실입금액', '선급금액', '잔액', '통화']
        st.dataframe(settle_df[display_cols].sort_values('발주번호', ascending=False), use_container_width=True)

        # [복구] 상세 내역 편집기 (한화환산액 컬럼 포함)
        st.subheader("📝 상세 내역 수정")
        # 편집 시 불필요한 dt 컬럼 제외, 한화환산액은 참고용으로 배치
        edit_cols = ['id', '유형', '발주번호', '거래처명', '상품명', '입금일', '통화', '실입금액', '선급금액', '한화환산액', '메모']
        edited_p = st.data_editor(
            filtered[edit_cols].sort_values('입금일', ascending=False), 
            hide_index=True, 
            use_container_width=True,
            column_config={
                "한화환산액": st.column_config.NumberColumn("한화환산액(참고)", format="%d"),
                "실입금액": st.column_config.NumberColumn(format="%.2f"),
                "선급금액": st.column_config.NumberColumn(format="%.2f")
            }
        )
        
        if st.button("수정 내용 클라우드 동기화 저장"):
            # 수정한 데이터를 클라우드에 업데이트 (환산액은 DB 컬럼에 맞게 처리)
            upsert_supabase_data("payments", edited_p.to_dict(orient='records'))
            st.success("수정사항이 클라우드에 반영되었습니다.")
            st.rerun()

        # [복구] 하단 메트릭 요약
        st.divider()
        m1, m2, m3 = st.columns(3)
        m1.metric("총 환산액 합계", f"{filtered['한화환산액'].sum():,.0f} 원")
        m2.metric("USD 합계", f"${filtered[filtered['통화']=='USD']['실입금액'].sum():,.2f}")
        m3.metric("CNY 합계", f"¥{filtered[filtered['통화']=='CNY']['실입금액'].sum():,.2f}")
