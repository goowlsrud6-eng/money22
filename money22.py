import streamlit as st
import pandas as pd
import re
import os
from datetime import datetime, timedelta
import plotly.graph_objects as go
from supabase import create_client, Client

# ==============================================================================
# 1. 초기 설정 및 Supabase 연결 (v136_Cloud_Full)
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
    """클라우드에서 데이터를 안전하게 가져오는 표준 함수"""
    try:
        res = supabase.table(table_name).select("*").execute()
        if not res.data: return pd.DataFrame()
        return pd.DataFrame(res.data)
    except Exception as e:
        return pd.DataFrame()

def upsert_supabase_data(table_name, data):
    """데이터 저장 및 수정 (Upsert)"""
    try:
        if not data: return True
        supabase.table(table_name).upsert(data).execute()
        return True
    except Exception as e:
        st.error(f"{table_name} 저장 실패: {e}")
        return False

def get_multiple_available_ids(count):
    """[복구] v136의 핵심: 삭제된 번호를 찾아주는 ID 재사용 로직"""
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

def process_ecount_v136_cloud(file):
    """[복구] 이카운트 발주서 엑셀 정밀 분석기 (원본 로직 100% 동일)"""
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
            
        v_fixed = match.iloc[0]['거래처명']
        v_type = match.iloc[0]['기본유형']
        
        f6_val = str(df.iloc[5, 5]) if len(df) > 5 else ""
        curr = "USD" if "USD" in f6_val else ("CNY" if any(x in f6_val for x in ["중국", "CNY"]) else "한화")
        
        prods = df.iloc[6:, 1 if curr == "한화" else 2].dropna().astype(str).tolist()
        prod_n = (prods[0].split("[")[0].strip() + (f" 외 {len(prods)-1}건" if len(prods)>1 else "")) if prods else "품목미상"
        
        l_idx = df.iloc[:, 5].last_valid_index()
        total = to_float(df.iloc[l_idx, 5]) if curr != "한화" and l_idx else to_float(str(df.iloc[4, 0]).split(":")[-1])
        
        upsert_supabase_data("orders", {
            "발주번호": raw_oid, "발주일": odate, "거래처명": v_fixed, 
            "상품명": prod_n, "유형": v_type, "통화": curr, "발주총액": total, "마감여부": 0
        })
        return True, None
    except Exception as e: return False, str(e)

# ==============================================================================
# 3. 유틸리티 함수 (smart_date 등 v136 원본 보존)
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
# 4. 메인 UI 및 탭별 로직 (Tab 0 ~ Tab 4 완전체)
# ==============================================================================
tabs = st.tabs(["입금 등록", "발주서 등록", "상세내역 및 정산", "거래처 관리", "환율 분석"])

# --- [Tab 0] 입금 내역 등록 (지능형 CSV 업로드 포함) ---
with tabs[0]:
    st.header("입금 내역 등록 및 관리")
    v_master, o_data = get_supabase_data("vendors"), get_supabase_data("orders")
    o_active = o_data[o_data['마감여부'] == 0] if not o_data.empty else pd.DataFrame()
    
    col_input, col_excel = st.columns([1.5, 1])
    with col_input:
        st.subheader("1. 수기 직접 입력")
        with st.form("manual_pay_form", clear_on_submit=True):
            p_oid = st.selectbox("발주번호 연동", ["없음"] + (list(o_active['발주번호']) if not o_active.empty else []))
            p_date = st.date_input("입금일자", value=datetime.now())
            auto_prod = o_active[o_active['발주번호'] == p_oid]['상품명'].values[0] if p_oid != "없음" else ""
            p_vn = st.selectbox("거래처 선택", ["선택"] + (list(v_master['거래처명']) if not v_master.empty else []))
            p_ct, p_pr = st.selectbox("유형 분류", CATEGORIES), st.text_input("상품명", value=auto_prod)
            r3c1, r3c2, r3c3 = st.columns(3)
            p_dep, p_pre, p_cur = r3c1.number_input("실입금액"), r3c2.number_input("선급금액"), r3c3.selectbox("거래통화", ["한화", "USD", "CNY"])
            p_memo = st.text_input("비고 (송금 사유 등)")
            if st.form_submit_button("입금 내역 저장"):
                if p_vn == "선택": st.error("거래처를 선택하세요.")
                else:
                    vi = v_master[v_master['거래처명']==p_vn].iloc[0]
                    upsert_supabase_data("payments", {"id": get_multiple_available_ids(1)[0], "발주번호": p_oid if p_oid != "없음" else None, "입금일": p_date.strftime("%Y-%m-%d"), "유형": p_ct, "거래처명": p_vn, "상품명": p_pr, "통화": p_cur, "실입금액": p_dep, "선급금액": p_pre, "메모": p_memo, "은행": vi['은행'], "계좌번호": vi['계좌번호'], "예금주": vi['예금주']})
                    st.success("저장 완료"); st.rerun()

    with col2:
        st.subheader("2. CSV 일괄 업로드 (v136 지능형 매칭)")
        csv_template = pd.DataFrame(columns=["발주번호", "거래처", "유형", "상품명", "입금일", "실입금액", "선급금액", "송금사유"])
        st.download_button("양식 다운로드", csv_template.to_csv(index=False).encode('utf-8-sig'), "payment_template.csv")
        up_pay = st.file_uploader("CSV 선택", type=['csv'], key=f"pay_up_{st.session_state.pay_up_key}")
        if up_pay and st.button("파일 일괄 저장 실행"):
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

# --- [Tab 1] 발주서 등록 및 관리 (이카운트 분석 복구) ---
with tabs[1]:
    st.header("발주서 등록 및 마감")
    c1, c2 = st.columns([1, 1.5])
    with c1:
        st.subheader("1. 발주 분석 및 등록")
        o_files = st.file_uploader("이카운트 엑셀 선택", type=['xlsx'], accept_multiple_files=True, key=f"ord_f_{st.session_state.order_up_key}")
        if o_files and st.button("발주서 일괄 분석 실행"):
            for f in o_files: process_ecount_v136_cloud(f)
            st.session_state.order_up_key += 1; st.rerun()
        with st.form("manual_ord"):
            st.write("**수기 발주 입력**")
            m_oid, m_vn = st.text_input("발주번호"), st.selectbox("거래처", ["선택"] + (list(v_master['거래처명']) if not v_master.empty else []))
            m_amt, m_cur = st.number_input("발주총액"), st.selectbox("통화", ["한화", "USD", "CNY"])
            if st.form_submit_button("발주 저장"):
                if m_oid and m_vn != "선택":
                    upsert_supabase_data("orders", {"발주번호": m_oid, "발주일": datetime.now().strftime("%Y-%m-%d"), "거래처명": m_vn, "발주총액": m_amt, "통화": m_cur, "마감여부": 0})
                    st.rerun()
    with c2:
        st.subheader("2. 발주 목록 및 소급 수정")
        if not o_data.empty:
            ev_o = st.data_editor(o_data.sort_values('발주일', ascending=False), hide_index=True, use_container_width=True)
            if st.button("수정 내용 및 연동 정보 저장"):
                upsert_supabase_data("orders", ev_o.to_dict(orient='records'))
                for _, r in ev_o.iterrows():
                    supabase.table("payments").update({"거래처명": r['거래처명'], "유형": r['유형'], "상품명": r['상품명']}).eq("발주번호", r['발주번호']).execute()
                st.success("동기화 완료"); st.rerun()

# --- [Tab 2] 상세 및 정산 (한화환산액 & 정밀 잔액 복구) ---
with tabs[2]:
    st.header("상세 내역 및 통합 정산")
    p_all, o_all, ex_rates = get_supabase_data("payments"), get_supabase_data("orders"), get_supabase_data("exchange_rates")
    if not p_all.empty:
        p_all['dt'] = pd.to_datetime(p_all['입금일'])
        f_c1, f_c2, f_c3, f_c4 = st.columns(4)
        target_m = f_c2.selectbox("조회 월", ["전체"] + list(range(1, 13)))
        search_key = f_c4.text_input("업체/상품 통합 검색")
        filtered = p_all.copy()
        if target_m != "전체": filtered = filtered[filtered['dt'].dt.month == int(target_m)]
        if search_key: filtered = filtered[filtered['거래처명'].str.contains(search_key, case=False, na=False) | filtered['상품명'].str.contains(search_key, case=False, na=False)]
        
        def get_v136_conversion(row):
            if row['통화'] == '한화': return to_float(row['실입금액'])
            ym_key, curr_key = str(row['입금일'])[:7], row['통화'].lower()
            if not ex_rates.empty:
                ex_rates['ym'] = pd.to_datetime(ex_rates['날짜']).dt.strftime('%Y-%m')
                avg = ex_rates[ex_rates['ym'] == ym_key][curr_key].mean()
                if not pd.isna(avg) and avg > 0: return to_float(row['실입금액']) * avg
            return to_float(row['실입금액']) * (1350.0 if row['통화'] == 'USD' else 190.0)

        filtered['한화환산액'] = filtered.apply(get_v136_conversion, axis=1)
        st.subheader("📊 발주번호별 정산 및 잔액 (v136 정밀 공식)")
        pay_agg = p_all.groupby('발주번호').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
        settle_df = pd.merge(o_all, pay_agg, on='발주번호', how='left').fillna(0)
        settle_df['잔액'] = settle_df['발주총액'] - (settle_df['실입금액'] + settle_df['선급금액'])
        settle_df['상태'] = settle_df['마감여부'].apply(lambda x: "✅ 마감" if x == 1 else "⏳ 진행")
        st.dataframe(settle_df[['발주번호','상태','거래처명','상품명','발주총액','실입금액','선급금액','잔액','통화']].sort_values('발주번호', ascending=False), use_container_width=True)

        st.subheader("📝 상세 내역 수정")
        edit_cols = ['id', '유형', '발주번호', '거래처명', '상품명', '입금일', '통화', '실입금액', '선급금액', '한화환산액', '메모']
        edited_p = st.data_editor(filtered[edit_cols].sort_values('입금일', ascending=False), hide_index=True, use_container_width=True)
        if st.button("수정 내용 클라우드 저장"):
            upsert_supabase_data("payments", edited_p.to_dict(orient='records')); st.rerun()
        st.divider()
        m1, m2, m3 = st.columns(3)
        m1.metric("총 환산액 합계", f"{filtered['한화환산액'].sum():,.0f} 원")
        m2.metric("USD 합계", f"${filtered[filtered['통화']=='USD']['실입금액'].sum():,.2f}")
        m3.metric("CNY 합계", f"¥{filtered[filtered['통화']=='CNY']['실입금액'].sum():,.2f}")
    else: st.info("데이터가 없습니다.")

# --- [Tab 3] 거래처 관리 (소급 업데이트 복구) ---
with tabs[3]:
    st.header("거래처 정보 관리")
    v_orig = get_supabase_data("vendors")
    with st.form("new_v_form"):
        vn, vt = st.text_input("거래처명"), st.selectbox("유형", CATEGORIES)
        if st.form_submit_button("거래처 등록"):
            if vn: upsert_supabase_data("vendors", {"거래처명": vn, "기본유형": vt}); st.rerun()
    if not v_orig.empty:
        ev_v = st.data_editor(v_orig, hide_index=True, use_container_width=True)
        if st.button("정보 저장 및 데이터 일괄 연동"):
            for i, r in ev_v.iterrows():
                if i < len(v_orig) and v_orig.iloc[i]['거래처명'] != r['거래처명']:
                    old_n = v_orig.iloc[i]['거래처명']
                    supabase.table("payments").update({"거래처명": r['거래처명'], "유형": r['기본유형']}).eq("거래처명", old_n).execute()
                    supabase.table("orders").update({"거래처명": r['거래처명'], "유형": r['기본유형']}).eq("거래처명", old_n).execute()
            upsert_supabase_data("vendors", ev_v.to_dict(orient='records')); st.success("동기화 완료"); st.rerun()

# --- [Tab 4] 환율 관리 ---
with tabs[4]:
    st.header("환율 분석")
    def up_ex(u, cur):
        df = pd.read_csv(u)
        upsert_supabase_data("exchange_rates", [{"날짜": smart_date(r['날짜']), cur.lower(): to_float(r['종가'])} for _, r in df.iterrows()])
    c1, c2 = st.columns(2)
    with c1:
        u_u = st.file_uploader("USD CSV", type=['csv'])
        if u_u and st.button("USD 업로드"): up_ex(u_u, "USD"); st.rerun()
    with c2:
        u_c = st.file_uploader("CNY CSV", type=['csv'])
        if u_c and st.button("CNY 업로드"): up_ex(u_c, "CNY"); st.rerun()
    ex_db = get_supabase_data("exchange_rates")
    if not ex_db.empty:
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=ex_db['날짜'], y=ex_db['usd'], name="USD"))
        fig.add_trace(go.Scatter(x=ex_db['날짜'], y=ex_db['cny'], name="CNY"))
        st.plotly_chart(fig, use_container_width=True)

st.sidebar.success("☁️ v136 Full Cloud Connected")