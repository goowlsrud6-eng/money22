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
    
    # 변수 이름을 col_input, col_excel로 정의
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
                    # v136 ID 재사용 로직 적용
                    upsert_supabase_data("payments", {"id": get_multiple_available_ids(1)[0], "발주번호": p_oid if p_oid != "없음" else None, "입금일": p_date.strftime("%Y-%m-%d"), "유형": p_ct, "거래처명": p_vn, "상품명": p_pr, "통화": p_cur, "실입금액": p_dep, "선급금액": p_pre, "메모": p_memo, "은행": vi['은행'], "계좌번호": vi['계좌번호'], "예금주": vi['예금주']})
                    st.success("저장 완료"); st.rerun()

    # 에러 방지: 위에서 정의한 col_excel 사용
    with col_excel:
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
                # 지능형 필드 매칭 로직
                up_list.append({"id": ids[i], "발주번호": oid_v or None, "입금일": smart_date(r.get('입금일')), "유형": match_o['유형'] if match_o is not None else (to_str(r.get('유형')) or "사입"), "거래처명": vn_f, "상품명": match_o['상품명'] if match_o is not None else to_str(r.get('상품명')), "통화": match_o['통화'] if match_o is not None else "한화", "실입금액": to_float(r.get('실입금액')), "선급금액": to_float(r.get('선급금액')), "메모": to_str(r.get('송금사유')), "은행": vi['은행'] if vi is not None else "", "계좌번호": vi['계좌번호'] if vi is not None else "", "예금주": vi['예금주'] if vi is not None else ""})
            if upsert_supabase_data("payments", up_list): st.session_state.pay_up_key += 1; st.rerun()

# --- [Tab 1] 발주서 등록 및 관리 (v136 마감 및 삭제 로직 보강) ---
with tabs[1]:
    st.header("발주서 등록 및 마감 관리")
    v_master, o_data = get_supabase_data("vendors"), get_supabase_data("orders")
    
    c1, c2 = st.columns([1, 1.5])
    with c1:
        st.subheader("1. 발주 분석 및 등록")
        # 이카운트 엑셀 업로드
        o_files = st.file_uploader("이카운트 엑셀 선택", type=['xlsx'], accept_multiple_files=True, key=f"ord_f_{st.session_state.order_up_key}")
        if o_files and st.button("발주서 일괄 분석 실행"):
            for f in o_files: 
                success, msg = process_ecount_v136_cloud(f)
                if not success: st.error(msg)
            st.session_state.order_up_key += 1; st.success("분석 완료"); st.rerun()
        
        st.divider()
        
        with st.form("manual_ord_form", clear_on_submit=True):
            st.write("**수기 발주 입력**")
            m_oid = st.text_input("발주번호 (필수)")
            m_vn = st.selectbox("거래처 선택", ["선택"] + (list(v_master['거래처명']) if not v_master.empty else []))
            col_m1, col_m2 = st.columns(2)
            m_amt = col_m1.number_input("발주총액", format="%.2f")
            m_cur = col_m2.selectbox("통화", ["한화", "USD", "CNY"])
            m_item = st.text_input("상품명 (선택)")
            
            if st.form_submit_button("발주 저장"):
                if m_oid and m_vn != "선택":
                    # 신규 등록 시 기본값 설정
                    v_type = v_master[v_master['거래처명']==m_vn].iloc[0]['기본유형'] if not v_master.empty else "기타"
                    upsert_supabase_data("orders", {
                        "발주번호": m_oid, 
                        "발주일": datetime.now().strftime("%Y-%m-%d"), 
                        "거래처명": m_vn, 
                        "상품명": m_item or "수기입력",
                        "유형": v_type,
                        "발주총액": m_amt, 
                        "통화": m_cur, 
                        "마감여부": 0
                    })
                    st.success("저장되었습니다."); st.rerun()
                else:
                    st.warning("발주번호와 거래처를 확인하세요.")

    with c2:
        st.subheader("2. 발주 목록 및 마감 처리")
        if not o_data.empty:
            # v136의 핵심: 마감여부를 체크박스로 직관적으로 관리
            ev_o = st.data_editor(
                o_data.sort_values('발주일', ascending=False), 
                hide_index=True, 
                use_container_width=True,
                column_config={
                    "마감여부": st.column_config.CheckboxColumn("마감", help="마감 시 입금등록 목록에서 제외", default=0),
                    "발주총액": st.column_config.NumberColumn(format="%.2f")
                },
                disabled=["발주번호"] # 발주번호 수정 방지
            )
            
            col_btn1, col_btn2 = st.columns(2)
            if col_btn1.button("수정 내용 및 마감 상태 저장"):
                upsert_supabase_data("orders", ev_o.to_dict(orient='records'))
                # 거래처명이나 상품명이 바뀌었을 경우 입금내역도 소급 수정
                for _, r in ev_o.iterrows():
                    supabase.table("payments").update({
                        "거래처명": r['거래처명'], 
                        "유형": r['유형'], 
                        "상품명": r['상품명']
                    }).eq("발주번호", r['발주번호']).execute()
                st.success("동기화 완료"); st.rerun()
                
            if col_btn2.button("⚠️ 선택된 발주 삭제"):
                # 에디터에서 행 삭제 기능 대신, 체크박스나 필터를 활용한 삭제 로직 보강 가능
                st.warning("삭제는 Supabase 대시보드에서 직접 수행하거나 별도 삭제 버튼 로직이 필요합니다.")
        else:
            st.info("등록된 발주 내역이 없습니다.")
            
# --- [Tab 2] 상세 내역 및 통합 정산 (연도 범위 조회 기능 포함) ---
with tabs[2]:
    st.header("상세 내역 및 통합 정산")
    p_all, o_all, ex_rates = get_supabase_data("payments"), get_supabase_data("orders"), get_supabase_data("exchange_rates")
    
    if not p_all.empty:
        p_all['dt'] = pd.to_datetime(p_all['입금일'])
        
        # 1. 연도 범위 필터 (25년~26년 등 이어서 보기 가능)
        f_c1, f_c2, f_c3, f_c4 = st.columns(4)
        years = sorted(p_all['dt'].dt.year.unique())
        
        # 시작 연도와 종료 연도를 선택하여 범위를 만듬
        start_y = f_c1.selectbox("시작 연도", years, index=0)
        end_y = f_c1.selectbox("종료 연도", years, index=len(years)-1)
        
        target_m = f_c2.selectbox("조회 월", ["전체"] + list(range(1, 13)))
        filter_cat = f_c3.selectbox("유형 필터", ["전체"] + CATEGORIES)
        search_key = f_c4.text_input("업체/상품 검색")
        
        # 필터링 적용: 시작 연도 <= 데이터 연도 <= 종료 연도
        filtered = p_all[(p_all['dt'].dt.year >= start_y) & (p_all['dt'].dt.year <= end_y)]
        
        if target_m != "전체": 
            filtered = filtered[filtered['dt'].dt.month == int(target_m)]
        if filter_cat != "전체": 
            filtered = filtered[filtered['유형'] == filter_cat]
        if search_key: 
            filtered = filtered[filtered['거래처명'].str.contains(search_key, case=False, na=False) | 
                              filtered['상품명'].str.contains(search_key, case=False, na=False)]
        
        # 2. 한화 환산 로직 (v136 월평균 환율 적용)
        def get_v136_conversion(row):
            if row['통화'] == '한화': return to_float(row['실입금액'])
            ym_key, curr_key = str(row['입금일'])[:7], row['통화'].lower()
            if not ex_rates.empty:
                ex_rates['ym'] = pd.to_datetime(ex_rates['날짜']).dt.strftime('%Y-%m')
                avg = ex_rates[ex_rates['ym'] == ym_key][curr_key].mean()
                if not pd.isna(avg) and avg > 0: return to_float(row['실입금액']) * avg
            return to_float(row['실입금액']) * (1350.0 if row['통화'] == 'USD' else 190.0)

        filtered['한화환산액'] = filtered.apply(get_v136_conversion, axis=1)

        # 3. 유형별 요약 테이블
        st.subheader(f"📊 {start_y}년~{end_y}년 지출 요약")
        if not filtered.empty:
            summary = filtered.groupby('유형').agg({
                '실입금액': 'sum', 
                '선급금액': 'sum', 
                '한화환산액': 'sum'
            }).reset_index()
            st.table(summary.style.format({
                '실입금액': '{:,.2f}', '선급금액': '{:,.2f}', '한화환산액': '{:,.0f}'
            }))

        # 4. 발주번호별 정산 및 잔액 (이 부분은 모든 연도의 발주를 대조해야 하므로 p_all 기준)
        st.subheader("🔍 발주별 정산 및 미수금 현황 (전체 기간)")
        pay_agg = p_all.groupby('발주번호').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
        settle_df = pd.merge(o_all, pay_agg, on='발주번호', how='left').fillna(0)
        settle_df['잔액'] = settle_df['발주총액'] - (settle_df['실입금액'] + settle_df['선급금액'])
        settle_df['상태'] = settle_df['마감여부'].apply(lambda x: "✅ 마감" if x == 1 else "⏳ 진행")
        
        st.dataframe(settle_df[['발주번호','상태','거래처명','상품명','발주총액','실입금액','선급금액','잔액','통화']].sort_values('발주번호', ascending=False), use_container_width=True)

        # 5. 상세 내역 수정
        st.subheader("📝 상세 내역 수정")
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
            upsert_supabase_data("payments", edited_p.to_dict(orient='records'))
            st.success("수정사항이 반영되었습니다."); st.rerun()

        # 6. 하단 메트릭
        st.divider()
        m1, m2, m3 = st.columns(3)
        m1.metric(f"선택 범위 총 환산액", f"{filtered['한화환산액'].sum():,.0f} 원")
        m2.metric("선택 범위 USD 합계", f"${filtered[filtered['통화']=='USD']['실입금액'].sum():,.2f}")
        m3.metric("선택 범위 CNY 합계", f"¥{filtered[filtered['통화']=='CNY']['실입금액'].sum():,.2f}")
    else:
        st.info("데이터가 없습니다. 먼저 입금 내역을 등록해 주세요.")

# --- [Tab 3] 거래처 관리 (모든 필드 복구 및 데이터 연동) ---
with tabs[3]:
    st.header("거래처 정보 관리")
    v_orig = get_supabase_data("vendors")
    
    # 1. 신규 등록 폼 (모든 입력 칸 복구)
    with st.form("new_v_form_full", clear_on_submit=True):
        st.subheader("신규 거래처 등록")
        col_v1, col_v2 = st.columns(2)
        vn = col_v1.text_input("거래처명 (필수)")
        vt = col_v2.selectbox("기본 유형", CATEGORIES)
        
        col_v3, col_v4, col_v5 = st.columns(3)
        vb = col_v3.text_input("은행")
        va = col_v4.text_input("계좌번호")
        vh = col_v5.text_input("예금주")
        
        if st.form_submit_button("거래처 등록 저장"):
            if vn:
                upsert_supabase_data("vendors", {
                    "거래처명": vn, 
                    "기본유형": vt, 
                    "은행": vb, 
                    "계좌번호": va, 
                    "예금주": vh
                })
                st.success(f"[{vn}] 등록 완료!"); st.rerun()
            else:
                st.error("거래처명은 필수입니다.")

    st.divider()

    # 2. 기존 목록 수정 및 동기화
    if not v_orig.empty:
        st.subheader("등록된 거래처 목록 (수정 후 저장 시 전체 데이터 연동)")
        # 데이터 에디터에서 모든 컬럼을 편집 가능하게 표시
        ev_v = st.data_editor(v_orig, hide_index=True, use_container_width=True)
        
        if st.button("수정 내용 저장 및 과거 데이터 일괄 동기화"):
            # 이름 변경 시 입금/발주 데이터까지 싹 바꿔주는 v136 핵심 로직
            for i, r in ev_v.iterrows():
                if i < len(v_orig) and v_orig.iloc[i]['거래처명'] != r['거래처명']:
                    old_n = v_orig.iloc[i]['거래처명']
                    # 입금 내역(payments)과 발주 내역(orders)의 거래처명도 함께 변경
                    supabase.table("payments").update({"거래처명": r['거래처명'], "유형": r['기본유형']}).eq("거래처명", old_n).execute()
                    supabase.table("orders").update({"거래처명": r['거래처명'], "유형": r['기본유형']}).eq("거래처명", old_n).execute()
            
            # 거래처 마스터 정보 최종 업데이트
            upsert_supabase_data("vendors", ev_v.to_dict(orient='records'))
            st.success("거래처 정보 및 관련 내역이 모두 업데이트되었습니다.")
            st.rerun()
    else:
        st.info("등록된 거래처가 없습니다.")

# --- [Tab 4] 환율 분석 (차트 꼬임 방지 + 좌우 대칭 레이아웃 완벽 복구) ---
with tabs[4]:
    st.header("📈 환율 데이터 분석 및 관리")
    
    # 1. 환율 데이터 업로드 로직
    def up_ex(u, cur):
        try:
            df_ex = pd.read_csv(u)
            df_ex.columns = [c.strip() for c in df_ex.columns]
            data_list = []
            for _, r in df_ex.iterrows():
                data_list.append({
                    "날짜": smart_date(r['날짜']), 
                    cur.lower(): to_float(r['종가'])
                })
            upsert_supabase_data("exchange_rates", data_list)
        except Exception as e:
            st.error(f"업로드 에러: {e}")

    up1, up2 = st.columns(2)
    with up1:
        u_u = st.file_uploader("USD CSV 업로드", type=['csv'], key="usd_up")
        if u_u and st.button("USD 데이터 동기화"):
            up_ex(u_u, "USD"); st.rerun()
    with up2:
        u_c = st.file_uploader("CNY CSV 업로드", type=['csv'], key="cny_up")
        if u_c and st.button("CNY 데이터 동기화"):
            up_ex(u_c, "CNY"); st.rerun()

    st.divider()

    # 2. 메인 분석 영역 (좌 USD / 우 CNY 대칭 구조)
    ex_db = get_supabase_data("exchange_rates")
    
    if not ex_db.empty:
        # 데이터 전처리
        ex_db['날짜'] = pd.to_datetime(ex_db['날짜'])
        ex_db['연도'] = ex_db['날짜'].dt.year
        ex_db['월'] = ex_db['날짜'].dt.month
        
        # 2025~2026년 비교용 데이터 필터링
        df_target = ex_db[ex_db['연도'].isin([2025, 2026])]

        # 좌우 레이아웃 분할
        col_left, col_right = st.columns(2)

        for i, curr in enumerate(['usd', 'cny']):
            target_col = col_left if i == 0 else col_right
            
            with target_col:
                st.subheader(f"💱 {curr.upper()} 분석")
                
                # [상단] 차트 배치 (정렬 추가로 선 꼬임 방지)
                if curr in ex_db.columns:
                    # 날짜순 정렬을 해야 선이 꼬이지 않습니다.
                    chart_df = ex_db[['날짜', curr]].dropna().sort_values('날짜')
                    
                    fig = go.Figure()
                    fig.add_trace(go.Scatter(
                        x=chart_df['날짜'], 
                        y=chart_df[curr], 
                        name=curr.upper(), 
                        mode='lines',
                        line=dict(color='blue' if curr=='usd' else 'red', width=2)
                    ))
                    fig.update_layout(
                        height=300, 
                        margin=dict(l=0, r=0, t=10, b=0), 
                        showlegend=False,
                        hovermode="x unified"
                    )
                    st.plotly_chart(fig, use_container_width=True)
                
                # [하단] 월별/연도별 비교표 배치
                m_avg = df_target.groupby(['연도', '월'])[curr].mean().reset_index()
                if not m_avg.empty:
                    pivot = m_avg.pivot(index='월', columns='연도', values=curr)
                    # 컬럼명 존재 여부 확인 후 이름 변경
                    pivot.columns = [f"{int(c)}년" for c in pivot.columns]
                    
                    # 2025년과 2026년 데이터가 모두 있을 때만 증감 계산
                    c25, c26 = "2025년", "2026년"
                    if c25 in pivot.columns and c26 in pivot.columns:
                        pivot['증감'] = pivot[c26] - pivot[c25]
                        pivot['%'] = (pivot['증감'] / pivot[c25] * 100)
                    
                    st.write(f"**{curr.upper()} 월별 평균 대조 (2025 vs 2026)**")
                    st.table(pivot.sort_index().style.format("{:,.2f}"))
                else:
                    st.info(f"{curr.upper()} 비교 분석을 위한 2025/2026 데이터가 부족합니다.")

        st.divider()
        
        # 3. 데이터 원본 수정 (하단 배치)
        with st.expander("데이터 관리 및 원본 수정"):
            display_db = ex_db.copy().sort_values('날짜', ascending=False)
            display_db['날짜'] = display_db['날짜'].dt.strftime('%Y-%m-%d')
            # 표시할 컬럼만 선택
            cols = [c for c in ['날짜', 'usd', 'cny'] if c in display_db.columns]
            edited_ex = st.data_editor(display_db[cols], hide_index=True, use_container_width=True)
            
            if st.button("수정 내용 저장"):
                try:
                    upsert_supabase_data("exchange_rates", edited_ex.to_dict(orient='records'))
                    st.success("저장되었습니다."); st.rerun()
                except Exception as e:
                    st.error(f"저장 중 오류 발생: {e}")
    else:
        st.info("데이터가 없습니다. 상단의 업로더를 통해 환율 CSV 파일을 등록해 주세요.")
