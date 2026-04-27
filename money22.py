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

# --- [Tab 1] 발주서 등록 및 관리 (누락된 소급 수정 로직 복구본) ---
with tabs[1]:
    st.header("📦 발주서 등록 및 마감 관리")
    
    v_master = get_supabase_data("vendors")
    o_data = get_supabase_data("orders")
    
    c1, c2 = st.columns([1, 1.8]) 
    
    with c1:
        st.subheader("1. 발주 분석 및 등록")
        o_files = st.file_uploader("이카운트 엑셀 선택", type=['xlsx'], accept_multiple_files=True, key=f"ord_f_{st.session_state.order_up_key}")
        if o_files and st.button("🚀 발주서 일괄 분석 실행", use_container_width=True):
            for f in o_files: 
                success, msg = process_ecount_v136_cloud(f)
                if not success: st.error(msg)
            st.session_state.order_up_key += 1; st.success("분석 완료!"); st.rerun()
        
        st.divider()
        
        with st.form("manual_ord_form", clear_on_submit=True):
            st.write("**수기 발주 입력**")
            m_oid = st.text_input("발주번호 (필수)")
            m_vn = st.selectbox("거래처 선택", ["선택"] + (list(v_master['거래처명']) if not v_master.empty else []))
            
            col_m1, col_m2 = st.columns(2)
            m_amt = col_m1.number_input("발주총액", format="%.2f")
            m_cur = col_m2.selectbox("통화", ["한화", "USD", "CNY"])
            m_item = st.text_input("상품명 (선택)")
            
            if st.form_submit_button("➕ 발주 저장", use_container_width=True):
                if m_oid and m_vn != "선택":
                    v_type = v_master[v_master['거래처명']==m_vn].iloc[0]['기본유형'] if not v_master.empty else "기타"
                    upsert_supabase_data("orders", {
                        "발주번호": m_oid, "발주일": datetime.now().strftime("%Y-%m-%d"), 
                        "거래처명": m_vn, "상품명": m_item or "수기입력",
                        "유형": v_type, "발주총액": m_amt, "통화": m_cur, "마감여부": 0
                    })
                    st.success("저장 완료!"); st.rerun()

        st.divider()
        st.subheader("🗑️ 발주 데이터 삭제")
        order_list = list(o_data['발주번호'].unique()) if not o_data.empty and '발주번호' in o_data.columns else []
        del_oid = st.selectbox("삭제할 발주번호 선택", ["선택"] + order_list, key="del_box")
        if st.button("❌ 선택한 발주 삭제 실행", type="secondary", use_container_width=True):
            if del_oid != "선택":
                # 발주서 삭제 및 입금내역 연결 해제
                supabase.table("orders").delete().eq("발주번호", del_oid).execute()
                supabase.table("payments").update({"발주번호": None}).eq("발주번호", del_oid).execute()
                st.error(f"[{del_oid}] 삭제됨"); st.rerun()

    with c2:
        st.subheader("2. 발주 목록 및 소급 수정")
        
        if not o_data.empty and '발주번호' in o_data.columns and '발주일' in o_data.columns:
            ev_o = st.data_editor(
                o_data.sort_values('발주일', ascending=False), 
                hide_index=True, 
                use_container_width=True,
                key=f"editor_{len(o_data)}", 
                column_config={
                    "마감여부": st.column_config.CheckboxColumn("마감"),
                    "발주총액": st.column_config.NumberColumn("총액", format="%.2f"),
                    "거래처명": st.column_config.SelectboxColumn("거래처명", options=list(v_master['거래처명']) if not v_master.empty else []),
                    "유형": st.column_config.SelectboxColumn("유형", options=CATEGORIES)
                },
                disabled=["발주번호"] 
            )
            
            # --- [복구된 핵심 기능] 저장 시 입금 내역 소급 수정 ---
            if st.button("💾 수정 내용 저장 및 내역 동기화", use_container_width=True):
                # 1. Orders 테이블 업데이트
                upsert_supabase_data("orders", ev_o.to_dict(orient='records'))
                
                # 2. Payments 테이블 동기화 (발주번호가 같은 모든 입금 내역의 업체/상품명 수정)
                for _, r in ev_o.iterrows():
                    supabase.table("payments").update({
                        "거래처명": r['거래처명'], 
                        "유형": r['유형'], 
                        "상품명": r['상품명']
                    }).eq("발주번호", r['발주번호']).execute()
                
                st.success("✅ 발주 수정 및 관련 입금내역 동기화 완료!"); st.rerun()
        else:
            st.info("💡 등록된 발주 내역이 없습니다. 왼쪽 메뉴를 이용해 주세요.")
            
# --- [Tab 2] 상세 내역 및 통합 정산 (비율 조정 및 가시성 최적화 버전) ---
with tabs[2]:
    st.header("📋 상세 내역 및 통합 정산")
    
    # 데이터 로드
    p_all = get_supabase_data("payments")
    o_all = get_supabase_data("orders")
    ex_rates = get_supabase_data("exchange_rates")
    
    if not p_all.empty:
        p_all['dt'] = pd.to_datetime(p_all['입금일'])
        
        # 1. 필터 섹션 (너비 조절을 위해 columns 배치)
        f_c1, f_c2, f_c3, f_c4 = st.columns([1, 1, 1, 1.5])
        years = sorted(p_all['dt'].dt.year.unique())
        
        # 시작/종료 연도 선택
        start_y = f_c1.selectbox("시작 연도", years, index=0)
        end_y = f_c1.selectbox("종료 연도", years, index=len(years)-1)
        
        target_m = f_c2.selectbox("조회 월", ["전체"] + list(range(1, 13)))
        filter_cat = f_c3.selectbox("유형 필터", ["전체"] + CATEGORIES)
        search_key = f_c4.text_input("🔍 업체/상품 검색 (검색어 입력)")
        
        # 필터링 적용
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

        # 3. 유형별 요약 테이블 (가운데 배치를 위해 여백 컬럼 활용)
        st.markdown(f"### 📊 {start_y}년 ~ {end_y}년 지출 요약")
        if not filtered.empty:
            summary = filtered.groupby('유형').agg({
                '실입금액': 'sum', 
                '선급금액': 'sum', 
                '한화환산액': 'sum'
            }).reset_index()
            
            sum_c1, sum_c2, sum_c3 = st.columns([0.1, 0.8, 0.1])
            with sum_c2:
                st.table(summary.style.format({
                    '실입금액': '{:,.2f}', '선급금액': '{:,.2f}', '한화환산액': '{:,.0f}'
                }))

        st.divider()

        # 4. 발주번호별 정산 및 잔액 (가시성 최적화 적용)
        st.subheader("🔍 발주별 정산 및 미수금 현황 (전체 기간)")
        pay_agg = p_all.groupby('발주번호').agg({'실입금액':'sum', '선급금액':'sum'}).reset_index()
        settle_df = pd.merge(o_all, pay_agg, on='발주번호', how='left').fillna(0)
        settle_df['잔액'] = settle_df['발주총액'] - (settle_df['실입금액'] + settle_df['선급금액'])
        settle_df['상태'] = settle_df['마감여부'].apply(lambda x: "✅ 마감" if x == 1 else "⏳ 진행")
        
        # 중요: 표 너비 비율 조정 (column_config)
        st.data_editor(
            settle_df[['발주번호','상태','거래처명','상품명','발주총액','실입금액','잔액','통화']].sort_values('발주번호', ascending=False),
            hide_index=True,
            use_container_width=True,
            key="settle_editor",
            column_config={
                "발주번호": st.column_config.TextColumn("발주번호", width="small"),
                "상태": st.column_config.TextColumn("상태", width="small"),
                "거래처명": st.column_config.TextColumn("거래처명", width="medium"),
                "상품명": st.column_config.TextColumn("상품명", width="large"), # 상품명을 넓게
                "발주총액": st.column_config.NumberColumn("발주총액", format="%.2f", width="medium"),
                "잔액": st.column_config.NumberColumn("미수잔액", format="%.2f", width="medium"),
                "통화": st.column_config.TextColumn("통화", width="small")
            }
        )

        st.divider()

        # 5. 상세 내역 수정 (가시성 최적화 적용)
        st.subheader("📝 상세 내역 수정 및 관리")
        edit_cols = ['id', '유형', '발주번호', '거래처명', '상품명', '입금일', '통화', '실입금액', '선급금액', '한화환산액', '메모']
        edited_p = st.data_editor(
            filtered[edit_cols].sort_values('입금일', ascending=False), 
            hide_index=True, 
            use_container_width=True,
            key="detail_editor",
            column_config={
                "id": st.column_config.TextColumn("ID", width="small"),
                "유형": st.column_config.SelectboxColumn("유형", options=CATEGORIES, width="small"),
                "거래처명": st.column_config.TextColumn("거래처명", width="medium"),
                "상품명": st.column_config.TextColumn("상품명", width="large"),
                "입금일": st.column_config.DateColumn("입금일", width="medium"),
                "한화환산액": st.column_config.NumberColumn("한화환산액(참고)", format="%d", width="medium"),
                "실입금액": st.column_config.NumberColumn("입금액", format="%.2f", width="medium"),
                "메모": st.column_config.TextColumn("비고/메모", width="large")
            }
        )
        
        if st.button("💾 수정 내용 클라우드 저장"):
            upsert_supabase_data("payments", edited_p.to_dict(orient='records'))
            st.success("수정사항이 클라우드에 안전하게 반영되었습니다."); st.rerun()

        # 6. 하단 메트릭 요약
        st.divider()
        m1, m2, m3 = st.columns(3)
        m1.metric(f"선택 범위 총 환산액", f"{filtered['한화환산액'].sum():,.0f} 원")
        m2.metric("선택 범위 USD 합계", f"${filtered[filtered['통화']=='USD']['실입금액'].sum():,.2f}")
        m3.metric("선택 범위 CNY 합계", f"¥{filtered[filtered['통화']=='CNY']['실입금액'].sum():,.2f}")
    else:
        st.info("데이터가 없습니다. 먼저 입금 내역을 등록해 주세요.")

# --- [Tab 3] 거래처 관리 (가시성 최적화 및 비율 조정 버전) ---
with tabs[3]:
    st.header("🏢 거래처 정보 관리")
    v_orig = get_supabase_data("vendors")
    
    # 상단 입력부: 수기 입력과 CSV 업로드를 적절한 비율로 배치
    col_v_in, col_v_csv = st.columns([1.5, 1])
    
    with col_v_in:
        st.subheader("1. 신규 거래처 수기 등록")
        with st.form("new_v_form_full", clear_on_submit=True):
            v_c1, v_c2 = st.columns([2, 1]) # 거래처명을 더 넓게
            vn = v_c1.text_input("거래처명 (필수)")
            vt = v_c2.selectbox("기본 유형", CATEGORIES)
            
            v_c3, v_c4, v_c5 = st.columns([1, 2, 1]) # 계좌번호를 더 넓게
            vb = v_c3.text_input("은행")
            va = v_c4.text_input("계좌번호")
            vh = v_c5.text_input("예금주")
            
            if st.form_submit_button("➕ 거래처 정보 저장"):
                if vn:
                    upsert_supabase_data("vendors", {
                        "거래처명": vn, "기본유형": vt, "은행": vb, "계좌번호": va, "예금주": vh
                    })
                    st.success(f"✅ [{vn}] 등록 완료!"); st.rerun()
                else:
                    st.error("⚠️ 거래처명은 필수 입력 항목입니다.")

    with col_v_csv:
        st.subheader("2. CSV 일괄 등록")
        # 거래처 전용 CSV 양식 생성
        v_template = pd.DataFrame(columns=["거래처명", "기본유형", "은행", "계좌번호", "예금주"])
        st.download_button(
            "📥 등록 양식(CSV) 다운로드", 
            v_template.to_csv(index=False).encode('utf-8-sig'), 
            "vendor_template.csv",
            use_container_width=True
        )
        
        up_vendor = st.file_uploader("파일 선택", type=['csv'], key="v_up_file")
        if up_vendor and st.button("🚀 일괄 저장 실행", use_container_width=True):
            try:
                df_v_up = pd.read_csv(up_vendor)
                df_v_up.columns = [str(c).strip().replace('\ufeff', '') for c in df_v_up.columns]
                
                v_list = []
                for _, r in df_v_up.iterrows():
                    if to_str(r.get('거래처명')):
                        v_list.append({
                            "거래처명": to_str(r.get('거래처명')),
                            "기본유형": to_str(r.get('기본유형')) or "기타",
                            "은행": to_str(r.get('은행')),
                            "계좌번호": to_str(r.get('계좌번호')),
                            "예금주": to_str(r.get('예금주'))
                        })
                
                if v_list:
                    upsert_supabase_data("vendors", v_list)
                    st.success(f"✨ {len(v_list)}건의 거래처 등록 완료!"); st.rerun()
            except Exception as e:
                st.error(f"❌ 업로드 중 오류 발생: {e}")

    st.divider()

    # 2. 기존 목록 수정 및 가시성 최적화
    if not v_orig.empty:
        st.subheader("📋 등록된 거래처 목록")
        st.info("💡 거래처명을 수정하면 과거의 입금/발주 내역의 이름도 모두 함께 변경됩니다.")
        
        # 가시성 핵심: 컬럼별 너비 지정
        ev_v = st.data_editor(
            v_orig.sort_values('거래처명'), 
            hide_index=True, 
            use_container_width=True,
            key="vendor_editor",
            column_config={
                "거래처명": st.column_config.TextColumn("거래처명", width="large", help="수정 시 모든 내역에 소급 적용"),
                "기본유형": st.column_config.SelectboxColumn("기본 유형", options=CATEGORIES, width="small"),
                "은행": st.column_config.TextColumn("은행", width="small"),
                "계좌번호": st.column_config.TextColumn("계좌번호", width="medium"),
                "예금주": st.column_config.TextColumn("예금주", width="small")
            }
        )
        
        col_save1, col_save2 = st.columns([1, 3])
        if col_save1.button("💾 변경사항 동기화 저장", use_container_width=True):
            # 동기화 로직 (이름 변경 시 연동 데이터 소급 수정)
            for i, r in ev_v.iterrows():
                # 기존 데이터와 비교하여 이름이 바뀌었는지 확인
                target_id = r.get('id')
                old_row = v_orig[v_orig['id'] == target_id]
                
                if not old_row.empty and old_row.iloc[0]['거래처명'] != r['거래처명']:
                    old_n = old_row.iloc[0]['거래처명']
                    # 연동된 payments, orders 테이블 일괄 업데이트
                    supabase.table("payments").update({"거래처명": r['거래처명'], "유형": r['기본유형']}).eq("거래처명", old_n).execute()
                    supabase.table("orders").update({"거래처명": r['거래처명'], "유형": r['기본유형']}).eq("거래처명", old_n).execute()
            
            # 최종 마스터 정보 업데이트
            upsert_supabase_data("vendors", ev_v.to_dict(orient='records'))
            st.success("✅ 거래처 정보 및 과거 내역 동기화가 완료되었습니다."); st.rerun()
    else:
        st.info("📢 등록된 거래처 정보가 없습니다. 신규 거래처를 등록해 주세요.")

# --- [Tab 4] 환율 분석 (가시성 최적화 및 좌우 대칭 레이아웃) ---
with tabs[4]:
    st.header("📈 환율 데이터 분석 및 관리")
    
    # 1. 환율 데이터 업로드 섹션
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

    up_c1, up_c2 = st.columns(2)
    with up_c1:
        u_u = st.file_uploader("USD CSV 업로드", type=['csv'], key="usd_up")
        if u_u and st.button("USD 데이터 동기화", use_container_width=True):
            up_ex(u_u, "USD"); st.rerun()
    with up_c2:
        u_c = st.file_uploader("CNY CSV 업로드", type=['csv'], key="cny_up")
        if u_c and st.button("CNY 데이터 동기화", use_container_width=True):
            up_ex(u_c, "CNY"); st.rerun()

    st.divider()

    # 2. 메인 분석 영역 (좌 USD / 우 CNY 대칭 구조)
    ex_db = get_supabase_data("exchange_rates")
    
    if not ex_db.empty:
        ex_db['날짜'] = pd.to_datetime(ex_db['날짜'])
        ex_db['연도'] = ex_db['날짜'].dt.year
        ex_db['월'] = ex_db['날짜'].dt.month
        df_target = ex_db[ex_db['연도'].isin([2025, 2026])]

        # 가시성을 위해 좌우 여백을 살짝 둔 2컬럼 레이아웃
        main_l, main_r = st.columns([1, 1], gap="large")

        for i, curr in enumerate(['usd', 'cny']):
            target_col = main_l if i == 0 else main_r
            
            with target_col:
                st.subheader(f"💱 {curr.upper()} 분석 리포트")
                
                # [상단] 차트 배치 (정렬로 꼬임 방지)
                chart_df = ex_db[['날짜', curr]].dropna().sort_values('날짜')
                if not chart_df.empty:
                    fig = go.Figure()
                    fig.add_trace(go.Scatter(
                        x=chart_df['날짜'], 
                        y=chart_df[curr], 
                        mode='lines',
                        line=dict(color='blue' if curr=='usd' else 'red', width=2)
                    ))
                    fig.update_layout(
                        height=250, 
                        margin=dict(l=0, r=0, t=10, b=0), 
                        showlegend=False,
                        hovermode="x unified"
                    )
                    st.plotly_chart(fig, use_container_width=True)
                
                # [하단] 월별 평균 비교표 (가시성 최적화)
                m_avg = df_target.groupby(['연도', '월'])[curr].mean().reset_index()
                if not m_avg.empty:
                    pivot = m_avg.pivot(index='월', columns='연도', values=curr)
                    pivot.columns = [f"{int(c)}년" for c in pivot.columns]
                    
                    # 증감 로직
                    c25, c26 = "2025년", "2026년"
                    if c25 in pivot.columns and c26 in pivot.columns:
                        pivot['증감'] = pivot[c26] - pivot[c25]
                        pivot['%'] = (pivot['증감'] / pivot[c25] * 100)
                    
                    st.write(f"**{curr.upper()} 월별 평균 대조**")
                    # 표가 너무 가로로 찢어지지 않도록 설정
                    st.dataframe(
                        pivot.sort_index().style.format("{:,.2f}"),
                        use_container_width=True,
                        column_config={
                            "월": st.column_config.TextColumn("월", width="small"),
                            "2025년": st.column_config.NumberColumn("25년", width="small"),
                            "2026년": st.column_config.NumberColumn("26년", width="small"),
                        }
                    )
                else:
                    st.info(f"{curr.upper()} 데이터 부족")

        st.divider()
        
        # 3. 데이터 원본 관리
        with st.expander("🛠️ 환율 데이터 원본 관리 및 수정"):
            sub_c1, sub_c2, sub_c3 = st.columns([0.1, 0.8, 0.1])
            with sub_c2:
                display_db = ex_db.copy().sort_values('날짜', ascending=False)
                display_db['날짜'] = display_db['날짜'].dt.strftime('%Y-%m-%d')
                cols = [c for c in ['날짜', 'usd', 'cny'] if c in display_db.columns]
                
                edited_ex = st.data_editor(
                    display_db[cols], 
                    hide_index=True, 
                    use_container_width=True,
                    column_config={
                        "날짜": st.column_config.TextColumn("날짜", width="medium"),
                        "usd": st.column_config.NumberColumn("USD", format="%.2f", width="small"),
                        "cny": st.column_config.NumberColumn("CNY", format="%.2f", width="small")
                    }
                )
                
                if st.button("💾 수정 내용 저장", use_container_width=True):
                    try:
                        upsert_supabase_data("exchange_rates", edited_ex.to_dict(orient='records'))
                        st.success("저장 완료!"); st.rerun()
                    except Exception as e:
                        st.error(f"저장 실패: {e}")
    else:
        st.info("환율 데이터를 업로드해 주세요.")
