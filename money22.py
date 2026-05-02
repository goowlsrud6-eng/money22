import streamlit as st
import pandas as pd
import numpy as np
import re
import os
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
import plotly.graph_objects as go
from supabase import create_client, Client

# ==============================================================================
# 1. 초기 설정 및 Supabase 연결 (v136_Cloud_Full)
# ==============================================================================
st.set_page_config(page_title="💳 입금·발주 관리", layout="wide")

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
        if not data:
            return True

        if isinstance(data, dict):
            data = [data]

        data = [
            {k: ("" if v is None else v) for k, v in row.items()}
            for row in data
        ]

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

def to_money(val):
    try:
        if val is None or pd.isna(val) or str(val).strip() == "":
            return 0

        s = str(val).strip()
        s = (
            s.replace(",", "")
             .replace("₩", "")
             .replace("원", "")
             .replace(" ", "")
        )

        return int(Decimal(s))
    except:
        return 0

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

# --- [Tab 0] 입금 내역 등록 (자동 연동 강화 버전) ---
with tabs[0]:
    st.header("입금 내역 등록 및 관리")

    v_master = get_supabase_data("vendors")
    o_data = get_supabase_data("orders")
    o_active = o_data[o_data['마감여부'] == 0] if not o_data.empty else pd.DataFrame()

    col_input, col_excel = st.columns([1.5, 1])

    # -------------------------------
    # 🔵 수기 입력 (자동 연동 핵심)
    # -------------------------------
    with col_input:
        st.subheader("1. 수기 직접 입력")

        with st.form("manual_pay_form", clear_on_submit=True):

            # 발주번호 선택
            p_oid = st.selectbox(
                "발주번호 연동",
                ["없음"] + (list(o_active['발주번호']) if not o_active.empty else [])
            )

            # 🔥 발주 자동 매칭
            if p_oid != "없음":
                match = o_active[o_active['발주번호'] == p_oid].iloc[0]
                auto_vn = match['거래처명']
                auto_type = match['유형']
                auto_prod = match['상품명']
                auto_cur = match['통화']
            else:
                auto_vn = "선택"
                auto_type = "선택"
                auto_prod = ""
                auto_cur = "한화"

            # 입력 필드
            p_date = st.date_input("입금일자", value=datetime.now())

            vn_list = ["선택"] + (list(v_master['거래처명'].unique()) if not v_master.empty else [])

            p_vn = st.selectbox(
                "거래처",
                vn_list,
                index=vn_list.index(auto_vn) if auto_vn in vn_list else 0
            )

            p_ct = st.selectbox(
                "유형 분류",
                ["선택"] + CATEGORIES,
                index=(["선택"] + CATEGORIES).index(auto_type) if auto_type in CATEGORIES else 0
            )

            p_pr = st.text_input("상품명", value=auto_prod)

            r3c1, r3c2, r3c3 = st.columns(3)

            p_dep = r3c1.number_input("실입금액", step=1)
            p_pre = r3c2.number_input("선급금액", step=1)

            cur_list = ["한화", "USD", "CNY"]
            p_cur = r3c3.selectbox(
                "거래통화",
                cur_list,
                index=cur_list.index(auto_cur) if auto_cur in cur_list else 0
            )

            p_memo = st.text_input("비고 (송금 사유 등)")

            # 저장
            if st.form_submit_button("입금 내역 저장"):

                if p_vn == "선택":
                    st.error("거래처를 선택하세요.")

                elif p_ct == "선택":
                    st.error("유형을 선택하세요.")

                elif p_dep == 0 and p_pre == 0:
                    st.error("금액을 입력하세요.")

                else:
                    vi = v_master[v_master['거래처명'] == p_vn].iloc[0]

                    upsert_supabase_data("payments", {
                        "id": get_multiple_available_ids(1)[0],
                        "발주번호": p_oid if p_oid != "없음" else None,
                        "입금일": p_date.strftime("%Y-%m-%d"),
                        "유형": p_ct,
                        "거래처명": p_vn,
                        "상품명": p_pr,
                        "통화": p_cur,
                        "실입금액": to_money(p_dep),
                        "선급금액": to_money(p_pre),
                        "메모": p_memo,
                        "은행": vi['은행'],
                        "계좌번호": vi['계좌번호'],
                        "예금주": vi['예금주']
                    })

                    st.success("저장 완료")
                    st.rerun()

    # -------------------------------
    # 🔵 CSV 업로드
    # -------------------------------
    with col_excel:
        st.subheader("2. CSV 일괄 업로드 (v136 지능형 매칭)")

        csv_template = pd.DataFrame(columns=[
            "발주번호", "거래처", "유형", "상품명",
            "입금일", "실입금액", "선급금액", "송금사유"
        ])

        st.download_button(
            "양식 다운로드",
            csv_template.to_csv(index=False).encode('utf-8-sig'),
            "payment_template.csv"
        )

        up_pay = st.file_uploader("CSV 선택", type=['csv'], key=f"pay_up_{st.session_state.pay_up_key}")

        if up_pay and st.button("파일 일괄 저장 실행"):

            df_up = pd.read_csv(up_pay, dtype=str, keep_default_na=False)
            df_up.columns = [str(c).strip().replace('\ufeff', '') for c in df_up.columns]

            ids = get_multiple_available_ids(len(df_up))
            up_list = []

            for i, r in df_up.iterrows():

                oid_v = to_str(r.get('발주번호'))
                vn_v = to_str(r.get('거래처'))

                match_o = o_data[o_data['발주번호'] == oid_v].iloc[0] if oid_v and not o_data[o_data['발주번호'] == oid_v].empty else None

                vn_f = match_o['거래처명'] if match_o is not None else vn_v

                vi = v_master[v_master['거래처명'].str.lower() == vn_f.lower()].iloc[0] \
                    if not v_master.empty and vn_f and not v_master[v_master['거래처명'].str.lower() == vn_f.lower()].empty else None

                up_list.append({
                    "id": ids[i],
                    "발주번호": oid_v or None,
                    "입금일": smart_date(r.get('입금일')),
                    "유형": match_o['유형'] if match_o is not None else (to_str(r.get('유형')) or "사입"),
                    "거래처명": vn_f,
                    "상품명": match_o['상품명'] if match_o is not None else to_str(r.get('상품명')),
                    "통화": match_o['통화'] if match_o is not None else "한화",
                    "실입금액": to_money(r.get('실입금액')),
                    "선급금액": to_money(r.get('선급금액')),
                    "메모": to_str(r.get('송금사유')),
                    "은행": vi['은행'] if vi is not None else "",
                    "계좌번호": vi['계좌번호'] if vi is not None else "",
                    "예금주": vi['예금주'] if vi is not None else ""
                })

            if upsert_supabase_data("payments", up_list):
                st.session_state.pay_up_key += 1
                st.rerun()

# --- [Tab 1] 발주서 등록 및 관리 ---
with tabs[1]:
    st.header("📦 발주서 등록 및 관리")

    if 'order_notice' in st.session_state:
        notice_type = st.session_state.order_notice.get("type", "success")
        notice_msg = st.session_state.order_notice.get("msg", "")

        if notice_type == "success":
            st.success(notice_msg)
        elif notice_type == "warning":
            st.warning(notice_msg)
        elif notice_type == "error":
            st.error(notice_msg)
        else:
            st.info(notice_msg)

        del st.session_state.order_notice

    v_master = get_supabase_data("vendors")
    o_data = get_supabase_data("orders")

    c1, c2 = st.columns([1, 1.8])

    with c1:
        st.subheader("발주 등록")

        o_files = st.file_uploader(
            "이카운트 엑셀",
            type=['xlsx'],
            accept_multiple_files=True,
            key=f"ord_up_{st.session_state.order_up_key}"
        )

        if o_files and st.button("🚀 분석 실행", use_container_width=True):
            success_count = 0
            fail_msgs = []

            for f in o_files:
                success, msg = process_ecount_v136_cloud(f)
                if success:
                    success_count += 1
                else:
                    fail_msgs.append(f"[{f.name}] {msg}")

            st.session_state.order_up_key += 1

            if fail_msgs:
                st.session_state.order_notice = {
                    "type": "warning",
                    "msg": f"분석 완료: 성공 {success_count}건, 실패 {len(fail_msgs)}건\n" + "\n".join(fail_msgs)
                }
            else:
                st.session_state.order_notice = {
                    "type": "success",
                    "msg": f"발주서 분석 및 등록 완료: {success_count}건"
                }

            st.rerun()

        st.divider()

        with st.form("manual_order_form", clear_on_submit=True):
            m_oid = st.text_input("발주번호")
            m_step = st.text_input("차수")

            vn_list = ["선택"] + list(v_master['거래처명'].unique()) if not v_master.empty else ["선택"]
            m_vn = st.selectbox("거래처", vn_list)

            m_prod = st.text_input("상품명")

            col_m1, col_m2 = st.columns(2)
            m_amt = col_m1.number_input("총액", min_value=0.0, step=0.01, format="%.2f")
            m_cur = col_m2.selectbox("통화", ["한화", "USD", "CNY"])

            if st.form_submit_button("저장"):
                if m_oid and m_vn != "선택":
                    v_type = (
                        v_master[v_master['거래처명'] == m_vn].iloc[0]['기본유형']
                        if not v_master.empty
                        else "기타"
                    )

                    new_order = {
                        "발주번호": str(m_oid).strip(),
                        "발주일": datetime.now().strftime("%Y-%m-%d"),
                        "발주차수": str(m_step).strip(),
                        "거래처명": str(m_vn).strip(),
                        "상품명": str(m_prod).strip(),
                        "유형": v_type,
                        "발주총액": float(m_amt),
                        "통화": str(m_cur),
                        "마감여부": 0,
                        "삭제여부": 0
                    }

                    if upsert_supabase_data("orders", new_order):
                        st.session_state.order_notice = {
                            "type": "success",
                            "msg": f"발주서 [{m_oid}] 등록 완료"
                        }
                        st.rerun()
                else:
                    st.error("발주번호와 거래처를 입력하세요.")

    with c2:
        st.subheader("발주 목록")

        if not o_data.empty:
            show_deleted = st.checkbox("삭제된 발주 보기")
            show_closed = st.checkbox("마감된 발주 포함")

            order_search = st.text_input(
                "🔍 발주 검색",
                placeholder="발주번호, 차수, 거래처명, 상품명, 유형, 통화 검색"
            )

            disp_o = o_data.copy()

            if show_deleted:
                disp_o = disp_o[disp_o['삭제여부'] == 1]
            else:
                disp_o = disp_o[disp_o['삭제여부'] == 0]

            if not show_closed:
                disp_o = disp_o[disp_o['마감여부'] == 0]

            if order_search:
                search_cols = [
                    c for c in ['발주번호', '발주차수', '거래처명', '상품명', '유형', '통화']
                    if c in disp_o.columns
                ]

                search_mask = pd.Series(False, index=disp_o.index)

                for col in search_cols:
                    search_mask = search_mask | disp_o[col].astype(str).str.contains(
                        order_search,
                        case=False,
                        na=False
                    )

                disp_o = disp_o[search_mask]

            st.caption(f"조회 결과: {len(disp_o)}건")

            disp_o['상태'] = disp_o.apply(
                lambda r: "🗑️" if r['삭제여부'] == 1 else ("🔴" if r['마감여부'] == 1 else "🟢"),
                axis=1
            )

            disp_o['삭제'] = False
            disp_o = disp_o.drop(columns=['삭제여부'], errors='ignore')
            disp_o = disp_o.sort_values(by=["마감여부", "발주일"], ascending=[True, False])

            ev_o = st.data_editor(
                disp_o,
                hide_index=True,
                use_container_width=True,
                height=min(700, 60 + len(disp_o) * 35),
                key="editor_orders",
                column_config={
                    "삭제": st.column_config.CheckboxColumn("삭제"),
                    "상태": st.column_config.TextColumn("상태", width="small"),
                    "마감여부": st.column_config.CheckboxColumn("마감"),
                    "발주총액": st.column_config.NumberColumn("총액", format="%,.2f"),
                    "발주차수": st.column_config.TextColumn("차수")
                },
                disabled=["상태", "발주번호", "발주일", "유형"]
            )

            col_btn1, col_btn2, col_btn3 = st.columns(3)

            with col_btn1:
                if st.button("💾 저장", use_container_width=True):
                    final_save = ev_o.drop(columns=['상태', '삭제'], errors='ignore')
                    clean_data = final_save.fillna("").to_dict(orient='records')

                    if upsert_supabase_data("orders", clean_data):
                        for _, r in ev_o.iterrows():
                            sync_payload = {
                                "거래처명": str(r['거래처명']).strip(),
                                "상품명": str(r['상품명']).strip(),
                                "유형": str(r['유형']).strip()
                            }

                            supabase.table("payments") \
                                .update(sync_payload) \
                                .eq("발주번호", str(r['발주번호'])) \
                                .execute()

                        st.session_state.order_notice = {
                            "type": "success",
                            "msg": f"발주 목록 변경사항 저장 완료: {len(clean_data)}건"
                        }
                        st.rerun()

            with col_btn2:
                if st.button("🗑️ 삭제", use_container_width=True):
                    del_list = ev_o[ev_o['삭제'] == True]

                    if del_list.empty:
                        st.info("삭제할 발주를 선택하세요.")
                    else:
                        for oid in del_list['발주번호']:
                            supabase.table("orders") \
                                .update({"삭제여부": 1}) \
                                .eq("발주번호", oid) \
                                .execute()

                        st.session_state.order_notice = {
                            "type": "success",
                            "msg": f"선택한 발주 삭제 완료: {len(del_list)}건"
                        }
                        st.rerun()

            with col_btn3:
                if show_deleted:
                    if st.button("♻️ 복구", use_container_width=True):
                        restore_list = ev_o[ev_o['삭제'] == True]

                        if restore_list.empty:
                            st.info("복구할 발주를 선택하세요.")
                        else:
                            for oid in restore_list['발주번호']:
                                supabase.table("orders") \
                                    .update({"삭제여부": 0}) \
                                    .eq("발주번호", oid) \
                                    .execute()

                            st.session_state.order_notice = {
                                "type": "success",
                                "msg": f"선택한 발주 복구 완료: {len(restore_list)}건"
                            }
                            st.rerun()

        else:
            st.info("내역 없음")


# --- [Tab 2] 상세 내역 및 통합 정산 ---
with tabs[2]:
    st.header("📋 상세 내역 및 통합 정산")

    p_all = get_supabase_data("payments")
    o_all = get_supabase_data("orders")
    ex_rates = get_supabase_data("exchange_rates")

    if not p_all.empty:
        if '삭제' not in p_all.columns:
            p_all['삭제'] = False

        p_all['삭제'] = p_all['삭제'].fillna(False).astype(bool)

        if not o_all.empty:
            ref_dict = o_all.set_index('발주번호')[['거래처명', '상품명', '유형', '발주차수']].to_dict('index')

            def fill_info(row):
                oid = row.get('발주번호')
                if oid in ref_dict:
                    if not to_str(row.get('거래처명')):
                        row['거래처명'] = ref_dict[oid]['거래처명']
                    if not to_str(row.get('상품명')):
                        row['상품명'] = ref_dict[oid]['상품명']
                    if not to_str(row.get('유형')):
                        row['유형'] = ref_dict[oid]['유형']
                    row['발주차수'] = ref_dict[oid].get('발주차수', '-')
                return row

            p_all = p_all.apply(fill_info, axis=1)

        if '메모' in p_all.columns and '송금사유' not in p_all.columns:
            p_all['송금사유'] = p_all['메모']

        p_all['dt'] = pd.to_datetime(p_all['입금일'], errors='coerce')
        p_all = p_all.dropna(subset=['dt'])

        st.subheader("🔎 필터")
        f_left, f_right = st.columns([1.2, 1])

        with f_left:
            f1, f2 = st.columns(2)
            f3, f4 = st.columns(2)

            years = sorted(p_all['dt'].dt.year.unique())

            start_y = f1.selectbox("시작 연도", years, index=0)
            end_y = f2.selectbox("종료 연도", years, index=len(years) - 1)
            start_m = f3.selectbox("시작 월", list(range(1, 13)), index=0)
            end_m = f4.selectbox("종료 월", list(range(1, 13)), index=11)

            filter_cat = st.selectbox("유형 필터", ["전체 유형"] + CATEGORIES)
            search_vendor = st.text_input("🔍 업체 검색")
            search_product = st.text_input("🔍 상품 검색")
            search_order = st.text_input("🔍 발주차수 검색")

        start_date = pd.to_datetime(f"{start_y}-{start_m:02d}-01")
        end_date = pd.to_datetime(f"{end_y}-{end_m:02d}-01") + pd.offsets.MonthEnd(1)

        filtered = p_all[
            (p_all['dt'] >= start_date) &
            (p_all['dt'] <= end_date)
        ].copy()

        if filter_cat != "전체 유형":
            filtered = filtered[filtered['유형'] == filter_cat]
        if search_vendor:
            filtered = filtered[filtered['거래처명'].str.contains(search_vendor, case=False, na=False)]
        if search_product:
            filtered = filtered[filtered['상품명'].str.contains(search_product, case=False, na=False)]
        if search_order:
            filtered = filtered[filtered['발주차수'].astype(str).str.contains(search_order, case=False, na=False)]

        # 계산용 데이터: 삭제된 입금은 항상 제외
        calc_filtered = filtered[filtered['삭제'] != True].copy()
        calc_p_all = p_all[p_all['삭제'] != True].copy()

        if not ex_rates.empty:
            ex_rates['ym'] = pd.to_datetime(ex_rates['날짜'], errors='coerce').dt.strftime('%Y-%m')

        def get_conv(row):
            try:
                amount = float(row.get('실입금액', 0))
                currency = str(row.get('통화', '한화'))

                if currency == '한화':
                    return amount

                ym = str(row.get('입금일'))[:7]

                if not ex_rates.empty:
                    rate_df = ex_rates[ex_rates['ym'] == ym]
                    curr = currency.lower()

                    if curr in rate_df.columns:
                        avg = rate_df[curr].mean()
                        if pd.notna(avg) and avg > 0:
                            return amount * float(avg)

                rates = {'USD': 1350.0, 'CNY': 190.0}
                return amount * rates.get(currency, 0)

            except:
                return 0

        filtered['한화환산액'] = filtered.apply(get_conv, axis=1)
        calc_filtered['한화환산액'] = calc_filtered.apply(get_conv, axis=1)

        with f_right:
            st.subheader("📊 필터 요약")

            order_base = o_all.copy() if not o_all.empty else pd.DataFrame(columns=['유형', '발주총액'])

            if not order_base.empty:
                if filter_cat != "전체 유형":
                    order_base = order_base[order_base['유형'] == filter_cat]
                if search_vendor:
                    order_base = order_base[order_base['거래처명'].astype(str).str.contains(search_vendor, case=False, na=False)]
                if search_product:
                    order_base = order_base[order_base['상품명'].astype(str).str.contains(search_product, case=False, na=False)]
                if search_order:
                    order_base = order_base[order_base['발주차수'].astype(str).str.contains(search_order, case=False, na=False)]

                order_base['발주총액'] = pd.to_numeric(order_base['발주총액'], errors='coerce').fillna(0)
                order_sum = order_base.groupby('유형')['발주총액'].sum().reset_index()
            else:
                order_sum = pd.DataFrame(columns=['유형', '발주총액'])

            summary = calc_filtered.groupby('유형').agg({
                '실입금액': 'sum',
                '선급금액': 'sum',
                '한화환산액': 'sum'
            }).reset_index()

            summary = pd.merge(summary, order_sum, on='유형', how='left').fillna(0)
            summary['총지급액'] = summary['실입금액']
            summary['선급금액'] = summary['선급금액']

            summary = summary.rename(columns={
                '발주총액': '총발주액',
                '한화환산액': '한환산액'
            })[['유형', '총발주액', '총지급액', '선급금액', '한환산액']]

            st.dataframe(
                summary.style.format('{:,.0f}', subset=['총발주액', '총지급액', '선급금액', '한환산액']),
                hide_index=True,
                use_container_width=True,
                height=220
            )

        st.divider()

        st.subheader("🔍 발주별 정산 및 미수금 현황")

        p_agg = calc_p_all.groupby('발주번호').agg({
            '실입금액': 'sum',
            '선급금액': 'sum'
        }).reset_index()

        s_df = pd.merge(o_all, p_agg, on='발주번호', how='left').fillna(0)
        s_df['미수잔액'] = s_df['발주총액'] - (s_df['실입금액'] + s_df['선급금액'])
        s_df['진행상태'] = s_df['마감여부'].apply(lambda x: "✅ 마감" if x == 1 else "⏳ 진행")

        disp_s = s_df[[
            '발주번호', '발주차수', '진행상태', '거래처명', '상품명',
            '발주총액', '실입금액', '선급금액', '미수잔액', '통화'
        ]]

        def highlight_row(row):
            style = [''] * len(row)
            if row['진행상태'] == "✅ 마감":
                style = ['background-color: #f2f2f2; color: #999;'] * len(row)
            if row['선급금액'] > 0:
                style[row.index.get_loc('선급금액')] = 'color: red;'
            if row['미수잔액'] > 0:
                style[row.index.get_loc('미수잔액')] = 'color: blue;'
            return style

        st.dataframe(
            disp_s.style.apply(highlight_row, axis=1).format(
                '{:,.0f}',
                subset=['발주총액', '실입금액', '선급금액', '미수잔액']
            ),
            hide_index=True,
            use_container_width=True
        )

        st.divider()

        st.subheader("📝 입금 상세 내역")

        show_deleted = st.checkbox("삭제된 내역 보기")

        filtered['상태'] = filtered['삭제'].apply(lambda x: '삭제됨' if x else '')

        if '메모' in filtered.columns and '송금사유' not in filtered.columns:
            filtered['송금사유'] = filtered['메모']

        display_cols = [
            'id', '발주번호', '발주차수', '유형', '거래처명', '상품명',
            '통화', '입금일', '실입금액', '선급금액', '한화환산액',
            '송금사유', '삭제', '상태'
        ]

        display_p = filtered[
            [c for c in display_cols if c in filtered.columns]
        ].sort_values('입금일', ascending=False).reset_index(drop=True)

        if not show_deleted:
            display_p = display_p[display_p['삭제'] != True].reset_index(drop=True)

        edited_p = st.data_editor(
            display_p,
            key="payment_editor_v10",
            hide_index=True,
            use_container_width=True,
            column_config={
                "id": st.column_config.NumberColumn("ID", disabled=True),
                "삭제": st.column_config.CheckboxColumn("삭제"),
                "실입금액": st.column_config.NumberColumn("실입금액", format="%,.0f", step=1),
                "선급금액": st.column_config.NumberColumn("선급금액", format="%,.0f", step=1),
                "한화환산액": st.column_config.NumberColumn("한화환산액", format="%,.0f", disabled=True),
                "입금일": st.column_config.TextColumn("입금일"),
                "송금사유": st.column_config.TextColumn("송금사유")
            }
        )

        b1, b2, b3 = st.columns(3)

        if b1.button("💾 상세 수정 저장", use_container_width=True):
            state = st.session_state.get("payment_editor_v10", {})
            edited_rows = state.get("edited_rows", {})

            if edited_rows:
                try:
                    success_count = 0

                    for idx, changes in edited_rows.items():
                        tid = int(display_p.iloc[int(idx)]["id"])
                        up_data = {}

                        if "실입금액" in changes:
                            up_data["실입금액"] = to_money(changes["실입금액"])
                        if "선급금액" in changes:
                            up_data["선급금액"] = to_money(changes["선급금액"])
                        if "입금일" in changes:
                            up_data["입금일"] = str(changes["입금일"])
                        if "송금사유" in changes:
                            up_data["메모"] = str(changes["송금사유"])

                        if up_data:
                            supabase.table("payments").update(up_data).eq("id", tid).execute()
                            success_count += 1

                    st.success(f"{success_count}건 저장 완료! 페이지를 새로고침합니다.")
                    st.rerun()

                except Exception as e:
                    st.error(f"저장 실패: {e}")
            else:
                st.info("수정된 내용이 없습니다.")

        if b2.button("🗑️ 선택 삭제 실행", use_container_width=True):
            try:
                delete_ids = edited_p[edited_p['삭제'] == True]['id'].tolist()

                if not delete_ids:
                    st.info("삭제할 내역을 선택하세요.")
                else:
                    for tid in delete_ids:
                        supabase.table("payments").update({"삭제": True}).eq("id", int(tid)).execute()

                    st.success(f"{len(delete_ids)}건 삭제 처리 완료")
                    st.rerun()

            except Exception as e:
                st.error(f"삭제 실패: {e}")

        if b3.button("♻️ 선택 복구 실행", use_container_width=True):
            try:
                restore_ids = edited_p[edited_p['삭제'] == True]['id'].tolist()

                if not restore_ids:
                    st.info("복구할 내역을 선택하세요.")
                else:
                    for tid in restore_ids:
                        supabase.table("payments").update({"삭제": False}).eq("id", int(tid)).execute()

                    st.success(f"{len(restore_ids)}건 복구 완료")
                    st.rerun()

            except Exception as e:
                st.error(f"복구 실패: {e}")

        st.divider()

        m1, m2, m3 = st.columns(3)
        m1.metric("총 지급액 (KRW)", f"{calc_filtered['한화환산액'].sum():,.0f} 원")
        m2.metric("총 지급액 (USD)", f"${calc_filtered[calc_filtered['통화'] == 'USD']['실입금액'].sum():,.0f}")
        m3.metric("총 지급액 (CNY)", f"¥{calc_filtered[calc_filtered['통화'] == 'CNY']['실입금액'].sum():,.0f}")

    else:
        st.info("입금 내역 없음")
        
# --- [Tab 3] 거래처 관리 ---
with tabs[3]:
    st.header("🏢 거래처 정보 관리")
    
    # 1. 데이터 로드 및 정렬
    v_orig = get_supabase_data("vendors")
    if not v_orig.empty:
        v_orig = v_orig.sort_values('거래처명').reset_index(drop=True)
    
    col_v_in, col_v_csv = st.columns([1.5, 1])
    
    # --- 상단: 등록 섹션 ---
    with col_v_in:
        st.subheader("1. 신규 거래처 수기 등록")
        with st.form("new_v_form_full", clear_on_submit=True):
            v_c1, v_c2 = st.columns([2, 1])
            vn = v_c1.text_input("거래처명 (필수)")
            vt = v_c2.selectbox("기본 유형", CATEGORIES)
            
            v_c3, v_c4, v_c5 = st.columns([1, 2, 1])
            vb = v_c3.text_input("은행")
            va = v_c4.text_input("계좌번호")
            vh = v_c5.text_input("예금주")
            
            if st.form_submit_button("➕ 거래처 정보 저장", use_container_width=True):
                if vn:
                    upsert_supabase_data("vendors", {
                        "거래처명": vn, "기본유형": vt, "은행": vb, "계좌번호": va, "예금주": vh
                    })
                    st.success(f"✅ [{vn}] 등록 완료!"); st.rerun()
                else:
                    st.error("⚠️ 거래처명은 필수 입력 항목입니다.")

    with col_v_csv:
        st.subheader("2. CSV 일괄 등록")
        v_template = pd.DataFrame(columns=["거래처명", "기본유형", "은행", "계좌번호", "예금주"])
        st.download_button("📥 등록 양식(CSV) 다운로드", v_template.to_csv(index=False).encode('utf-8-sig'), "vendor_template.csv", use_container_width=True)
        up_vendor = st.file_uploader("파일 선택", type=['csv'], key="v_up_file")
        if up_vendor and st.button("🚀 일괄 저장 실행", use_container_width=True):
            try:
                df_v_up = pd.read_csv(up_vendor)
                df_v_up.columns = [str(c).strip().replace('\ufeff', '') for c in df_v_up.columns]
                v_list = [r.to_dict() for _, r in df_v_up.iterrows() if to_str(r.get('거래처명'))]
                if v_list:
                    upsert_supabase_data("vendors", v_list)
                    st.success(f"✨ {len(v_list)}건 등록 완료!"); st.rerun()
            except Exception as e:
                st.error(f"❌ 오류: {e}")

    st.divider()

    # --- 하단: 목록 수정 및 검색 (표 크기 최적화 핵심) ---
    if not v_orig.empty:
        st.subheader("📋 등록된 거래처 목록")
        
        v_search = st.text_input("🔍 거래처 검색 (이름 또는 은행)", placeholder="찾으시는 거래처명을 입력하세요...")
        
        display_v = v_orig.copy()
        if v_search:
            display_v = display_v[display_v['거래처명'].str.contains(v_search, case=False, na=False) | 
                                  display_v['은행'].str.contains(v_search, case=False, na=False)]

        # ✅ [UI 최적화] 데이터 양에 따른 가변 높이 설정
        # 헤더(약 40px) + 행당(약 35px). 최대 600px까지만 확장
        v_height = min(600, 45 + len(display_v) * 37)

        ev_v = st.data_editor(
            display_v, 
            hide_index=True, 
            use_container_width=True,
            height=v_height,  # 🔥 표가 너무 커지지 않게 자동 조절
            key="vendor_editor_v2",
            column_config={
                "거래처명": st.column_config.TextColumn("거래처명", width="medium"), # width 오타 수정
                "기본유형": st.column_config.SelectboxColumn("기본 유형", options=CATEGORIES, width="small"),
                "은행": st.column_config.TextColumn("은행", width="small"),
                "계좌번호": st.column_config.TextColumn("계좌번호", width="medium"), # width 오타 수정
                "예금주": st.column_config.TextColumn("예금주", width="small"),
            }
        )
        
        if st.button("💾 변경사항 동기화 저장", use_container_width=True):
            for i, r in ev_v.iterrows():
                target_id = r.get('id')
                if target_id:
                    old_row = v_orig[v_orig['id'] == target_id]
                    if not old_row.empty and old_row.iloc[0]['거래처명'] != r['거래처명']:
                        old_n = old_row.iloc[0]['거래처명']
                        # 연관 테이블(payments, orders) 동기화
                        supabase.table("payments").update({"거래처명": r['거래처명'], "유형": r['기본유형']}).eq("거래처명", old_n).execute()
                        supabase.table("orders").update({"거래처명": r['거래처명'], "유형": r['기본유형']}).eq("거래처명", old_n).execute()
            
            upsert_supabase_data("vendors", ev_v.to_dict(orient='records'))
            st.success("✅ 동기화 완료!"); st.rerun()
    else:
        st.info("📢 등록된 거래처 정보가 없습니다.")

# --- [Tab 4] 환율 분석 ---
with tabs[4]:
    st.header("📈 환율 데이터 분석 및 관리")
    
    # -------------------------------
    # 1. 업로드
    # -------------------------------
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

    # -------------------------------
    # 2. 분석
    # -------------------------------
    ex_db = get_supabase_data("exchange_rates")
    
    if not ex_db.empty:
        ex_db['날짜'] = pd.to_datetime(ex_db['날짜'])
        ex_db['연도'] = ex_db['날짜'].dt.year
        ex_db['월'] = ex_db['날짜'].dt.month

        df_target = ex_db[ex_db['연도'].isin([2025, 2026])]

        main_l, main_r = st.columns(2, gap="large")

        for i, curr in enumerate(['usd', 'cny']):
            target_col = main_l if i == 0 else main_r
            
            with target_col:
                st.subheader(f"💱 {curr.upper()} 분석 리포트")

                # -------------------------------
                # 📈 차트
                # -------------------------------
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

                # -------------------------------
                # 📊 월별 분석
                # -------------------------------
                m_avg = df_target.groupby(['연도', '월'])[curr].mean().reset_index()

                if not m_avg.empty:

                    import numpy as np  # 안전하게 여기서도 선언

                    # 전체 시계열 정렬
                    m_avg_sorted = m_avg.sort_values(['연도', '월']).copy()

                    # 전월 값
                    m_avg_sorted['전월값'] = m_avg_sorted[curr].shift(1)

                    # 전월 대비 %
                    m_avg_sorted['지난달대비(%)'] = (
                        (m_avg_sorted[curr] - m_avg_sorted['전월값']) /
                        m_avg_sorted['전월값'].replace(0, np.nan)
                    ) * 100

                    # pivot 생성
                    pivot = m_avg_sorted.pivot(index='월', columns='연도', values=curr)
                    pivot.columns = [f"{int(c)}년" for c in pivot.columns]

                    c25, c26 = "2025년", "2026년"

                    # 전년동월 대비
                    if c25 in pivot.columns and c26 in pivot.columns:
                        pivot['전년동월대비(%)'] = (
                            (pivot[c26] - pivot[c25]) /
                            pivot[c25].replace(0, np.nan)
                        ) * 100

                    # index → 컬럼
                    pivot = pivot.reset_index()

                    # 전월대비 merge
                    prev_df = m_avg_sorted[m_avg_sorted['연도'] == 2026][['월','지난달대비(%)']]
                    pivot = pivot.merge(prev_df, on='월', how='left')

                    # 컬럼 정리
                    cols = ['월']
                    if c25 in pivot.columns:
                        cols.append(c25)
                    if c26 in pivot.columns:
                        cols.append(c26)
                    cols += ['전년동월대비(%)', '지난달대비(%)']

                    pivot = pivot[cols]

                    # 제목
                    st.write(f"**{curr.upper()} 월별 환율 추이 분석**")

                    # 출력 (🔥 퍼센트 적용)
                    st.dataframe(
                        pivot.style.format({
                            c25: "{:,.2f}",
                            c26: "{:,.2f}",
                            '전년동월대비(%)': "{:.2f}%",
                            '지난달대비(%)': "{:.2f}%",
                        }),
                        use_container_width=True
                    )

                else:
                    st.info(f"{curr.upper()} 데이터 부족")

        st.divider()

        # -------------------------------
        # 🛠️ 원본 관리
        # -------------------------------
        with st.expander("🛠️ 환율 데이터 원본 관리 및 수정"):
            display_db = ex_db.copy().sort_values('날짜', ascending=False)
            display_db['날짜'] = display_db['날짜'].dt.strftime('%Y-%m-%d')

            cols = [c for c in ['날짜', 'usd', 'cny'] if c in display_db.columns]

            edited_ex = st.data_editor(
                display_db[cols],
                hide_index=True,
                use_container_width=True,
                column_config={
                    "날짜": st.column_config.TextColumn("날짜"),
                    "usd": st.column_config.NumberColumn("USD", format="%.2f"),
                    "cny": st.column_config.NumberColumn("CNY", format="%.2f")
                }
            )

            if st.button("💾 수정 내용 저장", use_container_width=True):
                try:
                    upsert_supabase_data("exchange_rates", edited_ex.to_dict(orient='records'))
                    st.success("저장 완료!")
                    st.rerun()
                except Exception as e:
                    st.error(f"저장 실패: {e}")

    else:
        st.info("환율 데이터를 업로드해 주세요.")
