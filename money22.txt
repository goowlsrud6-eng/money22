import streamlit as st
import pandas as pd

# 1. 초기 데이터 설정 (예시)
if 'ledger' not in st.session_state:
    # 예시 데이터: 업체, 품목, 발주총액, 입금내역리스트
    st.session_state.ledger = {
        "우일코리아/에어메쉬/초도": {
            "total_order": 1000000,
            "history": []
        }
    }

st.title("💰 자금 관리 프로그램 (선급금 직관 관리)")

# 2. 작성 구역 (팀원용)
st.subheader("📝 입금 및 선급금 기록")
with st.expander("내역 추가하기", expanded=True):
    target = st.selectbox("품목 선택", list(st.session_state.ledger.keys()))
    col1, col2, col3 = st.columns(3)
    
    with col1:
        deposit = st.number_input("실제 입금 금액 (통장 출금액)", value=0, step=10000)
    with col2:
        # 사용자님의 핵심 요구사항: (+)는 적립, (-)는 차감
        advance = st.number_input("선급금 처리 (+적립 / -차감)", value=0, step=10000)
    with col3:
        reason = st.text_input("송금 사유", value="정산")

    if st.button("기록 저장"):
        st.session_state.ledger[target]["history"].append({
            "date": pd.Timestamp.now().strftime("%Y-%m-%d"),
            "deposit": deposit,
            "advance": advance,
            "reason": reason
        })
        st.success("기록되었습니다!")

st.divider()

# 3. 가로형 대시보드 (사용자님/대표님용)
st.subheader("📊 전체 현황 (가로 보기)")

summary_list = []
for title, data in st.session_state.ledger.items():
    history_df = pd.DataFrame(data["history"])
    
    # 계산 로직
    total_paid = history_df["deposit"].sum() if not history_df.empty else 0
    # 선급금 잔액 = 모든 선급금 처리(+와 -)의 합산
    adv_balance = history_df["advance"].sum() if not history_df.empty else 0
    # 최종 잔액 = 발주총액 - 실제입금총액
    remaining = data["total_order"] - total_paid

    summary_list.append({
        "업체/품목/차수": title,
        "발주총액": f"{data['total_order']:,}",
        "실제 입금 합계": f"{total_paid:,}",
        "선급금 잔액": f"{adv_balance:,}", # 0원이 되면 정산 완료를 직관적으로 확인
        "미지급 잔액": f"{remaining:,}"
    })

st.table(pd.DataFrame(summary_list))

# 4. 상세 히스토리 (엑셀처럼 가로로 길게 보기 가능)
if st.checkbox("상세 입금 히스토리 보기"):
    for title, data in st.session_state.ledger.items():
        if data["history"]:
            st.write(f"📍 {title} 상세 내역")
            st.dataframe(pd.DataFrame(data["history"]))