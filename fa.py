import streamlit as st
import yfinance as yf
import pandas as pd
import io
from typing import Any

# --- 1. 페이지 설정 및 제목 ---
st.set_page_config(
    page_title="📈 yfinance 재무 데이터 엑셀 추출기 (PER, PBR, ROE 포함)", 
    layout="wide"
)
st.title("💰 yfinance 재무 데이터 일괄 다운로더")
st.markdown("---")

# --- 2. 데이터 추출 및 엑셀 생성 함수 ---

@st.cache_data(ttl=3600)
def fetch_and_create_excel(ticker: str) -> io.BytesIO | None:
    """
    yfinance에서 연간/분기 재무 데이터와 주요 통계를 가져와 메모리 내 엑셀 파일을 생성합니다.
    (yfinance의 속성명 변경에 대응하여 'financials'와 'quarterly_financials'를 활용합니다.)
    """
    try:
        stock = yf.Ticker(ticker)
        # 티커가 유효한지 확인하기 위한 기본 정보 요청
        if not stock.info:
            return None
    except Exception:
        return None

    # 데이터 수집 (Sheet Name: DataFrame 구조)
    financial_data: dict[str, pd.DataFrame] = {}
    
    # ----------------------------------------------------
    # 재무 데이터 수집 및 시트 이름 명확화 (연간/분기)
    # yfinance의 최신 버전(0.2.x) 호환성을 위해 속성명을 변경합니다.
    # ----------------------------------------------------

    # 연간 데이터
    # stock.income_stmt 대신 financials 사용
    if not stock.financials.empty:
        financial_data["Income_Statement (연간)"] = stock.financials
    if not stock.balance_sheet.empty:
        financial_data["Balance_Sheet (연간)"] = stock.balance_sheet
    if not stock.cashflow.empty:
        financial_data["Cash_Flow (연간)"] = stock.cashflow

    # 분기별 데이터
    # stock.quarterly_income_stmt 대신 quarterly_financials 사용
    if not stock.quarterly_financials.empty:
        financial_data["Income_Statement (분기)"] = stock.quarterly_financials
    if not stock.quarterly_balance_sheet.empty:
        financial_data["Balance_Sheet (분기)"] = stock.quarterly_balance_sheet
    if not stock.quarterly_cashflow.empty:
        financial_data["Cash_Flow (분기)"] = stock.quarterly_cashflow
        
    # ----------------------------------------------------
    # 주요 통계 데이터 정리 (Key_Statistics)
    # ----------------------------------------------------
    info: dict[str, Any] = stock.info
    key_stats_raw = {
        "Full Name (전체 이름)": info.get('longName', 'N/A'), # 종목 이름 추가
        "Market Cap (시가총액)": info.get('marketCap'),
        "Trailing P/E (PER)": info.get('trailingPE'), 
        "Price/Book (PBR)": info.get('priceToBook'), 
        "Return On Equity (ROE)": info.get('returnOnEquity'),
        "5Y EPS Growth (5년 EPS 성장률)": info.get('fiveYearAvgProfitGrowth'), 
        "Dividend Yield (배당수익률)": info.get('dividendYield'),
        "Beta (시장 민감도)": info.get('beta'),
        "Forward P/E (선행 PER)": info.get('forwardPE'),
        "Shares Outstanding (총 발행 주식수)": info.get('sharesOutstanding'),
    }

    stats_df = pd.DataFrame.from_dict(key_stats_raw, orient='index', columns=['Value'])
    stats_df.index.name = 'Metric'
    financial_data["Key_Statistics"] = stats_df # 통계 시트 추가

    output = io.BytesIO()
    is_data_present = False
    
    try:
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            for sheet_name, df in financial_data.items():
                
                # 데이터프레임이 존재하고 비어있지 않은지 확인
                if df is not None and not df.empty:
                    
                    if sheet_name == "Key_Statistics":
                        # 통계 시트는 Transpose하지 않음
                        df.to_excel(writer, sheet_name=sheet_name, index=True)
                    else:
                        # 재무 3표는 날짜를 가로(컬럼)로 만들기 위해 Transpose
                        # 인덱스 이름 설정하여 엑셀에서 명확하게 표시
                        df.T.index.name = "Date"
                        df.T.to_excel(writer, sheet_name=sheet_name, index=True)
                        
                    is_data_present = True

    except Exception as e:
        # 오류 발생 시 디버깅을 위해 에러 로그 출력 가능
        st.error(f"Excel 파일 생성 중 오류 발생: {e}")
        return None

    if not is_data_present:
        return None
    
    output.seek(0)
    return output

# --- 3. Streamlit UI 구현 (단일 페이지) ---

st.header("⬇️ 개별 티커 데이터 다운로드")
st.info("재무 3표 데이터와 주요 통계 지표(PER, PBR, ROE 포함)를 하나의 엑셀 파일로 추출합니다. 연간/분기 데이터가 시트 이름으로 명확히 구분됩니다.")

ticker_input = st.text_input(
    "분석할 주식 티커를 입력하고 Enter를 누르세요 (예: TSLA, AAPL, 005930.KS)", 
    "", 
    key="download_ticker"
).upper()

if ticker_input:
    st.markdown(f"### '{ticker_input}' 데이터 추출 중...")
    
    # 데이터 추출 및 엑셀 생성
    with st.spinner("재무 데이터 및 통계 수집 중..."):
        excel_buffer = fetch_and_create_excel(ticker_input)

    if excel_buffer:
        today_str = pd.Timestamp.now().strftime("%Y%m%d")
        download_filename = f"{ticker_input}_Financials_Stats_{today_str}.xlsx"
        
        st.success(f"✅ '{ticker_input}' 엑셀 파일 생성이 완료되었습니다!")
        st.download_button(
            label="⬇️ 엑셀 파일 다운로드 (.xlsx)",
            data=excel_buffer,
            file_name=download_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    else:
        st.error(f"⚠️ **{ticker_input}**에 대한 유효한 재무 데이터를 찾을 수 없거나 파일 생성에 실패했습니다. 티커를 확인해 주세요. (예: 한국 주식은 **005930.KS**와 같이 '.KS'를 붙여야 합니다.)")

st.markdown("---")
st.caption("Powered by yfinance & Streamlit")
