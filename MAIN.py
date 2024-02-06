# crtfc 인증키
key = "39adf78493b2846aba6e77ffc54edb841ab81c8e"


# 작성한 file들의 함수들을 import
import stock_code as stock_code # 고유번호 및 종목번호
import financial_statements as statements #재무제표
import naver_finance as finance #투자정보
import industry_breif as breif #산업정보
import naver_international as international #주가지수 정보

"""
bsns_year: 사업연도(4자리) - 2015년 이후 지원 - str 자료형
fs_div: 개별/연결 구분 - OFS:재무제표, CFS:연결재무제표 - str 자료형
report_code: 보고서 코드 - str 자료형
             11013: 1분기 보고서
             11012: 반기 보고서
             11014: 3분기 보고서
             11011: 사업 보고서
"""

# 분석할 기업명
name = "삼성전자"

### 고유번호 및 종목번호 ###

# 기업 이름으로부터 종목 코드를 받아옴
company_code, unique_code = stock_code.get_stock_code(name)

# Excel로 저장이 필요할 때 주석 해제 후 실행




### 투자지표  ###

company_industry = finance.find_industry(name) #산업군 분류
industry_per = finance.industry_per(company_code) #업종 평균 PER
fundamentals = finance.naver_fundamentals(company_code) #주요 투자지표
price = finance.naver_price(company_code) #시세 및 주주현황
holdings = finance.naver_holdings(company_code) #주요 주주

print("산업군:", company_industry) #산업군 분류
print("업종 평균 PER:", industry_per, end="\n\n") #업종 평균 PER
print(fundamentals, end="\n\n") #주요 투자지표
print(price, end="\n\n") #시세 및 주주현황
print(holdings, end = "\n\n") #주요 주주



### 산업분석 ###

# 산업군 이름으로 해당 산업 요약 제공. 대형 산업(제약 등)의 경우 속도가 느릴 수 있음
industry_summary = breif.industry_summary(company_industry)

print(industry_summary.head(), end="\n\n")



### 재무제표 및 주요 개정 ###

# 위의 값을 선택해 넣어서 재무제표 크롤링
df_book = statements.get_company_book(crtfc_key= key, corp_code = unique_code, 
                                      bsns_year = "2022", reprt_code = "11011", fs_div = "CFS" )

# 재무제표 주요 정보(당기/전기/전전기)
current, past, pastpast = statements.information(df_book)



### 전세계 주요 주가지수 ###

# 주요 지수 요약본
index = international.major_index()

print(index.head(), end = "\n\n")

### 최초 1회 실행 - 데이터 다운로드 후 Excel 저장 ###

# 유가증권시장/코스닥 산업별 종목 분류. 실행 시간 오래걸림. 1회 실행 이후 excel 로딩하여 사용 추천
# finance.as_excel(finance.company_find(finance.industry_find())

# DART 오픈API로부터 기업명과 종목코드를 받아 Excel로 저장.
# stock_code.save_as_excel(stock_code.get_unique_number())



### Excel 저장 구역 ###
# 1. 종목분석 결과 Excel로 저장시 주석 해제 후 실행
finance.naver_as_excel(fundamentals, price, holdings)

# 2. 산업분석 결과 Excel로 저장시 주석 해제 후 실행
breif.summary_as_excel(industry_summary)

# 3. 재무제표 Excel로 저장시 주석 해제 후 실행
statements.save_as_excel(df_book)

# 4. 재무제표 주요 정보 Excel로 저장시 주석 해제 후 실행
statements.info_as_excel(current, past, pastpast)

# 5. 전세계 주요 주가지수 Excel로 저장시 주석 해제 후 실행
international.index_as_excel(index)