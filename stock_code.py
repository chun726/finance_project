# DART 오픈 API 인증키
crtfc_key = "39adf78493b2846aba6e77ffc54edb841ab81c8e"

# Import libraries
import requests
import request_error_info as request_error
import random
import zipfile
import io
import xml.etree.ElementTree as et
import pandas as pd
import os

# 기업 고유번호 크롤링
def get_unique_number(crtfc_key = "39adf78493b2846aba6e77ffc54edb841ab81c8e"):

    # https://opendart.fss.or.kr/guide/detail.do?apiGrpCd=DS001&apiId=2019018
    # DART 오픈 API 고유번호 개발가이드 참고
    params = {"crtfc_key":crtfc_key}
    items = ["corp_code", "corp_name", "stock_code", "modify_date"]
    item_names = ["고유번호", "정식명칭", "종목코드", "최종변경일자"]
    url = "https://opendart.fss.or.kr/api/corpCode.xml"

    # Request library를 통한 호출
    res = requests.get(url, params = params)
    req_status = res.status_code
    t = random.randint(0, len(request_error._codes[req_status]) - 1)
    print("Request status")
    print(req_status, ":", request_error._codes[req_status][t] ,end="\n\n")

    # Zipfile로 결과를 받아오고 열기
    zfile = zipfile.ZipFile(io.BytesIO(res.content))
    fin = zfile.open(zfile.namelist()[0])

    # UTF-8 decoding
    root = et.fromstring(fin.read().decode("utf-8"))
    data = []

    # 데이터 저장
    for child in root: # root의 요소들에 의해
        if len(child.find("stock_code").text.strip()) > 1: # stock_code가 존재하면
            data.append([]) # 저장할 빈 list를 만들어 준 이후
            for item in items:
                data[-1].append(child.find(item).text) # items의 항목들을 저장
    
    df = pd.DataFrame(data, columns=item_names) # Pandas dataframe으로 변환
    df.set_index("정식명칭", inplace = True)
    return df

# 기업명으로 종목코드 확인
def get_stock_code(name):
    try:
        df = get_unique_number()
        names = ["종목", "종목코드", "고유번호"]

        code1 = df.loc[name, "종목코드"] # 해당 이름을 가지는 회사의 종목코드를 받아옴
        code2 = df.loc[name, "고유번호"]
        data = [name, code1, code2]

        company = pd.DataFrame([data], columns = names)
        print(company, end = "\n\n")

    except: # 에러처리
        print("=========ERROR=========")
        print("유효하지 않은 이름입니다!")
    
    return code1, code2

def save_as_excel(frame):
    try:
        output_directory = r"D:\DART\EXCEL"
        file_path = os.path.join(output_directory, "stock_code.xlsx")
        frame.to_excel(file_path)

        print("stock_code.xlsx saved!")
    except:
        print("DataFrame이 올바르지 않습니다!")