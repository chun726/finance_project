import pandas as pd
from bs4 import BeautifulSoup as bs
import requests
import os

industry = "자동차"

def industry_summary(industry):
    df = pd.read_excel("D:\DART\EXCEL\industry_code_for_naver.xlsx")

    selected_row = df[df["산업"] == industry]
    code = str(selected_row["코드"].tolist()[0])

    base = "https://finance.naver.com/sise/sise_group_detail.naver?type=upjong&no="
    finance_base = "https://navercomp.wisereport.co.kr/v2/company/c1010001.aspx?cmp_cd="
    finance_final = "&cn="

    url = base + code
    response = requests.get(url)
    soup = bs(response.content, "html.parser")

    table = soup.find("table", summary = "업종별 시세 리스트")
    tbody = table.find("tbody")
    companies = tbody.find_all("tr", onmouseover = "mouseOver(this)")

    company_link = {}

    for company in companies:
        company_name = company.find("div", class_ = "name_area")
        
        company_href = company_name.a["href"].strip("/item/main.naver?code=")
        company_name = company_name.get_text(strip = True)

        company_link[company_name] = finance_base + company_href + finance_final

    industry_info = ["종목명", "시가총액(억원)", "수익률(1M)", "수익률(3M)", "수익률(6M)", "수익률(1Y)"]

    industry_data = {}

    i = 0
    for key in company_link.keys():
        url = company_link[key]
        response = requests.get(url)
        soup = bs(response.content, "html.parser")

        name = [key]

        table = soup.find("table", id = "cTB11")
        if table:
            tbody = table.find("tbody")

            trs = tbody.find_all("tr")
            
            for tr in trs:
                th = tr.find("th")
                if th.get_text(strip = True) == "시가총액":
                    market_cap = [tr.find("td").get_text(strip = True).strip("억원")]
                elif th.get_text(strip = True) == "수익률(1M/3M/6M/1Y)":
                    rate_of_return = tr.find("td").get_text(strip = True).split("/")
            
            data = name + market_cap + rate_of_return
        else:
            data = name + ["N/A", "N/A", "N/A", "N/A", "N/A"]
        
        industry_data[i] = data
        i += 1

    df_industry_data = pd.DataFrame.from_dict(industry_data, orient = "index")
    df_industry_data.columns = industry_info

    return df_industry_data

def summary_as_excel(df_industry_data):
   output_directory = r"D:\DART\EXCEL"
   file_path = os.path.join(output_directory, "2_industry_summary.xlsx")
   with pd.ExcelWriter(file_path) as writer:
      df_industry_data.to_excel(writer, index = False)
      print("2_industry_summary.xlsx saved!")


