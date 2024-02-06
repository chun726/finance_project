import requests
import pandas as pd
from bs4 import BeautifulSoup as bs
import os
import re

def isEven(num):
    if num == 0:
        raise Exception("0")
    elif num%2 == 0:
        return True
    elif num%2 == 1:
        return False
    
def industry_find():
    # URL
    url = "https://finance.naver.com/sise/sise_group.naver?type=upjong"
    industry = {}
    industry_code = {}
    # Getting response
    response = requests.get(url)
    soup = bs(response.text, "html.parser")

    # Find a table
    table = soup.find("table", class_= "type_1")

    if table:
        rows = table.find_all("tr")
        for row in rows:
            tds = row.find_all("td")
            if tds and len(tds) >= 2:
                link_td = tds[0]
                if link_td.a:
                    href = link_td.a["href"]
                    text = link_td.a.get_text()
                    code = href[-3:].strip("=")

                    industry[text] = href
                    industry_code[text] = code
    
    df_industry_code = pd.DataFrame(list(industry_code.items()), columns = ["산업", "코드"])
    output_directory = r"D:\DART\EXCEL"
    file_path = os.path.join(output_directory, "industry_code_for_naver.xlsx")    

    with pd.ExcelWriter(file_path) as writer:
        df_industry_code.to_excel(writer)
        print("industry_code_for_naver.xlsx saved!")   
    return industry



def company_find(industry):
    company = {}
    for key in industry.keys():
        temp_company = []

        header = "https://finance.naver.com/"
        value = industry[key]
        url = header + value

        response = requests.get(url)
        soup = bs(response.text, "html.parser")

        table = soup.find("table", class_= "type_5")
        tbody = table.find("tbody")
        rows = tbody.find_all("tr")

        for row in rows:
            name_td = row.find("td", class_ = "name")
            if name_td: 
                company_name = name_td.a.get_text(strip = True)
                temp_company.append(company_name)
        
        company[key] = temp_company
    
    return company
    

def as_excel(company):
    df_company = pd.DataFrame(list(company.items()), columns = ["산업", "기업"])
    output_directory = r"D:\DART\EXCEL"
    file_path = os.path.join(output_directory, "industry_company.xlsx")

    with pd.ExcelWriter(file_path) as writer:
        df_company.to_excel(writer)
        print("industry_company.xlsx saved!")

def find_industry(name):
    if name == "현대자동차":
        name = "현대차"
    df = pd.read_excel("D:\DART\EXCEL\industry_company.xlsx")
    for index, row in df.iterrows():
        if name in row["기업"]:
            return row["산업"]
    
    return "오류"

def industry_per(code):
    base = "https://finance.naver.com/item/main.naver?code="
    url = base + code
    response = requests.get(url)
    soup = bs(response.text, "html.parser")
    element = soup.find("table", summary="동일업종 PER 정보")
    table = element.find("tr", class_ = "strong")
    data = table.find("em").text.strip()

    return data


def naver_fundamentals(code):
    base = "https://navercomp.wisereport.co.kr/v2/company/c1010001.aspx?cmp_cd="
    final = "&cn="
    url = base + code + final
    response = requests.get(url)
    soup = bs(response.content, "html.parser")
    table = soup.find("table", class_="gHead03")

    left_elements = table.find_all("th", class_="left")
    fundamentals = {}

    heads = table.find_all("th", style="padding-left:0; padding-right:0")
    for i in range(len(heads)):
        heads[i] = heads[i].get_text(strip=True)

    i = 0
    for left_element in left_elements:
        category = left_element.get_text(strip=True)
        num_element_left = left_element.find_next_sibling("td", class_="num").get_text(strip=True)
        num_element_right = left_element.find_next_sibling("td", class_="num noline-right").get_text(strip=True)
        temp_element = [category, num_element_left, num_element_right]
        fundamentals[i] = temp_element
        i += 1


    df_fundamentals = pd.DataFrame.from_dict(fundamentals, orient = "index")
    df_fundamentals.columns = heads
    
    return df_fundamentals


def naver_price(code):
    base = "https://navercomp.wisereport.co.kr/v2/company/c1010001.aspx?cmp_cd="
    final = "&cn="
    url = base + code + final
    response = requests.get(url)
    soup = bs(response.content, "html.parser")
    table = soup.find("table", class_="gHead")

    titles = table.find_all("th", scope = "row")
    price = {}

    date = soup.find("dd", class_ = "header-table-cell unit").get_text(strip = True).strip("[""]")

    i = 0
    for title in titles:
        name = title.get_text(strip=True)
        #print(name)

        content = title.find_next_sibling("td", class_ = "num").get_text(strip = True)
        cleaned_content = re.sub(r"\s+", " ", content)
        #print(content)

        price[i] = [name, cleaned_content]
        i += 1

    head = ["주요지표", date]
    df_price = pd.DataFrame.from_dict(price, orient = "index")
    df_price.columns = head

    return df_price

def naver_holdings(code):
    base = "https://navercomp.wisereport.co.kr/v2/company/c1010001.aspx?cmp_cd="
    final = "&cn="
    url = base + code + final
    response = requests.get(url)
    soup = bs(response.content, "html.parser")
    table = soup.find("table", id = "cTB13")
    members = table.find_all("tr", class_=["p_sJJ10", "p_sJJ20"])

    head = table.find("caption", class_ = "blind").get_text(strip = True).split(",")

    i = 0
    holdings = {}
    for member in members:
        holder = member.find("td", class_ = "line txt").get_text(strip = True)
        holding = member.find("td", class_ = "line num").get_text(strip = True)
        percentage = member.find("td", class_ = "noline-right num").get_text(strip = True) + "%"

        half_length = len(holder)//2

        if isEven(len(holder)) == True and holder[:half_length] == holder[half_length:]:
            holder = holder[:half_length]

        holdings[i] = [holder, holding, percentage]
        i += 1

    df_holdings = pd.DataFrame.from_dict(holdings, orient = "index")
    df_holdings.columns = head

    return df_holdings

def naver_as_excel(fundamentals, price, holdings):
   output_directory = r"D:\DART\EXCEL"
   file_path = os.path.join(output_directory, "1_analysis.xlsx")
   with pd.ExcelWriter(file_path) as writer:
      fundamentals.to_excel(writer, sheet_name="펀더멘털", index = False)
      price.to_excel(writer, sheet_name = "시세 및 주주현황", index = False)
      holdings.to_excel(writer, sheet_name = "주요주주", index = False)
      print("1_analysis.xlsx saved!")