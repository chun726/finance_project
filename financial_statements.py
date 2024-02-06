# DART 오픈 API 인증키
crtfc_key = "39adf78493b2846aba6e77ffc54edb841ab81c8e"

# Import libraries
import requests
import pandas as pd
import json
import os

def convertFnltt(url, items, item_names, params):
    res = requests.get(url, params)
    json_dict = json.loads(res.text)
    data = []
    if json_dict['status'] == "000":
      for line in json_dict['list']:
        data.append([])
        for itm in items:
          if itm in line.keys():
            data[-1].append(line[itm])
          else: data[-1].append('')
    df = pd.DataFrame(data,columns=item_names)
    return df

    
def get_company_book(crtfc_key, corp_code, bsns_year, reprt_code, fs_div):
    # 오픈 API 개발 가이드 참고
    # https://opendart.fss.or.kr/guide/detail.do?apiGrpCd=DS003&apiId=2019020

    url = "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json"
    
    # json의 출력값(items)에 해당하는 dataframe column명(item_names) 지정
    items = ["rcept_no","reprt_code","bsns_year","corp_code","sj_div","sj_nm", 
             "account_id","account_nm","account_detail","thstrm_nm", "thstrm_amount",
             "thstrm_add_amount","frmtrm_nm","frmtrm_amount", "frmtrm_q_nm","frmtrm_q_amount",
             "frmtrm_add_amount","bfefrmtrm_nm", "bfefrmtrm_amount","ord", "currency"] 
    item_names = ["접수번호","보고서코드","사업연도","고유번호","재무제표구분", "재무제표명",
                  "계정ID","계정명","계정상세","당기명","당기금액", "당기누적금액","전기명","전기금액","전기명(분/반기)", 
                  "전기금액(분/반기)","전기누적금액","전전기명","전전기금액", "계정과목정렬순서", "통화 단위"] 
    
    params = {"crtfc_key":crtfc_key, "corp_code":corp_code, "bsns_year":bsns_year, 
              "reprt_code":reprt_code, "fs_div":fs_div}
    
    df_book = convertFnltt(url, items, item_names, params)

    to_drop = ["접수번호", "보고서코드", "고유번호", "재무제표구분", "계정ID", "당기누적금액", 
               "전기누적금액", "계정과목정렬순서", "통화 단위", "전기금액(분/반기)", "전기명(분/반기)"]
    df_book = df_book.drop(labels = to_drop, axis = 1)

    

       
    return df_book

def save_as_excel(df_book):
    output_directory = r"D:\DART\EXCEL"
    file_path = os.path.join(output_directory, "3_financial_statement.xlsx")

    financial_dataframe = {}
    unique_df = df_book["재무제표명"].unique()
    try:
      with pd.ExcelWriter(file_path) as writer:
         for table in unique_df:
            financial_dataframe[table] = df_book[df_book["재무제표명"] == table]
            financial_dataframe[table].to_excel(writer, sheet_name = table, index = False)
      print("3_financial_statement.xlsx saved!")
    except:
       print("해당 공시가 존재하지 않습니다!")


def information(df_book):
   manufacture = True
   try:  
      total_revenue_row = df_book.loc[df_book["계정명"].str.contains("매출")]
      current_total_revenue = total_revenue_row["당기금액"].iloc[0]
      past_total_revenue = total_revenue_row["전기금액"].iloc[0]
      pastpast_total_revenue = total_revenue_row["전전기금액"].iloc[0]

      gross_profit_row = df_book.loc[df_book["계정명"] == "매출총이익"]
      current_gross_profit = gross_profit_row["당기금액"].iloc[0]
      past_gross_profit = gross_profit_row["전기금액"].iloc[0]
      pastpast_gross_profit = gross_profit_row["전전기금액"].iloc[0]

      SGnA_expense_row = df_book.loc[df_book["계정명"].str.contains("판매비와")]
      current_SGnA_expense = SGnA_expense_row["당기금액"].iloc[0]
      past_SGnA_expense = SGnA_expense_row["전기금액"].iloc[0]
      pastpast_SGnA_expense = SGnA_expense_row["전전기금액"].iloc[0]
   except:
      manufacture = False
      operating_revenue_row = df_book.loc[df_book["계정명"].str.contains("영업수익")]
      current_operating_revenue = operating_revenue_row["당기금액"].iloc[0]
      past_operating_revenue = operating_revenue_row["전기금액"].iloc[0]
      pastpast_operating_revenue = operating_revenue_row["전전기금액"].iloc[0]

      operating_expense_row = df_book.loc[df_book["계정명"].str.contains("영업비용")]
      current_operating_expense = operating_expense_row["당기금액"].iloc[0]
      past_operating_expense = operating_expense_row["전기금액"].iloc[0]
      pastpast_operating_expense = operating_expense_row["전전기금액"].iloc[0]
      

   operating_profit_row = df_book.loc[df_book["계정명"].str.contains("영업이익")]
   current_operating_profit = operating_profit_row["당기금액"].iloc[0]
   past_operating_profit = operating_profit_row["전기금액"].iloc[0]
   pastpast_operating_profit = operating_profit_row["전전기금액"].iloc[0]

   try:
      earning_beforetax_row = df_book.loc[df_book["계정명"].str.contains("법인세비용차감전")]
      current_earning_beforetax = earning_beforetax_row["당기금액"].iloc[0]
      past_earning_beforetax = earning_beforetax_row["전기금액"].iloc[0]
      pastpast_earning_beforetax = earning_beforetax_row["전전기금액"].iloc[0]
   except:
      earning_beforetax_row = df_book.loc[df_book["계정명"].str.contains("법인세차감전")]
      current_earning_beforetax = earning_beforetax_row["당기금액"].iloc[0]
      past_earning_beforetax = earning_beforetax_row["전기금액"].iloc[0]
      pastpast_earning_beforetax = earning_beforetax_row["전전기금액"].iloc[0]

   net_profit_row = df_book.loc[df_book["계정명"].str.contains("당기순이익")]
   current_net_profit = net_profit_row["당기금액"].iloc[0]
   past_net_profit = net_profit_row["전기금액"].iloc[0]
   pastpast_net_profit = net_profit_row["전전기금액"].iloc[0]


   if manufacture == True:
      current = {"매출액": current_total_revenue, "매출총이익":current_gross_profit, "판관비":current_SGnA_expense, "영업이익":current_operating_profit,
               "법인세비용차감전순이익":current_earning_beforetax, "순이익":current_net_profit}
      
      past = {"매출액": past_total_revenue, "매출총이익":past_gross_profit, "판관비":past_SGnA_expense, "영업이익":past_operating_profit,
               "법인세비용차감전순이익":past_earning_beforetax, "순이익":past_net_profit}
      
      pastpast = {"매출액": pastpast_total_revenue, "매출총이익":pastpast_gross_profit, "판관비":pastpast_SGnA_expense, "영업이익":pastpast_operating_profit,
               "법인세비용차감전순이익":pastpast_earning_beforetax, "순이익":pastpast_net_profit}
   else:
      current = {"영업수익":current_operating_revenue, "영업비용":current_operating_expense, "영업이익":current_operating_profit,
               "법인세비용차감전순이익":current_earning_beforetax, "순이익":current_net_profit}
      
      past = {"영업수익":past_operating_revenue, "영업비용":past_operating_expense, "영업이익":past_operating_profit,
               "법인세비용차감전순이익":past_earning_beforetax, "순이익":past_net_profit}
      
      pastpast = {"영업수익":pastpast_operating_revenue, "영업비용":past_operating_expense, "영업이익":pastpast_operating_profit,
               "법인세비용차감전순이익":pastpast_earning_beforetax, "순이익":pastpast_net_profit}      
   
   return current, past, pastpast

def info_as_excel(current, past, pastpast):
   df_current = pd.DataFrame(list(current.items()), columns = ["개정과목", "금액"])
   df_past = pd.DataFrame(list(past.items()), columns = ["개정과목", "금액"])
   df_pastpast = pd.DataFrame(list(pastpast.items()), columns = ["개정과목", "금액"])


   output_directory = r"D:\DART\EXCEL"
   file_path = os.path.join(output_directory, "4_breif_statement.xlsx")
   with pd.ExcelWriter(file_path) as writer:
      df_current.to_excel(writer, sheet_name="당기", index = False)
      df_past.to_excel(writer, sheet_name = "전기", index = False)
      df_pastpast.to_excel(writer, sheet_name = "전전기", index = False)
      print("4_breif_statement.xlsx saved!")




