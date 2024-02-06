from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd
import os

def major_index():
    stock = {}
    i = 0

    url = "https://finance.naver.com/sise/"
    column = ["종목명", "현재가", "전일비/등락률"]

    driver = webdriver.Chrome()
    driver.get(url)
    time.sleep(2)

    koreas = driver.find_element(By.CLASS_NAME, "box_top_submain2")
    korea = koreas.find_element(By.TAG_NAME, "ul")

    li_elements = korea.find_elements(By.TAG_NAME, "li")
    li_elements.pop()

    for li_element in li_elements:
        prices = li_element.find_elements(By.TAG_NAME, "span")
        
        name = prices[0].text
        price = prices[1].text
        delta = prices[2].text.replace("\n", " ").split(" ")
        delta.pop()

        new_delta = []
        if delta[1][0] == "+":
            new_delta.append("+" + delta[0])
        else:
            new_delta.append("-" + delta[0])
        
        new_delta.append(delta[1])

        stock[i] = [name, price, new_delta[0] + "/" + new_delta[1]]
        i += 1

    url = "https://finance.naver.com/world/"

    driver = webdriver.Chrome()
    driver.get(url)
    time.sleep(2)

    market1 = driver.find_element(By.XPATH, '//*[@id="worldIndexColumn1"]')
    li_elements = market1.find_elements(By.TAG_NAME, "li")
    for li_element in li_elements:
        name = li_element.find_element(By.TAG_NAME, "dt").text

        data = li_element.find_element(By.CLASS_NAME, "point_status")
        price = data.find_element(By.TAG_NAME, "strong").text
        delta = data.find_element(By.TAG_NAME, "em").text
        percentage = data.find_element(By.TAG_NAME, "span").text

        if percentage[0] == "+":
            delta = "+" + delta
        else:
            delta = "-" + delta

        stock[i] = [name, price, delta + "/" + percentage]
        i += 1

    market2 = driver.find_element(By.XPATH, '//*[@id="worldIndexColumn2"]')
    li_elements = market2.find_elements(By.TAG_NAME, "li")
    for li_element in li_elements:
        name = li_element.find_element(By.TAG_NAME, "dt").text

        data = li_element.find_element(By.CLASS_NAME, "point_status")
        price = data.find_element(By.TAG_NAME, "strong").text
        delta = data.find_element(By.TAG_NAME, "em").text
        percentage = data.find_element(By.TAG_NAME, "span").text

        if percentage[0] == "+":
            delta = "+" + delta
        else:
            delta = "-" + delta

        stock[i] = [name, price, delta + "/" + percentage]
        i += 1

    market3 = driver.find_element(By.XPATH, '//*[@id="worldIndexColumn3"]')
    li_elements = market3.find_elements(By.TAG_NAME, "li")
    for li_element in li_elements:
        name = li_element.find_element(By.TAG_NAME, "dt").text

        data = li_element.find_element(By.CLASS_NAME, "point_status")
        price = data.find_element(By.TAG_NAME, "strong").text
        delta = data.find_element(By.TAG_NAME, "em").text
        percentage = data.find_element(By.TAG_NAME, "span").text

        if percentage[0] == "+":
            delta = "+" + delta
        else:
            delta = "-" + delta

        stock[i] = [name, price, delta + "/" + percentage]
        i += 1

    df_stock = pd.DataFrame.from_dict(stock, orient = "index")
    df_stock.columns = column

    return df_stock

def index_as_excel(df_stock):
   output_directory = r"D:\DART\EXCEL"
   file_path = os.path.join(output_directory, "5_today_index.xlsx")
   with pd.ExcelWriter(file_path) as writer:
      df_stock.to_excel(writer, index = False)
      print("5_today_index.xlsx saved!")
