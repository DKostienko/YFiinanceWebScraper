class Stock_Financials:

    def dcf_yfinance():
        import pandas as pd
        from bs4 import BeautifulSoup
        import lxml
        import requests
        import openpyxl
        from openpyxl import load_workbook
        import time
        import os
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        import datetime

        index = input("Input the ticker of the company you'd like to see the financials of: ")

        headers = { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/71.0.3578.98 Safari/537.36' }

        url_is = 'https://finance.yahoo.com/quote/'+index+'/financials?p='+index
        url_bs = 'https://finance.yahoo.com/quote/'+index+'/balance-sheet?p='+index
        url_cf = 'https://finance.yahoo.com/quote/'+index+'/cash-flow?p='+index
        url_info = 'https://finance.yahoo.com/quote/'+index+'?p='+index

        def get_annual_data(url):
            read_data = requests.get(url,headers=headers, timeout=5)
            content = read_data.content
            soup = BeautifulSoup(content,'lxml')

            data = pd.DataFrame(
                [e.stripped_strings for e in soup.select('[data-test="fin-row"]')],
                columns=soup.select_one('div:has(>[data-test="fin-row"])').previous_sibling.stripped_strings
            )
            return data

        def get_quarterly_data(url):
            options = webdriver.ChromeOptions()
            options = Options()
            options.add_argument("--headless")
            driver = webdriver.Chrome(options=options)
            options.add_argument('headless')

            driver.get(url)

            element = driver.find_element('xpath','//*[@name="agree"]')
            element.click()

            time.sleep(5)

            element = driver.find_element('xpath','//*[@id="myLightboxContainer"]/section/button[1]')
            element.click()

            time.sleep(5)

            element = driver.find_element('xpath','//*[@id="Col1-1-Financials-Proxy"]/section/div[1]/div[2]/button')
            element.click()

            time.sleep(5)

            html = driver.page_source
            soup = BeautifulSoup(html, "html.parser")

            data = pd.DataFrame(
                [e.stripped_strings for e in soup.select('[data-test="fin-row"]')],
                columns=soup.select_one('div:has(>[data-test="fin-row"])').previous_sibling.stripped_strings
            )
            return data

        def stock_info(url):
            read_data_info = requests.get(url_info,headers=headers, timeout=5)
            content_info = read_data_info.content
            soup_info = BeautifulSoup(content_info,'lxml')

            ls = []
            for l in soup_info.find_all('div') and soup_info.find_all('span') and soup_info.find_all('td'):
                ls.append(l.string)
            new_ls = list(filter(None,ls))
            new_ls = new_ls[0:]
            info_data = list(zip(*[iter(new_ls)]*2))
            Stock_Info = pd.DataFrame(info_data[0:])
            return Stock_Info

        def stock_ticker(ticker):
            stock_name = pd.DataFrame({'Ticker': [index.upper()]})
            return stock_name

        Today = datetime.date.today()
        Todays_date_pd = pd.DataFrame({"Date": [Today]})


        Income_Statement_Annual = get_annual_data(url_is)
        Balance_Sheet_Annual = get_annual_data(url_bs)
        Cash_Flow_Annual = get_annual_data(url_cf)
        Income_statement_Quarterly = get_quarterly_data(url_is)
        Balance_sheet_Quarterly = get_quarterly_data(url_bs)
        Cash_Flow_Quarterly = get_quarterly_data(url_cf)
        Stock_Info = stock_info(url_info)
        Stock_Name = stock_ticker(index)

        path = '/Users/apple/Desktop/DCF Models/DCF_Model.xlsx'

        ExcelWorkbook = load_workbook(path)

        writer = pd.ExcelWriter(path, engine='openpyxl')

        writer.book = ExcelWorkbook

        Income_Statement_Annual.to_excel(writer,sheet_name="Annual Income Statement")
        Balance_Sheet_Annual.to_excel(writer,sheet_name="Annual Balance Sheet")
        Cash_Flow_Annual.to_excel(writer,sheet_name="Annual Cash Flow")
        Income_statement_Quarterly.to_excel(writer,sheet_name = "Quarterly Income Statement")
        Balance_sheet_Quarterly.to_excel(writer,sheet_name = "Quarterly Balance Sheet")
        Cash_Flow_Quarterly.to_excel(writer,sheet_name = "Quarterly Annual Cash Flow")
        Stock_Info.to_excel(writer,sheet_name = "Stock Info")
        Stock_Name.to_excel(writer,sheet_name = "Stock Info",startcol=1,startrow=17)
        Todays_date_pd.to_excel(writer,sheet_name = "Stock Info",startcol=1,startrow=19)

        writer.save()
        ExcelWorkbook.save(f'/Users/apple/Desktop/DCF Models/DCF-Model ${index.upper()}.xlsx')

        print(f"DataFrame is exported successfully to {path} Excel File.")
    dcf_yfinance()
