import urllib2

from lxml import html
import requests
from bs4 import BeautifulSoup
import xlsxwriter
import re
from datetime import datetime
import json
import time




from selenium import webdriver
from bs4 import BeautifulSoup

browser=webdriver.Chrome()
for i in range(1,5):
    url = 'http://ca.louisvuitton.com/eng-ca/women/handbags/to-%s' %i
    browser.get(url)


    soup=BeautifulSoup(browser.page_source, "lxml")
    data_list = soup.find_all("a", {"data-sku":True}, class_="product-item tagClick")

    data = []
    for d in data_list:
        output = {}
        output["Code"] = d["data-sku"]
        product_name = d.find("div",class_="productName toMinimize")
        if product_name is not None:
            output["Name"] = d.find("div",class_="productName toMinimize").text
        product_price = d.find("div",class_="productPrice")
        if product_price is not None:
            output["Price"] = d.find("div",class_="productPrice").text
        data.append(output)

    print data
    time.sleep(5)
browser.quit()






# data = []
# first_output = [] 
# store_list = json_data["stores"]
# for i in store_list:
# 	output = {}
# 	output["Store Number"] = i["store_id"]
# 	soup = BeautifulSoup(i["4"], "lxml")
# 	output["Store Name"] = soup.find("body").find("span", class_="name").text
# 	output["City"] = soup.find("body").find("span", class_="city").text
# 	output["Street Adress"] = soup.find("body").find("span", class_="address").text
# 	output["State"] = soup.find("body").find("span", class_="prov_state").text
# 	output["Zip Code"] = soup.find("body").find("span", class_="postal_zip").text

# 	data.append(output)
# print data 





def write_to_excel(workbook,worksheet,data):
        
        # w = tzwhere.tzwhere()
        bold = workbook.add_format({'bold': True})
        bold_italic = workbook.add_format({'bold': True, 'italic':True})
        border_bold = workbook.add_format({'border':True,'bold':True})
        border_bold_grey = workbook.add_format({'border':True,'bold':True,'bg_color':'#d3d3d3'})
        border = workbook.add_format({'border':True,'bold':True})
        
        #worksheet = workbook.add_worksheet('%s_%s'%(a,j))
        worksheet.set_column('B:D', 22)
        worksheet.set_column('E:F', 33)
        row = 0
        col = 0


        worksheet.write(row,col,'Store List',bold)
        row = row + 1

        row = row + 2

        worksheet.write(row,col,'Sl No',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Name',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Code',border_bold_grey)
        col = col + 1
        worksheet.write(row,col,'Price',border_bold_grey)
       
    

        row = row + 1
        i = 0


        """{'City': u' HOMER', 'Store Name': u' TECH CONNECT, INC', 
        'Zip Code': u' 99603', 'Street Adress': u' 432 EAST PIONEER AVE #C', 
        'State': u' AK', 'Store Number': u'2766977'}"""

        for output in data:
                
            i = i + 1
            col = 0
            worksheet.write(row, col, i, border)
            col = col + 1
            worksheet.write(row, col, output["Name"] if output.has_key('Name') else '',border)
            col = col + 1
            worksheet.write(row, col, output["Code"] if output.has_key('Code') else '',border)
            col = col + 1
            worksheet.write(row, col, output["Price"] if output.has_key('Price') else '',border)

            col = col + 1
            row = row + 1

workbook = xlsxwriter.Workbook('store.xlsx')
worksheet = workbook.add_worksheet('Hand Bags')
write_to_excel(workbook,worksheet,data)
workbook.close()
