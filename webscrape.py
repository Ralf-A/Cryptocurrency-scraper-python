from bs4 import BeautifulSoup
import requests
import xlsxwriter
import time
from datetime import datetime
import numpy as np
import pandas as pd

url = "https://coinmarketcap.com/all/views/all/"
def scrape():
    leht = requests.get(url)
    soup = BeautifulSoup(leht.text, "html.parser")   
    tablerow = soup.find_all("tr", attrs={"class":"cmc-table-row"})
    i = 0
    hind={}
    for row in tablerow:
        if i == 20:
            break
        else:
            i += 1
        nimirida = row.find("td", attrs={"class":"cmc-table__cell cmc-table__cell--sticky cmc-table__cell--sortable cmc-table__cell--left cmc-table__cell--sort-by__name"})
        nimi = nimirida.find("a", attrs={"class":"cmc-table__column-name--name cmc-link"}).text.strip()
        cryptohind = row.find("td", attrs={"class":"cmc-table__cell cmc-table__cell--sortable cmc-table__cell--right cmc-table__cell--sort-by__price"}).text.strip().strip("$").replace(",", "")
        hind.update({nimi: cryptohind})
    return hind

rida = 0
tulp = 1
hinnarida = 1
hinnatulp = 1
korduseid = 1
workbook = xlsxwriter.Workbook("andmed.xlsx")
worksheet = workbook.add_worksheet()
worksheet.write(0, 0, "Aeg")

for i in scrape():
    worksheet.write(rida, tulp, i)
    tulp+=1
while(True):
    try:
        rida = 0
        hind = scrape()
        print(hind)
        for i in hind:
            worksheet.write(hinnarida, hinnatulp, hind[i])
            hinnatulp += 1
            time_now = datetime.now()
            worksheet.write(hinnarida, 0, time_now.strftime("%H:%M"))
        time.sleep(15)
        hinnarida += 1
        hinnatulp = 1
        korduseid += 1
    except KeyboardInterrupt:
        workbook.close()
        break

coin_data=pd.read_excel("./andmed.xlsx",sheet_name="Sheet1")
excel_file_path="./graafikud.xlsx"
workbook=xlsxwriter.Workbook(excel_file_path)
coin_worksheet=workbook.add_worksheet()
date_format=workbook.add_format({"num_format": "hh:mm"})
for i,col_name in enumerate(coin_data.columns):
    coin_worksheet.write(0,i,col_name)
    if(i==0):
        coin_worksheet.write_column(1,i,coin_data[col_name],date_format)
    else:
        coin_worksheet.write_column(1, i, coin_data[col_name])
        chart=workbook.add_chart({"type":"scatter","subtype":"straight"})
        col_letter=xlsxwriter.utility.xl_col_to_name(i)
        chart.add_series({"categories":"=Sheet1!$A$2:$A$"+str(2+len(coin_data[col_name]-1)),
                          "values":"=Sheet1!$"+col_letter+"$2:$"+col_letter+"$755",
                          "name": col_name})
        chart.set_x_axis({"name":"Korduseid","min":1,"max":korduseid})
        chart.set_title({"name": col_name})
        chart.set_y_axis({"name":"Hind"})
        coin_worksheet.insert_chart("B"+str(17*(i)),chart)
workbook.close()
    


