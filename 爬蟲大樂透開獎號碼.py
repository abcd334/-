from selenium import webdriver
from datetime import datetime
from bs4 import BeautifulSoup
import openpyxl

wb = openpyxl.load_workbook('C:\\Users\Arlen\Desktop\GitHub\win-the-lottery\大樂透開獎號碼-新版.xlsx')
ws = wb['工作表1']
k=3

date_format = "%y/%m/%d"

driver = webdriver.Chrome("./goolemapSpider/chromedriver.exe")
driver.get('https://www.taiwanlottery.com.tw/lotto/Lotto649/history.aspx')

'''
time.sleep(2)
flag=driver.find_element("name",'Lotto649Control_history$txtNO')

flag.send_keys("103000001")
time.sleep(1)

driver.find_element("id",'Lotto649Control_history_btnSubmit').click()
time.sleep(2)
'''
soup = BeautifulSoup(driver.page_source, "lxml")
#tables = soup.find_all("table", {"class": "table_org td_hm"})

DrawTerm=soup.find("span", {"id": "Lotto649Control_history_dlQuery_L649_DrawTerm_0"}).text
DDate=soup.find("span", {"id": "Lotto649Control_history_dlQuery_L649_DDate_0"}).text
SellAmount=soup.find("span", {"id": "Lotto649Control_history_dlQuery_L649_SellAmount_0"}).text
TotalAmount=soup.find("span", {"id": "Lotto649Control_history_dlQuery_Total_0"}).text
ws.cell(k,1,int(DrawTerm))
ws.cell(k,2,DDate.replace(year=DDate.year+1911))
ws.cell(k,3,int(SellAmount.replace(",", "")))
ws.cell(k,4,int(TotalAmount.replace(",", "")))
print(DDate.replace(year=DDate.year+1911))

for i in range(1,7):
    locals()['SNo'+str(i)]=soup.find("span", {"id": "Lotto649Control_history_dlQuery_SNo" + str(i) +"_0"}).text
    ws.cell(k, 4+i,int(locals()['SNo'+str(i)]))

SNo7=soup.find("span", {"id": "Lotto649Control_history_dlQuery_No7_0"}).text
ws.cell(k, 11,int(SNo7))

for i in 'ABC':
    locals()['Categ'+str(i)]=soup.find("span", {"id": "Lotto649Control_history_dlQuery_L649_Categ" + str(i) +"3_0"}).text
ws.cell(k, 12,int(CategA.replace(",", "")))
ws.cell(k, 13,int(CategB.replace(",", "")))
ws.cell(k, 14,int(CategC.replace(",", "")))

for i in range(2,7):
    locals()['label'+str(i)]=soup.find("span", {"id": "Lotto649Control_history_dlQuery_Label" + str(i) +"_0"}).text
    ws.cell(k, 13+i,int(locals()['label'+str(i)].replace(",", "")))

categA_1=soup.find_all("span", {"id": "Lotto649Control_history_dlQuery_L649_CategA4_8"})
categB_1=soup.find_all("span", {"id": "Lotto649Control_history_dlQuery_Label7_8"})
categC_1=soup.find_all("span", {"id": "Lotto649Control_history_dlQuery_Label8_8"})
label2_1=soup.find_all("span", {"id": "Lotto649Control_history_dlQuery_Label9_8"})
label3_1=soup.find_all("span", {"id": "Lotto649Control_history_dlQuery_Label10_8"})
label4_1=soup.find_all("span", {"id": "Lotto649Control_history_dlQuery_Label11_8"})
label5_1=soup.find_all("span", {"id": "Lotto649Control_history_dlQuery_Label12_8"})
label6_1=soup.find_all("span", {"id": "Lotto649Control_history_dlQuery_Label13_8"})

categA_2=soup.find_all("span", {"id": "Lotto649Control_history_dlQuery_L649_CategA5_8"})
categB_2=soup.find_all("span", {"id": "Lotto649Control_history_dlQuery_Label14_8"})
categC_2=soup.find_all("span", {"id": "Lotto649Control_history_dlQuery_Label15_8"})
label2_2=soup.find_all("span", {"id": "Lotto649Control_history_dlQuery_Label16_8"})



'''
for j in range(1, 32):
    ws.cell(k, j, int(table.cell(i, 8 + j).text))
'''
k += 1
wb.save('大樂透開獎號碼-新版.xlsx')
#tables=pd.read_html(chrome.page_source)
#print(DrawTerm,DDate,SellAmount,TotalAmount)
#print(SNo1,SNo2,SNo3,SNo4,SNo5,SNo7)
#print(CategA,CategB,CategC,label2,label3,label4,label5,label6)
#print(tables)