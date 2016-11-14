from datetime import datetime
import win32com.client
import win32api
from win32con import MB_SYSTEMMODAL
import re
from time import sleep
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from statistics import median

now = datetime.now()

c = 2

Excel = win32com.client.Dispatch("Excel.Application")

wb = Excel.Workbooks.Add()
ws = wb.ActiveSheet

browser = webdriver.Firefox()

brand = 'Быструмгель'

for regcode in range(1,99):
    
    prices_30 = []
    prices_50 = []
    prices_100 = []
    
    browser.get('http://poisklekarstv.ru/?region=' + str(regcode))

    lekElem = browser.find_element_by_name('lek_name')
    lekElem.send_keys(brand)
    submitBtn = browser.find_element_by_class_name('sub')
    submitBtn.click()

    html = browser.page_source
    soup = BeautifulSoup(html, 'html.parser')
    regname = re.search('(<a href="#" id="click-elem">)(.+)(<i></i></a>)', str(soup.find("span", class_ = 'tptCity').a)).group(2)
    if regname == 'не выбран':
        continue

    try:
        warningElem = browser.find_element_by_class_name('warning')
        if re.match('В настоящее время по данному запросу информация отсутствует.*', warningElem.text):
            for i in range(3):
                ws.Cells(c+i,1).Value = str(now)[:11]
                ws.Cells(c+i,3).Value = regname
                
            ws.Cells(c,2).Value = brand + ' 2,5% 30г'
            ws.Cells(c+1,2).Value = brand + ' 2,5% 50г'
            ws.Cells(c+2,2).Value = brand + ' 2,5% 100г'

            ws.Cells(c,4).Value = 'NA'
            ws.Cells(c+1,4).Value = 'NA'
            ws.Cells(c+2,4).Value = 'NA'

            c += 3
            continue
    except:
        pass
        
    try:
        showAllElem = browser.find_element_by_link_text('все результаты')
        showAllElem.click()
    except:
        pass

    try:
        element = WebDriverWait(browser, 10).until(EC.presence_of_element_located((By.ID, "example_next")))
        nextBtn = browser.find_element_by_link_text('следующая')
    except:
        sleep(1)
        try:
            warningElem = browser.find_element_by_class_name('warning')
            if re.match('В настоящее время по данному запросу информация отсутствует.*', warningElem.text):
                for i in range(3):
                    ws.Cells(c+i,1).Value = str(now)[:11]
                    ws.Cells(c+i,3).Value = regname
                    
                ws.Cells(c,2).Value = brand + ' 2,5% 30г'
                ws.Cells(c+1,2).Value = brand + ' 2,5% 50г'
                ws.Cells(c+2,2).Value = brand + ' 2,5% 100г'

                ws.Cells(c,4).Value = 'NA'
                ws.Cells(c+1,4).Value = 'NA'
                ws.Cells(c+2,4).Value = 'NA'

                c += 3
                continue
        except:
            pass

    while True:
        
        html = browser.page_source
        soup = BeautifulSoup(html, 'html.parser')

        for row in soup.find_all("tr", class_ = 'odd'):
            price = row.contents[-3].string
            if price != None:
                if re.search('100', str(row.contents[0].b)) != None:
                    prices_100.append(float(price))
                elif re.search('50', str(row.contents[0].b)) != None:
                    prices_50.append(float(price))
                elif re.search('30', str(row.contents[0].b)) != None:
                    prices_30.append(float(price))

        for row in soup.find_all("tr", class_ = 'even'):
            price = row.contents[-3].string
            if price != None:
                if re.search('100', str(row.contents[0].b)) != None:
                    prices_100.append(float(price))
                elif re.search('50', str(row.contents[0].b)) != None:
                    prices_50.append(float(price))
                elif re.search('30', str(row.contents[0].b)) != None:
                    prices_30.append(float(price))

        if nextBtn.get_attribute("class") == 'next paginate_button paginate_button_disabled':
            break

        nextBtn.click()
        
        nextBtn = browser.find_element_by_link_text('следующая')
                
    if prices_30 == []:
        median_price_30 = 'NA'
    else:
        median_price_30 = median(prices_30)

    if prices_50 == []:
        median_price_50 = 'NA'
    else:
        median_price_50 = median(prices_50)

    if prices_100 == []:
        median_price_100 = 'NA'
    else:
        median_price_100 = median(prices_100)

    for i in range(3):
        ws.Cells(c+i,1).Value = str(now)[:11]
        ws.Cells(c+i,3).Value = regname
        
    ws.Cells(c,2).Value = brand + ' 2,5% 30г'
    ws.Cells(c+1,2).Value = brand + ' 2,5% 50г'
    ws.Cells(c+2,2).Value = brand + ' 2,5% 100г'

    ws.Cells(c,4).Value = median_price_30
    ws.Cells(c+1,4).Value = median_price_50
    ws.Cells(c+2,4).Value = median_price_100

    c += 3

ws.Name = 'Data'
ws.Cells(1,1).Value = 'Date'
ws.Cells(1,2).Value = 'SKU'
ws.Cells(1,3).Value = 'Region'
ws.Cells(1,4).Value = 'Median price'

for i in range(1,5):
    ws.Cells(1,i).Font.Bold = True
    ws.Columns(i).ColumnWidth = 30

endtime = datetime.now()
rtime = endtime - now

browser.quit()

Excel.Visible = True
response = win32api.MessageBox(0, "Script completed successfully.\nElapsed time: " + str(int(rtime.total_seconds() // 60)) + " min " + str(int(rtime.total_seconds() % 60)) + " sec", "Python 3", MB_SYSTEMMODAL)
