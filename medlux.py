from datetime import datetime
import re
import win32com.client
import win32api
from win32con import MB_SYSTEMMODAL
from urllib.parse   import quote
from urllib.request import urlopen
from bs4 import BeautifulSoup

now = datetime.now()

Excel = win32com.client.Dispatch("Excel.Application")
wb = Excel.Workbooks.Add()
ws = wb.ActiveSheet

brands_e = ['Пенталгин', 'Нурофен', 'Линекс']
codes_e = []
total_ph = {}
for b in brands_e:
    page = urlopen('http://www.medlux.ru/' + quote("лекарства/" + b))
    soup = BeautifulSoup(page.read(), 'html.parser')
    subs = soup.find_all("ul", "sublist")
    for sub in subs:
        for i in sub:
            if not i.has_attr('class'):
                a = i['id']
                code = re.search('(g_id)(.+)', a)
                codes_e.append(code.group(2))

for code in codes_e:
    page = urlopen('http://www.medlux.ru/?_grp[]=' + code)
    soup = BeautifulSoup(page.read(), 'html.parser')
    storedata = soup.find_all("td", class_ = 'storedata')
    for i in storedata:
        data = re.search('(/?store=)([0-9]+)(&drugs=)(.+)', i.find('a')['href'])
        total_ph[data.group(2)] = i.parent['region']

dt = {}
for code in total_ph:
    dt[total_ph[code]] = dt.get(total_ph[code], 0) + 1

c = 1
    
brands = ['Капотен', 'Акридерм', 'Галазолин', 'Диакарб', 'Сопелка', 'Тригрим', 'Азелик', 'Фитолизин', 'Менорил', 'Овипол Клио', 'Пантодерм',
          'Гипосарт', 'Нолодатак', 'Эспиро', 'Максигра', 'Баклосан', 'Сумамигрен', 'Нормобакт', 'Венолайф', 'Joyskin', 'Клиндацин', 'Быструмгель', 'Быструмкапс', 'Сиресп', 'Аквадетрим', 'Боботик', 'Фастум', 'Вольтарен Эмульгель', 'Кетонал']
for brand in brands:
    c_b = c+1
    d = {}
    total_ph = {}
    dbt = {}
    page = urlopen('http://www.medlux.ru/' + quote("лекарства/" + brand))
    soup = BeautifulSoup(page.read(), 'html.parser')
    subs = soup.find_all("ul", "sublist")
    for sub in subs:
        for i in sub:
            if not i.has_attr('class'):
                a = i['title']
                b = i['id']
                code = re.search('(g_id)(.+)', b)
                title = re.search('(<span>)(.+)(</span><span>)(.+)(</span>)', a) 
                d[title.group(2)+' '+title.group(4)] = code.group(2)
    for sku, code in sorted(d.items()):
        page = urlopen('http://www.medlux.ru/?_grp[]=' + code)
        soup = BeautifulSoup(page.read(), 'html.parser')
        stores = soup.find_all("tr", class_ = re.compile('^stores_list'))
        storedata = soup.find_all("td", class_ = 'storedata')
        dc = {}
        dqty = {}
        dvalue = {}
        ddvalue = {}
        for i in storedata:
            data = re.search('(/?store=)([0-9]+)(&drugs=)(.+)', i.find('a')['href'])
            total_ph[data.group(2)] = i.parent['region']
        for store in stores:
            if re.match('([0-9]+)(;)([0-9]+)', store['pres_drugs_quant'].replace(" ","")):
                qs = re.search('([0-9]+)(;)([0-9]+)', store['pres_drugs_quant'].replace(" ",""))
                ps = re.search('(.+)(р\.)(;)(.+)(р\.)', store['pres_drugs_price'].replace(" ",""))
                dps = re.search('(.+)(р\.)(;)(.+)(р\.)', store['pres_drugs_discount_price'].replace(" ",""))
                q1, q2 = int(qs.group(1)), int(qs.group(3))
                qty_add = q1 + q2
                value_add = (float(ps.group(1).replace(",",".").replace(" ",""))*q1 + float(ps.group(4).replace(",",".").replace(" ",""))*q2)
                dvalue_add = (float(dps.group(1).replace(",",".").replace(" ",""))*q1 + float(dps.group(4).replace(",",".").replace(" ",""))*q2)
            else:
                qty_add = int(float(store['pres_drugs_quant']))
                value_add = float(store['pres_drugs_price'].replace(",",".").replace(" ","")[:-2])*qty_add
                dvalue_add = float(store['pres_drugs_discount_price'].replace(",",".").replace(" ","")[:-2])*qty_add
            dc[store['region']] = dc.get(store['region'], 0) + 1
            dqty[store['region']] = dqty.get(store['region'], 0) + qty_add
            dvalue[store['region']] = dvalue.get(store['region'], 0) + value_add
            ddvalue[store['region']] = ddvalue.get(store['region'], 0) + dvalue_add
        for reg in dt:
            c += 1
            ws.Cells(c,1).Value = str(now)[:11]
            ws.Cells(c,2).Value = brand
            ws.Cells(c,3).Value = sku
            ws.Cells(c,4).Value = reg
            ws.Cells(c,5).Value = dc.get(reg, 0)
            ws.Cells(c,6).Value = dt[reg]
            ws.Cells(c,7).Value = dqty.get(reg, 0)
            ws.Cells(c,8).Value = round(dvalue.get(reg, 0), 2)
            ws.Cells(c,9).Value = round(ddvalue.get(reg, 0), 2)
    for code in total_ph:
        dbt[total_ph[code]] = dbt.get(total_ph[code], 0) + 1
    j = 0
    for reg in dt:
        j += 1
        ws.Cells(c+j, 1).Value = str(now)[:11]
        ws.Cells(c+j, 2).Value = brand
        ws.Cells(c+j, 3).Value = brand + ' Total'
        ws.Cells(c+j, 4).Value = reg
        ws.Cells(c+j, 5).Value = dbt.get(reg, 0)
        ws.Cells(c+j, 6).Value = dt[reg]
        ws.Cells(c+j, 7).Value = 0
        ws.Cells(c+j, 8).Value = 0
        ws.Cells(c+j, 9).Value = 0
        for i in range(c_b, c+1):
            if ws.Cells(i,4).Value == ws.Cells(c+j, 4).Value:
                ws.Cells(c+j, 7).Value += ws.Cells(i, 7).Value
                ws.Cells(c+j, 8).Value += ws.Cells(i, 8).Value
                ws.Cells(c+j, 9).Value += ws.Cells(i, 9).Value
    c += j
        

ws.Name = str(now)[:11]
ws.Cells(1,1).Value = 'Date'
ws.Cells(1,2).Value = 'Brand'
ws.Cells(1,3).Value = 'SKU'
ws.Cells(1,4).Value = 'Territory'
ws.Cells(1,5).Value = 'Pharmacies, distribution'
ws.Cells(1,6).Value = 'Pharmacies, total number'
ws.Cells(1,7).Value = 'Sum of Units'
ws.Cells(1,8).Value = 'Sum of Value Ret'
ws.Cells(1,9).Value = 'Sum of Value Ret Disc'

for i in range(1,10):
    ws.Cells(1,i).Font.Bold = True
    ws.Columns(i).ColumnWidth = 12
ws.Range("A1:I1").WrapText = True

endtime = datetime.now()
rtime = endtime - now

Excel.Visible = True
response = win32api.MessageBox(0, "Script completed successfully.\nElapsed time: " + str(int(rtime.total_seconds() // 60))
                               + " min " + str(int(rtime.total_seconds() % 60)) + " sec", "Python 3", MB_SYSTEMMODAL)
