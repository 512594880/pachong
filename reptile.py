import asyncio

import aiohttp

import time

import os 

import xlwt
import xlrd
from xlutils.copy import copy


from bs4 import BeautifulSoup
# valueArray = []

sheetsize = 0
row = 1 
excelName = '/Users/anne/Desktop/'+"药品注册与受理" +'.xls'
# path = os.path.join(excelName)
# wbk = xlrd.open_workbook(excelName,formatting_info=True)
# wbknew = copy(wbk)
# wbknew.encoding = 'utf-8'
# wbk = wbknew
wbk = xlwt.Workbook(encoding = 'utf-8')
sheet = wbk.add_sheet('sheet 1')
def saveExcel(str,row,colum):
	global sheetsize
	# if row >= 50000:
	# 	sheetsize +=1
	sheet = wbk.get_sheet(sheetsize)
		# row = 1
	# else:
	# 	sheet = wbk.get_sheet(sheetsize)
    
    
	sheet.write(row, colum,  str)
	return

def saveInExcel(name,str,row,colum):


    excelName = '/Users/anne/Desktop/'+name +'.xls'
    path = os.path.join(excelName)
    exists = os.path.exists(path)

    if exists:
        wbk = xlrd.open_workbook(excelName,formatting_info=True)
        wbknew = copy(wbk)
        wbknew.encoding = 'utf-8'
        wbk = wbknew
        sheet = wbk.get_sheet(0)
    else:
        wbk = xlwt.Workbook(encoding = 'utf-8')
        sheet = wbk.add_sheet('sheet 1')
        sheet = wbk.get_sheet(0)
    sheet.write(row, colum,  str)
    
    return


start = time.time()

def handlResult(result):
	print("---------")
	colum = 0
	global row
	soup = BeautifulSoup(result)
	liResutl = soup.findAll('tbody')
	for li in liResutl:
		trlist = li.findAll('tr')
		for tr in trlist:
			# valuelist = []
			tdlist = tr.findAll('td')
			thlist = tr.findAll('th')
			for th in thlist:
				value = ""
				if 'href' in th:	
					name = th.find('a',attrs = {"class" : "cl-blue"})
					value= name.get_text()
				else:
					value = th.get_text()
				# valuelist.append(value)
				saveExcel(value,row,colum)
				colum = colum +1
			for td in tdlist: 
				value = ""
				if 'href' in td:	
					name = td.find('a',attrs = {"class" : "cl-blue"})
					value= name.get_text()
				else:
					value = td.get_text()
				# valuelist.append(value)
				saveExcel(value,row,colum)
				colum = colum +1
			colum = 0
			row = row +1
			# valueArray.append(valuelist)
			print("----" + str(row) +"--------")

async def get(url):

	headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36'}

	session = aiohttp.ClientSession(headers=headers)

	response = await session.get(url)

	result = await response.text()

	session.close()

	return result

async def requestBynum(i):

	url = 'https://db.yaozh.com/zhuce?p='+str(i+1)+'&pageSize=20'

	print('Waiting for',url)

	result = await get(url)

	handlResult(result)

	# print( 'Get response from', url, 'Result:', result)
# urlname = ["xinyao","zhuangrang","linchuangshiyan","inn","yaozhreport"]
# urlPageSize = [10,3,10,3,129]

# for k in range(10):

#     new_thread = threading.Thread(target=booth,args=(k,))
#     new_thread.start()
shuzu = range(67,100)




# for y in range(len(shuzu)-1):
# 	tasks = [asyncio.ensure_future(requestBynum(i)) for i in range(shuzu[y]*100,shuzu[y+1]*100)]


# 	loop = asyncio.get_event_loop()

# 	loop.run_until_complete(asyncio.wait(tasks))

# 	wbk.save(excelName)
for y in range (10):
	tasks = [asyncio.ensure_future(requestBynum(i)) for i in range(y,y+1)]
	# tasks = asyncio.ensure_future(requestBynum(4))
	loop = asyncio.get_event_loop()

	loop.run_until_complete(asyncio.wait(tasks))

	wbk.save(excelName)

end = time.time()

print( 'Cost time:', end - start)