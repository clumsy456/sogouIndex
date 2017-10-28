import requests
from bs4 import BeautifulSoup
from urllib.parse import quote
from random import choice, randint
from userAgents import agents
import re
import json
import xlwt

keyword = '煤改电'
time = 'MONTH' # YEAR, MONTH, WEEK
searchType = 'SEARCH_ALL' # SEARCH_ALL, SEARCH_PC, SEARCH_WAP, MEDIA_WECHAT

url = 'http://zhishu.sogou.com/index/searchHeat?kwdNamesStr=' + \
	quote(keyword, encoding='utf8') + '&timePeriodType=' + time + \
	'&dataType=' + searchType + '&queryType=INPUT'

try:
	session = requests.Session()
	agent = choice(agents)
	headers = {'User-Agent': agent}
	req = session.get(url, headers=headers, timeout=60)
	bsObj = BeautifulSoup(req.content, 'html5lib')
	scriptStr = bsObj.script.get_text()
	data = re.findall(r'root\.SG\.data\s=\s(.+?);', scriptStr)[0]
	wholeData = re.findall(r'root\.SG\.wholedata\s=\s(.+?);', scriptStr)[0]
	print('数据读取成功')
	dataJson = json.loads(data)
	wholeDataJason = json.loads(wholeData)

	dataList=[]
	timeList=[]
	wholeDataList=[]
	wholeTimeList=[]

	workbook = xlwt.Workbook()
	worksheetData = workbook.add_sheet('data')
	worksheetWholeData = workbook.add_sheet('wholeData')

	i = 0
	for dataItem in dataJson['pvList'][0]:
		dataOne = dataItem['pv']
		timeOne = dataItem['date']
		dataList.append(dataOne)
		timeList.append(timeOne)
		worksheetData.write(i, 0, label=dataOne)
		worksheetData.write(i, 1, label=timeOne)
		i = i + 1

	j = 0
	for wholeDataItem in wholeDataJason['pvList'][0]:
		wholeDataOne = wholeDataItem['pv']
		wholeTimeOne = wholeDataItem['date']
		wholeDataList.append(wholeDataOne)
		wholeTimeList.append(wholeTimeOne)
		worksheetWholeData.write(j, 0, label=wholeDataOne)
		worksheetWholeData.write(j, 1, label=wholeTimeOne)
		j = j + 1

	workbook.save(keyword+'-'+time+'-'+searchType+'.xls')
	print('写入成功')
except Exception as e:
	print(e)
	if bsObj.find('div', {'class': 'noresult'}):
		print('未收录关键词')
finally:
	print('结束')