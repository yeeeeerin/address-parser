# -*- coding: utf-8 -*-
import requests
import json
from openpyxl import load_workbook
import csv

API_HOST = 'http://dapi.kakao.com/v2/local/search/address.json'
headers = {'Authorization': 'KakaoAK 0199154d7a2cce1682f40ad58d85a5e0'}

#####엑셀#######
FROM = 2
TO = 5862
NAME = 'H'
ADDRESS = 'N'
###############

def req(address, method, data={}):
    url = API_HOST +'?query='+address;

    if method == 'GET':
        return requests.get(url, headers=headers)
    else:
        return requests.post(url, headers=headers, data=data)

def toCSV(mnet_list):
    file = open('./schoolList.csv', 'wt', encoding='utf-8', newline='')
    csvfile = csv.writer(file)
    for row in mnet_list :
        csvfile.writerow(row)
    file.close()

load_wb = load_workbook("./schoolList.xlsx", data_only=True)
load_ws = load_wb['middle-high-school']

schoolList = []
for num in range(FROM,To):
	name = load_ws[NAME+str(num)].value
	address = load_ws[ADDRESS+str(num)].value
	resp = req(address, 'GET')
	if(resp.status_code == 200):
		parser = (json.loads(resp.text))
		if(parser["documents"]):
			data = [name, parser["documents"][0]["address"]["x"] ,parser["documents"][0]["address"]["y"] ]
			schoolList.append(data);
			print(str(num) + " ["+data[0]+" ," + data[1] +" ,"+ data[2] + "]")
	

toCSV(schoolList)

# resp = req('서울특별시 종로구 필운대로1길 34 배화여자중학교', 'GET')

# parser = (json.loads(resp.text))
# data = parser["documents"][0]["address"]["address_name"] + "," + parser["documents"][0]["address"]["x"] +","+parser["documents"][0]["address"]["y"] 

# print(data)
	
# print("response status:\n%d" % resp.status_code)
# print("response headers:\n%s" % resp.headers)






