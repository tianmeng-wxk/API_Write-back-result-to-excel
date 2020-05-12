import requests
import xlrd

file = xlrd.open_workbook("号码归属地和天气查询.xlsx")


print("get请求")
url = file.sheet_by_index(0).cell_value(1, 3)
data = {"phone": "13429667914", "key": "ffe950dd3e85ee36de34b33575d1e177"}
respone1 = requests.get(url=url, params=data)
print(respone1.text)
print("post请求")
url2 = 'http://apis.juhe.cn/simpleWeather/query'
data2 = {"city": "深圳", "key": "371e78e1c2d2eeabb361294723857e8d"}
respone2 = requests.post(url=url2,data=data2)
print(respone2.text)
