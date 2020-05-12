import requests,json,xlrd
#取得Excel某一行所有列数据，返回dict数据
def xlsx_getRow(sheet, row):
    object = {}
    object["method"] = sheet.cell_value(row, 2)
    object["url"] = sheet.cell_value(row, 3)
    object["params"] = sheet.cell_value(row, 4)
    object["header"] = sheet.cell_value(row, 5)
    object["code"] = sheet.cell_value(row, 6)
    #第八列数据如果是float类型则取到它的int类型数据
    if(type(sheet.cell_value(row, 7)))=='float':
        object["expected_code"] = int(sheet.cell_value(row, 7))
    else:
        object["expected_code"] = sheet.cell_value(row, 7)
    if object["header"] == "":
        object["header"] = None
    return object

def xlsx_request(object):
    try:
        if object["method"]=="post":
            response = requests.post(object["url"], data=datatodict(object["params"]), headers=object["header"])
            result = response.json()
        elif object["method"]=="get":
            response = requests.get(object["url"], params=datatodict(object["params"]), headers=object["header"])
            result = response.json()
        else:
            print("Unknown method" + object["method"])
    except requests.exceptions.ConnectTimeout as e:
        result = {object["code"]: "timeout"}
    print(result)
    return result#返回json格式的数据进行下面的断言


def datatodict(datastr):
    datastr = datastr.replace('{','').replace('}','')#分别吧{和}替换成‘’（空）
    datalist = datastr.split(',')#把参数以，分隔开变为[city:深圳，key:123456]列表
    res={}
    for d in datalist:
        k,v = d.split(':')
        res[k] = v
    return res

def xlsx_set(sheet,row,col,value,red=False):
    sheet.write(row, col, value)

def dosheet(brd, bwt, sheetIndex, cookies=None):
    brd_sheet = brd.sheets()[sheetIndex]
    bwt_sheet = bwt.get_sheet(sheetIndex)
    count = brd_sheet.nrows#获取Excel的行数
    for i in range(1, count):
        object = xlsx_getRow(brd_sheet, i)
        object["cookies"] = cookies
        result = xlsx_request(object)
        if result.get(object["code"]) == (object["expected_code"]):
            xlsx_set(bwt_sheet, i, 8, "pass", False)
            xlsx_set(bwt_sheet, i, 9, result[object["code"]], False)
        else:
            xlsx_set(bwt_sheet, i, 8, "fail", True)
            xlsx_set(bwt_sheet, i, 9, result[object["code"]], False)

def open_file(path):
    brd = xlrd.open_workbook(path)
    return brd




