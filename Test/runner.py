from Test.keywork import dosheet, open_file
from xlutils.copy import copy
import xlrd
brd = open_file("号码归属地和天气查询.xlsx")
# brd = xlrd.open_workbook("号码归属地和天气查询.xlsx")
bwt = copy(brd)
dosheet(brd, bwt, 0)

bwt.save("E:\\关键字驱动\\Test\\新号码归属地和天气查询.xlsx")#保存为csv格式文件在此路径会出现乱码情况

