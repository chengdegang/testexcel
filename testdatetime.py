from datetime import datetime
import os
from openpyxl.workbook import Workbook
import time

wbc = Workbook()
log_path = os.path.dirname(os.path.abspath('.')) + '/testexcel/file/'
t = time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))
suffix = '.xlsx'  # 文件类型
newfile = t + suffix
path = log_path + t + suffix
wbc.save(path)
print(log_path+newfile)

#/Users/jackrechard/PycharmProjects/testexcel

# date = datetime.strptime("2019-01-01",'%Y-%m-%d')
# print(date)

list1 = []