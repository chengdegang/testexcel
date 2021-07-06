import openpyxl
import xlrd
# import win32com.client as win32
import time
from openpyxl.cell import MergedCell
from openpyxl.worksheet.worksheet import Worksheet

file1 = ''
file_xxx = '/Users/jackrechard/PycharmProjects/testexcel/file/xxx.xls'
file_xxx2 = '/Users/jackrechard/PycharmProjects/testexcel/file/xxx2.xlsx'

"""
筛选数据并写到列表里,三重判断，分别判断列指标1："value"，行指标1："_1"，行指标2："_2"
支持的excel格式为xls
"""
def xlrd_read_xls(file):
    wb = xlrd.open_workbook(file)
    table = wb.sheet_by_name('Sheet1')
    #获取行数列数
    nrow = table.nrows
    ncol = table.ncols
    need_data =[]
    need_data2 = []
    #获取标题
    for row in range(nrow):
        for col in range(ncol):
            if "title" in str(table.cell(row, col).value):
                need_data.append(table.row_values(row))
                break
    #获取需要的行，判断条件为包含'_1'
    for row in range(nrow):
        for col in range(ncol):
            #判断是否包含'value'，是的话选取整列
            if "value" in str(table.cell(row,col).value):
                #判断这列是否包含'_1'，是的话取整行
                # print(table.col_values(col))
                #遍历这一列的所有数据，并标记存在'_1'的行
                for i in range(len(table.col_values(col))):
                    if "_1" in str(table.cell(i,col).value):
                        # print("存在")
                        need_data.append(table.row_values(i))

    # 获取需要的行，判断条件为包含'_2'
    # for i in range(len(need_data)):
    #     for j in range(len(need_data[0])):
    #         if "_2" in need_data[i][j]:
    #             need_data2.append(need_data[i])
    print(need_data)
    return need_data

"""
支持xlsx格式的读,输入要读的文件路径
"""
def openpy_read_xlsx(file):
    wb = openpyxl.load_workbook(file)
    table = wb['Sheet2']
    ws = wb.active
    #获取行数列数
    nrow = table.max_row
    ncol = table.max_column
    need_data2 = []

    #获取标题
    for row in range(1,nrow+1):
        for col in range(1,ncol+1):
            #如果这一行中存在title字样，遍历这一行的所有数据，添加到列表中
            if "title" in str(table.cell(row, col).value):
                for i in range(1,ncol+1):
                    # print(table.cell(row, i).value)
                    need_data2.append(table.cell(row, i).value)
                # print(need_data2)
                break

    #获取需要的行，判断条件为包含'_1'
    for row in range(1,nrow+1):
        for col in range(1,ncol+1):
            #判断是否包含'value'，是的话选取整列
            if "value" in str(table.cell(row,col).value):
                #判断这列是否包含'_1'，是的话取整行
                # print(table.cell(row, col).value)
                #遍历这一列的所有数据，并标记存在'_1'的行
                for i in range(1,nrow+1):
                    if "_1" in str(table.cell(i,col).value):
                        # print(table.cell(i,col).value)
                        #将找到的这一行所有数据写入列表
                        for j in range(1,ncol+1):
                            # print(table.cell(i,j).value)
                            need_data2.append(table.cell(i,j).value)

    #将改列表按照列数切成多个列表
    per_list_len = ncol
    list_of_group = zip(*(iter(need_data2),) * per_list_len)
    end_list = [list(i) for i in list_of_group]  # i is a tuple
    count = len(need_data2) % per_list_len
    end_list.append(need_data2[-count:]) if count != 0 else end_list
    print(end_list)
    return end_list

"""
写入均只支持xlsx，两种写法，用第二种即可
"""
#写法1，列为手动写的，有长度限制
# data = [['title1', 'title2', 'title_value3'], ['数据4', '数据5', '数据6_1'], ['数据10', '数据11', '数据12_1']]
def write_excel_(file,data):
    #内存中创建一个空表格
    # wb = openpyxl.Workbook()
    # sheet = wb.active
    # sheet.title = 'ccc_sheet'
    # print(wb.sheetnames)

    #写到原来的表格中，新建一个sheet,只支持xlsx文件
    wb = openpyxl.load_workbook(file)
    wb.create_sheet(index=0,title='ccc_sheet')
    sheet = wb['ccc_sheet']
    abc = "0ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    for i in range(len(data)):
        for j in range(len(data[i])):
            sheet[abc[j+1]+str(i+1)] = data[i][j]
            # sheet.cell(row=i+1,column=j+1,value=data[i][j])
            print(f'字母是{abc[j+1]} 数字是{i+1} 数据是{data[i][j]}')
    wb.save(file)

data2 = [['John Brown', 18, 'New York No. 1 Lake Park'],['John Brown2', 11, 'New York No. 1 Lake Park2']]
#写法2，较好的写法
def write_excel(file,data):
    wb = openpyxl.load_workbook(file)
    #取一个时间戳作为sheet的名字
    tt = '0705'
    #默认写到第一个sheet中，index为0
    wb.create_sheet(index=0,title=f'ccc_sheet_{tt}')
    sheet = wb[f'ccc_sheet_{tt}']
    #好的写法写入excel
    for row_index, row_item in enumerate(data):
        for col_index, col_item in enumerate(row_item):
            sheet.cell(row=row_index + 1, column=col_index + 1, value=col_item)
    wb.save(file)
    print('---写入完成---')

def merged_deal_xlsx(file):
    """
    处理xlsx的excel,获取文件路径，重读excel文件并输出每一个单元格重写了数据的列表（主要针对有合并单元格的情况）
    """
    newdata = []
    wb = openpyxl.load_workbook(file)
    table = wb["Sheet2"]
    nrow = table.max_row
    ncol = table.max_column

    for row_index in range(1, nrow + 1):
        for col_index in range(1, ncol + 1):
            cell = table.cell(row=row_index, column=col_index)
            if isinstance(cell, MergedCell):  # 判断该单元格是否为合并单元格
                for merged_range in table.merged_cell_ranges:  # 循环查找该单元格所属的合并区域
                    if cell.coordinate in merged_range:
                        # 获取合并区域左上角的单元格作为该单元格的值返回
                        cell_ = table.cell(row=merged_range.min_row, column=merged_range.min_col)
                        newdata.append(cell_.value)
                        break
            else:
                cell_ = table.cell(row=row_index, column=col_index)
                # newdata.append(cell_.value)
                newdata.append(cell_.value)

    per_list_len = ncol
    list_of_group = zip(*(iter(newdata),) * per_list_len)
    end_list = [list(i) for i in list_of_group]  # i is a tuple
    count = len(newdata) % per_list_len
    end_list.append(newdata[-count:]) if count != 0 else end_list
    print(end_list)
    return end_list

#转换功能先忽略
# def xls_to_xlsx():
#     fname = "/Users/jackrechard/PycharmProjects/testexcel/file/xxx.xls"
#     excel = win32.gencache.EnsureDispatch('Excel.Application')
#     wb = excel.Workbooks.Open(fname)
#
#     wb.SaveAs(fname + "x", FileFormat=51)  # FileFormat = 51 is for .xlsx extension
#     wb.Close()  # FileFormat = 56 is for .xls extension
#     excel.Application.Quit()

#读xls的excel
# xlrd_read_xls(file_xxx)
#读xlsx的excel
# openpy_read_xlsx(file_xxx2)
#将xls筛选的数据写入到excel中
# write_excel_(file_xxx2,xlrd_read_xls(file_xxx))
#将xlsx取出的数据写入excel
# write_excel(file_xxx2,openpy_read_xlsx(file_xxx2))

#读合并单元格
# hb_excel(file_xxx2)

if __name__ == "__main__":
    merged_deal_xlsx(file_xxx2)
