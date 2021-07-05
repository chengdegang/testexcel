import openpyxl
import xlrd
# import win32com.client as win32

file1 = ''
file2 = '/Users/jackrechard/PycharmProjects/testexcel/file/xxx.xls'
file3 = '/Users/jackrechard/PycharmProjects/testexcel/file/xxx2.xlsx'

"""
筛选数据并写到列表里,三重判断，分别判断列指标1："value"，行指标1："_1"，行指标2："_2"
支持的excel格式为xls
"""
def xlrd_read_xls():
    wb = xlrd.open_workbook(file2)
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
支持xlsx格式的读
"""
def openpy_read_xlsx():
    wb = openpyxl.load_workbook(file3)
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
    per_list_len = 12
    list_of_group = zip(*(iter(need_data2),) * per_list_len)
    end_list = [list(i) for i in list_of_group]  # i is a tuple
    count = len(need_data2) % per_list_len
    end_list.append(need_data2[-count:]) if count != 0 else end_list
    print(end_list)
    return end_list

#写法1
# data = [['title1', 'title2', 'title_value3'], ['数据4', '数据5', '数据6_1'], ['数据10', '数据11', '数据12_1']]
def write_excel_(data):
    #内存中创建一个空表格
    # wb = openpyxl.Workbook()
    # sheet = wb.active
    # sheet.title = 'ccc_sheet'
    # print(wb.sheetnames)

    #写到原来的表格中，新建一个sheet,只支持xlsx文件
    wb = openpyxl.load_workbook(file3)
    wb.create_sheet(index=0,title='ccc_sheet')
    sheet = wb['ccc_sheet']
    abc = "0ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    for i in range(len(data)):
        for j in range(len(data[i])):
            sheet[abc[j+1]+str(i+1)] = data[i][j]
            # sheet.cell(row=i+1,column=j+1,value=data[i][j])
            print(f'字母是{abc[j+1]} 数字是{i+1} 数据是{data[i][j]}')
    wb.save(file3)

data2 = [['John Brown', 18, 'New York No. 1 Lake Park'],['John Brown2', 11, 'New York No. 1 Lake Park2']]
#写法2
def write_excel():
    wb = openpyxl.load_workbook(file3)
    wb.create_sheet(index=0,title='ccc_sheet_222')
    sheet = wb['ccc_sheet_222']
    #好的写法写入excel
    for row_index, row_item in enumerate(data2):
        for col_index, col_item in enumerate(row_item):
            sheet.cell(row=row_index + 1, column=col_index + 1, value=col_item)
    wb.save(file3)

#转换功能先忽略
# def xls_to_xlsx():
#     fname = "/Users/jackrechard/PycharmProjects/testexcel/file/xxx.xls"
#     excel = win32.gencache.EnsureDispatch('Excel.Application')
#     wb = excel.Workbooks.Open(fname)
#
#     wb.SaveAs(fname + "x", FileFormat=51)  # FileFormat = 51 is for .xlsx extension
#     wb.Close()  # FileFormat = 56 is for .xls extension
#     excel.Application.Quit()

# xlrd_read_xls()
#将筛选的数据写入到excel中
# write_excel_(xlrd_read())
# xls_to_xlsx()
# exc_two()
openpy_read_xlsx()
