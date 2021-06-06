import openpyxl
import xlrd

file1 = ''
file2 = '/Users/jackrechard/PycharmProjects/testexcel/file/xxx.xls'
file3 = '/Users/jackrechard/PycharmProjects/testexcel/file/xxx2.xlsx'

def openpyxl_read():
    # print(type(wb))
    # print(wb.sheetnames)
    # #获取excel中的活动表
    # k_sheet = wb.active
    # print(k_sheet)
    wb = openpyxl.load_workbook(file1)
    sheet = wb['sheet1']
    # 获取指定位置的数据
    # print(sheet['B4'].value)
    # 获取指定行列的值
    # print(sheet.cell(row=3,column=13).value)
    # 获取excel中的行数，列数
    # print(sheet.max_row)
    # print(sheet.max_column)

    # 遍历专业列的所有数据
    # for row in range(2,sheet.max_row+1):
    #     zhuanye = sheet['M'+str(row)].value
    #     print(zhuanye)

    # 不可用，需要自增的是字母不是数字
    # for column in range(3,sheet.max_column+1):
    #     col = sheet['A'+str(column)].value
    #     print(col)

    # 获取整个表格的最大行数及列数简单方法
    ws = wb['sheet1']
    # print(ws.dimensions)
    # 按行遍历，但不显示value
    # for row in ws.rows:
    #     print(row)
    # 遍历全部，好些的写法
    for row_obj in sheet[ws.dimensions]:
        for cell_obj in row_obj:
            # cell_obj.coordinate获取当前的单元格属性
            print(cell_obj.coordinate, cell_obj.value)
        print('end of row')

#逻辑有问题，先忽略
def xlrd_read_():
    wb = xlrd.open_workbook(file2)
    #获取所有工作表的名称
    # print(wb.sheet_names())
    # 根据工作表的名称获取工作表的内容，常用在已知需要的sheet名称时
    table = wb.sheet_by_name('Sheet2')
    # 打印工作表的名称、行数和列数
    # print(table.name, table.nrows, table.ncols)
    #另一种方式获取表sheet，常用在不知道需要的sheet名称时，可以取第n个
    # index = wb.sheet_names()[1]
    # print(index)
    #获取索引行的数据
    # print(table.row_values(1))
    #获取索引列的数据
    # print(table.col_values(0))
    # print(table.row(0))

    #遍历所有的数据
    #获取该sheet的行数i及列数l
    nrow = table.nrows
    ncol = table.ncols
    for i in range(nrow):
        # print(table.row_values(i))
        #判断每个单元格的数据，是否包含value
        for l in range(ncol):
            if "value" in table.row_values(i)[l]:
                # print(table.row_values(i))
                #取该列的所有数据
                print(table.col_values(l))
                #判断该列的数据包含_1则取整行
                # print(len(table.col_values(l))-1)
                if "_1" in table.col_values(l)[nrow-1]:
                    #若包含_1获取行数，并取改行的所有数据
                    anadata1 = []

def xlrd_read():
    wb = xlrd.open_workbook(file2)
    table = wb.sheet_by_name('Sheet2')
    nrow = table.nrows
    ncol = table.ncols
    # print(nrow-1)
    # print(ncol-1)
    # print(table.cell(nrow-1,ncol-1))
    need_data =[]
    #获取标题
    for row in range(nrow):
        for col in range(ncol):
            if "title" in table.cell(row, col).value:
                need_data.append(table.row_values(row))
                break
    #获取需要的行
    for row in range(nrow):
        for col in range(ncol):
            #判断是否包含'value'，是的话选取整列
            if "value" in table.cell(row,col).value:
                #判断这列是否包含'_1'，是的话取整行
                # print(table.col_values(col))
                #遍历这一列的所有数据，并标记存在'_1'的行
                for i in range(len(table.col_values(col))):
                    if "_1" in table.cell(i,col).value:
                        # print("存在")
                        need_data.append(table.row_values(i))
    print(need_data)

data = [['title1', 'title2', 'title_value3'], ['数据4', '数据5', '数据6_1'], ['数据10', '数据11', '数据12_1']]

#逻辑有问题忽略
def write_excel_():
    #内存中创建一个空表格
    # wb = openpyxl.Workbook()
    # sheet = wb.active
    # sheet.title = 'ccc_sheet'
    # print(wb.sheetnames)

    #写到原来的表格中，新建一个sheet,只支持xlsx文件
    wb = openpyxl.load_workbook(file3)
    wb.create_sheet(index=0,title='ccc_sheet')
    sheet = wb['ccc_sheet']
    # sheet['A1'] = 'test2'
    # print(sheet['A1'].value)

    #???
    # for i in data:
    #     print(i)
    #取数据
    abc = "0ABCDEFG"
    j = 0
    # for l in range(3):
    #     print(abc[l+1])
    # print(abc[1+1])
    for i in range(len(data)):
        j = j+1
        for data_re in data[i]:
            # j=j+1
            print(f'data_re: {data_re}')

            sheet[abc[i+1]+str(i+1)] = data_re
            print(f'字母是{abc[i+1]} 数字是{str(i+1)} 数据是{data_re}')
            # print(sheet[abc[i+1]+str(i+1)].value)

    # wb.save(file3)

data2 = [['John Brown', 18, 'New York No. 1 Lake Park'],['John Brown2', 11, 'New York No. 1 Lake Park2']]
#另一种写入方式
def write_excel():
    wb = openpyxl.load_workbook(file3)
    wb.create_sheet(index=0,title='ccc_sheet_222')
    sheet = wb['ccc_sheet_222']
    #好的写法写入excel
    for row_index, row_item in enumerate(data):
        for col_index, col_item in enumerate(row_item):
            sheet.cell(row=row_index + 1, column=col_index + 1, value=col_item)
    wb.save(file3)

# xlrd_read()
write_excel()