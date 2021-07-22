import xlrd

from deal_xlsx import openpy_read_xlsx
from test_excel import write_excel

file_xxx = '/Users/jackrechard/PycharmProjects/testexcel/file/xxx.xls'
file_xxx2 = '/Users/jackrechard/PycharmProjects/testexcel/file/xxx2.xlsx'

"""
筛选数据并写到列表里,三重判断，分别判断列指标1："value"，行指标1："_1"，行指标2："_2"
支持的excel格式为xls
"""
def xlrd_read_xls(file,sheetname = 'Sheet1'):
    wb = xlrd.open_workbook(file)
    table = wb.sheet_by_name(sheetname)
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
合并单元格的处理，传入行列坐标输出值，需要重写
"""
def merged_deal_xls(row,col):
    cell_value = None
    # print(table.merged_cells)
    for (rlow, rhigh, clow, chigh) in table.merged_cells:  # 遍历表格中所有合并单元格位置信息
        # print(rlow,rhigh,clow,chigh)
        if (row >= rlow and row < rhigh):  # 行坐标判断
            if (col >= clow and col < chigh):  # 列坐标判断
                # 如果满足条件，就把合并单元格第一个位置的值赋给其它合并单元格
                cell_value = table.cell_value(rlow, clow)
                # print('合并单元格')
                break  #不符合条件跳出循环，防止覆盖
            else:
                # print('普通单元格')
                cell_value = table.cell_value(row, col)
    return cell_value

"""
从xls文件中读取指定sheet数据，并将合并单元格补全,返回一个list
"""
if __name__ == "__main__":
    newdata = []
    file = file_xxx
    wb = xlrd.open_workbook(file, formatting_info=True)
    table = wb.sheet_by_name('Sheet1')
    # 获取行数列数
    nrow = table.nrows
    ncol = table.ncols
    #(1,3)代表第二行第四列
    # print(merged_deal_xls(3,2))
    for row_index in range(nrow):
        for col_index in range(ncol):
            # print(merged_deal_xls(row_index,col_index))
            newdata.append(merged_deal_xls(row_index,col_index))
    per_list_len = ncol
    list_of_group = zip(*(iter(newdata),) * per_list_len)
    end_list = [list(i) for i in list_of_group]  # i is a tuple
    count = len(newdata) % per_list_len
    end_list.append(newdata[-count:]) if count != 0 else end_list
    end_list2 = []
    for element in end_list:
        if element not in end_list2:
            end_list2.append(element)
    print(end_list2)

    #第一步，读取xls的file文件，并将其合并补全每一个合并单元格的数据，返回list，并写入到xlsx文件；
    #第二步，读取第一步保存的文件并筛选指定的数据返回一个list
    #第三步，将这个list数据写入到文件中，命名一个sheetname
    # write_excel(file=file_xxx2,data=end_list2,sheetname='ccc')
    # data = openpy_read_xlsx(file=file_xxx2,sheetname='ccc_sheet_ccc')
    # write_excel(file=file_xxx2,data=data,sheetname='ccc2')
