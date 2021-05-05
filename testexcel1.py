import csv
import matplotlib.pyplot as plt
import datetime

filename = 'file/sitka_weather_07-2018_simple.csv'
file2 = 'file/goal.txt'
try:
    with open(filename) as f:
        #reader处理以逗号分隔的第一行数据，并将每项数据存储在列表中
        r = csv.reader(f)
        header_now = next(r)
        #获取最高温
        highs,dates = [],[]
        date2s = []
        for row in r:
            #???
            # date = datetime.strptime(row[2],'%Y-%m-%d')
            date2 = row[2]
            high = int(row[5])
            highs.append(high)
            date2s.append(date2)
            # date.append(date)
except FileNotFoundError:
    print("file not found,please check ur fileload...")
else:
    #可视化
    plt.style.use('classic')
    fig, ax = plt.subplots()
    ax.plot(date2s,highs,c='red')
    ax.set_title("temperature in july")
    ax.set_xlabel('')
    ax.set_ylabel("temp(f)")
    ax.tick_params(axis='both',which='major')
    plt.show()

    print(highs)
    # 将highs写入某个文件
    # with open(file2,'a') as f2:
    #     for l in highs:
    #         f2.write(f"{l}   jjj\n")
    #打印第一行的索引及其值用于检查索引顺序
    # for index,column_header in enumerate(header_now):
    #     print(index,column_header)
    print(header_now)
    # print(r)



