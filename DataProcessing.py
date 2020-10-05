import xlrd

#获取文件路径
book = xlrd.open_workbook("../任务二数据.xlsx")

#获取表
table1 = book.sheet_by_name("信息类一")
table2 = book.sheet_by_name("信息类二")

#获取总行数总列数
row_Num1 = table1.nrows
col_Num1 = table1.ncols
row_Num2 = table2.nrows
col_Num2 = table2.ncols

info1 = []
info2 = []
key1 = table1.row_values(1)#找出字典的key值
key2 = table2.row_values(1)

j = 1
for i in range(row_Num1-1):
    d ={}
    values = table1.row_values(j)
    for x in range(col_Num1):
        # 把key值对应的value赋值给key，每行循环
        d[key1[x]]=values[x]
    j+=1
    # 把字典加到列表中
    info1.append(d)
info1.pop(0)

k = 1
for i in range(row_Num2-1):
    d ={}
    values = table2.row_values(k)
    for x in range(col_Num2):
        # 把key值对应的value赋值给key，每行循环
        d[key2[x]]=values[x]
    k+=1
    # 把字典加到列表中
    info2.append(d)
info2.pop(0)