import xlrd
import jieba

#获取文件路径
book = xlrd.open_workbook("./任务二数据.xlsx")

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

# 计算jaccard系数
def Jaccrad(model, reference):  # terms_reference为源对象，terms_model为候选对象
    terms_reference = jieba.cut(reference)  # 默认精准模式
    terms_model = jieba.cut(model)
    grams_reference = set(terms_reference)  # 去重；如果不需要就改为list
    grams_model = set(terms_model)
    temp = 0
    for i in grams_reference:
        if i in grams_model:
            temp = temp + 1
    fenmu = len(grams_model) + len(grams_reference) - temp  # 并集
    jaccard_coefficient = float(temp / fenmu)  # 交集
    return jaccard_coefficient

result = []

for i in info1:
    for j in info2:
        if Jaccrad(j['企业名称'],i['企业名称'])>0.3:
            if i['法人代表'] == j['企业法人']:
                result.append([i['企业名称'],i['许可名称'],i['法人代表'],i['许可机关'],j['企业名称'],j['企业法人'],j['企业类型'],j['登记机关'],j['企业属地']])

res_excel = open('./对齐结果.xls', 'w', encoding='gbk')
res_excel.write('企业名称（信息类一）\t许可名称（信息类一）\t法人代表（信息类一）\t许可机关（信息类一）\t企业名称（信息类二）\t企业法人（信息类二）\t企业类型（信息类二）\t登记机关（信息类二）\t企业属地（信息类二）\n')
for m in range(len(result)):
    for n in range(len(result[m])):
        res_excel.write(str(result[m][n]))
        res_excel.write('\t')
    res_excel.write('\n')
res_excel.close()