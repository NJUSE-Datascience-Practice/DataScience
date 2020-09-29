import xlrd
filename = "../任务二数据.xlsx"

book = xlrd.open_workbook(filename)

sheet1 = book.sheet_by_name("信息类一")
sheet2 = book.sheet_by_name("信息类二")

nrows1 = sheet1.nrows
nrows2 = sheet2.nrows

row_list1 = []
row_list2 = []

for i in range(2,nrows1):
    row_data = sheet1.row_values(i)
    row_list1.append(row_data)
# for i in row_list1:
#     print(i)

for i in range(2,nrows2):
    row_data = sheet2.row_values(i)
    row_list2.append(row_data)
# for i in row_list2:
#     print(i)