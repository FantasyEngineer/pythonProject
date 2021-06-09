from openpyxl import *

# 加载表格
wb = load_workbook('副本部门成员信息new.xlsx')

# 获取指定的sheet
# sheetsName = wb.sheetnames
# print(sheetsName)

sheet = wb["0602"]
# 这是一个二元数组
# cells = sheet0602['a8':'N8']
# for cell in cells[0]:
#     print(cell.value)

# 获取整行数据
data_all = sheet.rows

# 元数组
data_row = list(data_all)
for row in data_row:
    # 获取二级部门，下标是1
    secDepartment: str = row[1].value;
    print(secDepartment)
    # 从列表中删除空，删除人事，业务，二级部门这几行
    if '人事'.__eq__(secDepartment) or '业务'.__eq__(secDepartment) or '二级部门'.__eq__(
            secDepartment) or secDepartment is None:
        data_row.remove(row)

print('-------------------')

for row in data_row:
    # 获取二级部门，下标是1
    secDepartment: str = row[1].value;
    print(secDepartment)
# print("结束", len(data_row))
#
# # 取出最后一个看看数据
# lastone = data_row[145]
# for one in lastone:
#     print(one.value)

# # 将每一行都添加到了一个listrow中
# for row in listRow:
#     print(row.value)

# 先算b列的部门，去除人事以及业务
# dataSecDepartment = sheet['B']
# dataSecDepartment = tuple(dataSecDepartment)
# # 获取的是所有的成员名称
# for i in range(dataSecDepartment.__len__()):
#     print(dataSecDepartment[i].value)
