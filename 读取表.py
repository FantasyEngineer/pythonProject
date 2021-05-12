import xlrd

# 读取整个表格，生  成对象
book = xlrd.open_workbook("/Users/jimmy/text.xlsx")
# 获取表单数据
print(f"包含表单数量{book.nsheets}")
print(f"包含表单名称{book.sheet_names()}")

# 读取对应表单的对象(可以根据名字，或者根据下标)
excelIndex2 = book.sheet_by_index(2)
excelName2018 = book.sheet_by_name("2018")
# 获取所有的表单
sheetsAll = book.sheets()

print(f"所有的表单数据：{sheetsAll}")
print(f"下标为2的表单：{excelIndex2}")
print(f"名称为2018的表单：{excelName2018}")

# 输出所有的行数, 输出所有的列数
print(
    f"2018这个表格里共有多少行？：{excelName2018.nrows}，"
    f"共有多少列？：{excelName2018.ncols},"
    f"表单的索引？：{excelName2018.number},"
    f"表单的名字？：{excelName2018.name}")

# -----------获取单个单元格中的数据-----------
excelName2018 = book.sheet_by_name("2018")
# 获取1行1列的数据，也就是"月份"
print(excelName2018.cell_value(0, 0))
# 获取11行，第2列的数据，为55。0
print(excelName2018.cell_value(rowx=10, colx=1))

# 获取整个一行的单元格中的数据,['月份', '收入'],可以指定开始的位置，结束的位置
print(excelName2018.row_values(0, start_colx=0, end_colx=2))
# 获取一整列的单元格中的内容 ['月份', 1.0, '*2', 3.0, 4.0, 5.0, 6.0, '*7', 8.0, 9.0, 10.0]
print(excelName2018.col_values(0))

# 计算整个2018年的收入
inComeList = excelName2018.col_values(colx=1, start_rowx=1)
# 将收入累加
income = sum(inComeList)
print(f"2018年的总收入是：{income}")

# 将月份上带星号的总收入数据计算出来
starSum = 0
# 获取当前第一行的月份list
monthes = excelName2018.col_values(colx=0)
for row, month in enumerate(monthes):
    # 如果月份里面是字符串，并且是带星号的，这些月份的收入相加一下
    if type(month) is str and month.__contains__("*"):
        print(row, month)
        monthData = excelName2018.cell_value(rowx=row, colx=1)
        starSum = starSum + monthData
print(f"带星号的月份的总收入{starSum}")

# --------------------那么去掉*的总收入是--------------------
remainderInCome = income - starSum
print(f"去掉带星号月份总收入为：{remainderInCome}")
