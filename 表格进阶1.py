import xlrd

book = xlrd.open_workbook("/Users/jimmy/text.xlsx")
sheets = book.sheets()
s = book.sheet_by_index(0)

for index, sheet in enumerate(sheets):
    print(f"---------------当前执行的表单：{sheet.name}---------------")
    # 查找每个sheet的第一列的标题
    sheetTitle = sheet.row_values(rowx=0)
    print(f"所有的标题{sheetTitle}")
    # 查找收入所在的列
    incomeColum = sheetTitle.index("收入")
    print(f"收入所在的列：{incomeColum}")
    # 计算收入
    incomeYearList = sheet.col_values(colx=incomeColum, start_rowx=1)
    print(f"总的收入list：{incomeYearList}")
    inComeYearData = sum(incomeYearList)
    print(f"{sheet.name}总收入：{inComeYearData}")

    # 以下为找到带*号的月份，然后计算*月份的总收入
    # 然后找出月份所在的那一列
    monthColum = sheetTitle.index("月份")
    print(f"月份所在的列：{monthColum}")
    # 输出月份所在的那一列的所有数据
    months = sheet.col_values(colx=monthColum)
    print(f"月份这一列所有的数据：{months}")

    # 查找月份这一列带星号的月份
    starIncome = 0;
    for row, month in enumerate(months):
        if type(month) is str and month.__contains__("*"):
            当前star月的数据 = sheet.cell_value(rowx=row, colx=incomeColum)
            print(row, month, 当前star月的数据)
            starIncome += 当前star月的数据
    print(f"{sheet.name}你那星号月总收入：{starIncome}")
    #     求出本月总金额-星号月的总金额
    inComeYear = inComeYearData - starIncome
    print(f"{sheet.name}年的总收入{inComeYear}")
