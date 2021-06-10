from openpyxl import *
import time  # 引入time模块
import calendar  # 引入日历
import chinese_calendar
from chinese_calendar import is_workday
from datetime import datetime

# 加载表格
wb = load_workbook('副本部门成员信息new.xlsx')

# ----------------------需要填写的数据下--------------------------
# 需要整理的sheet名称
sheet = wb["0602"]
# 需要整理的是几月的数据
year = 2021
month = 5
print('---------------------时间计算----------------------------')
# 获取到日历
cal = calendar.month(year, month)
print(f"{year}年{month}月日历:")
print(cal)
# 计算当前月有多少天，用以获取最后一天的时间
monthLen = calendar._monthlen(year, month)
# 所以能计算出当前月的第一天，以及最后一天
monthStartData = f"{year}/{month}/1"
monthEndData = f"{year}/{month}/{monthLen}"
# 将时间格式化成date格式
monthStartData = datetime.strptime(monthStartData, '%Y/%m/%d')
monthEndData = datetime.strptime(monthEndData, '%Y/%m/%d')

# 获取所有的上班时间
workDayAll = []
# 获取当月的所有的上班日
for i in range(1, monthLen + 1):
    day = f"{year}/{month}/{i}"
    day = datetime.strptime(day, '%Y/%m/%d')
    isWorkDay = chinese_calendar.is_workday(day)
    if isWorkDay:
        workDayAll.append(day)

# 每周共有多少工作日
workDays = len(workDayAll)
# 每周的每人的工作总小时数，总工作日*8
totalHourOnePersonMonth = workDays * 8

print('当月开始时间：' + monthStartData.__str__(), '当月结束时间：' + monthEndData.__str__())

print('---------------------操作表格---------------------------')

# 获取整行数据
data_all = sheet.rows
# 正编的小时数
zhengbianTotalHour: int = 0
# 外包的小时数
waobaoTotalHour: int = 0
# 符合要求的所有的人数
totalPersonNum: int = 0

# 元数组
data_row = list(data_all)
for row in data_row:
    # 获取二级部门，下标是1，编制在4，入职时间12，离职时间是13
    secDepartment: str = row[1].value;
    # 姓名
    name: str = row[2].value
    # 编制
    bianzhi: str = row[4].value
    # 入职时间
    enterTime: datetime = row[12].value
    # 离职时间
    outTime: datetime = row[13].value

    # print(secDepartment)
    # 当不是这几个部门的时间，再开始计算这个人的当月的小时数
    if not '人事'.__eq__(secDepartment) and not '业务'.__eq__(secDepartment) and not '二级部门'.__eq__(
            secDepartment) and secDepartment is not None:
        # 计算一下人数，也可以看到当前执行的位置
        totalPersonNum += 1
        # 输出当前行的所有数据
        # print(secDepartment, name, bianzhi, enterTime, outTime)
        if bianzhi.__eq__("正编"):  # 正编
            # 当是在职状态下的时候，直接将在职改成下个月当天的时间。便于处理
            if outTime.__str__().strip().__eq__("在职"):
                outTime = datetime.now()

            if outTime >= monthEndData and enterTime < monthStartData:
                # print(name, "离职时间是大于当月的，或者是月末最后一天，但是全勤")
                zhengbianTotalHour += totalHourOnePersonMonth
            else:
                # 当月离职的。
                if enterTime < monthStartData <= outTime < monthEndData:
                    print(name, "当月离职的", outTime)
                    for i in range(len(workDayAll)):
                        if outTime >= workDayAll[i]:
                            print(name, outTime, workDayAll[i])
                            zhengbianTotalHour += 8
                # 说明是当月就职的,并且当月不离职
                elif monthStartData <= enterTime <= monthEndData < outTime:
                    # print(name, f"{month}月入职")
                    for i in range(len(workDayAll) - 1, 0, -1):
                        if enterTime <= workDayAll[i]:
                            # print(name, enterTime, workDayAll[i])
                            zhengbianTotalHour += 8
                # 当月就职，当月离职
                elif monthStartData <= enterTime <= monthEndData and outTime <= monthEndData:
                    # print(name, f"{month}月入职，当月离职了")
                    for i in range(len(workDayAll) - 1, 0, -1):
                        if enterTime <= workDayAll[i] <= outTime:
                            # print(name, enterTime, workDayAll[i])
                            zhengbianTotalHour += 8
        else:  # 外包
            # 当是在职状态下的时候，直接将在职改成下个月当天的时间。便于处理
            if outTime.__str__().strip().__eq__("在职"):
                outTime = datetime.now()

            if outTime >= monthEndData and enterTime < monthStartData:
                # print(name, "离职时间是大于当月的，或者是月末最后一天，但是全勤")
                waobaoTotalHour += totalHourOnePersonMonth
            else:
                # 当月离职的。
                if enterTime < monthStartData <= outTime < monthEndData:
                    # print(name, "当月离职的", outTime)
                    for i in range(len(workDayAll)):
                        if outTime >= workDayAll[i]:
                            # print(name, outTime, workDayAll[i])
                            waobaoTotalHour += 8
                # 说明是当月就职的,并且当月不离职
                elif monthStartData <= enterTime <= monthEndData < outTime:
                    # print(name, f"{month}月入职")
                    for i in range(len(workDayAll) - 1, 0, -1):
                        if enterTime <= workDayAll[i]:
                            # print(name, enterTime, workDayAll[i])
                            waobaoTotalHour += 8
                # 当月就职，当月离职
                elif monthStartData <= enterTime <= monthEndData and outTime <= monthEndData:
                    # print(name, f"{month}月入职，当月离职了")
                    for i in range(len(workDayAll) - 1, 0, -1):
                        if enterTime <= workDayAll[i] <= outTime:
                            # print(name, enterTime, workDayAll[i])
                            waobaoTotalHour += 8

print(f'-------去除人事以及业务总人数是{totalPersonNum}----------')
print(f'-------{month}月满勤共有{workDays}天上班时间，满勤总小时数是：{totalHourOnePersonMonth}------------')
print(f'-------{month}月正编工时数{zhengbianTotalHour}------------')
print(f'-------{month}月外包工时数{waobaoTotalHour}------------')
print(f'-------{month}月总工时数{waobaoTotalHour + zhengbianTotalHour}------------')

# for row in data_row:
#     # 获取二级部门，下标是1
#     secDepartment: str = row[1].value;
#     print(secDepartment)
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
