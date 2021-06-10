import datetime

# 范围时间
d_time = datetime.datetime.strptime(str(datetime.datetime.now().date()) + '9:00', '%Y-%m-%d%H:%M')
d_time1 = datetime.datetime.strptime(str(datetime.datetime.now().date()) + '18:00', '%Y-%m-%d%H:%M')

# 当前时间
n_time = datetime.datetime.now()

# 判断当前时间是否在范围时间内
if n_time > d_time and n_time < d_time1:
    print(True)
else:
    print(False)