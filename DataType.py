# -----------------基本数据类型的定义与赋值
a = 1
b = 0.1
c = 'hello,world'
d = True
# 多个变量赋值
e = f = g = 99
print(e)
print(f)
print(g)
# 也可以分别赋值
h, i, j, k = 1, 2, 3, 4
print(h)
print(i)
print(j)
print(k)

# -----------------输出参数类型
print(type(a))
print(type(b))
print(type(c))
print(type(d))
# <class 'int'>
# <class 'float'>
# <class 'str'>
# <class 'bool'>
# 判定参数类型isinstance
aIsInt = isinstance(a, int)
print("a是int类型?" + str(aIsInt))

# -----------------基本类型之间的转换
print("将int型转换为字符串类型：" + str(a))
print("将float转换为字符串类型：" + str(b))
print("将bool转换为字符串类型：" + str(d))
num = "0.1"
# 将num str类型转换成float
numf = float(num)
print(numf)

# 如果将非字符串转换成float呢，以下会报错
# numstr = 'nihao'
# numstrInt = int(numstr)
# print(numstrInt)

# 这里输出是0，这里不存在四舍五入，这里是只取最前面的int类型
print(int(b))

# ------------------字符串的操作
str = 'Runoob'

print(str)  # 输出字符串
print(str[0:-1])  # 输出第一个到倒数第二个的所有字符
print(str[0])  # 输出字符串第一个字符
print(str[2:5])  # 输出从第三个开始到第五个的字符
print(str[2:])  # 输出从第三个开始的后的所有字符
print(str * 2)  # 输出字符串两次，也可以写成 print (2 * str)
print(str + "TEST")  # 连接字符串

print('Ru\noob')
print(r'Ru\noob')
