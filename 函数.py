def add(a, b):
    return a + b


# 不定长参数
def max(*args):
    return args


print("-------")
c = add(1, 2)
print(c)

print(max(1, 2, 3, 4))


# 定义函数
def printme(str):
    # 打印任何传入的字符串
    print(str)
    return


# 调用函数
printme("我要调用用户自定义函数!")

printme("再次调用同一函数")


# 查看地址变化
def change(a):
    print(id(a))  # 指向的是同一个对象
    a = 10
    print(id(a))  # 一个新对象


a = 1
print(id(a))
change(a)


# 可写函数说明
def changeme(mylist):
    mylist.append([1, 2, 3, 4])
    print("函数内取值: ", mylist)
    return


# 调用changeme函数
mylist = [10, 20, 30]
changeme(mylist)
print("函数外取值: ", mylist)


# 参数默认值
def printinfo(name, age=35):
    print("名字: ", name)
    print("年龄: ", age)
    return


# 调用printinfo函数
print("------------------------")

printinfo(age=50, name="runoob")
print("------------------------")
printinfo(name="runoob")


# 可写函数说明
def printinfo(arg1, *vartuple):
    print("输出: ", end="")
    print(arg1)
    for var in vartuple:
        print(var)
    return


# 调用printinfo 函数
printinfo(10)
printinfo(70, 60, 50)

# lambda表达式
sum = lambda arg1, arg2: arg1 + arg2

# 调用sum函数
print("相加后的值为 : ", sum(10, 20))
print("相加后的值为 : ", sum(20, 20))


def f(a, b, /, c, d, *, e, f):
    print(a, b, c, d, e, f)


f(10, 20, 30, 40, e=50, f=60)


# def(**kwargs) 把N个关键字参数转化为字典:
def func(country, province, **b):
    print(country, province, b)


func("China", "Sichuan", city="Chengdu", section="JingJiang")


# 参数字典
def addDictionary(**a):
    print(a)


addDictionary(name="houjiguo", sex="male")

# 幂运算
g = lambda x, y: x ** 2 + y ** 2

print(g(1, 2))

# lambda表达式可以进行参数默认值
g = lambda x=0, y=0: x ** 2 + y ** 2

print(g(y=1))
