# List（列表） 是 Python 中使用最频繁的数据类型。
#
# 列表可以完成大多数集合类的数据结构实现。列表中元素的类型可以不相同，它支持数字，字符串甚至可以包含列表（所谓嵌套）。
#
# 列表是写在方括号 [] 之间、用逗号分隔开的元素列表。
#
# 和字符串一样，列表同样可以被索引和截取，列表被截取后返回一个包含所需元素的新列表。
from typing import List

t = ['a', 'b', 'c', 'd', 'e']
tinylist = [123, 'runoob']

# 输出整个列表
print(t)

# 输出下标为1的数据
print(t[1])

# 输出下标1到3的数据，不包含下标为3
print(t[1:3])

# 输出下标为2以后的所有数据
print(t[1:])

# 原有数组扩容3倍
conList = tinylist * 3
print(conList)


def reverseWords(input):
    # 通过空格将字符串分隔符，把各个单词分隔为列表
    inputWords = input.split(" ")

    # 翻转字符串
    # 假设列表 list = [1,2,3,4],
    # list[0]=1, list[1]=2 ，而 -1 表示最后一个元素 list[-1]=4 ( 与 list[3]=4 一样)
    # inputWords[-1::-1] 有三个参数
    # 第一个参数 -1 表示最后一个元素
    # 第二个参数为空，表示移动到列表末尾
    # 第三个参数为步长，-1 表示逆向
    inputWords = inputWords[-1::-1]
    print(inputWords)

    # 重新组合字符串
    output = ' '.join(inputWords)

    return output


input = 'I like runoob'
rw = reverseWords(input)
print(rw)

a = [66.25, 333, 333, 1, 1234.5]
print("------------")
# 返回数据的个数
print(a.count(333), a.count(1))
# 按位置添加
a.insert(0, -1)
print("-----insert-------")
print(a)
# 添加
a.append(333)
print(a)

# 翻转
a.reverse()
print(a)

# 返回当前数据在list中的位置
print(a.index(333, 4, len(a)))

# 排序
a.sort()
print(a)

# 删除最前面一个数据
a.remove(333)
a.remove(333)

print(a)

# 将列表当做堆栈使用
stack = [1, 2, 3, 4, 5]
stack.append(6)
stack.append(7)
print(stack)
stack.pop()
print(stack)

# 将列表当作队列使用
from collections import deque

queue = deque(["Eric", "John", "Michael"])
print(queue)

queue.append("terry")
queue.append("graham")
print(queue)
queue.popleft()
print(queue)
queue.pop()
print(queue)



