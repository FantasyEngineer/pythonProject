# 创建一个文件，并写入内容
f = open("/Users/jimmy/foo.txt", "w+")

f.write("Python 是一个非常好的语言。\n是的，的确非常好!!\n")

# 关闭打开的文件
f.close()

# 打开一个文件，并读取内容
f = open("/Users/jimmy/foo.txt", "r")
s = f.readline()
print(s)

f.close()
