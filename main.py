import sys

print('================Python import mode==========================')

print('命令行参数为:')
print(sys.argv)
print("数组长度" + len(sys.argv).__str__())
for i in sys.argv:
    print(i)
print('\n python 路径为', sys.path)






