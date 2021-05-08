# 列表推导式


vec = [2, 4, 6]
b = [3 * x for x in vec]
print(b)

c = [[x, x ** 2] for x in vec]
print(c)

freshfruit = ['  banana', '  loganberry ', 'passion fruit  ']

# 删除前后空格
d = [x.strip() for x in freshfruit]
print(d)

print([3 * x for x in vec if x > 3])
print([3 * x for x in vec if x < 2])