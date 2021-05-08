list = [1, 2, 3, 4]
print(list)

it = iter(list)
print(next(it))

for x in it:
    print(x, end=" ")


