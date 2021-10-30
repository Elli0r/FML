dict = {0: 0}
result = 0
for n in range (1, 107):
    if (n % 2 == 0):
        result = 3 * dict.get(n/2) + 2
    else:
        result = 3*n + dict.get(n-1)
    dict[n] = result

print(dict[106])
