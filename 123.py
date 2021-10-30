temp_1 = '0'
temp_2 = '0'
max_length = 0
counter = 0
f = open('c:/24-2.txt', 'r')
for symb in f.read():
    temp_1 = temp_2
    temp_2 = symb
    if (temp_1 == 'Y') & (temp_2 == 'Z'):
        if counter > max_length:
            max_length = counter
        counter = 0
    counter += 1
print(max_length)

