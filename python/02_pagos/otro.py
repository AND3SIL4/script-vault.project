lista = [1, 2, 3, 4, 5]

long = len(lista)
numero = 1
new = []
for i in range(long - 1):
    numero += 1
    new.append(numero)

print(new)  # [2, 3, 4, 5, 6]
