
linha = [0] * 2
rga = 0
matriz = []
t = 0
while t <=2:
    matriz = matriz + [linha] * 3
    rg = len(matriz)
    for l in range(rg - rga):
        linha = []
        string1 = str("Matriz {},{} :".format(l + rga, 0))
        string2 = str("Matriz {},{} :".format(l + rga, 1))
        linha.append(string1)
        linha.append(string2)
        matriz[l + rga] = linha
    t = t + 1
    rga = rg
print(len(matriz))

logger_file = open('./log/logger.log', "a", encoding='utf-8')
hm = len(matriz)
for l in range(hm):
    stringFile = str(matriz[l])
    logger_file.write(stringFile + '\n')

logger_file.close()     

logger_file = open("./log/logger.log", "r+")

linhas = logger_file.readlines()

matriz = linhas[1].split(",")

teste2 = " " + matriz[4].replace("'","").replace("]\n","").strip()

teste = teste2.strip()[:3]

#logger_file.truncate(0)

logger_file.close()

