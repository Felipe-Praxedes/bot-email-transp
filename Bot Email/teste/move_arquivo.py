from mimetypes import init
import os
import schedule
import pathlib
import shutil
import time
import sys


dirBaixados = os.getcwd() + '\\Baixados'
dirEnviados = os.getcwd() + '\\Enviados'
list_file = []
new_list_file = []
new_file = []
# função que lista todos os arquivos da pasta
def lista_arquivos_inicial(dirBaixados):
    lista = []
    for root, dirs, files in os.walk(dirBaixados):
        for file in files:
            lista.append(os.path.join(root,file))
    return lista
# valida se o dirBaixados do diretorio é valido
def valida_dir(self):
    if os.path.isdir(pathlib.Path(self)):
        return True
    else:
        print('Diretório inserido não é valido!!')
        return False
class TransferenciaArquivos:
    
    def iniciar(self, old_dir,new_dir):
       new_list_file =  self.lista_arquivos_inicial(old_dir)
       #  new_file = self.separa_arquivos_novos(list_file, new_list_file)
       self.copia_arquivos(new_dir,new_list_file)
    # função que lista todos os arquivos dentro da pasta
    def lista_arquivos_inicial(self, dirBaixados):
        lista = []
        for root, dirs, files in os.walk(dirBaixados):
            for file in files:
                lista.append(os.path.join(root,file))
        return lista
    # função que extrai todos os arquivos recentes que não estavam na primeira consulta
    def separa_arquivos_novos(self, old_list, new_list):
        lista =[]
        for file in new_list:
            if file not in old_list:
                lista.append(file)
        return lista
    # função que transfere os arquivos para a pasta de destino
    def copia_arquivos(self, dirBaixados, list):
        if len(list) == 0:
            pass
        else:
            for file in list:
                name_arq = pathlib.Path(file).name
                new_dir = os.path.join(dirBaixados,name_arq)
                list_file.append(file)
                try:
                    shutil.move(file, new_dir)
                    print(f'O arquivo {file} foi transferido com sucesso')
                except Exception as e:
                    print(f'Não foi possivel copiar o arquivo {name_arq}')
                    continue
print('     =================================================================')
print('     Pressione ctrl + C no prompt de comando para cancelar o programa')
print('     =================================================================')
#================= INICIO DO PROGRAMA ==============================='
# insere dirBaixados de arquivos
try:
    origem = dirBaixados # input("Digite o diretório de origem: ").strip()
    var = valida_dir(origem)
    if var == False:
        sys.exit()
    destino = dirEnviados # input("Digite o diretório de Destino: ").strip()
    var = valida_dir(destino)
    if var == False:
        sys.exit()
    
except KeyboardInterrupt:
    print()
    print('O Programa foi finalizado')
    sys.exit()
# lista todos os arquivos dentro da pasta
list_file = lista_arquivos_inicial(dirBaixados)
# instancia classe
program = TransferenciaArquivos()
# definição de rotina (frequencia que sera executado)
program.iniciar(origem,destino)
#schedule.every(20).seconds.do(program.iniciar,origem,destino)
# loop de repetição de rotina
#while 1:
#    try:
#        schedule.run_pending()
#        time.sleep(1)
#    except KeyboardInterrupt:
#        print('Programa finalizado')
#        sys.exit()